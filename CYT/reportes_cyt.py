# -*- coding: utf-8 -*-
"""Generación de reportes Excel de Ciencia y Tecnología."""
import re
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter

from config_cyt import (
    CARPETA_RESULTADOS, AREA, NOMBRE_CARPETA_REPORTES,
    COMPETENCIAS, COLOR_CABECERA, COLOR_INICIO, COLOR_PROCESO,
    COLOR_LOGRADO, COLOR_SECUNDARIO
)

CARPETA_REPORTES = CARPETA_RESULTADOS / NOMBRE_CARPETA_REPORTES
NIVELES_LOGRO = ["INICIO", "PROCESO", "LOGRADO"]

def nombre_seguro(texto):
    texto = re.sub(r'[\\/*?:"<>|]', "_", str(texto))
    texto = re.sub(r"\s+", "_", texto.strip())
    return texto[:90] or "SIN_INSTITUCION"

def _nivel_fill(nivel):
    return {
        "INICIO": COLOR_INICIO,
        "PROCESO": COLOR_PROCESO,
        "LOGRADO": COLOR_LOGRADO
    }.get(str(nivel).upper(), "FFFFFF")

def _estilizar_hoja(ws, titulo=None):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A4"
    if titulo:
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(2, ws.max_column))
        c = ws.cell(1, 1, titulo)
        c.font = Font(name="Arial", size=15, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=COLOR_CABECERA)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 28

def _escribir_dataframe(ws, df, start_row=3):
    for j, col in enumerate(df.columns, 1):
        c = ws.cell(start_row, j, str(col))
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=COLOR_CABECERA)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for i, row in enumerate(df.itertuples(index=False, name=None), start_row + 1):
        for j, value in enumerate(row, 1):
            if pd.isna(value):
                value = None
            elif hasattr(value, "item"):
                value = value.item()
            c = ws.cell(i, j, value)
            c.font = Font(name="Arial", size=9)
            if i % 2 == 0:
                c.fill = PatternFill("solid", fgColor="F2F2F2")
            if isinstance(value, float) and ("PORCENTAJE" in str(df.columns[j-1])):
                c.number_format = "0.0%"
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(df.columns))}{start_row + len(df)}"
    for j, col in enumerate(df.columns, 1):
        sample = [len(str(col))]
        sample.extend(len(str(x)) for x in df[col].head(100).fillna("").tolist())
        ws.column_dimensions[get_column_letter(j)].width = min(max(max(sample) + 2, 12), 34)
    ws.row_dimensions[start_row].height = 32

def _resumen_competencias(df):
    filas = []
    for comp in COMPETENCIAS:
        cid = comp["id"]
        fila = {
            "NIVEL": "TODOS",
            "COMPETENCIA": comp["nombre"],
            "ESTUDIANTES": len(df),
            "ACIERTOS_ACUMULADOS": int(df[f"{cid}_ACIERTOS"].sum()),
            "PREGUNTAS_POSIBLES_ACUMULADAS": int(df[f"{cid}_TOTAL"].sum()),
            "PORCENTAJE_GLOBAL": (
                df[f"{cid}_ACIERTOS"].sum() / df[f"{cid}_TOTAL"].sum()
                if df[f"{cid}_TOTAL"].sum() else 0
            )
        }
        filas.append(fila)
        for nivel in ["PRIMARIA", "SECUNDARIA"]:
            sub = df[df["NIVEL"] == nivel]
            total = sub[f"{cid}_TOTAL"].sum()
            filas.append({
                "NIVEL": nivel,
                "COMPETENCIA": comp["nombre"],
                "ESTUDIANTES": len(sub),
                "ACIERTOS_ACUMULADOS": int(sub[f"{cid}_ACIERTOS"].sum()),
                "PREGUNTAS_POSIBLES_ACUMULADAS": int(total),
                "PORCENTAJE_GLOBAL": sub[f"{cid}_ACIERTOS"].sum() / total if total else 0
            })
    return pd.DataFrame(filas)

def _resumen_grados(df):
    filas = []
    for (nivel, grado), sub in df.groupby(["NIVEL", "GRADO"], dropna=False):
        filas.append({
            "NIVEL": nivel, "GRADO": int(grado), "ESTUDIANTES": len(sub),
            "PROMEDIO_PORCENTAJE": sub["PORCENTAJE_GLOBAL"].mean() / 100,
            "INICIO": int((sub["NIVEL_LOGRO_GLOBAL"] == "INICIO").sum()),
            "PROCESO": int((sub["NIVEL_LOGRO_GLOBAL"] == "PROCESO").sum()),
            "LOGRADO": int((sub["NIVEL_LOGRO_GLOBAL"] == "LOGRADO").sum())
        })
    return pd.DataFrame(filas).sort_values(["NIVEL", "GRADO"])

def _resumen_instituciones(df):
    filas = []
    for institucion, sub in df.groupby("INSTITUCION", dropna=False):
        filas.append({
            "INSTITUCION": institucion or "SIN INSTITUCION",
            "ESTUDIANTES": len(sub),
            "PROMEDIO_PORCENTAJE": sub["PORCENTAJE_GLOBAL"].mean() / 100,
            "INICIO": int((sub["NIVEL_LOGRO_GLOBAL"] == "INICIO").sum()),
            "PROCESO": int((sub["NIVEL_LOGRO_GLOBAL"] == "PROCESO").sum()),
            "LOGRADO": int((sub["NIVEL_LOGRO_GLOBAL"] == "LOGRADO").sum())
        })
    return pd.DataFrame(filas).sort_values("INSTITUCION")

def _crear_dashboard(ws, df):
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:F1")
    ws["A1"] = f"DASHBOARD - {AREA} - TUKUYRICUY 2026"
    ws["A1"].font = Font(name="Arial", size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=COLOR_CABECERA)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 30

    total = len(df)
    datos = [
        ("TOTAL ESTUDIANTES", total),
        ("PRIMARIA", int((df["NIVEL"] == "PRIMARIA").sum())),
        ("SECUNDARIA", int((df["NIVEL"] == "SECUNDARIA").sum())),
        ("PROMEDIO GLOBAL", float(df["PORCENTAJE_GLOBAL"].mean() / 100)),
    ]
    for r, (etiqueta, valor) in enumerate(datos, 3):
        ws.cell(r, 1, etiqueta).font = Font(bold=True, color="FFFFFF")
        ws.cell(r, 1).fill = PatternFill("solid", fgColor=COLOR_CABECERA)
        ws.cell(r, 2, valor)
        ws.cell(r, 2).font = Font(bold=True, size=12)
        if etiqueta == "PROMEDIO GLOBAL":
            ws.cell(r, 2).number_format = "0.0%"

    conteo = df["NIVEL_LOGRO_GLOBAL"].value_counts()
    ws["A9"] = "NIVEL DE LOGRO GLOBAL"
    ws["A9"].font = Font(bold=True, color="FFFFFF")
    ws["A9"].fill = PatternFill("solid", fgColor=COLOR_CABECERA)
    ws["A10"], ws["B10"] = "NIVEL", "ESTUDIANTES"
    for c in ws[10][:2]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=COLOR_CABECERA)
    for i, nivel in enumerate(NIVELES_LOGRO, 11):
        ws.cell(i, 1, nivel)
        ws.cell(i, 1).fill = PatternFill("solid", fgColor=_nivel_fill(nivel))
        ws.cell(i, 2, int(conteo.get(nivel, 0)))

    chart = PieChart()
    chart.title = "Distribución por nivel de logro"
    chart.add_data(Reference(ws, min_col=2, min_row=10, max_row=13), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=11, max_row=13))
    chart.height, chart.width = 8, 13
    ws.add_chart(chart, "D3")

    comp_df = _resumen_competencias(df)
    comp_df = comp_df[comp_df["NIVEL"] == "TODOS"]
    ws["A16"] = "RESULTADOS POR COMPETENCIA"
    ws["A16"].font = Font(bold=True, color="FFFFFF")
    ws["A16"].fill = PatternFill("solid", fgColor=COLOR_CABECERA)
    headers = ["COMPETENCIA", "PORCENTAJE GLOBAL"]
    for j, h in enumerate(headers, 1):
        ws.cell(17, j, h).font = Font(bold=True, color="FFFFFF")
        ws.cell(17, j).fill = PatternFill("solid", fgColor=COLOR_CABECERA)
    for i, row in enumerate(comp_df.itertuples(index=False), 18):
        ws.cell(i, 1, row.COMPETENCIA)
        ws.cell(i, 2, row.PORCENTAJE_GLOBAL).number_format = "0.0%"
    ws.column_dimensions["A"].width = 48
    ws.column_dimensions["B"].width = 22

    bar = BarChart()
    bar.type = "bar"
    bar.style = 10
    bar.title = "Porcentaje de aciertos por competencia"
    bar.y_axis.title = "Competencia"
    bar.x_axis.title = "Porcentaje"
    bar.add_data(Reference(ws, min_col=2, min_row=17, max_row=20), titles_from_data=True)
    bar.set_categories(Reference(ws, min_col=1, min_row=18, max_row=20))
    bar.height, bar.width = 8, 16
    ws.add_chart(bar, "D19")

def construir_libro(df, ruta, subtitulo):
    wb = Workbook()
    ws = wb.active
    ws.title = "DASHBOARD"
    _crear_dashboard(ws, df)

    hojas = [
        ("RESUMEN_COMPETENCIAS", _resumen_competencias(df)),
        ("RESUMEN_GRADOS", _resumen_grados(df)),
        ("RESUMEN_INSTITUCIONES", _resumen_instituciones(df)),
    ]
    for nombre, tabla in hojas:
        sh = wb.create_sheet(nombre)
        _estilizar_hoja(sh, f"{AREA} - {nombre.replace('_', ' ')}")
        _escribir_dataframe(sh, tabla)

    detalle = df.copy()
    sh = wb.create_sheet("DETALLE_ESTUDIANTES")
    _estilizar_hoja(sh, f"{AREA} - DETALLE DE ESTUDIANTES")
    _escribir_dataframe(sh, detalle)

    for sh in wb.worksheets[1:]:
        sh.sheet_view.showGridLines = False
    Path(ruta).parent.mkdir(parents=True, exist_ok=True)
    wb.save(ruta)

def generar_reporte_general(df):
    ruta = CARPETA_REPORTES / "REPORTE_GENERAL_CYT.xlsx"
    construir_libro(df, ruta, "Consolidado de todas las instituciones")
    print(f"Reporte general: {ruta}")

def generar_reportes_por_institucion(df):
    carpeta = CARPETA_REPORTES / "POR_INSTITUCION"
    carpeta.mkdir(parents=True, exist_ok=True)
    archivos = []
    for institucion, sub in df.groupby("INSTITUCION", dropna=False):
        nombre = nombre_seguro(institucion or "SIN_INSTITUCION")
        ruta = carpeta / f"REPORTE_CYT_{nombre}.xlsx"
        construir_libro(sub.copy(), ruta, str(institucion))
        archivos.append(ruta)
    zip_path = CARPETA_REPORTES / "REPORTES_CYT_POR_INSTITUCION.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for archivo in archivos:
            zf.write(archivo, arcname=archivo.name)
    print(f"Reportes por institución: {len(archivos):,}")
    print(f"ZIP por institución: {zip_path}")

def generar_todos_los_reportes(df):
    CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
    generar_reporte_general(df)
    generar_reportes_por_institucion(df)
    print(f"Todos los reportes están en: {CARPETA_REPORTES}")
