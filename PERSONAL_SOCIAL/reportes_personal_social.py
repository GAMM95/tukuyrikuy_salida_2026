import re
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

from config_personal_social import (
    AREA, CARPETA_RESULTADOS, CARPETA_REPORTES_IE,
    NOMBRE_ARCHIVO_GENERAL, NOMBRE_ARCHIVO_ZIP,
    COMPETENCIAS, NIVELES, COLORES, COLUMNAS_IDENTIFICACION,
)


def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


FILL_HEADER = fill(COLORES["CABECERA"])
FILL_LEVEL = {n: fill(COLORES[n]) for n in NIVELES}
THIN_BORDER = Border(
    left=Side(style="thin", color=COLORES["BORDE"]),
    right=Side(style="thin", color=COLORES["BORDE"]),
    top=Side(style="thin", color=COLORES["BORDE"]),
    bottom=Side(style="thin", color=COLORES["BORDE"]),
)


def safe_filename(value):
    value = re.sub(r'[\\/*?:"<>|]', "_", str(value))
    value = re.sub(r"\s+", "_", value.strip())
    return value[:100] or "SIN_NOMBRE"


def preparar_hoja(ws, titulo, subtitulo=None):
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:H1")
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, color=COLORES["BLANCO"], size=15)
    ws["A1"].fill = FILL_HEADER
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28
    if subtitulo:
        ws.merge_cells("A2:H2")
        ws["A2"] = subtitulo
        ws["A2"].font = Font(italic=True, color=COLORES["TEXTO"], size=10)
        ws["A2"].alignment = Alignment(horizontal="center")


def escribir_dataframe(ws, df, start_row=4, start_col=1):
    for j, nombre in enumerate(df.columns, start_col):
        celda = ws.cell(start_row, j, str(nombre))
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = FILL_HEADER
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        celda.border = THIN_BORDER

    for i, fila in enumerate(df.itertuples(index=False, name=None), start_row + 1):
        for j, valor in enumerate(fila, start_col):
            if pd.isna(valor):
                valor = None
            celda = ws.cell(i, j, valor)
            celda.border = THIN_BORDER
            celda.alignment = Alignment(vertical="center", wrap_text=True)
            nombre_col = str(df.columns[j - start_col]).upper()
            if nombre_col.startswith("%") or "PORCENTAJE" in nombre_col:
                celda.number_format = "0.0%"
            elif "PUNTAJE" in nombre_col:
                celda.number_format = "0.0"
            elif nombre_col in ("ESTUDIANTES", "CORRECTAS", "INCORRECTAS"):
                celda.number_format = "#,##0"

    ultima_fila = start_row + len(df)
    ultima_col = start_col + len(df.columns) - 1
    if len(df.columns):
        ws.auto_filter.ref = (
            f"{get_column_letter(start_col)}{start_row}:"
            f"{get_column_letter(ultima_col)}{ultima_fila}"
        )
    ws.freeze_panes = f"A{start_row + 1}"

    for col in range(start_col, ultima_col + 1):
        letra = get_column_letter(col)
        max_len = 12
        for row in range(start_row, min(ultima_fila, start_row + 100) + 1):
            valor = ws.cell(row, col).value
            if valor is not None:
                max_len = max(max_len, min(len(str(valor)) + 2, 42))
        ws.column_dimensions[letra].width = max_len

    # Colorear las columnas que contienen niveles de logro.
    for j, nombre in enumerate(df.columns, start_col):
        if str(nombre).endswith("_NIVEL") or str(nombre) == "NIVEL_LOGRO_GLOBAL":
            letra = get_column_letter(j)
            for nivel, color in FILL_LEVEL.items():
                ws.conditional_formatting.add(
                    f"{letra}{start_row+1}:{letra}{ultima_fila}",
                    CellIsRule(
                        operator="equal",
                        formula=[f'"{nivel}"'],
                        fill=color,
                    ),
                )


def resumen_competencias(df):
    filas = []
    for c in COMPETENCIAS:
        col_nivel = f"{c['codigo']}_NIVEL"
        col_pct = f"{c['codigo']}_PORCENTAJE"
        fila = {
            "COMPETENCIA": c["nombre"],
            "ESTUDIANTES": len(df),
            "ACIERTOS_PROMEDIO": df[f"{c['codigo']}_ACIERTOS"].mean(),
            "PORCENTAJE_PROMEDIO": df[col_pct].mean(),
        }
        for nivel in NIVELES:
            cantidad = int((df[col_nivel] == nivel).sum())
            fila[f"{nivel}"] = cantidad
            fila[f"% {nivel}"] = cantidad / len(df) if len(df) else 0
        filas.append(fila)
    return pd.DataFrame(filas)


def tabla_resumen(df, claves):
    filas = []
    agrupador = claves[0] if len(claves) == 1 else claves
    for clave, grupo in df.groupby(agrupador, dropna=False):
        valores = clave if isinstance(clave, tuple) else (clave,)
        fila = {k: v for k, v in zip(claves, valores)}
        fila["ESTUDIANTES"] = len(grupo)

        for nivel in NIVELES:
            n = int((grupo["NIVEL_LOGRO_GLOBAL"] == nivel).sum())
            fila[nivel] = n
            fila[f"% {nivel}"] = n / len(grupo) if len(grupo) else 0

        fila["PUNTAJE_GLOBAL_PROMEDIO"] = grupo["PUNTAJE_GLOBAL"].mean()
        for c in COMPETENCIAS:
            fila[f"% PROM. {c['corto']}"] = grupo[f"{c['codigo']}_PORCENTAJE"].mean()
        filas.append(fila)

    return pd.DataFrame(filas)


def detalle_reporte(df):
    columnas = [c for c in COLUMNAS_IDENTIFICACION if c in df.columns]
    columnas += ["CORRECTAS", "INCORRECTAS", "PORCENTAJE"]
    for c in COMPETENCIAS:
        columnas += [
            f"{c['codigo']}_ACIERTOS",
            f"{c['codigo']}_TOTAL",
            f"{c['codigo']}_PORCENTAJE",
            f"{c['codigo']}_NIVEL",
        ]
    columnas += ["PUNTAJE_GLOBAL", "NIVEL_LOGRO_GLOBAL"]
    return df[[c for c in columnas if c in df.columns]].copy()


def agregar_grafico_niveles(ws, df_resumen, titulo, fila=4, columna=10):
    if df_resumen.empty or not all(n in df_resumen.columns for n in NIVELES):
        return
    # Tabla auxiliar con conteos por nivel, fuera del área principal.
    ws.cell(fila, columna, "NIVEL")
    ws.cell(fila, columna + 1, "ESTUDIANTES")
    for r, nivel in enumerate(NIVELES, fila + 1):
        ws.cell(r, columna, nivel)
        ws.cell(r, columna + 1, int(df_resumen[nivel].sum()))
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = titulo
    chart.y_axis.title = "Nivel de logro"
    chart.x_axis.title = "Estudiantes"
    data = Reference(ws, min_col=columna + 1, min_row=fila, max_row=fila + len(NIVELES))
    cats = Reference(ws, min_col=columna, min_row=fila + 1, max_row=fila + len(NIVELES))
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.height = 6
    chart.width = 11
    ws.add_chart(chart, "J4")


def construir_libro(df, ruta, nombre):
    wb = Workbook()
    ws = wb.active
    ws.title = "DASHBOARD"
    preparar_hoja(ws, f"REPORTE DE {AREA} - PRIMARIA", nombre)

    total = len(df)
    instituciones = df["INSTITUCION"].nunique()
    promedio = df["PUNTAJE_GLOBAL"].mean()
    logrado = (df["NIVEL_LOGRO_GLOBAL"] == "LOGRADO").mean()

    indicadores = [
        ("TOTAL ESTUDIANTES", total),
        ("INSTITUCIONES", instituciones),
        ("PUNTAJE GLOBAL PROMEDIO", promedio),
        ("% EN NIVEL LOGRADO", logrado),
    ]
    for idx, (etiqueta, valor) in enumerate(indicadores, 4):
        ws.cell(idx, 1, etiqueta).font = Font(bold=True, color="FFFFFF")
        ws.cell(idx, 1).fill = FILL_HEADER
        ws.cell(idx, 2, float(valor) if isinstance(valor, (int, float)) else valor)
        ws.cell(idx, 2).font = Font(bold=True, size=12)
        if etiqueta.startswith("%"):
            ws.cell(idx, 2).number_format = "0.0%"
        elif "PROMEDIO" in etiqueta:
            ws.cell(idx, 2).number_format = "0.0"

    comp = resumen_competencias(df)
    ws_comp = wb.create_sheet("RESUMEN_COMPETENCIAS")
    preparar_hoja(ws_comp, "RESULTADOS POR COMPETENCIA")
    escribir_dataframe(ws_comp, comp)

    grado = tabla_resumen(df, ["GRADO"])
    ws_grado = wb.create_sheet("RESUMEN_GRADOS")
    preparar_hoja(ws_grado, "RESULTADOS POR GRADO")
    escribir_dataframe(ws_grado, grado)

    inst = tabla_resumen(df, ["INSTITUCION"])
    ws_inst = wb.create_sheet("RESUMEN_INSTITUCIONES")
    preparar_hoja(ws_inst, "RESULTADOS POR INSTITUCIÓN")
    escribir_dataframe(ws_inst, inst)

    detalle = detalle_reporte(df)
    ws_det = wb.create_sheet("DETALLE_ESTUDIANTES")
    preparar_hoja(ws_det, "DETALLE DE ESTUDIANTES")
    escribir_dataframe(ws_det, detalle)

    # Gráfico de distribución global por niveles en dashboard.
    conteos = df["NIVEL_LOGRO_GLOBAL"].value_counts()
    ws["A11"] = "NIVEL"
    ws["B11"] = "ESTUDIANTES"
    for r, nivel in enumerate(NIVELES, 12):
        ws.cell(r, 1, nivel)
        ws.cell(r, 2, int(conteos.get(nivel, 0)))
        ws.cell(r, 1).fill = FILL_LEVEL[nivel]
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Distribución global por niveles"
    chart.y_axis.title = "Estudiantes"
    chart.x_axis.title = "Nivel"
    chart.add_data(Reference(ws, min_col=2, min_row=11, max_row=14), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=12, max_row=14))
    chart.height = 7
    chart.width = 13
    ws.add_chart(chart, "D4")

    # Indicadores de niveles por competencia en una hoja independiente.
    ws_comp.freeze_panes = "A5"
    wb.save(ruta)


def generar_reporte_general(df):
    ruta = CARPETA_RESULTADOS / NOMBRE_ARCHIVO_GENERAL
    construir_libro(df, ruta, "Consolidado de todas las instituciones")
    print(f"Reporte general generado: {ruta}")
    return ruta


def generar_reportes_instituciones(df):
    CARPETA_REPORTES_IE.mkdir(parents=True, exist_ok=True)
    rutas = []
    for institucion, grupo in df.groupby("INSTITUCION", dropna=False):
        nombre = safe_filename(institucion)
        ruta = CARPETA_REPORTES_IE / f"REPORTE_PERSONAL_SOCIAL_{nombre}.xlsx"
        construir_libro(grupo.copy(), ruta, str(institucion))
        rutas.append(ruta)

    zip_path = CARPETA_RESULTADOS / NOMBRE_ARCHIVO_ZIP
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as archivo_zip:
        for ruta in rutas:
            archivo_zip.write(ruta, arcname=ruta.name)

    print(f"Reportes individuales generados: {len(rutas)}")
    print(f"ZIP generado: {zip_path}")
    return rutas, zip_path


def generar_todos_los_reportes(df):
    if df is None or df.empty:
        raise ValueError("No hay datos para generar reportes.")
    df = df.sort_values(
        ["INSTITUCION", "GRADO", "SECCION", "APELLIDOS", "NOMBRES"],
        kind="stable"
    )
    generar_reporte_general(df)
    generar_reportes_instituciones(df)
    print("\nGENERACIÓN DE REPORTES FINALIZADA.")
