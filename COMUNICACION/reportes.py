import re
import zipfile
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

from config_comunicacion import CARPETA_RESULTADOS, AREA
from procesamiento import resumen_niveles, resumen_por_grado, resumen_competencia


def nombre_archivo_seguro(texto):
    texto = str(texto).strip()
    texto = re.sub(r"[^\w\- ]", "", texto, flags=re.UNICODE)
    texto = re.sub(r"\s+", "_", texto)
    return texto[:90] or "INSTITUCION_SIN_NOMBRE"


def ajustar_hoja(ws):
    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    for columna in ws.columns:
        letra = columna[0].column_letter
        max_len = 0
        for celda in columna[:250]:
            valor = "" if celda.value is None else str(celda.value)
            max_len = max(max_len, len(valor))
        ws.column_dimensions[letra].width = min(max(max_len + 2, 12), 42)


def escribir_dataframe(ws, df, tabla_nombre=None):
    for fila in [df.columns.tolist()] + df.astype(object).where(pd.notna(df), None).values.tolist():
        ws.append(fila)
    if len(df.columns) and len(df) > 0 and tabla_nombre:
        ref = f"A1:{ws.cell(row=ws.max_row, column=ws.max_column).coordinate}"
        tabla = Table(displayName=tabla_nombre, ref=ref)
        tabla.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showFirstColumn=False,
            showLastColumn=False, showRowStripes=True, showColumnStripes=False
        )
        ws.add_table(tabla)
    ajustar_hoja(ws)


def formato_porcentajes(ws, nombres_columnas):
    encabezados = {c.value: c.column for c in ws[1]}
    for nombre in nombres_columnas:
        if nombre in encabezados:
            col = encabezados[nombre]
            for fila in range(2, ws.max_row + 1):
                ws.cell(fila, col).number_format = "0.00%"


def crear_dashboard_general(ruta, df):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "DASHBOARD"
    ws.sheet_view.showGridLines = False

    ws["A1"] = "TUKUYRICUY 2026 - REPORTE DE COMUNICACIÓN"
    ws["A1"].font = Font(size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="17365D")
    ws.merge_cells("A1:F1")
    ws["A3"] = "FILTRO DE NIVEL"
    ws["A3"].font = Font(bold=True)
    ws["B3"] = "TODOS"
    dv = DataValidation(type="list", formula1='"TODOS,PRIMARIA,SECUNDARIA"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(ws["B3"])
    ws["C3"] = "Selecciona TODOS, PRIMARIA o SECUNDARIA"
    ws["C3"].font = Font(italic=True, color="666666")

    # Los cálculos se hacen con fórmulas para que el filtro funcione al abrir Excel.
    ws["A5"], ws["B5"] = "INDICADOR", "RESULTADO"
    indicadores = [
        ("Estudiantes evaluados", '=IF($B$3="TODOS",COUNTA(DETALLE_ESTUDIANTES!$A:$A)-1,COUNTIF(DETALLE_ESTUDIANTES!$C:$C,$B$3))'),
        ("Instituciones educativas", '=IF($B$3="TODOS",COUNTA(RESUMEN_IE!$A:$A)-1,COUNTIFS(DETALLE_ESTUDIANTES!$C:$C,$B$3,DETALLE_ESTUDIANTES!$D:$D,"<>"))'),
        ("Promedio de aciertos", '=IFERROR(IF($B$3="TODOS",AVERAGE(DETALLE_ESTUDIANTES!$N:$N),AVERAGEIF(DETALLE_ESTUDIANTES!$C:$C,$B$3,DETALLE_ESTUDIANTES!$N:$N)),0)'),
        ("Promedio porcentual", '=IFERROR(IF($B$3="TODOS",AVERAGE(DETALLE_ESTUDIANTES!$O:$O),AVERAGEIF(DETALLE_ESTUDIANTES!$C:$C,$B$3,DETALLE_ESTUDIANTES!$O:$O)),0)'),
    ]
    for i, (etiqueta, formula) in enumerate(indicadores, start=6):
        ws.cell(i, 1, etiqueta)
        ws.cell(i, 2, formula)
    ws["B9"].number_format = "0.00%"

    ws["A12"], ws["B12"], ws["C12"] = "NIVEL DE LOGRO", "ESTUDIANTES", "PORCENTAJE"
    niveles = ["INICIO", "PROCESO", "LOGRADO", "REVISAR"]
    for i, nivel in enumerate(niveles, start=13):
        ws.cell(i, 1, nivel)
        ws.cell(i, 2, f'=IF($B$3="TODOS",COUNTIF(DETALLE_ESTUDIANTES!$P:$P,A{i}),COUNTIFS(DETALLE_ESTUDIANTES!$C:$C,$B$3,DETALLE_ESTUDIANTES!$P:$P,A{i}))')
        ws.cell(i, 3, f'=IFERROR(B{i}/$B$6,0)')
        ws.cell(i, 3).number_format = "0.00%"

    ws["E12"], ws["F12"] = "NIVEL", "ESTUDIANTES"
    for i, nivel in enumerate(niveles[:3], start=13):
        ws.cell(i, 5, nivel)
        ws.cell(i, 6, f"=B{i}")
    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Estudiantes por nivel de logro"
    chart.y_axis.title = "Estudiantes"
    chart.x_axis.title = "Nivel de logro"
    chart.height = 8
    chart.width = 15
    chart.add_data(Reference(ws, min_col=6, min_row=12, max_row=15), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=5, min_row=13, max_row=15))
    ws.add_chart(chart, "E5")

    for cell in ws[5] + ws[12]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F4E78")
    for col, width in {"A": 29, "B": 20, "C": 20, "D": 4, "E": 20, "F": 18}.items():
        ws.column_dimensions[col].width = width
    ws.freeze_panes = "A5"

    # Resumen auxiliar por institución.
    ie = wb.create_sheet("RESUMEN_IE")
    resumen_ie = (
        df.groupby(["INSTITUCION", "NIVEL"], dropna=False)
          .size().unstack(fill_value=0).reset_index()
    )
    escribir_dataframe(ie, resumen_ie, "TablaResumenIE")

    grado = wb.create_sheet("POR_GRADO")
    escribir_dataframe(grado, resumen_por_grado(df), "TablaPorGrado")

    comp = wb.create_sheet("COMPETENCIA")
    escribir_dataframe(comp, resumen_competencia(df), "TablaCompetencia")
    formato_porcentajes(comp, ["PORCENTAJE_GLOBAL"])

    detalle = wb.create_sheet("DETALLE_ESTUDIANTES")
    columnas_preferidas = [
        "ID_REGISTRO", "FECHA_HORA", "NIVEL", "INSTITUCION", "GRADO",
        "SECCION", "APELLIDOS", "NOMBRES", "DNI", "SEXO",
        *[f"P{i}" for i in range(1, 17)],
        "CORRECTAS", "INCORRECTAS", "TOTAL_PREGUNTAS",
        "PORCENTAJE_REPORTE", "COMP_1_NOMBRE", "COMP_1_ACIERTOS",
        "COMP_1_TOTAL", "COMP_1_PORCENTAJE", "NIVEL_LOGRO",
        "OBSERVACION_DATOS",
    ]
    columnas = [c for c in columnas_preferidas if c in df.columns]
    detalle_df = df[columnas].copy()
    escribir_dataframe(detalle, detalle_df, "TablaDetalle")
    formato_porcentajes(detalle, ["PORCENTAJE_REPORTE", "COMP_1_PORCENTAJE"])

    # Facilita que Excel recalcule fórmulas al abrir el archivo.
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass

    wb.save(ruta)


def crear_reporte_institucion(df_ie, institucion, ruta):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "RESUMEN"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "TUKUYRICUY 2026 - COMUNICACIÓN"
    ws["A1"].font = Font(size=15, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor="17365D")
    ws.merge_cells("A1:D1")
    ws["A3"] = "INSTITUCIÓN EDUCATIVA"
    ws["B3"] = institucion
    ws["A4"] = "ESTUDIANTES EVALUADOS"
    ws["B4"] = len(df_ie)
    ws["A5"] = "PROMEDIO DE ACIERTOS"
    ws["B5"] = pd.to_numeric(df_ie["CORRECTAS"], errors="coerce").mean()
    ws["A6"] = "PROMEDIO PORCENTUAL"
    ws["B6"] = pd.to_numeric(df_ie["PORCENTAJE_REPORTE"], errors="coerce").mean()
    ws["B6"].number_format = "0.00%"

    ws["A8"], ws["B8"], ws["C8"] = "NIVEL DE LOGRO", "ESTUDIANTES", "PORCENTAJE"
    for i, nivel in enumerate(["INICIO", "PROCESO", "LOGRADO", "REVISAR"], start=9):
        cantidad = int((df_ie["NIVEL_LOGRO"] == nivel).sum())
        ws.cell(i, 1, nivel)
        ws.cell(i, 2, cantidad)
        ws.cell(i, 3, cantidad / len(df_ie) if len(df_ie) else 0)
        ws.cell(i, 3).number_format = "0.00%"

    chart = BarChart()
    chart.type = "col"
    chart.style = 10
    chart.title = "Estudiantes por nivel de logro"
    chart.y_axis.title = "Estudiantes"
    chart.x_axis.title = "Nivel"
    chart.add_data(Reference(ws, min_col=2, min_row=8, max_row=11), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=9, max_row=11))
    chart.height = 7
    chart.width = 13
    ws.add_chart(chart, "E3")
    for row in [8]:
        for cell in ws[row]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
    ws.column_dimensions["A"].width = 29
    ws.column_dimensions["B"].width = 43
    ws.column_dimensions["C"].width = 18

    grado = wb.create_sheet("POR_GRADO")
    escribir_dataframe(grado, resumen_por_grado(df_ie), "TablaGradoIE")

    comp = wb.create_sheet("COMPETENCIA")
    escribir_dataframe(comp, resumen_competencia(df_ie), "TablaCompetenciaIE")
    formato_porcentajes(comp, ["PORCENTAJE_GLOBAL"])

    detalle = wb.create_sheet("DETALLE_ESTUDIANTES")
    columnas = [
        c for c in [
            "ID_REGISTRO", "NIVEL", "GRADO", "SECCION", "APELLIDOS",
            "NOMBRES", "DNI", "SEXO", "CORRECTAS", "INCORRECTAS",
            "TOTAL_PREGUNTAS", "PORCENTAJE_REPORTE", "NIVEL_LOGRO",
            "COMP_1_ACIERTOS", "COMP_1_TOTAL", "COMP_1_PORCENTAJE",
            "OBSERVACION_DATOS",
        ] + [f"P{i}" for i in range(1, 17)] if c in df_ie.columns
    ]
    # Recalcula lista sin depender de orden ni campos opcionales.
    columnas = [c for c in [
        "ID_REGISTRO", "NIVEL", "GRADO", "SECCION", "APELLIDOS",
        "NOMBRES", "DNI", "SEXO", "CORRECTAS", "INCORRECTAS",
        "TOTAL_PREGUNTAS", "PORCENTAJE_REPORTE", "NIVEL_LOGRO",
        "COMP_1_ACIERTOS", "COMP_1_TOTAL", "COMP_1_PORCENTAJE",
        "OBSERVACION_DATOS", *[f"P{i}" for i in range(1, 17)]
    ] if c in df_ie.columns]
    escribir_dataframe(detalle, df_ie[columnas].copy(), "TablaDetalleIE")
    formato_porcentajes(detalle, ["PORCENTAJE_REPORTE", "COMP_1_PORCENTAJE"])

    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"
    except Exception:
        pass
    wb.save(ruta)


def generar_todos_los_reportes(df):
    CARPETA_RESULTADOS.mkdir(parents=True, exist_ok=True)
    carpeta_ie = CARPETA_RESULTADOS / "REPORTES_INSTITUCIONES"
    carpeta_ie.mkdir(parents=True, exist_ok=True)

    # Elimina solo reportes anteriores generados por este programa.
    for anterior in carpeta_ie.glob("*.xlsx"):
        anterior.unlink()

    ruta_general = CARPETA_RESULTADOS / "REPORTE_GENERAL_COMUNICACION.xlsx"
    crear_dashboard_general(ruta_general, df)

    rutas_ie = []
    for institucion, grupo in df.groupby("INSTITUCION", dropna=False):
        nombre = nombre_archivo_seguro(institucion)
        ruta = carpeta_ie / f"{nombre}.xlsx"
        crear_reporte_institucion(grupo.copy(), institucion, ruta)
        rutas_ie.append(ruta)

    ruta_zip = CARPETA_RESULTADOS / "REPORTES_INSTITUCIONES.zip"
    with zipfile.ZipFile(ruta_zip, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for ruta in rutas_ie:
            zf.write(ruta, arcname=ruta.name)

    # Guarda también una base procesada para auditoría y reutilización.
    ruta_procesada = CARPETA_RESULTADOS / "BASE_PROCESADA_COMUNICACION.xlsx"
    with pd.ExcelWriter(ruta_procesada, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="RESULTADOS_PROCESADOS", index=False)

    return ruta_general, carpeta_ie, ruta_zip, ruta_procesada
