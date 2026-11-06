import re
import zipfile
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

from config_ccss import (
    AREA, CARPETA_RESULTADOS, CARPETA_REPORTES_IE,
    NOMBRE_ARCHIVO_GENERAL, NOMBRE_ARCHIVO_ZIP,
    COMPETENCIAS, NIVELES, COLORES, COLUMNAS_IDENTIFICACION,
)

def relleno(color):
    return PatternFill("solid", fgColor=color)

FILL_HEADER = relleno(COLORES["CABECERA"])
FILL_NIVELES = {n: relleno(COLORES[n]) for n in NIVELES}
BORDE = Border(
    left=Side(style="thin", color=COLORES["BORDE"]),
    right=Side(style="thin", color=COLORES["BORDE"]),
    top=Side(style="thin", color=COLORES["BORDE"]),
    bottom=Side(style="thin", color=COLORES["BORDE"]),
)

def nombre_seguro(valor):
    return re.sub(r"\\s+", "_", re.sub(r'[\\\\/*?:"<>|]', "_", str(valor).strip()))[:100] or "SIN_DATO"

def preparar_hoja(ws, titulo, subtitulo=""):
    ws.sheet_view.showGridLines = False
    ws.merge_cells("A1:H1")
    ws["A1"] = titulo
    ws["A1"].font = Font(bold=True, color="FFFFFF", size=15)
    ws["A1"].fill = FILL_HEADER
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.row_dimensions[1].height = 28
    if subtitulo:
        ws.merge_cells("A2:H2")
        ws["A2"] = subtitulo
        ws["A2"].alignment = Alignment(horizontal="center")
        ws["A2"].font = Font(
            name="Arial",
            size=10,
            italic=True,
            color="666666"
        )
def escribir_tabla(ws, df, fila_inicio=4):
    for j, col in enumerate(df.columns, 1):
        celda = ws.cell(fila_inicio, j, str(col))
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = FILL_HEADER
        celda.border = BORDE
        celda.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for i, fila in enumerate(df.itertuples(index=False, name=None), fila_inicio + 1):
        for j, valor in enumerate(fila, 1):
            if pd.isna(valor):
                valor = None
            celda = ws.cell(i, j, valor)
            celda.border = BORDE
            celda.alignment = Alignment(vertical="center", wrap_text=True)
            col = str(df.columns[j-1]).upper()
            if col.startswith("%") or "PORCENTAJE" in col:
                celda.number_format = "0.0%"
            elif "PUNTAJE" in col or "ACIERTOS_PROMEDIO" in col:
                celda.number_format = "0.0"

    ultima = fila_inicio + len(df)
    if len(df.columns):
        ws.auto_filter.ref = f"A{fila_inicio}:{get_column_letter(len(df.columns))}{ultima}"
    ws.freeze_panes = f"A{fila_inicio+1}"

    for j, col in enumerate(df.columns, 1):
        letra = get_column_letter(j)
        ancho = min(max([len(str(col))] + [
            len(str(x)) for x in df[col].dropna().head(100)
        ]) + 2, 42)
        ws.column_dimensions[letra].width = max(12, ancho)
        if str(col).endswith("_NIVEL") or str(col) == "NIVEL_LOGRO_GLOBAL":
            for nivel, color in FILL_NIVELES.items():
                ws.conditional_formatting.add(
                    f"{letra}{fila_inicio+1}:{letra}{ultima}",
                    CellIsRule(operator="equal", formula=[f'"{nivel}"'], fill=color)
                )

def resumen_competencias(df):
    filas = []
    for c in COMPETENCIAS:
        cod = c["codigo"]
        fila = {
            "COMPETENCIA": c["nombre"],
            "PREGUNTAS": ", ".join(f"P{i}" for i in c["preguntas"]),
            "ESTUDIANTES": len(df),
            "ACIERTOS_PROMEDIO": df[f"{cod}_ACIERTOS"].mean(),
            "PORCENTAJE_PROMEDIO": df[f"{cod}_PORCENTAJE"].mean(),
        }
        for nivel in NIVELES:
            n = int((df[f"{cod}_NIVEL"] == nivel).sum())
            fila[nivel] = n
            fila[f"% {nivel}"] = n / len(df) if len(df) else 0
        filas.append(fila)
    return pd.DataFrame(filas)

def tabla_resumen(df, claves):
    filas = []
    agrupador = claves[0] if len(claves) == 1 else claves
    for clave, grupo in df.groupby(agrupador, dropna=False):
        valores = clave if isinstance(clave, tuple) else (clave,)
        fila = dict(zip(claves, valores))
        fila["ESTUDIANTES"] = len(grupo)
        for nivel in NIVELES:
            n = int((grupo["NIVEL_LOGRO_GLOBAL"] == nivel).sum())
            fila[nivel] = n
            fila[f"% {nivel}"] = n / len(grupo)
        fila["PUNTAJE_GLOBAL_PROMEDIO"] = grupo["PUNTAJE_GLOBAL"].mean()
        for c in COMPETENCIAS:
            fila[f"% PROM. {c['corto']}"] = grupo[f"{c['codigo']}_PORCENTAJE"].mean()
        filas.append(fila)
    return pd.DataFrame(filas)

def detalle_estudiantes(df):
    cols = [c for c in COLUMNAS_IDENTIFICACION if c in df.columns]
    cols += ["CORRECTAS", "INCORRECTAS", "PORCENTAJE"]
    for c in COMPETENCIAS:
        cod = c["codigo"]
        cols += [f"{cod}_ACIERTOS", f"{cod}_TOTAL", f"{cod}_PORCENTAJE", f"{cod}_NIVEL"]
    cols += ["PUNTAJE_GLOBAL", "NIVEL_LOGRO_GLOBAL"]
    return df[[c for c in cols if c in df.columns]].copy()

def construir_libro(df, ruta, subtitulo):
    wb = Workbook()
    ws = wb.active
    ws.title = "DASHBOARD"
    preparar_hoja(ws, f"REPORTE DE {AREA} - SECUNDARIA", subtitulo)

    indicadores = [
        ("TOTAL ESTUDIANTES", len(df)),
        ("INSTITUCIONES", df["INSTITUCION"].nunique()),
        ("PUNTAJE GLOBAL PROMEDIO", df["PUNTAJE_GLOBAL"].mean()),
        ("% EN NIVEL LOGRADO", (df["NIVEL_LOGRO_GLOBAL"] == "LOGRADO").mean()),
    ]
    for fila, (etiqueta, valor) in enumerate(indicadores, 4):
        ws.cell(fila, 1, etiqueta).font = Font(bold=True, color="FFFFFF")
        ws.cell(fila, 1).fill = FILL_HEADER
        ws.cell(fila, 2, float(valor))
        if etiqueta.startswith("%"):
            ws.cell(fila, 2).number_format = "0.0%"
        elif "PROMEDIO" in etiqueta:
            ws.cell(fila, 2).number_format = "0.0"

    ws_comp = wb.create_sheet("RESUMEN_COMPETENCIAS")
    preparar_hoja(ws_comp, "RESULTADOS POR COMPETENCIA")
    escribir_tabla(ws_comp, resumen_competencias(df))

    ws_grado = wb.create_sheet("RESUMEN_GRADOS")
    preparar_hoja(ws_grado, "RESULTADOS POR GRADO")
    escribir_tabla(ws_grado, tabla_resumen(df, ["GRADO"]))

    ws_ie = wb.create_sheet("RESUMEN_INSTITUCIONES")
    preparar_hoja(ws_ie, "RESULTADOS POR INSTITUCIÓN")
    escribir_tabla(ws_ie, tabla_resumen(df, ["INSTITUCION"]))

    ws_det = wb.create_sheet("DETALLE_ESTUDIANTES")
    preparar_hoja(ws_det, "DETALLE DE ESTUDIANTES")
    escribir_tabla(ws_det, detalle_estudiantes(df))

    conteos = df["NIVEL_LOGRO_GLOBAL"].value_counts()
    ws["A11"], ws["B11"] = "NIVEL", "ESTUDIANTES"
    for fila, nivel in enumerate(NIVELES, 12):
        ws.cell(fila, 1, nivel).fill = FILL_NIVELES[nivel]
        ws.cell(fila, 2, int(conteos.get(nivel, 0)))

    grafico = BarChart()
    grafico.type = "col"
    grafico.style = 10
    grafico.title = "Distribución global por niveles"
    grafico.y_axis.title = "Estudiantes"
    grafico.x_axis.title = "Nivel"
    grafico.add_data(Reference(ws, min_col=2, min_row=11, max_row=14), titles_from_data=True)
    grafico.set_categories(Reference(ws, min_col=1, min_row=12, max_row=14))
    grafico.height, grafico.width = 7, 13
    ws.add_chart(grafico, "D4")
    wb.save(ruta)

def generar_reporte_general(df):
    ruta = CARPETA_RESULTADOS / NOMBRE_ARCHIVO_GENERAL
    construir_libro(df, ruta, "Consolidado de todas las instituciones")
    print(f"Reporte general generado: {ruta}")

def generar_reportes_instituciones(df):
    CARPETA_REPORTES_IE.mkdir(parents=True, exist_ok=True)
    rutas = []
    for institucion, grupo in df.groupby("INSTITUCION", dropna=False):
        ruta = CARPETA_REPORTES_IE / f"REPORTE_CCSS_{nombre_seguro(institucion)}.xlsx"
        construir_libro(grupo.copy(), ruta, str(institucion))
        rutas.append(ruta)

    zip_path = CARPETA_RESULTADOS / NOMBRE_ARCHIVO_ZIP
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as z:
        for ruta in rutas:
            z.write(ruta, arcname=ruta.name)
    print(f"Reportes individuales generados: {len(rutas)}")
    print(f"ZIP generado: {zip_path}")

def generar_todos_los_reportes(df):
    if df is None or df.empty:
        raise ValueError("No hay datos para generar reportes.")
    df = df.sort_values(
        ["INSTITUCION", "GRADO", "SECCION", "APELLIDOS", "NOMBRES"],
        kind="stable"
    )
    generar_reporte_general(df)
    generar_reportes_instituciones(df)
    print("GENERACIÓN DE REPORTES DE CCSS FINALIZADA.")
