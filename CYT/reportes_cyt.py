# ================================================================
# TUKUYRICUY 2026
# GENERADOR DE REPORTES - CIENCIA Y TECNOLOGÍA
# Archivo: reportes_cyt.py
# ================================================================

import re
import zipfile
import pandas as pd

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

from config_cyt import (
    AREA,
    CARPETA_RESULTADOS,
    CARPETA_REPORTES_IE,
    NOMBRE_ARCHIVO_GENERAL,
    NOMBRE_ARCHIVO_ZIP,
    COMPETENCIAS,
    NIVELES,
    COLORES,
    COLUMNAS_IDENTIFICACION,
    COLUMNAS_DETALLE_INSTITUCION,
)


# ================================================================
# 1. ESTILOS
# ================================================================

def relleno(color):
    return PatternFill(fill_type="solid", fgColor=color)


FILL_CABECERA = relleno(COLORES["CABECERA"])    # #404040
FILL_GRIS = relleno(COLORES["GRIS_CLARO"])
FILLS_NIVELES = {n: relleno(COLORES[n]) for n in NIVELES}

BORDE = Border(bottom=Side(style="thin", color=COLORES["GRIS_MEDIO"]))


# ================================================================
# 2. FUNCIONES GENERALES
# ================================================================

def limpiar_nombre_archivo(nombre):
    nombre = re.sub(r'[\\/*?:"<>|]', "_", str(nombre).strip())
    nombre = re.sub(r"\s+", "_", nombre)

    return nombre[:120] or "SIN_NOMBRE"


def escribir_cabecera(ws, fila, columnas):
    for numero, nombre in enumerate(columnas, start=1):
        celda = ws.cell(row=fila, column=numero, value=nombre)
        celda.fill = FILL_CABECERA
        celda.font = Font(bold=True, color=COLORES["BLANCO"], size=10)
        celda.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )
        celda.border = BORDE

    ws.row_dimensions[fila].height = 35


def escribir_titulo(ws, titulo, subtitulo=None, hasta="H"):
    ws.merge_cells(f"A1:{hasta}1")
    ws["A1"] = titulo
    ws["A1"].fill = FILL_CABECERA
    ws["A1"].font = Font(bold=True, color=COLORES["BLANCO"], size=16)
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    if subtitulo:
        ws.merge_cells(f"A2:{hasta}2")
        ws["A2"] = subtitulo
        ws["A2"].font = Font(italic=True, color=COLORES["TEXTO"], size=10)
        ws["A2"].alignment = Alignment(horizontal="center")


def es_columna_nivel(nombre):
    return (
        nombre == "NIVEL_LOGRO_GLOBAL"
        or (nombre.startswith("COMP_") and nombre.endswith("_NIVEL"))
    )


def es_columna_porcentaje(nombre):
    """Columnas que guardan una fracción (0-1) y se muestran como 43.75 %."""
    return (
        nombre in ("PORCENTAJE_FUENTE_NORMALIZADO", "PORCENTAJE_CALCULADO")
        or (nombre.startswith("COMP_") and nombre.endswith("_PORCENTAJE"))
    )


def escribir_tabla(ws, df):
    """Escribe un DataFrame desde A1 con cabecera #404040."""
    columnas = list(df.columns)
    escribir_cabecera(ws, 1, columnas)

    datos = df.astype(object).where(df.notna(), None)

    for fila in datos.itertuples(index=False, name=None):
        ws.append(list(fila))

    ultima = ws.max_row

    for numero, columna in enumerate(columnas, start=1):
        letra = get_column_letter(numero)

        if columna.startswith("%"):
            formato = "0.0%"
        elif es_columna_porcentaje(columna):
            formato = "0.00%"
        elif columna.startswith("PUNTAJE"):
            formato = "0.0"
        else:
            formato = None

        # Este bloque va FUERA del if/elif/else.
        if formato:
            for celda in ws[letra][1:]:
                celda.number_format = formato

        if es_columna_nivel(columna) and ultima >= 2:
            rango = f"{letra}2:{letra}{ultima}"

            for nivel in NIVELES:
                ws.conditional_formatting.add(
                    rango,
                    CellIsRule(
                        operator="equal",
                        formula=[f'"{nivel}"'],
                        fill=FILLS_NIVELES[nivel],
                        font=Font(bold=True, color=COLORES["TEXTO"]),
                    ),
                )

        muestra = datos[columna].head(200)
        largo = max(
            [len(str(columna)) // 2]
            + [len(str(v)) for v in muestra if v is not None]
        )
        ws.column_dimensions[letra].width = min(max(largo + 3, 12), 45)

    ws.freeze_panes = "A2"
    ws.sheet_view.showGridLines = False

    if ultima >= 2:
        ws.auto_filter.ref = ws.dimensions


# ================================================================
# 3. TABLAS DE RESUMEN
# ================================================================

def resumen_competencias(df):
    """Una fila por competencia más una fila GLOBAL."""
    total = len(df)
    filas = []

    for c in COMPETENCIAS:
        niveles = df[f"{c['codigo']}_NIVEL"]

        fila = {"COMPETENCIA": c["nombre"], "ESTUDIANTES": total}

        for nivel in NIVELES:
            fila[nivel] = int(niveles.eq(nivel).sum())

        for nivel in NIVELES:
            fila[f"% {nivel}"] = fila[nivel] / total if total else 0

        fila["% LOGRO PROMEDIO"] = (
            df[f"{c['codigo']}_PORCENTAJE"].mean() / 100 if total else 0
        )
        filas.append(fila)

    fila = {"COMPETENCIA": "RESULTADO GLOBAL", "ESTUDIANTES": total}

    for nivel in NIVELES:
        fila[nivel] = int(df["NIVEL_LOGRO_GLOBAL"].eq(nivel).sum())

    for nivel in NIVELES:
        fila[f"% {nivel}"] = fila[nivel] / total if total else 0

    fila["% LOGRO PROMEDIO"] = (
        df["PUNTAJE_GLOBAL"].mean() / 100 if total else 0
    )
    filas.append(fila)

    return pd.DataFrame(filas)


def tabla_resumen(df, claves, etiqueta=None):
    """Resultado global por grupo (grado, sección, institución...)."""
    d = df.sort_values(claves).copy()

    for nivel in NIVELES:
        d[f"_{nivel}"] = d["NIVEL_LOGRO_GLOBAL"].eq(nivel)

    g = d.groupby(claves, sort=False)
    tabla = g.size().rename("ESTUDIANTES").to_frame()

    for nivel in NIVELES:
        tabla[nivel] = g[f"_{nivel}"].sum().astype(int)

    for nivel in NIVELES:
        tabla[f"% {nivel}"] = tabla[nivel] / tabla["ESTUDIANTES"]

    tabla["PUNTAJE_PROMEDIO"] = g["PUNTAJE_GLOBAL"].mean()

    for c in COMPETENCIAS:
        tabla[f"% PROM. {c['corto']}"] = (
            g[f"{c['codigo']}_PORCENTAJE"].mean() / 100
        )

    tabla = tabla.reset_index()

    if etiqueta:
        tabla.insert(
            0,
            etiqueta,
            tabla[claves].astype(str).agg(" - ".join, axis=1),
        )

    return tabla


def preparar_detalle(df, completo=False):
    if completo:
        columnas = list(COLUMNAS_DETALLE_INSTITUCION)
    else:
        columnas = [c for c in COLUMNAS_IDENTIFICACION if c in df.columns]
        columnas += ["CORRECTAS", "INCORRECTAS"]

        for c in COMPETENCIAS:
            columnas += [f"{c['codigo']}_ACIERTOS", f"{c['codigo']}_NIVEL"]

        columnas += ["PUNTAJE_GLOBAL", "NIVEL_LOGRO_GLOBAL"]

    detalle = df[[c for c in columnas if c in df.columns]].copy()

    # COMP_x_PORCENTAJE viene de 0 a 100; se pasa a fracción para mostrarlo en %.
    for c in COMPETENCIAS:
        col = f"{c['codigo']}_PORCENTAJE"

        if col in detalle.columns:
            detalle[col] = detalle[col] / 100

    return detalle


# ================================================================
# 4. DASHBOARD
# ================================================================

def agregar_grafico(ws_dash, ws_datos, tabla, col_categorias, titulo, ancla):
    """Barras apiladas al 100 %: rojo, amarillo y verde pastel."""
    columnas = list(tabla.columns)
    col_ini = columnas.index("% INICIO") + 1
    col_fin = columnas.index("% LOGRADO") + 1
    ultima = len(tabla) + 1

    grafico = BarChart()
    grafico.type = "col"
    grafico.grouping = "percentStacked"
    grafico.overlap = 100
    grafico.gapWidth = 60
    grafico.title = titulo
    grafico.width = 26
    grafico.height = 9
    grafico.legend.position = "b"

    grafico.add_data(
        Reference(ws_datos, min_col=col_ini, max_col=col_fin,
                  min_row=1, max_row=ultima),
        titles_from_data=True,
    )
    grafico.set_categories(
        Reference(ws_datos, min_col=col_categorias,
                  min_row=2, max_row=ultima)
    )

    for serie, nivel in zip(grafico.series, NIVELES):
        serie.graphicalProperties.solidFill = COLORES[nivel]
        serie.graphicalProperties.line.solidFill = COLORES["BORDE"]

    grafico.dataLabels = DataLabelList()
    grafico.dataLabels.showVal = True
    grafico.dataLabels.showSerName = False
    grafico.dataLabels.showCatName = False
    grafico.dataLabels.showLegendKey = False

    grafico.x_axis.delete = False
    grafico.y_axis.delete = False
    grafico.y_axis.number_format = "0%"

    ws_dash.add_chart(grafico, ancla)


def construir_dashboard(ws, df, titulo, ws_comp, comp, ws_grupos, grupos):
    escribir_titulo(ws, titulo, f"Resultados de {AREA} - TUKUYRICUY 2026")

    for letra in "ABCDEFGH":
        ws.column_dimensions[letra].width = 18

    indicadores = [
        ("A", "B", "ESTUDIANTES", len(df), "#,##0"),
        ("C", "D", "INSTITUCIONES", df["INSTITUCION"].nunique(), "#,##0"),
        ("E", "F", "PUNTAJE PROMEDIO (de 100)",
         df["PUNTAJE_GLOBAL"].mean(), "0.0"),
        ("G", "H", "% EN LOGRADO",
         df["NIVEL_LOGRO_GLOBAL"].eq("LOGRADO").mean(), "0.0%"),
    ]

    for ini, fin, etiqueta, valor, formato in indicadores:
        ws.merge_cells(f"{ini}4:{fin}4")
        ws.merge_cells(f"{ini}5:{fin}5")

        ws[f"{ini}4"] = etiqueta
        ws[f"{ini}4"].fill = FILL_GRIS
        ws[f"{ini}4"].font = Font(bold=True, color=COLORES["TEXTO"], size=10)
        ws[f"{ini}4"].alignment = Alignment(horizontal="center")

        ws[f"{ini}5"] = valor
        ws[f"{ini}5"].number_format = formato
        ws[f"{ini}5"].fill = FILL_GRIS
        ws[f"{ini}5"].font = Font(bold=True, color=COLORES["TEXTO"], size=18)
        ws[f"{ini}5"].alignment = Alignment(horizontal="center")

    ws.row_dimensions[5].height = 32

    agregar_grafico(
        ws, ws_comp, comp, 1,
        "Niveles de logro por competencia", "A8",
    )
    agregar_grafico(
        ws, ws_grupos, grupos, 1,
        "Nivel de logro global por nivel y grado", "A28",
    )

    ws.sheet_view.showGridLines = False


# ================================================================
# 5. CREAR UN LIBRO COMPLETO
# ================================================================

def crear_libro(df, titulo, ruta, hojas_extra, detalle_completo=False):
    """hojas_extra: lista de (nombre_hoja, columnas_de_agrupacion)."""
    wb = Workbook()

    ws_dash = wb.active
    ws_dash.title = "DASHBOARD"

    comp = resumen_competencias(df)
    ws_comp = wb.create_sheet("RESUMEN_COMPETENCIAS")
    escribir_tabla(ws_comp, comp)

    grupos = tabla_resumen(df, ["NIVEL", "GRADO"], etiqueta="GRUPO")
    ws_grupos = wb.create_sheet("RESUMEN_GRADOS")
    escribir_tabla(ws_grupos, grupos)

    for nombre, claves in hojas_extra:
        ws = wb.create_sheet(nombre)
        escribir_tabla(ws, tabla_resumen(df, claves))

    ws_detalle = wb.create_sheet("DETALLE_ESTUDIANTES")
    escribir_tabla(ws_detalle, preparar_detalle(df, completo=detalle_completo))

    construir_dashboard(ws_dash, df, titulo, ws_comp, comp, ws_grupos, grupos)

    ruta.parent.mkdir(parents=True, exist_ok=True)

    wb.save(ruta)


# ================================================================
# 6. REPORTE GENERAL
# ================================================================

def generar_reporte_general(df):
    ruta = CARPETA_RESULTADOS / NOMBRE_ARCHIVO_GENERAL

    crear_libro(
        df,
        f"REPORTE GENERAL - {AREA}",
        ruta,
        hojas_extra=[("RESUMEN_INSTITUCIONES", ["INSTITUCION", "NIVEL"])],
    )

    print(f"Reporte general: {ruta}")

    return ruta


# ================================================================
# 7. REPORTES POR INSTITUCIÓN
# ================================================================

def generar_reportes_instituciones(df):
    CARPETA_REPORTES_IE.mkdir(parents=True, exist_ok=True)

    instituciones = sorted(df["INSTITUCION"].dropna().unique())

    for numero, institucion in enumerate(instituciones, start=1):
        grupo = df[df["INSTITUCION"] == institucion]

        nombre = f"REPORTE_{limpiar_nombre_archivo(institucion)}_CYT.xlsx"
        ruta = CARPETA_REPORTES_IE / nombre

        crear_libro(
            grupo,
            f"{institucion} - {AREA}",
            ruta,
            hojas_extra=[("RESUMEN_SECCIONES", ["NIVEL", "GRADO", "SECCION"])],
            detalle_completo=True,
        )

        print(f"[{numero}/{len(instituciones)}] {nombre}")

    return len(instituciones)


def comprimir_reportes():
    ruta_zip = CARPETA_RESULTADOS / NOMBRE_ARCHIVO_ZIP

    with zipfile.ZipFile(ruta_zip, "w", zipfile.ZIP_DEFLATED) as zf:
        for archivo in sorted(CARPETA_REPORTES_IE.glob("*.xlsx")):
            zf.write(archivo, arcname=archivo.name)

    print(f"ZIP de instituciones: {ruta_zip}")

    return ruta_zip


# ================================================================
# 8. FUNCIÓN PRINCIPAL
# ================================================================

def generar_todos_los_reportes(df):
    orden = [
        c for c in ["INSTITUCION", "NIVEL", "GRADO", "SECCION", "APELLIDOS"]
        if c in df.columns
    ]
    df = df.sort_values(orden).reset_index(drop=True)

    generar_reporte_general(df)
    total = generar_reportes_instituciones(df)
    comprimir_reportes()

    print(f"\nReportes institucionales generados: {total}")


if __name__ == "__main__":
    from procesamiento_cyt import procesar_datos

    generar_todos_los_reportes(procesar_datos())