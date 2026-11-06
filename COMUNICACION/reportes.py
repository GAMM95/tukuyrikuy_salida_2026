
import re
import zipfile
from pathlib import Path

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

from config_comunicacion import CARPETA_RESULTADOS, AREA
from procesamiento import (
    resumen_niveles,
    resumen_por_grado,
    resumen_competencia
)


# ============================================================
# TUKUYRICUY 2026
# GENERADOR DE REPORTES POR INSTITUCIÓN
# ÁREA: COMUNICACIÓN
#
# GENERA:
# 1. REPORTE GENERAL
# 2. REPORTE INDIVIDUAL POR INSTITUCIÓN
# 3. RESUMEN POR GRADO
# 4. RESUMEN POR COMPETENCIA
# 5. DETALLE DE ESTUDIANTES
# 6. RESPUESTAS P1-P16
# 7. VALORACIONES V1-V16
# 8. COMPETENCIAS EN LA MISMA TABLA
# 9. ARCHIVO ZIP
# ============================================================


# ------------------------------------------------------------
# CONFIGURACIÓN
# ------------------------------------------------------------

MAX_PREGUNTAS = 16

NIVELES_LOGRO = [
    "INICIO",
    "PROCESO",
    "LOGRADO",
    "REVISAR"
]


# ------------------------------------------------------------
# FUNCIONES GENERALES
# ------------------------------------------------------------

def nombre_archivo_seguro(texto):

    texto = str(texto).strip()

    texto = re.sub(
        r"[^\w\- ]",
        "",
        texto,
        flags=re.UNICODE
    )

    texto = re.sub(
        r"\s+",
        "_",
        texto
    )

    return texto[:90] or "INSTITUCION_SIN_NOMBRE"


def ajustar_hoja(ws):

    ws.freeze_panes = "A2"

    ws.sheet_view.showGridLines = False

    for cell in ws[1]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="1F4E78"
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

    ws.row_dimensions[1].height = 35

    for columna in ws.columns:

        letra = columna[0].column_letter

        max_len = 0

        for celda in columna[:250]:

            valor = (
                ""
                if celda.value is None
                else str(celda.value)
            )

            max_len = max(
                max_len,
                len(valor)
            )

        ws.column_dimensions[letra].width = min(
            max(max_len + 2, 12),
            42
        )


def escribir_dataframe(
    ws,
    df,
    tabla_nombre=None
):

    if df is None:
        return

    if len(df.columns) == 0:
        return

    # Encabezados
    ws.append(df.columns.tolist())

    # Registros
    for fila in df.astype(object).where(
        pd.notna(df),
        None
    ).values.tolist():

        ws.append(fila)

    # Crear tabla Excel
    if len(df) > 0 and tabla_nombre:

        ref = (
            f"A1:"
            f"{ws.cell(row=ws.max_row, column=ws.max_column).coordinate}"
        )

        tabla = Table(
            displayName=tabla_nombre,
            ref=ref
        )

        tabla.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )

        ws.add_table(tabla)

    ajustar_hoja(ws)


def formato_porcentajes(
    ws,
    nombres_columnas
):

    encabezados = {
        c.value: c.column
        for c in ws[1]
    }

    for nombre in nombres_columnas:

        if nombre in encabezados:

            col = encabezados[nombre]

            for fila in range(
                2,
                ws.max_row + 1
            ):

                ws.cell(
                    fila,
                    col
                ).number_format = "0.00%"


def preparar_valoracion(valor):

    if pd.isna(valor):
        return None

    try:

        numero = float(
            str(valor).strip().replace(",", ".")
        )

        if numero in (0, 1):
            return int(numero)

    except (ValueError, TypeError):
        pass

    texto = str(valor).strip().upper()

    if texto in ("CORRECTA", "ACIERTO"):
        return 1

    if texto in ("INCORRECTA", "ERROR"):
        return 0

    return None


# ------------------------------------------------------------
# COLUMNAS DEL DETALLE DE ESTUDIANTES
# ------------------------------------------------------------

def obtener_columnas_detalle(df):

    columnas_preferidas = [

        # IDENTIFICACIÓN
        "ID_REGISTRO",
        "FECHA_HORA",
        "NIVEL",
        "INSTITUCION",
        "GRADO",
        "SECCION",
        "APELLIDOS",
        "NOMBRES",
        "DNI",
        "SEXO",

        # RESULTADOS GENERALES
        "CORRECTAS",
        "INCORRECTAS",
        "TOTAL_PREGUNTAS",
        "PORCENTAJE",
        "PORCENTAJE_REPORTE",

        # COMPETENCIA
        "COMP_1_NOMBRE",
        "COMP_1_PREGUNTAS",
        "COMP_1_ACIERTOS",
        "COMP_1_TOTAL",
        "COMP_1_PORCENTAJE",
        "NIVEL_LOGRO",

        # PORCENTAJES DE CONTROL
        "PORCENTAJE_FUENTE_NORMALIZADO",
        "PORCENTAJE_CALCULADO",

        # ESTUDIANTE
        "ESTUDIANTE",

        # RESPUESTAS
        *[
            f"P{i}"
            for i in range(1, MAX_PREGUNTAS + 1)
        ],

        # VALORACIONES
        *[
            f"V{i}"
            for i in range(1, MAX_PREGUNTAS + 1)
        ],

        # OBSERVACIONES
        "OBSERVACION_DATOS"
    ]

    columnas = [
        c
        for c in columnas_preferidas
        if c in df.columns
    ]

    # Conservar otros campos disponibles
    for c in df.columns:

        if c not in columnas:
            columnas.append(c)

    return columnas


# ------------------------------------------------------------
# DETALLE DE COMPETENCIA
# ------------------------------------------------------------

def crear_detalle_competencias(df_ie):

    registros = []

    for _, fila in df_ie.iterrows():

        apellidos = fila.get(
            "APELLIDOS",
            ""
        )

        nombres = fila.get(
            "NOMBRES",
            ""
        )

        estudiante = fila.get(
            "ESTUDIANTE",
            None
        )

        if pd.isna(estudiante) or not estudiante:

            estudiante = (
                f"{apellidos} {nombres}"
            ).strip()

        competencia = fila.get(
            "COMP_1_NOMBRE",
            "Lee diversos tipos de textos escritos en su lengua materna"
        )

        for i in range(
            1,
            MAX_PREGUNTAS + 1
        ):

            pregunta = f"P{i}"

            columna_v = f"V{i}"

            respuesta = fila.get(
                pregunta,
                None
            )

            valoracion = preparar_valoracion(
                fila.get(columna_v, None)
            )

            if valoracion == 1:
                resultado = "CORRECTA"

            elif valoracion == 0:
                resultado = "INCORRECTA"

            else:
                resultado = "SIN VALORACIÓN"

            registros.append({

                "ID_REGISTRO": fila.get(
                    "ID_REGISTRO",
                    ""
                ),

                "NIVEL": fila.get(
                    "NIVEL",
                    ""
                ),

                "GRADO": fila.get(
                    "GRADO",
                    ""
                ),

                "SECCION": fila.get(
                    "SECCION",
                    ""
                ),

                "ESTUDIANTE": estudiante,

                "DNI": fila.get(
                    "DNI",
                    ""
                ),

                "COMPETENCIA": competencia,

                "PREGUNTA": pregunta,

                "RESPUESTA_ALUMNO": respuesta,

                "VALORACION": valoracion,

                "RESULTADO": resultado

            })

    return pd.DataFrame(registros)


# ------------------------------------------------------------
# RESUMEN DE VALORACIÓN POR COMPETENCIA
# ------------------------------------------------------------

def crear_resumen_valoracion_competencia(df_ie):

    registros = []

    for (nivel, grado), grupo in df_ie.groupby(
        ["NIVEL", "GRADO"],
        dropna=False
    ):

        valores = []

        for i in range(
            1,
            MAX_PREGUNTAS + 1
        ):

            columna = f"V{i}"

            if columna in grupo.columns:

                serie = grupo[columna].apply(
                    preparar_valoracion
                )

                valores.extend(
                    serie.dropna().tolist()
                )

        aciertos = sum(valores)

        preguntas_valoradas = len(valores)

        porcentaje = (
            aciertos / preguntas_valoradas
            if preguntas_valoradas
            else 0
        )

        registros.append({

            "NIVEL": nivel,

            "GRADO": grado,

            "COMPETENCIA": (
                "Lee diversos tipos de textos escritos "
                "en su lengua materna"
            ),

            "ESTUDIANTES": len(grupo),

            "ACIERTOS_ACUMULADOS": aciertos,

            "PREGUNTAS_VALORADAS": preguntas_valoradas,

            "PORCENTAJE_GLOBAL": porcentaje

        })

    return pd.DataFrame(registros)


# ------------------------------------------------------------
# DASHBOARD GENERAL
# ------------------------------------------------------------

def crear_dashboard_general(ruta, df):

    wb = Workbook()

    ws = wb.active

    ws.title = "DASHBOARD"

    ws.sheet_view.showGridLines = False

    ws["A1"] = (
        "TUKUYRICUY 2026 - "
        "REPORTE DE COMUNICACIÓN"
    )

    ws["A1"].font = Font(
        size=16,
        bold=True,
        color="FFFFFF"
    )

    ws["A1"].fill = PatternFill(
        "solid",
        fgColor="17365D"
    )

    ws.merge_cells("A1:F1")

    ws["A3"] = "FILTRO DE NIVEL"

    ws["A3"].font = Font(bold=True)

    ws["B3"] = "TODOS"

    dv = DataValidation(
        type="list",
        formula1='"TODOS,PRIMARIA,SECUNDARIA"',
        allow_blank=False
    )

    ws.add_data_validation(dv)

    dv.add(ws["B3"])

    ws["C3"] = (
        "Selecciona TODOS, PRIMARIA o SECUNDARIA"
    )

    ws["C3"].font = Font(
        italic=True,
        color="666666"
    )

    ws["A5"] = "INDICADOR"

    ws["B5"] = "RESULTADO"

    indicadores = [

        (
            "Estudiantes evaluados",
            '=IF($B$3="TODOS",COUNTA(DETALLE_ESTUDIANTES!$A:$A)-1,COUNTIF(DETALLE_ESTUDIANTES!$C:$C,$B$3))'
        ),

        (
            "Instituciones educativas",
            '=IF($B$3="TODOS",COUNTA(RESUMEN_IE!$A:$A)-1,COUNTIFS(DETALLE_ESTUDIANTES!$C:$C,$B$3,DETALLE_ESTUDIANTES!$D:$D,"<>"))'
        ),

        (
            "Promedio de aciertos",
            '=IFERROR(IF($B$3="TODOS",AVERAGE(DETALLE_ESTUDIANTES!$L:$L),AVERAGEIF(DETALLE_ESTUDIANTES!$C:$C,$B$3,DETALLE_ESTUDIANTES!$L:$L)),0)'
        ),

        (
            "Promedio porcentual",
            '=IFERROR(IF($B$3="TODOS",AVERAGE(DETALLE_ESTUDIANTES!$O:$O),AVERAGEIF(DETALLE_ESTUDIANTES!$C:$C,$B$3,DETALLE_ESTUDIANTES!$O:$O)),0)'
        )

    ]

    for i, (etiqueta, formula) in enumerate(
        indicadores,
        start=6
    ):

        ws.cell(i, 1, etiqueta)

        ws.cell(i, 2, formula)

    ws["B9"].number_format = "0.00%"

    # Niveles de logro

    ws["A12"] = "NIVEL DE LOGRO"

    ws["B12"] = "ESTUDIANTES"

    ws["C12"] = "PORCENTAJE"

    for i, nivel in enumerate(
        NIVELES_LOGRO,
        start=13
    ):

        ws.cell(i, 1, nivel)

        ws.cell(
            i,
            2,
            f'=IF($B$3="TODOS",COUNTIF(DETALLE_ESTUDIANTES!$T:$T,A{i}),COUNTIFS(DETALLE_ESTUDIANTES!$C:$C,$B$3,DETALLE_ESTUDIANTES!$T:$T,A{i}))'
        )

        ws.cell(
            i,
            3,
            f'=IFERROR(B{i}/$B$6,0)'
        )

        ws.cell(
            i,
            3
        ).number_format = "0.00%"

    # Gráfico

    ws["E12"] = "NIVEL"

    ws["F12"] = "ESTUDIANTES"

    for i, nivel in enumerate(
        NIVELES_LOGRO[:3],
        start=13
    ):

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

    chart.add_data(
        Reference(
            ws,
            min_col=6,
            min_row=12,
            max_row=15
        ),
        titles_from_data=True
    )

    chart.set_categories(
        Reference(
            ws,
            min_col=5,
            min_row=13,
            max_row=15
        )
    )

    ws.add_chart(
        chart,
        "E5"
    )

    for fila in [5, 12]:

        for cell in ws[fila]:

            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )

            cell.fill = PatternFill(
                "solid",
                fgColor="1F4E78"
            )

    for col, width in {
        "A": 29,
        "B": 20,
        "C": 20,
        "D": 4,
        "E": 20,
        "F": 18
    }.items():

        ws.column_dimensions[col].width = width

    ws.freeze_panes = "A5"

    # --------------------------------------------------------
    # RESUMEN POR INSTITUCIÓN
    # --------------------------------------------------------

    ie = wb.create_sheet("RESUMEN_IE")

    resumen_ie = (
        df.groupby(
            ["INSTITUCION", "NIVEL"],
            dropna=False
        )
        .size()
        .unstack(fill_value=0)
        .reset_index()
    )

    escribir_dataframe(
        ie,
        resumen_ie,
        "TablaResumenIE"
    )

    # --------------------------------------------------------
    # POR GRADO
    # --------------------------------------------------------

    grado = wb.create_sheet("POR_GRADO")

    escribir_dataframe(
        grado,
        resumen_por_grado(df),
        "TablaPorGrado"
    )

    # --------------------------------------------------------
    # COMPETENCIA
    # --------------------------------------------------------

    comp = wb.create_sheet("COMPETENCIA")

    escribir_dataframe(
        comp,
        resumen_competencia(df),
        "TablaCompetencia"
    )

    formato_porcentajes(
        comp,
        ["PORCENTAJE_GLOBAL"]
    )

    # --------------------------------------------------------
    # DETALLE ESTUDIANTES
    # --------------------------------------------------------

    detalle = wb.create_sheet(
        "DETALLE_ESTUDIANTES"
    )

    columnas = obtener_columnas_detalle(df)

    detalle_df = df[columnas].copy()

    escribir_dataframe(
        detalle,
        detalle_df,
        "TablaDetalle"
    )

    formato_porcentajes(
        detalle,
        [
            "PORCENTAJE_REPORTE",
            "PORCENTAJE",
            "COMP_1_PORCENTAJE"
        ]
    )

    try:

        wb.calculation.fullCalcOnLoad = True

        wb.calculation.forceFullCalc = True

        wb.calculation.calcMode = "auto"

    except Exception:
        pass

    wb.save(ruta)


# ------------------------------------------------------------
# REPORTE POR INSTITUCIÓN
# ------------------------------------------------------------

def crear_reporte_institucion(
    df_ie,
    institucion,
    ruta
):

    wb = Workbook()

    # ========================================================
    # HOJA 1: RESUMEN
    # ========================================================

    ws = wb.active

    ws.title = "RESUMEN"

    ws.sheet_view.showGridLines = False

    ws["A1"] = (
        "TUKUYRICUY 2026 - "
        "REPORTE DE COMUNICACIÓN"
    )

    ws["A1"].font = Font(
        size=15,
        bold=True,
        color="FFFFFF"
    )

    ws["A1"].fill = PatternFill(
        "solid",
        fgColor="17365D"
    )

    ws.merge_cells("A1:D1")

    ws["A3"] = "INSTITUCIÓN EDUCATIVA"

    ws["B3"] = institucion

    ws["A4"] = "ESTUDIANTES EVALUADOS"

    ws["B4"] = len(df_ie)

    ws["A5"] = "PROMEDIO DE ACIERTOS"

    ws["B5"] = pd.to_numeric(
        df_ie["CORRECTAS"],
        errors="coerce"
    ).mean()

    ws["A6"] = "PROMEDIO PORCENTUAL"

    ws["B6"] = pd.to_numeric(
        df_ie["PORCENTAJE_REPORTE"],
        errors="coerce"
    ).mean()

    ws["B6"].number_format = "0.00%"

    ws["A8"] = "NIVEL DE LOGRO"

    ws["B8"] = "ESTUDIANTES"

    ws["C8"] = "PORCENTAJE"

    for i, nivel in enumerate(
        NIVELES_LOGRO,
        start=9
    ):

        cantidad = int(
            (df_ie["NIVEL_LOGRO"] == nivel).sum()
        )

        ws.cell(i, 1, nivel)

        ws.cell(i, 2, cantidad)

        ws.cell(
            i,
            3,
            cantidad / len(df_ie) if len(df_ie) else 0
        )

        ws.cell(
            i,
            3
        ).number_format = "0.00%"

    chart = BarChart()

    chart.type = "col"

    chart.style = 10

    chart.title = "Estudiantes por nivel de logro"

    chart.y_axis.title = "Estudiantes"

    chart.x_axis.title = "Nivel"

    chart.add_data(
        Reference(
            ws,
            min_col=2,
            min_row=8,
            max_row=11
        ),
        titles_from_data=True
    )

    chart.set_categories(
        Reference(
            ws,
            min_col=1,
            min_row=9,
            max_row=11
        )
    )

    chart.height = 7

    chart.width = 13

    ws.add_chart(
        chart,
        "E3"
    )

    for cell in ws[8]:

        cell.font = Font(
            bold=True,
            color="FFFFFF"
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="1F4E78"
        )

    ws.column_dimensions["A"].width = 29

    ws.column_dimensions["B"].width = 43

    ws.column_dimensions["C"].width = 18

    # ========================================================
    # HOJA 2: POR GRADO
    # ========================================================

    grado = wb.create_sheet("POR_GRADO")

    escribir_dataframe(
        grado,
        resumen_por_grado(df_ie),
        "TablaGradoIE"
    )

    # ========================================================
    # HOJA 3: COMPETENCIA
    # ========================================================

    comp = wb.create_sheet("COMPETENCIA")

    escribir_dataframe(
        comp,
        resumen_competencia(df_ie),
        "TablaCompetenciaIE"
    )

    formato_porcentajes(
        comp,
        ["PORCENTAJE_GLOBAL"]
    )

    # ========================================================
    # HOJA 4: DETALLE ESTUDIANTES
    #
    # TODO EN UNA SOLA TABLA:
    # IDENTIFICACIÓN
    # RESULTADOS GENERALES
    # COMPETENCIA
    # RESPUESTAS P1-P16
    # VALORACIONES V1-V16
    # OBSERVACIONES
    # ========================================================

    detalle = wb.create_sheet(
        "DETALLE_ESTUDIANTES"
    )

    columnas_preferidas = [

        # IDENTIFICACIÓN
        "ID_REGISTRO",
        "FECHA_HORA",
        "NIVEL",
        "INSTITUCION",
        "GRADO",
        "SECCION",
        "APELLIDOS",
        "NOMBRES",
        "DNI",
        "SEXO",

        # RESULTADOS GENERALES
        "CORRECTAS",
        "INCORRECTAS",
        "TOTAL_PREGUNTAS",
        "PORCENTAJE",
        "PORCENTAJE_REPORTE",
        "NIVEL_LOGRO",

        # COMPETENCIA
        "COMP_1_NOMBRE",
        "COMP_1_PREGUNTAS",
        "COMP_1_ACIERTOS",
        "COMP_1_TOTAL",
        "COMP_1_PORCENTAJE",

        # RESPUESTAS DEL ALUMNO
        *[
            f"P{i}"
            for i in range(1, MAX_PREGUNTAS + 1)
        ],

        # VALORACIONES
        *[
            f"V{i}"
            for i in range(1, MAX_PREGUNTAS + 1)
        ],

        # VALIDACIÓN Y OBSERVACIONES
        "PORCENTAJE_FUENTE_NORMALIZADO",
        "PORCENTAJE_CALCULADO",
        "ESTUDIANTE",
        "OBSERVACION_DATOS"
    ]

    columnas = [
        c
        for c in columnas_preferidas
        if c in df_ie.columns
    ]

    detalle_df = df_ie[columnas].copy()

    escribir_dataframe(
        detalle,
        detalle_df,
        "TablaDetalleIE"
    )

    formato_porcentajes(
        detalle,
        [
            "PORCENTAJE",
            "PORCENTAJE_REPORTE",
            "COMP_1_PORCENTAJE",
            "PORCENTAJE_FUENTE_NORMALIZADO",
            "PORCENTAJE_CALCULADO"
        ]
    )

    # ========================================================
    # RECÁLCULO AUTOMÁTICO
    # ========================================================

    try:

        wb.calculation.fullCalcOnLoad = True

        wb.calculation.forceFullCalc = True

        wb.calculation.calcMode = "auto"

    except Exception:
        pass

    wb.save(ruta)


# ------------------------------------------------------------
# GENERAR TODOS LOS REPORTES
# ------------------------------------------------------------

def generar_todos_los_reportes(df):

    CARPETA_RESULTADOS.mkdir(
        parents=True,
        exist_ok=True
    )

    carpeta_ie = (
        CARPETA_RESULTADOS
        / "REPORTES_INSTITUCIONES"
    )

    carpeta_ie.mkdir(
        parents=True,
        exist_ok=True
    )

    # Eliminar reportes anteriores
    for anterior in carpeta_ie.glob("*.xlsx"):

        anterior.unlink()

    # --------------------------------------------------------
    # REPORTE GENERAL
    # --------------------------------------------------------

    ruta_general = (
        CARPETA_RESULTADOS
        / "REPORTE_GENERAL_COMUNICACION.xlsx"
    )

    crear_dashboard_general(
        ruta_general,
        df
    )

    # --------------------------------------------------------
    # REPORTES INDIVIDUALES
    # --------------------------------------------------------

    rutas_ie = []

    for institucion, grupo in df.groupby(
        "INSTITUCION",
        dropna=False
    ):

        if pd.isna(institucion):
            institucion = "SIN_INSTITUCION"

        nombre = nombre_archivo_seguro(
            institucion
        )

        ruta = (
            carpeta_ie
            / f"{nombre}.xlsx"
        )

        crear_reporte_institucion(
            grupo.copy(),
            institucion,
            ruta
        )

        rutas_ie.append(ruta)

        print(
            f"Reporte generado: {institucion}"
        )

    # --------------------------------------------------------
    # ZIP DE REPORTES
    # --------------------------------------------------------

    ruta_zip = (
        CARPETA_RESULTADOS
        / "REPORTES_INSTITUCIONES.zip"
    )

    with zipfile.ZipFile(
        ruta_zip,
        "w",
        compression=zipfile.ZIP_DEFLATED
    ) as zf:

        for ruta in rutas_ie:

            zf.write(
                ruta,
                arcname=ruta.name
            )

    # --------------------------------------------------------
    # BASE PROCESADA
    # --------------------------------------------------------

    ruta_procesada = (
        CARPETA_RESULTADOS
        / "BASE_PROCESADA_COMUNICACION.xlsx"
    )

    with pd.ExcelWriter(
        ruta_procesada,
        engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="RESULTADOS_PROCESADOS",
            index=False
        )

    print("\n========================================")

    print("TUKUYRICUY 2026")

    print("REPORTES GENERADOS CORRECTAMENTE")

    print("========================================")

    print(
        f"Instituciones: {len(rutas_ie)}"
    )

    print(
        f"Reporte general: {ruta_general}"
    )

    print(
        f"Reportes individuales: {carpeta_ie}"
    )

    print(
        f"Archivo ZIP: {ruta_zip}"
    )

    print(
        f"Base procesada: {ruta_procesada}"
    )

    return (
        ruta_general,
        carpeta_ie,
        ruta_zip,
        ruta_procesada
    )
