# import numpy as np
# import pandas as pd
#
# from config_comunicacion import (
#     AREA, TOTAL_PREGUNTAS_ESPERADAS, COLUMNAS_OBLIGATORIAS,
#     clasificar_nivel_logro,
# )
#
#
# def _normalizar_texto(serie):
#     return serie.fillna("").astype(str).str.strip()
#
#
# def _normalizar_porcentaje(valor):
#     """Convierte 0.75, 75 o '75%' a 0.75."""
#     if pd.isna(valor):
#         return np.nan
#     if isinstance(valor, str):
#         valor = valor.strip().replace("%", "").replace(",", ".")
#     try:
#         numero = float(valor)
#     except (TypeError, ValueError):
#         return np.nan
#     return numero / 100 if numero > 1 else numero
#
#
# def cargar_y_procesar(archivo, pestana="RESULTADOS"):
#     """
#     Lee la base exportada desde Google Sheets.
#     Conserva todas las columnas originales y agrega campos calculados.
#     No elimina registros ni modifica el archivo de origen.
#     """
#     try:
#         df = pd.read_excel(archivo, sheet_name=pestana, engine="openpyxl")
#     except ValueError as exc:
#         raise ValueError(
#             f"No se encontró la pestaña '{pestana}' en el archivo. "
#             "Verifica PESTANA_ORIGEN en config_comunicacion.py."
#         ) from exc
#
#     df.columns = [str(c).strip().upper() for c in df.columns]
#
#     faltantes = [c for c in COLUMNAS_OBLIGATORIAS if c not in df.columns]
#     if faltantes:
#         raise ValueError(
#             "Faltan columnas obligatorias en el Excel: " + ", ".join(faltantes)
#         )
#
#     # Normaliza campos de texto sin descartar las demás columnas.
#     for col in [
#         "NIVEL", "INSTITUCION", "GRADO", "SECCION",
#         "APELLIDOS", "NOMBRES", "COMP_1_NOMBRE"
#     ]:
#         if col in df.columns:
#             df[col] = _normalizar_texto(df[col])
#
#     df["NIVEL"] = df["NIVEL"].str.upper()
#     df["INSTITUCION"] = df["INSTITUCION"].str.replace(r"\s+", " ", regex=True)
#     df["GRADO"] = df["GRADO"].str.replace(r"\.0$", "", regex=True)
#
#     for col in [
#         "CORRECTAS", "INCORRECTAS", "TOTAL_PREGUNTAS",
#         "COMP_1_ACIERTOS", "COMP_1_TOTAL"
#     ]:
#         if col in df.columns:
#             df[col] = pd.to_numeric(df[col], errors="coerce")
#
#     # Puntaje de logro específico de Comunicación.
#     df["NIVEL_LOGRO"] = df["CORRECTAS"].apply(clasificar_nivel_logro)
#
#     # Conserva el porcentaje recibido y genera uno consistente con los aciertos.
#     if "PORCENTAJE" in df.columns:
#         df["PORCENTAJE_FUENTE_NORMALIZADO"] = df["PORCENTAJE"].apply(
#             _normalizar_porcentaje
#         )
#     else:
#         df["PORCENTAJE_FUENTE_NORMALIZADO"] = np.nan
#
#     df["PORCENTAJE_CALCULADO"] = (
#         df["CORRECTAS"] / df["TOTAL_PREGUNTAS"].replace(0, np.nan)
#     )
#     df["PORCENTAJE_REPORTE"] = df["PORCENTAJE_CALCULADO"]
#
#     df["ESTUDIANTE"] = (
#         df.get("APELLIDOS", "").fillna("").astype(str).str.strip()
#         + ", "
#         + df.get("NOMBRES", "").fillna("").astype(str).str.strip()
#     ).str.strip(", ")
#
#     # Observaciones de calidad: se informa, pero no se borra ni altera la fila.
#     df["OBSERVACION_DATOS"] = ""
#     df.loc[df["NIVEL_LOGRO"].eq("REVISAR"), "OBSERVACION_DATOS"] = (
#         "Puntaje ausente o fuera del rango 0-16"
#     )
#     df.loc[df["INSTITUCION"].eq(""), "OBSERVACION_DATOS"] += " | Institución vacía"
#     df.loc[
#         df["TOTAL_PREGUNTAS"].ne(TOTAL_PREGUNTAS_ESPERADAS),
#         "OBSERVACION_DATOS"
#     ] += " | Total de preguntas distinto de 16"
#
#     if "ID_REGISTRO" in df.columns:
#         duplicados = df["ID_REGISTRO"].duplicated(keep=False) & df["ID_REGISTRO"].notna()
#         df.loc[duplicados, "OBSERVACION_DATOS"] += " | ID_REGISTRO duplicado"
#
#     return df
#
#
# def resumen_niveles(df, filtro_nivel="TODOS"):
#     datos = df if filtro_nivel == "TODOS" else df[df["NIVEL"] == filtro_nivel]
#     orden = ["INICIO", "PROCESO", "LOGRADO", "REVISAR"]
#     conteo = datos["NIVEL_LOGRO"].value_counts()
#     total = len(datos)
#     return pd.DataFrame([
#         {
#             "NIVEL_LOGRO": nivel,
#             "ESTUDIANTES": int(conteo.get(nivel, 0)),
#             "PORCENTAJE": int(conteo.get(nivel, 0)) / total if total else 0,
#         }
#         for nivel in orden
#     ])
#
#
# def resumen_por_grado(df, filtro_nivel="TODOS"):
#     datos = df if filtro_nivel == "TODOS" else df[df["NIVEL"] == filtro_nivel]
#     columnas = [
#         "NIVEL", "GRADO", "ESTUDIANTES",
#         "INICIO", "PROCESO", "LOGRADO", "REVISAR",
#         "PROMEDIO_ACIERTOS", "PROMEDIO_PORCENTAJE"
#     ]
#     if datos.empty:
#         return pd.DataFrame(columns=columnas)
#
#     tabla = pd.crosstab(
#         [datos["NIVEL"], datos["GRADO"]], datos["NIVEL_LOGRO"]
#     ).reset_index()
#
#     for nivel in ["INICIO", "PROCESO", "LOGRADO", "REVISAR"]:
#         if nivel not in tabla.columns:
#             tabla[nivel] = 0
#
#     promedios = (
#         datos.groupby(["NIVEL", "GRADO"], dropna=False)
#         .agg(
#             ESTUDIANTES=("ID_REGISTRO", "size"),
#             PROMEDIO_ACIERTOS=("CORRECTAS", "mean"),
#             PROMEDIO_PORCENTAJE=("PORCENTAJE_REPORTE", "mean"),
#         )
#         .reset_index()
#     )
#     tabla = tabla.merge(promedios, on=["NIVEL", "GRADO"], how="left")
#     return tabla[columnas].sort_values(["NIVEL", "GRADO"]).reset_index(drop=True)
#
#
# def resumen_competencia(df):
#     """Comunicación se resume usando los campos COMP_1_* existentes."""
#     if df.empty:
#         return pd.DataFrame(columns=[
#             "NIVEL", "COMPETENCIA", "ESTUDIANTES",
#             "ACIERTOS_ACUMULADOS", "PREGUNTAS_POSIBLES_ACUMULADAS",
#             "PORCENTAJE_GLOBAL"
#         ])
#
#     datos = df.copy()
#     if "COMP_1_NOMBRE" not in datos.columns:
#         datos["COMP_1_NOMBRE"] = AREA
#
#     datos["COMP_1_NOMBRE"] = datos["COMP_1_NOMBRE"].fillna("").astype(str).str.strip()
#     datos.loc[datos["COMP_1_NOMBRE"].eq(""), "COMP_1_NOMBRE"] = AREA
#
#     return (
#         datos.groupby(["NIVEL", "COMP_1_NOMBRE"], dropna=False)
#         .agg(
#             ESTUDIANTES=("ID_REGISTRO", "size"),
#             ACIERTOS_ACUMULADOS=("COMP_1_ACIERTOS", "sum"),
#             PREGUNTAS_POSIBLES_ACUMULADAS=("COMP_1_TOTAL", "sum"),
#         )
#         .reset_index()
#         .rename(columns={"COMP_1_NOMBRE": "COMPETENCIA"})
#         .assign(
#             PORCENTAJE_GLOBAL=lambda x:
#                 x["ACIERTOS_ACUMULADOS"] /
#                 x["PREGUNTAS_POSIBLES_ACUMULADAS"].replace(0, np.nan)
#         )
#     )


# ================================================================
# TUKUYRICUY 2026
# PROCESAMIENTO DE RESULTADOS - COMUNICACIÓN
# Archivo: procesamiento_comunicacion.py
# ================================================================

import numpy as np
import pandas as pd
from pathlib import Path

from config_comunicacion import (
    ARCHIVO_ORIGEN,
    HOJAS_ORIGEN,
    CARPETA_RESULTADOS,
    NOMBRE_ARCHIVO_PROCESADO,
    COMPETENCIAS,
    TOTAL_PREGUNTAS,
    COLUMNAS_OBLIGATORIAS,
    COLUMNAS_IDENTIFICACION,
    COLUMNAS_PREGUNTAS,
    COLUMNAS_VALORACION,
    COLUMNAS_RESULTADOS,
    COLUMNAS_COMPETENCIAS,
    COLUMNAS_RESULTADO_GLOBAL,
    COLUMNAS_PORCENTAJES,
    clasificar_nivel_competencia,
    calcular_puntaje_ponderado,
    clasificar_nivel_global,
    crear_carpetas,
    validar_configuracion,
)


# 1. NORMALIZAR NOMBRES DE COLUMNAS
def normalizar_columnas(df):
    df = df.copy()

    df.columns = (
        df.columns.astype(str)
        .str.strip()
        .str.upper()
        .str.replace(" ", "_", regex=False)
        .str.replace("-", "_", regex=False)
    )

    return df


# 2. IDENTIFICAR NIVEL EDUCATIVO
def agregar_nivel(df, nombre_hoja):
    if "NIVEL" in df.columns:
        df["NIVEL"] = (
            df["NIVEL"]
            .fillna("SIN DATO")
            .astype(str)
            .str.strip()
            .str.upper()
            .replace("", "SIN DATO")
        )

    elif nombre_hoja.upper() == "PRIMARIA":
        df["NIVEL"] = "PRIMARIA"

    elif nombre_hoja.upper() == "SECUNDARIA":
        df["NIVEL"] = "SECUNDARIA"

    else:
        raise ValueError(
            "No se encontró la columna NIVEL en la hoja RESULTADOS."
        )

    return df


# 3. VALIDAR COLUMNAS OBLIGATORIAS
def validar_columnas_obligatorias(df, nombre_hoja):
    faltantes = [c for c in COLUMNAS_OBLIGATORIAS if c not in df.columns]

    if faltantes:
        raise ValueError(
            f"La hoja '{nombre_hoja}' no contiene las columnas "
            f"necesarias: {', '.join(faltantes)}."
        )


# 4. LIMPIAR IDENTIFICACIÓN
def limpiar_identificacion(df):
    df = df.copy()

    for columna in ["INSTITUCION", "SECCION"]:
        if columna in df.columns:
            df[columna] = (
                df[columna]
                .fillna("SIN DATO")
                .astype(str)
                .str.strip()
                .str.upper()
                .str.replace(r"\s+", " ", regex=True)
                .replace("", "SIN DATO")
            )

    if "GRADO" in df.columns:
        df["GRADO"] = (
            pd.to_numeric(df["GRADO"], errors="coerce")
            .fillna(0)
            .astype(int)
        )

    return df


# 5. QUITAR COLUMNAS CALCULADAS DEL ORIGEN
# Se conserva CORRECTAS (es la base de la clasificación);
# todo lo demás se recalcula con la misma regla.
def eliminar_columnas_calculadas(df):
    a_eliminar = [
        c for c in df.columns
        if c.startswith("COMP_")
           or c.startswith("PORCENTAJE_")
           or (c in COLUMNAS_RESULTADOS and c != "CORRECTAS")
           or c in COLUMNAS_RESULTADO_GLOBAL
           or c in ("NIVEL_LOGRO", "OBSERVACION_DATOS")
    ]

    return df.drop(columns=a_eliminar)


# 6. CALCULAR RESULTADOS
def calcular_resultados(df):
    df = df.copy()

    # --- CORRECTAS: entero entre 0 y 16 ---
    bruto = pd.to_numeric(df["CORRECTAS"], errors="coerce")
    valido = bruto.between(0, TOTAL_PREGUNTAS) & (bruto % 1 == 0)
    correctas = bruto.where(valido)

    invalidos = int((~valido).sum())

    if invalidos:
        print(
            f"AVISO: {invalidos} registros con CORRECTAS vacío o fuera de "
            f"0-{TOTAL_PREGUNTAS}; quedan como REVISAR."
        )

    # --- Resultado general ---
    df["CORRECTAS"] = correctas
    df["TOTAL_PREGUNTAS"] = TOTAL_PREGUNTAS
    df["INCORRECTAS"] = TOTAL_PREGUNTAS - correctas
    df["PORCENTAJE"] = (correctas / TOTAL_PREGUNTAS * 100).round(2)

    # --- Por competencia ---
    columnas_ponderadas = []

    for c in COMPETENCIAS:
        pref = c["codigo"]
        total = len(c["preguntas"])
        peso = c["peso"]

        # Comunicación tiene una sola competencia con las 16 preguntas,
        # por eso sus aciertos son las CORRECTAS.
        aciertos = correctas

        df[f"{pref}_NOMBRE"] = c["nombre"]
        df[f"{pref}_PREGUNTAS"] = ", ".join(f"P{p}" for p in c["preguntas"])
        df[f"{pref}_ACIERTOS"] = aciertos
        df[f"{pref}_TOTAL"] = total
        df[f"{pref}_PORCENTAJE"] = (aciertos / total * 100).round(2)
        df[f"{pref}_NIVEL"] = aciertos.map(clasificar_nivel_competencia)
        df[f"{pref}_PESO"] = peso
        df[f"{pref}_PUNTAJE_PONDERADO"] = aciertos.map(
            lambda a, peso=peso, total=total:
            calcular_puntaje_ponderado(a, peso, total)
        )

        columnas_ponderadas.append(f"{pref}_PUNTAJE_PONDERADO")

    # --- Global sobre 100 ---
    df["PUNTAJE_GLOBAL"] = (
        df[columnas_ponderadas].sum(axis=1, min_count=1).round(2)
    )
    df["NIVEL_LOGRO_GLOBAL"] = correctas.map(clasificar_nivel_global)

    return df

# PORCENTAJES DE CONTROL
def normalizar_porcentaje(valor):
    """Convierte 0.75, 75 o '75%' a 0.75."""
    if pd.isna(valor):
        return np.nan

    if isinstance(valor, str):
        valor = valor.strip().replace("%", "").replace(",", ".")

    try:
        numero = float(valor)
    except (TypeError, ValueError):
        return np.nan

    return numero / 100 if numero > 1 else numero


def agregar_porcentajes(df, porcentaje_origen=None):
    df = df.copy()

    if porcentaje_origen is not None:
        df["PORCENTAJE_FUENTE_NORMALIZADO"] = porcentaje_origen.map(
            normalizar_porcentaje
        )
    else:
        df["PORCENTAJE_FUENTE_NORMALIZADO"] = np.nan

    df["PORCENTAJE_CALCULADO"] = df["CORRECTAS"] / TOTAL_PREGUNTAS

    return df

# 7. OBSERVACIONES DE CALIDAD (no se elimina ningún registro)
def agregar_observaciones(df, total_origen=None):
    partes = pd.DataFrame(index=df.index)

    partes["a"] = np.where(
        df["NIVEL_LOGRO_GLOBAL"].eq("REVISAR"),
        "Puntaje ausente o fuera del rango 0-16", "",
    )
    partes["b"] = np.where(
        df["INSTITUCION"].eq("SIN DATO"), "Institución vacía", ""
    )

    if total_origen is not None:
        distinto = total_origen.notna() & total_origen.ne(TOTAL_PREGUNTAS)
        partes["c"] = np.where(
            distinto, "Total de preguntas distinto de 16", ""
        )

    if "ID_REGISTRO" in df.columns:
        duplicado = (
                df["ID_REGISTRO"].duplicated(keep=False) & df["ID_REGISTRO"].notna()
        )
        partes["d"] = np.where(duplicado, "ID_REGISTRO duplicado", "")

    df["OBSERVACION_DATOS"] = [
        " | ".join(x for x in fila if x)
        for fila in partes.itertuples(index=False, name=None)
    ]

    return df


# 8. ORDENAR COLUMNAS
def ordenar_columnas(df):
    orden = (
        COLUMNAS_IDENTIFICACION
        + COLUMNAS_PREGUNTAS
        + COLUMNAS_VALORACION
        + COLUMNAS_RESULTADOS
        + COLUMNAS_COMPETENCIAS
        + COLUMNAS_RESULTADO_GLOBAL
        + COLUMNAS_PORCENTAJES
    )

    principales = [c for c in orden if c in df.columns]
    otras = [c for c in df.columns if c not in orden]

    return df[principales + otras]


# 9. PROCESAR UNA HOJA
def procesar_hoja(nombre_hoja):
    print(f"Procesando hoja: {nombre_hoja}")

    try:
        df = pd.read_excel(ARCHIVO_ORIGEN, sheet_name=nombre_hoja, dtype=object)
    except ValueError as exc:
        raise ValueError(
            f"No se encontró la pestaña '{nombre_hoja}' en el archivo. "
            "Verifica HOJAS_ORIGEN en config_comunicacion.py."
        ) from exc

    df = normalizar_columnas(df)
    df = df.dropna(how="all")
    df = agregar_nivel(df, nombre_hoja)

    validar_columnas_obligatorias(df, nombre_hoja)

    # Datos del origen que se necesitan antes de recalcular
    total_origen = (
        pd.to_numeric(df["TOTAL_PREGUNTAS"], errors="coerce")
        if "TOTAL_PREGUNTAS" in df.columns else None
    )
    porcentaje_origen = (
        df["PORCENTAJE"].copy() if "PORCENTAJE" in df.columns else None
    )

    df = limpiar_identificacion(df)
    df = eliminar_columnas_calculadas(df)
    df = calcular_resultados(df)
    df = agregar_porcentajes(df, porcentaje_origen)
    df = agregar_observaciones(df, total_origen)

    return ordenar_columnas(df)

# 10. PROCESAR TODAS LAS HOJAS
def procesar_datos():
    crear_carpetas()
    validar_configuracion()

    if not Path(ARCHIVO_ORIGEN).exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de origen:\n{ARCHIVO_ORIGEN}\n"
            "Copia el Excel descargado de Google Sheets a la carpeta "
            "'resources' y nómbralo RESULTADOS_COMUNICACION.xlsx"
        )

    resultados = [procesar_hoja(hoja) for hoja in HOJAS_ORIGEN]

    df_final = pd.concat(resultados, ignore_index=True, sort=False)

    if "ID_REGISTRO" in df_final.columns:
        repetidos = int(df_final["ID_REGISTRO"].duplicated().sum())

        if repetidos:
            print(f"AVISO: {repetidos} registros con ID_REGISTRO repetido.")

    archivo_salida = CARPETA_RESULTADOS / NOMBRE_ARCHIVO_PROCESADO

    with pd.ExcelWriter(archivo_salida, engine="openpyxl") as writer:
        df_final.to_excel(writer, sheet_name="DETALLE_ESTUDIANTES", index=False)

    print("\n========================================")
    print("PROCESAMIENTO DE COMUNICACIÓN COMPLETADO")
    print("========================================")
    print(f"Total de registros: {len(df_final)}")
    print(df_final["NIVEL_LOGRO_GLOBAL"].value_counts().to_string())
    print(f"Archivo generado: {archivo_salida}")

    return df_final


if __name__ == "__main__":
    procesar_datos()
