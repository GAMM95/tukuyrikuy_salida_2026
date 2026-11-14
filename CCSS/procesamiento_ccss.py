
# ================================================================
# TUKUYRICUY 2026
# PROCESAMIENTO DE RESULTADOS - CIENCIAS SOCIALES (SECUNDARIA)
# Archivo: procesamiento_ccss.py
# ================================================================

import pandas as pd
import numpy as np
from pathlib import Path

from config_ccss import (
    ARCHIVO_ORIGEN,
    HOJAS_ORIGEN,
    NIVEL_OBJETIVO,
    GRADOS_OBJETIVO,
    CARPETA_RESULTADOS,
    NOMBRE_ARCHIVO_PROCESADO,
    COMPETENCIAS,
    TOTAL_PREGUNTAS,
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


# ================================================================
# 1. NORMALIZAR NOMBRES DE COLUMNAS
# ================================================================

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


# ================================================================
# 2. IDENTIFICAR NIVEL EDUCATIVO
# ================================================================

def agregar_nivel(df, nombre_hoja):
    df = df.copy()

    if "NIVEL" in df.columns:
        df["NIVEL"] = (
            df["NIVEL"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    elif nombre_hoja.upper() == "SECUNDARIA":
        df["NIVEL"] = "SECUNDARIA"

    else:
        raise ValueError(
            "No se encontró la columna NIVEL en la hoja "
            f"'{nombre_hoja}'."
        )

    return df


# ================================================================
# 3. VALIDAR COLUMNAS V1-V16
# ================================================================

def validar_columnas_valoracion(df, nombre_hoja):

    requeridas = ["NIVEL", "GRADO"] + COLUMNAS_VALORACION

    faltantes = [
        c for c in requeridas
        if c not in df.columns
    ]

    if faltantes:
        raise ValueError(
            f"La hoja '{nombre_hoja}' no contiene las columnas "
            f"necesarias: {', '.join(faltantes)}."
        )


# ================================================================
# 4. LIMPIAR IDENTIFICACIÓN
# ================================================================

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
                .replace("", "SIN DATO")
            )

    for columna in ["APELLIDOS", "NOMBRES", "SEXO"]:
        if columna in df.columns:
            df[columna] = (
                df[columna]
                .fillna("")
                .astype(str)
                .str.strip()
            )

    if "DNI" in df.columns:
        df["DNI"] = (
            df["DNI"]
            .fillna("")
            .astype(str)
            .str.replace(r"\.0$", "", regex=True)
            .str.strip()
        )

    if "GRADO" in df.columns:
        df["GRADO"] = (
            pd.to_numeric(df["GRADO"], errors="coerce")
            .fillna(0)
            .astype(int)
        )

    return df


# ================================================================
# 5. QUITAR COLUMNAS CALCULADAS DEL ORIGEN
# ================================================================
# Todo se recalcula desde V1-V16 para que los porcentajes,
# puntajes ponderados y niveles sigan las reglas de CCSS.


def eliminar_columnas_calculadas(df):

    a_eliminar = [
        c for c in df.columns
        if c.startswith("COMP_")
        or c.startswith("PORCENTAJE_")
        or c in COLUMNAS_RESULTADOS
        or c in COLUMNAS_RESULTADO_GLOBAL
        or c == "PUNTAJE_GLOBAL"
    ]

    return df.drop(columns=a_eliminar, errors="ignore")


# ================================================================
# 6. CALCULAR RESULTADOS
# ================================================================

def calcular_resultados(df):
    df = df.copy()

    # --- Valoraciones: solo 0 o 1 ---

    valores = df[COLUMNAS_VALORACION].apply(
        pd.to_numeric,
        errors="coerce"
    )

    validos = valores.isin([0, 1])
    invalidos = int((~validos).sum().sum())

    if invalidos:
        print(
            f"AVISO: {invalidos} valoraciones vacías o distintas "
            "de 0/1 se contaron como incorrectas."
        )

    df[COLUMNAS_VALORACION] = (
        valores.where(validos, 0).astype(int)
    )

    # --- Resultado general ---

    df["CORRECTAS"] = df[COLUMNAS_VALORACION].sum(axis=1)

    df["TOTAL_PREGUNTAS"] = TOTAL_PREGUNTAS

    df["INCORRECTAS"] = (
        TOTAL_PREGUNTAS - df["CORRECTAS"]
    )

    df["PORCENTAJE"] = (
        df["CORRECTAS"] / TOTAL_PREGUNTAS * 100
    ).round(2)

    # --- Por competencia ---

    columnas_ponderadas = []

    for c in COMPETENCIAS:

        pref = c["codigo"]

        columnas = [
            f"V{p}" for p in c["preguntas"]
        ]

        total = len(columnas)

        aciertos = df[columnas].sum(axis=1)

        peso = c["peso"]

        df[f"{pref}_NOMBRE"] = c["nombre"]

        df[f"{pref}_PREGUNTAS"] = ", ".join(
            f"P{p}" for p in c["preguntas"]
        )

        df[f"{pref}_ACIERTOS"] = aciertos

        df[f"{pref}_TOTAL"] = total

        df[f"{pref}_PORCENTAJE"] = (
            aciertos / total * 100
        ).round(2)

        # Cada competencia utiliza sus propios límites.
        df[f"{pref}_NIVEL"] = aciertos.map(
            lambda a, comp=c:
                clasificar_nivel_competencia(a, comp)
        )

        df[f"{pref}_PESO"] = peso

        # Puntaje ponderado sobre 100.
        df[f"{pref}_PUNTAJE_PONDERADO"] = aciertos.map(
            lambda a, comp=c:
                calcular_puntaje_ponderado(a, comp)
        )

        columnas_ponderadas.append(
            f"{pref}_PUNTAJE_PONDERADO"
        )

    # --- Resultado global sobre 100 ---

    df["PUNTAJE_GLOBAL_PONDERADO"] = (
        df[columnas_ponderadas]
        .sum(axis=1)
        .round(2)
    )

    df["PORCENTAJE_GLOBAL"] = (
        df["PUNTAJE_GLOBAL_PONDERADO"]
    )

    df["NIVEL_LOGRO_GLOBAL"] = (
        df["PUNTAJE_GLOBAL_PONDERADO"]
        .map(clasificar_nivel_global)
    )

    return df

# ================================================================
# PORCENTAJES DE CONTROL
# ================================================================

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
        df["PORCENTAJE_FUENTE_NORMALIZADO"] = (
            porcentaje_origen.reindex(df.index).map(normalizar_porcentaje)
        )
    else:
        df["PORCENTAJE_FUENTE_NORMALIZADO"] = np.nan

    df["PORCENTAJE_CALCULADO"] = df["CORRECTAS"] / TOTAL_PREGUNTAS

    return df

# ================================================================
# 7. ORDENAR COLUMNAS
# ================================================================

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

    principales = [
        c for c in orden
        if c in df.columns
    ]

    otras = [
        c for c in df.columns
        if c not in orden
    ]

    return df[principales + otras]

# ================================================================
# 8. PROCESAR UNA HOJA
# ================================================================

def procesar_hoja(nombre_hoja):

    print(f"Procesando hoja: {nombre_hoja}")

    df = pd.read_excel(
        ARCHIVO_ORIGEN,
        sheet_name=nombre_hoja,
        dtype=object
    )

    df = normalizar_columnas(df)

    df = df.dropna(how="all")

    df = agregar_nivel(df, nombre_hoja)

    validar_columnas_valoracion(df, nombre_hoja)

    # --- Filtrar secundaria ---

    df = df[
        df["NIVEL"] == NIVEL_OBJETIVO
    ].copy()

    # --- Limpiar identificación ---

    df = limpiar_identificacion(df)

    # --- Filtrar grados 1 a 5 ---

    df = df[
        df["GRADO"].isin(GRADOS_OBJETIVO)
    ].copy()

    if df.empty:
        raise ValueError(
            "No hay registros de SECUNDARIA "
            "de 1.º a 5.º grado."
        )

    # --- Recalcular resultados ---

    # Se guarda el porcentaje del origen antes de recalcular todo
    porcentaje_origen = (
        df["PORCENTAJE"].copy() if "PORCENTAJE" in df.columns else None
    )

    df = eliminar_columnas_calculadas(df)

    df = calcular_resultados(df)

    df = agregar_porcentajes(df, porcentaje_origen)

    return ordenar_columnas(df)


# ================================================================
# 9. PROCESAR TODAS LAS HOJAS
# ================================================================

def procesar_datos():

    crear_carpetas()

    validar_configuracion()

    if not Path(ARCHIVO_ORIGEN).exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de origen:\n{ARCHIVO_ORIGEN}"
        )

    # --- Identificar hoja de origen ---

    libro = pd.ExcelFile(ARCHIVO_ORIGEN)

    hoja_preferida = HOJAS_ORIGEN[0]

    if hoja_preferida in libro.sheet_names:
        hoja_seleccionada = hoja_preferida
    else:
        hoja_seleccionada = libro.sheet_names[0]

        print(
            f"AVISO: No existe la hoja '{hoja_preferida}'. "
            f"Se utilizará '{hoja_seleccionada}'."
        )

    # --- Procesar hoja seleccionada ---

    resultados = [
        procesar_hoja(hoja_seleccionada)
    ]

    df_final = pd.concat(
        resultados,
        ignore_index=True,
        sort=False
    )

    # --- Verificar registros repetidos ---

    if "ID_REGISTRO" in df_final.columns:

        repetidos = int(
            df_final["ID_REGISTRO"].duplicated().sum()
        )

        if repetidos:
            print(
                f"AVISO: {repetidos} registros "
                "con ID_REGISTRO repetido."
            )

    # --- Guardar archivo procesado ---

    archivo_salida = (
        CARPETA_RESULTADOS / NOMBRE_ARCHIVO_PROCESADO
    )

    with pd.ExcelWriter(
        archivo_salida,
        engine="openpyxl"
    ) as writer:

        df_final.to_excel(
            writer,
            sheet_name="DETALLE_ESTUDIANTES",
            index=False
        )

    print("\n========================================")
    print("PROCESAMIENTO DE CCSS COMPLETADO")
    print("========================================")

    print(f"Área: CIENCIAS SOCIALES")
    print(f"Nivel: {NIVEL_OBJETIVO}")
    print(f"Grados: {GRADOS_OBJETIVO}")

    print(f"Total de registros: {len(df_final):,}")

    print("\nRESULTADOS GLOBALES:")

    print(
        df_final["NIVEL_LOGRO_GLOBAL"]
        .value_counts()
        .to_string()
    )

    print(f"\nArchivo generado: {archivo_salida}")

    return df_final


if __name__ == "__main__":
    procesar_datos()
