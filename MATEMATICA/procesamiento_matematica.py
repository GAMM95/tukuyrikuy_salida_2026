# ================================================================
# TUKUYRICUY 2026
# PROCESAMIENTO DE RESULTADOS - MATEMÁTICA
# Archivo: procesamiento_matematica.py
# ================================================================

import pandas as pd
from pathlib import Path

from config_matematica import (
    ARCHIVO_ORIGEN,
    HOJAS_ORIGEN,
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
    clasificar_nivel_competencia,
    calcular_puntaje_ponderado,
    clasificar_nivel_global,
    crear_carpetas,
    validar_configuracion,
)


# ------------------------------------------------
# 1. NORMALIZAR NOMBRES DE COLUMNAS
# ------------------------------------------------

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


# ------------------------------------------------
# 2. IDENTIFICAR NIVEL EDUCATIVO
# ------------------------------------------------

def agregar_nivel(df, nombre_hoja):

    if "NIVEL" in df.columns:
        df["NIVEL"] = (
            df["NIVEL"]
            .astype(str)
            .str.strip()
            .str.upper()
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


# ------------------------------------------------
# 3. VALIDAR COLUMNAS V1-V16
# ------------------------------------------------

def validar_columnas_valoracion(df, nombre_hoja):
    faltantes = [c for c in COLUMNAS_VALORACION if c not in df.columns]

    if faltantes:
        raise ValueError(
            f"La hoja '{nombre_hoja}' no contiene las columnas de "
            f"valoración necesarias: {', '.join(faltantes)}."
        )


# ------------------------------------------------
# 4. LIMPIAR IDENTIFICACIÓN
# ------------------------------------------------

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

    if "GRADO" in df.columns:
        df["GRADO"] = (
            pd.to_numeric(df["GRADO"], errors="coerce")
            .fillna(0)
            .astype(int)
        )

    return df


# ------------------------------------------------
# 5. QUITAR COLUMNAS CALCULADAS DEL ORIGEN
# ------------------------------------------------
# Todo se recalcula desde V1-V16 para que los porcentajes
# (0-100) y los niveles sigan la misma regla.

def eliminar_columnas_calculadas(df):
    a_eliminar = [
        c for c in df.columns
        if c.startswith("COMP_")
        or c in COLUMNAS_RESULTADOS
        or c in COLUMNAS_RESULTADO_GLOBAL
    ]

    return df.drop(columns=a_eliminar)


# ------------------------------------------------
# 6. CALCULAR RESULTADOS
# ------------------------------------------------

def calcular_resultados(df):
    df = df.copy()

    # --- Valoraciones: solo 0 o 1 ---
    valores = df[COLUMNAS_VALORACION].apply(pd.to_numeric, errors="coerce")
    validos = valores.isin([0, 1])
    invalidos = int((~validos).sum().sum())

    if invalidos:
        print(
            f"AVISO: {invalidos} valoraciones vacías o distintas de 0/1 "
            "se contaron como incorrectas."
        )

    df[COLUMNAS_VALORACION] = valores.where(validos, 0).astype(int)

    # --- Resultado general ---
    df["CORRECTAS"] = df[COLUMNAS_VALORACION].sum(axis=1)
    df["TOTAL_PREGUNTAS"] = TOTAL_PREGUNTAS
    df["INCORRECTAS"] = TOTAL_PREGUNTAS - df["CORRECTAS"]
    df["PORCENTAJE"] = (df["CORRECTAS"] / TOTAL_PREGUNTAS * 100).round(2)

    # --- Por competencia ---
    columnas_ponderadas = []

    for c in COMPETENCIAS:
        pref = c["codigo"]
        columnas = [f"V{p}" for p in c["preguntas"]]
        total = len(columnas)
        aciertos = df[columnas].sum(axis=1)
        peso = c["peso"]

        df[f"{pref}_NOMBRE"] = c["nombre"]
        df[f"{pref}_PREGUNTAS"] = ", ".join(f"P{p}" for p in c["preguntas"])
        df[f"{pref}_ACIERTOS"] = aciertos
        df[f"{pref}_TOTAL"] = total
        df[f"{pref}_PORCENTAJE"] = (aciertos / total * 100).round(2)
        df[f"{pref}_NIVEL"] = aciertos.map(clasificar_nivel_competencia)
        df[f"{pref}_PESO"] = peso
        df[f"{pref}_PUNTAJE_PONDERADO"] = aciertos.map(
            lambda a, peso=peso: calcular_puntaje_ponderado(a, peso)
        )

        columnas_ponderadas.append(f"{pref}_PUNTAJE_PONDERADO")

    # --- Global sobre 100 ---
    df["PUNTAJE_GLOBAL"] = df[columnas_ponderadas].sum(axis=1).round(2)
    df["NIVEL_LOGRO_GLOBAL"] = df["PUNTAJE_GLOBAL"].map(clasificar_nivel_global)

    return df


# ------------------------------------------------
# 7. ORDENAR COLUMNAS
# ------------------------------------------------

def ordenar_columnas(df):
    orden = (
        COLUMNAS_IDENTIFICACION
        + COLUMNAS_PREGUNTAS
        + COLUMNAS_VALORACION
        + COLUMNAS_RESULTADOS
        + COLUMNAS_COMPETENCIAS
        + COLUMNAS_RESULTADO_GLOBAL
    )

    principales = [c for c in orden if c in df.columns]
    otras = [c for c in df.columns if c not in orden]

    return df[principales + otras]


# ------------------------------------------------
# 8. PROCESAR UNA HOJA
# ------------------------------------------------

def procesar_hoja(nombre_hoja):
    print(f"Procesando hoja: {nombre_hoja}")

    df = pd.read_excel(ARCHIVO_ORIGEN, sheet_name=nombre_hoja, dtype=object)

    df = normalizar_columnas(df)
    df = df.dropna(how="all")
    df = agregar_nivel(df, nombre_hoja)

    validar_columnas_valoracion(df, nombre_hoja)

    df = limpiar_identificacion(df)
    df = eliminar_columnas_calculadas(df)
    df = calcular_resultados(df)

    return ordenar_columnas(df)


# ------------------------------------------------
# 9. PROCESAR TODAS LAS HOJAS
# ------------------------------------------------

def procesar_datos():
    crear_carpetas()
    validar_configuracion()

    if not Path(ARCHIVO_ORIGEN).exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de origen:\n{ARCHIVO_ORIGEN}"
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
    print("PROCESAMIENTO DE MATEMÁTICA COMPLETADO")
    print("========================================")
    print(f"Total de registros: {len(df_final)}")
    print(df_final["NIVEL_LOGRO_GLOBAL"].value_counts().to_string())
    print(f"Archivo generado: {archivo_salida}")

    return df_final


if __name__ == "__main__":
    procesar_datos()