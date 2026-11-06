# -*- coding: utf-8 -*-
"""Procesamiento de resultados de Ciencia y Tecnología."""
import pandas as pd
import numpy as np

from config_cyt import (
    ARCHIVO_ENTRADA, HOJA_ENTRADA, ARCHIVO_BASE_PROCESADA,
    COLUMNAS_IDENTIFICACION, COLUMNAS_RESPUESTAS, COLUMNAS_VALORACIONES,
    COMPETENCIAS, NIVELES, clasificar_competencia, clasificar_global
)

def _normalizar_columnas(df):
    df.columns = [str(c).strip().upper() for c in df.columns]
    return df

def _normalizar_grado(valor):
    if pd.isna(valor):
        return np.nan
    try:
        return int(float(str(valor).strip()))
    except (ValueError, TypeError):
        return np.nan

def procesar_resultados():
    if not ARCHIVO_ENTRADA.exists():
        raise FileNotFoundError(f"No existe el archivo: {ARCHIVO_ENTRADA}")

    xls = pd.ExcelFile(ARCHIVO_ENTRADA)
    hoja = HOJA_ENTRADA if HOJA_ENTRADA in xls.sheet_names else xls.sheet_names[0]
    print(f"Procesando hoja: {hoja}")
    df = pd.read_excel(ARCHIVO_ENTRADA, sheet_name=hoja)
    df = _normalizar_columnas(df)

    requeridas = ["NIVEL", "GRADO", "INSTITUCION"] + COLUMNAS_VALORACIONES
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError("Faltan columnas obligatorias: " + ", ".join(faltantes))

    df["NIVEL"] = df["NIVEL"].astype(str).str.strip().str.upper()
    df["GRADO"] = df["GRADO"].apply(_normalizar_grado)

    # Se conservan únicamente los grados incluidos en la evaluación.
    mascara = (
        ((df["NIVEL"] == "PRIMARIA") & df["GRADO"].isin(NIVELES["PRIMARIA"]["grados"])) |
        ((df["NIVEL"] == "SECUNDARIA") & df["GRADO"].isin(NIVELES["SECUNDARIA"]["grados"]))
    )
    df = df.loc[mascara].copy()

    if df.empty:
        raise ValueError("No se encontraron registros de primaria (2.º–6.º) o secundaria (1.º–5.º).")

    for col in COLUMNAS_VALORACIONES:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0).clip(0, 1).astype(int)

    for col in COLUMNAS_IDENTIFICACION + COLUMNAS_RESPUESTAS:
        if col not in df.columns:
            df[col] = ""

    # Normalización de datos de identificación.
    for col in ["INSTITUCION", "SECCION", "APELLIDOS", "NOMBRES", "SEXO", "ID_REGISTRO"]:
        df[col] = df[col].fillna("").astype(str).str.strip()
    df["DNI"] = df["DNI"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
    df["GRADO"] = df["GRADO"].astype(int)

    df["CORRECTAS"] = df[COLUMNAS_VALORACIONES].sum(axis=1)
    df["INCORRECTAS"] = 16 - df["CORRECTAS"]
    df["TOTAL_PREGUNTAS"] = 16
    df["PORCENTAJE"] = df["CORRECTAS"] / 16

    porcentajes_ponderados = []
    for comp in COMPETENCIAS:
        columnas = [f"V{i}" for i in comp["preguntas"]]
        aciertos = df[columnas].sum(axis=1)
        total = len(columnas)
        df[f"{comp['id']}_NOMBRE"] = comp["nombre"]
        df[f"{comp['id']}_ACIERTOS"] = aciertos
        df[f"{comp['id']}_TOTAL"] = total
        df[f"{comp['id']}_PORCENTAJE"] = aciertos / total
        df[f"{comp['id']}_NIVEL"] = aciertos.apply(
            lambda x, c=comp: clasificar_competencia(int(x), c)
        )
        porcentajes_ponderados.append(df[f"{comp['id']}_PORCENTAJE"] * comp["peso"])

    df["PORCENTAJE_GLOBAL"] = sum(porcentajes_ponderados) * 100
    df["NIVEL_LOGRO_GLOBAL"] = df["PORCENTAJE_GLOBAL"].apply(clasificar_global)

    # Orden de columnas: identificación, respuestas, valoración y resultados.
    columnas_resultado = [
        "CORRECTAS", "INCORRECTAS", "TOTAL_PREGUNTAS", "PORCENTAJE",
        "PORCENTAJE_GLOBAL", "NIVEL_LOGRO_GLOBAL"
    ]
    for comp in COMPETENCIAS:
        columnas_resultado += [
            f"{comp['id']}_NOMBRE", f"{comp['id']}_ACIERTOS",
            f"{comp['id']}_TOTAL", f"{comp['id']}_PORCENTAJE",
            f"{comp['id']}_NIVEL"
        ]
    orden = (
        COLUMNAS_IDENTIFICACION + COLUMNAS_RESPUESTAS +
        COLUMNAS_VALORACIONES + columnas_resultado
    )
    df = df[[c for c in orden if c in df.columns]]

    ARCHIVO_BASE_PROCESADA.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(ARCHIVO_BASE_PROCESADA, index=False, engine="openpyxl")
    print("\\nPROCESAMIENTO DE CIENCIA Y TECNOLOGÍA COMPLETADO")
    print(f"Estudiantes procesados: {len(df):,}")
    print(df["NIVEL_LOGRO_GLOBAL"].value_counts().to_string())
    print(f"Archivo generado: {ARCHIVO_BASE_PROCESADA}")
    return df
