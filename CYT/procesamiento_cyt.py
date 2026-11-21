# ================================================================
# TUKUYRICUY 2026
# PROCESAMIENTO DE RESULTADOS - CIENCIA Y TECNOLOGÍA
# Archivo: procesamiento_cyt.py
# ================================================================

import pandas as pd
import numpy as np
from pathlib import Path

from config_cyt import (
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
    COLUMNAS_PORCENTAJES,

    obtener_competencias,
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
            "No se encontró la columna NIVEL "
            "en la hoja de resultados."
        )

    return df


# ================================================================
# 3. VALIDAR COLUMNAS V1-V16
# ================================================================

def validar_columnas_valoracion(df, nombre_hoja):
    faltantes = [
        c for c in COLUMNAS_VALORACION
        if c not in df.columns
    ]

    if faltantes:
        raise ValueError(
            f"La hoja '{nombre_hoja}' no contiene "
            "las columnas de valoración necesarias: "
            f"{', '.join(faltantes)}."
        )


# ================================================================
# 4. LIMPIAR IDENTIFICACIÓN
# ================================================================

def limpiar_identificacion(df):
    df = df.copy()

    # Normalizar campos de texto.

    columnas_texto = [
        "INSTITUCION",
        "SECCION",
        "APELLIDOS",
        "NOMBRES",
        "SEXO",
        "ID_REGISTRO",
    ]

    for columna in columnas_texto:

        if columna in df.columns:
            df[columna] = (
                df[columna]
                .fillna("")
                .astype(str)
                .str.strip()
            )

    # Normalizar institución.

    if "INSTITUCION" in df.columns:
        df["INSTITUCION"] = (
            df["INSTITUCION"]
            .str.upper()
            .replace("", "SIN DATO")
        )

    # Normalizar sección.

    if "SECCION" in df.columns:
        df["SECCION"] = (
            df["SECCION"]
            .str.upper()
            .replace("", "SIN DATO")
        )

    # Normalizar grado.

    if "GRADO" in df.columns:
        df["GRADO"] = (
            pd.to_numeric(
                df["GRADO"],
                errors="coerce"
            )
        )

    # Normalizar DNI.

    if "DNI" in df.columns:
        df["DNI"] = (
            df["DNI"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.replace(r"\.0$", "", regex=True)
        )

    return df


# ================================================================
# 5. QUITAR COLUMNAS CALCULADAS DEL ORIGEN
# ================================================================
#
# Todo se recalcula desde V1-V16.
#
# Esto permite mantener los mismos criterios de evaluación,
# porcentajes y niveles de logro en cada ejecución.
# ================================================================

def eliminar_columnas_calculadas(df):
    a_eliminar = [
        c for c in df.columns
        if c.startswith("COMP_")
           or c.startswith("PORCENTAJE_")
           or c in COLUMNAS_RESULTADOS
           or c in COLUMNAS_RESULTADO_GLOBAL
           or c in ["PUNTAJE_GLOBAL", "NIVEL_LOGRO_GLOBAL"]
    ]

    return df.drop(
        columns=a_eliminar,
        errors="ignore"
    )


# ================================================================
# 6. CALCULAR RESULTADOS GENERALES
# ================================================================

def calcular_resultados(df):
    df = df.copy()

    # ------------------------------------------------------------
    # 6.1. VALORACIONES: SOLO 0 O 1
    # ------------------------------------------------------------

    valores = df[COLUMNAS_VALORACION].apply(
        pd.to_numeric,
        errors="coerce"
    )

    validos = valores.isin([0, 1])

    invalidos = int(
        (~validos).sum().sum()
    )

    if invalidos:
        print(
            f"AVISO: {invalidos} valoraciones vacías "
            "o distintas de 0/1 se contaron como incorrectas."
        )

    df[COLUMNAS_VALORACION] = (
        valores
        .where(validos, 0)
        .astype(int)
    )

    # ------------------------------------------------------------
    # 6.2. RESULTADO GENERAL
    # ------------------------------------------------------------

    df["CORRECTAS"] = (
        df[COLUMNAS_VALORACION]
        .sum(axis=1)
    )

    df["TOTAL_PREGUNTAS"] = TOTAL_PREGUNTAS

    df["INCORRECTAS"] = (
            TOTAL_PREGUNTAS - df["CORRECTAS"]
    )

    df["PORCENTAJE"] = (
            df["CORRECTAS"] /
            TOTAL_PREGUNTAS * 100
    ).round(2)

    # ------------------------------------------------------------
    # 6.3. INICIALIZAR COLUMNAS DE COMPETENCIAS
    # ------------------------------------------------------------

    columnas_ponderadas = []

    for competencia in COMPETENCIAS:
        pref = competencia["codigo"]

        df[f"{pref}_NOMBRE"] = ""

        df[f"{pref}_PREGUNTAS"] = ""

        df[f"{pref}_ACIERTOS"] = 0

        df[f"{pref}_TOTAL"] = 0

        df[f"{pref}_PORCENTAJE"] = 0.0

        df[f"{pref}_NIVEL"] = "REVISAR"

        df[f"{pref}_PESO"] = competencia["peso"]

        df[f"{pref}_PUNTAJE_PONDERADO"] = 0.0

        columnas_ponderadas.append(
            f"{pref}_PUNTAJE_PONDERADO"
        )

    # ------------------------------------------------------------
    # 6.4. CALCULAR COMPETENCIAS SEGÚN NIVEL Y GRADO
    # ------------------------------------------------------------
    #
    # IMPORTANTE:
    # 1.° de secundaria tiene una distribución especial.
    #
    # Por eso se obtiene la distribución correspondiente
    # para cada grupo de estudiantes.
    # ------------------------------------------------------------

    grupos = df.groupby(
        ["NIVEL", "GRADO"],
        sort=False
    ).groups

    for (nivel, grado), indices in grupos.items():

        competencias_grupo = obtener_competencias(
            nivel,
            grado
        )

        print(
            f"Calculando competencias: "
            f"{nivel} - {grado}.° grado"
        )

        for competencia in competencias_grupo:
            pref = competencia["codigo"]

            preguntas = competencia["preguntas"]

            columnas = [
                f"V{p}"
                for p in preguntas
            ]

            total = len(columnas)

            peso = competencia["peso"]

            # Aciertos de la competencia.

            aciertos = (
                df.loc[indices, columnas]
                .sum(axis=1)
            )

            # Porcentaje de logro (0-100).

            porcentaje = (
                    aciertos / total * 100
            ).round(2)

            # Clasificación del nivel.

            niveles = aciertos.map(
                lambda a, c=competencia:
                clasificar_nivel_competencia(a, c)
            )

            # Puntaje ponderado.

            ponderado = aciertos.map(
                lambda a, t=total, p=peso:
                calcular_puntaje_ponderado(
                    a,
                    t,
                    p
                )
            )

            # Guardar resultados.

            df.loc[
                indices,
                f"{pref}_NOMBRE"
            ] = competencia["nombre"]

            df.loc[
                indices,
                f"{pref}_PREGUNTAS"
            ] = ", ".join(
                f"P{p}" for p in preguntas
            )

            df.loc[
                indices,
                f"{pref}_ACIERTOS"
            ] = aciertos

            df.loc[
                indices,
                f"{pref}_TOTAL"
            ] = total

            df.loc[
                indices,
                f"{pref}_PORCENTAJE"
            ] = porcentaje

            df.loc[
                indices,
                f"{pref}_NIVEL"
            ] = niveles

            df.loc[
                indices,
                f"{pref}_PESO"
            ] = peso

            df.loc[
                indices,
                f"{pref}_PUNTAJE_PONDERADO"
            ] = ponderado

    # ------------------------------------------------------------
    # 6.5. RESULTADO GLOBAL PONDERADO
    # ------------------------------------------------------------
    #
    # Cada competencia tiene un peso de 1/3.
    #
    # El resultado global es el promedio de los porcentajes
    # de logro de las tres competencias.
    # ------------------------------------------------------------

    df["PUNTAJE_GLOBAL_PONDERADO"] = (
        df[columnas_ponderadas]
        .sum(axis=1)
        .round(2)
    )

    # Alias para mantener compatibilidad con otros reportes.

    df["PUNTAJE_GLOBAL"] = (
        df["PUNTAJE_GLOBAL_PONDERADO"]
    )

    df["PORCENTAJE_GLOBAL"] = (
        df["PUNTAJE_GLOBAL_PONDERADO"]
    )

    # Clasificación global.

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

            + [
                "PUNTAJE_GLOBAL",
            ]

    )

    principales = [
        c for c in orden
        if c in df.columns
    ]

    otras = [
        c for c in df.columns
        if c not in orden
    ]

    return df[
        principales + otras
        ]


# ================================================================
# 8. PROCESAR UNA HOJA
# ================================================================

def procesar_hoja(nombre_hoja):
    print(f"\nProcesando hoja: {nombre_hoja}")

    df = pd.read_excel(ARCHIVO_ORIGEN, sheet_name=nombre_hoja, dtype=object)

    df = normalizar_columnas(df)

    df = df.dropna(
        how="all"
    )

    if df.empty:
        raise ValueError(
            f"La hoja '{nombre_hoja}' no contiene registros."
        )

    df = agregar_nivel(
        df,
        nombre_hoja
    )

    # Validar columnas de valoración.
    validar_columnas_valoracion(        df,        nombre_hoja    )
    df = limpiar_identificacion(df)

    # Validar nivel y grado.
    grados_validos = {
        nivel: grados
        for nivel, grados in [
            ("PRIMARIA", [2, 3, 4, 5, 6]),
            ("SECUNDARIA", [1, 2, 3, 4, 5]),
        ]
    }

    mascara = df.apply(
        lambda fila: (
                fila["NIVEL"] in grados_validos
                and pd.notna(fila["GRADO"])
                and int(fila["GRADO"])
                in grados_validos[fila["NIVEL"]]
        ),
        axis=1
    )

    descartados = int((~mascara).sum())

    if descartados:
        print(
            f"AVISO: {descartados} registros "
            "fueron excluidos por nivel o grado no válido."
        )

    df = df.loc[mascara].copy()

    if df.empty:
        raise ValueError(
            "No se encontraron registros válidos "
            "de primaria (2.º-6.º) o secundaria (1.º-5.º)."
        )

    df["GRADO"] = (
        df["GRADO"].astype(int)
    )

    # Asegurar columnas de identificación.

    for columna in COLUMNAS_IDENTIFICACION:

        if columna not in df.columns:
            df[columna] = ""

    # Asegurar columnas de respuestas P1-P16.

    for columna in COLUMNAS_PREGUNTAS:

        if columna not in df.columns:
            df[columna] = ""

    # Se guarda el porcentaje del origen antes de recalcular todo.

    porcentaje_origen = (
        df["PORCENTAJE"].copy() if "PORCENTAJE" in df.columns else None
    )

    # Quitar resultados calculados previamente.
    df = eliminar_columnas_calculadas(df)

    # Calcular resultados.
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
            f"No se encontró el archivo de origen:\n"
            f"{ARCHIVO_ORIGEN}"
        )

    # Procesar todas las hojas configuradas.

    resultados = [

        procesar_hoja(hoja)

        for hoja in HOJAS_ORIGEN

    ]

    # Consolidar resultados.

    df_final = pd.concat(
        resultados,
        ignore_index=True,
        sort=False
    )

    # ------------------------------------------------------------
    # VALIDAR REGISTROS DUPLICADOS
    # ------------------------------------------------------------

    if "ID_REGISTRO" in df_final.columns:

        ids = (
            df_final["ID_REGISTRO"]
            .astype(str)
            .str.strip()
        )

        ids_validos = (
                ids.notna()
                & ids.ne("")
        )

        repetidos = int(
            ids.loc[ids_validos].duplicated().sum()
        )

        if repetidos:
            print(
                f"AVISO: {repetidos} registros "
                "con ID_REGISTRO repetido."
            )

    # ------------------------------------------------------------
    # GUARDAR BASE PROCESADA
    # ------------------------------------------------------------

    archivo_salida = (
            CARPETA_RESULTADOS /
            NOMBRE_ARCHIVO_PROCESADO
    )

    archivo_salida.parent.mkdir(
        parents=True,
        exist_ok=True
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

    # ------------------------------------------------------------
    # RESUMEN DE EJECUCIÓN
    # ------------------------------------------------------------

    print("\n" + "=" * 64)

    print(
        "PROCESAMIENTO DE CIENCIA Y TECNOLOGÍA COMPLETADO"
    )

    print("=" * 64)

    print(
        f"Total de registros: {len(df_final):,}"
    )

    print("\nREGISTROS POR NIVEL:")

    print(
        df_final["NIVEL"]
        .value_counts()
        .to_string()
    )

    print("\nRESULTADOS GLOBALES:")

    print(
        df_final["NIVEL_LOGRO_GLOBAL"]
        .value_counts()
        .to_string()
    )

    print("\nPROMEDIO GLOBAL:")

    print(
        f"{df_final['PUNTAJE_GLOBAL_PONDERADO'].mean():.2f}%"
    )

    print("\nARCHIVO GENERADO:")

    print(archivo_salida)

    print("=" * 64)

    return df_final


# ================================================================
# 10. EJECUCIÓN PRINCIPAL
# ================================================================

if __name__ == "__main__":
    procesar_datos()
