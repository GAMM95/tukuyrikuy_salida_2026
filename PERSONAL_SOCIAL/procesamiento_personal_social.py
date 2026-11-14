# ================================================================
# TUKUYRICUY 2026
# PROCESAMIENTO DE RESULTADOS - PERSONAL SOCIAL (PRIMARIA)
# Archivo: procesamiento_personal_social.py
# ================================================================

import pandas as pd
import numpy as np

from config_personal_social import (
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

def agregar_nivel(df):
    df = df.copy()

    if "NIVEL" not in df.columns:
        raise ValueError(
            "El archivo de origen no contiene la columna NIVEL. "
            "No es posible identificar los registros de primaria."
        )

    df["NIVEL"] = (
        df["NIVEL"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    return df


# ================================================================
# 3. VALIDAR COLUMNAS DE VALORACIÓN V1-V16
# ================================================================

def validar_columnas_valoracion(df):
    requeridas = (
            ["NIVEL", "GRADO"]
            + COLUMNAS_VALORACION
    )

    faltantes = [
        columna
        for columna in requeridas
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Faltan columnas necesarias en el Excel: "
            + ", ".join(faltantes)
        )


# ================================================================
# 4. LIMPIAR DATOS DE IDENTIFICACIÓN
# ================================================================

def limpiar_identificacion(df):
    df = df.copy()

    # Crear columnas de identificación faltantes.
    for columna in COLUMNAS_IDENTIFICACION:

        if columna not in df.columns:
            df[columna] = ""

    columnas_texto = [
        "NIVEL",
        "INSTITUCION",
        "SECCION",
        "APELLIDOS",
        "NOMBRES",
        "SEXO",
    ]

    for columna in columnas_texto:

        if columna in df.columns:
            df[columna] = (
                df[columna]
                .fillna("")
                .astype(str)
                .str.strip()
            )

    # Normalizar nivel e institución.
    df["NIVEL"] = df["NIVEL"].str.upper()

    df["INSTITUCION"] = (
        df["INSTITUCION"]
        .str.upper()
        .replace("", "SIN DATO")
    )

    df["SECCION"] = (
        df["SECCION"]
        .str.upper()
        .replace("", "SIN DATO")
    )

    # Normalizar grado.
    df["GRADO"] = (
        pd.to_numeric(df["GRADO"], errors="coerce")
        .fillna(0)
        .astype(int)
    )

    # Normalizar DNI.
    df["DNI"] = (
        df["DNI"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )

    return df


# ================================================================
# 5. ELIMINAR COLUMNAS CALCULADAS ANTERIORMENTE
# ================================================================
def eliminar_columnas_calculadas(df):
    df = df.copy()

    columnas_calculadas = (
            COLUMNAS_RESULTADOS
            + COLUMNAS_COMPETENCIAS
            + COLUMNAS_RESULTADO_GLOBAL
            + ["PUNTAJE_GLOBAL"]
    )

    columnas_eliminar = [
        columna
        for columna in df.columns
        if (
                columna.startswith("COMP_")
                or columna.startswith("PORCENTAJE_")
                or columna in columnas_calculadas
        )
    ]

    return df.drop(
        columns=columnas_eliminar,
        errors="ignore"
    )


# ================================================================
# 6. CALCULAR RESULTADOS
# ================================================================

def calcular_resultados(df):
    df = df.copy()

    # ------------------------------------------------------------
    # 6.1. NORMALIZAR VALORACIONES V1-V16
    # ------------------------------------------------------------

    valores = df[COLUMNAS_VALORACION].apply(
        pd.to_numeric,
        errors="coerce"
    )

    validos = valores.isin([0, 1])

    invalidos = int((~validos).sum().sum())

    if invalidos > 0:
        print(
            f"ADVERTENCIA: Se encontraron {invalidos} "
            "valoraciones vacías o inválidas. "
            "Se contabilizarán como incorrectas."
        )

    df[COLUMNAS_VALORACION] = (
        valores
        .where(validos, 0)
        .fillna(0)
        .astype(int)
    )

    # ------------------------------------------------------------
    # 6.2. RESULTADOS GENERALES
    # ------------------------------------------------------------

    df["CORRECTAS"] = (
        df[COLUMNAS_VALORACION].sum(axis=1)
    )

    df["TOTAL_PREGUNTAS"] = TOTAL_PREGUNTAS

    df["INCORRECTAS"] = (
            TOTAL_PREGUNTAS - df["CORRECTAS"]
    )

    df["PORCENTAJE"] = (
            df["CORRECTAS"] / TOTAL_PREGUNTAS * 100
    ).round(2)

    # ------------------------------------------------------------
    # 6.3. RESULTADOS POR COMPETENCIA
    # ------------------------------------------------------------

    columnas_ponderadas = []

    for competencia in COMPETENCIAS:
        codigo = competencia["codigo"]

        preguntas = [
            f"V{i}"
            for i in competencia["preguntas"]
        ]

        total = len(preguntas)

        aciertos = df[preguntas].sum(axis=1)

        porcentaje = (
                aciertos / total * 100
        ).round(2)

        # Nombre de la competencia.
        df[f"{codigo}_NOMBRE"] = (
            competencia["nombre"]
        )

        # Preguntas correspondientes.
        df[f"{codigo}_PREGUNTAS"] = ", ".join(
            f"P{i}"
            for i in competencia["preguntas"]
        )

        # Cantidad de aciertos.
        df[f"{codigo}_ACIERTOS"] = aciertos

        # Total de preguntas.
        df[f"{codigo}_TOTAL"] = total

        # Porcentaje de aciertos.
        df[f"{codigo}_PORCENTAJE"] = porcentaje

        # Nivel de logro.
        df[f"{codigo}_NIVEL"] = aciertos.apply(
            lambda x: clasificar_nivel_competencia(
                x,
                competencia
            )
        )

        # Peso de la competencia.
        df[f"{codigo}_PESO"] = (
            competencia["peso"]
        )

        # Puntaje ponderado.
        df[f"{codigo}_PUNTAJE_PONDERADO"] = (
            aciertos.apply(
                lambda x: calcular_puntaje_ponderado(
                    x,
                    competencia
                )
            )
        )

        columnas_ponderadas.append(
            f"{codigo}_PUNTAJE_PONDERADO"
        )

    # ------------------------------------------------------------
    # 6.4. RESULTADO GLOBAL PONDERADO
    # ------------------------------------------------------------

    df["PUNTAJE_GLOBAL_PONDERADO"] = (
        df[columnas_ponderadas]
        .sum(axis=1)
        .round(2)
    )

    # El puntaje global utiliza una escala de 0 a 100.
    df["PORCENTAJE_GLOBAL"] = (
        df["PUNTAJE_GLOBAL_PONDERADO"]
    )

    # Nivel de logro global.
    df["NIVEL_LOGRO_GLOBAL"] = (
        df["PUNTAJE_GLOBAL_PONDERADO"]
        .apply(clasificar_nivel_global)
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

    existentes = [
        columna
        for columna in orden
        if columna in df.columns
    ]

    adicionales = [
        columna
        for columna in df.columns
        if columna not in existentes
    ]

    return df[existentes + adicionales]


# ================================================================
# 8. PROCESAR UNA HOJA
# ================================================================

def procesar_hoja(nombre_hoja):
    print(f"\nProcesando hoja: {nombre_hoja}")

    df = pd.read_excel(ARCHIVO_ORIGEN, sheet_name=nombre_hoja, dtype=object)
    df = normalizar_columnas(df)
    df = df.dropna(how="all").copy()

    if df.empty:
        print(
            f"La hoja {nombre_hoja} no contiene registros."
        )

        return pd.DataFrame()

    # 8.3. Identificar nivel educativo.
    df = agregar_nivel(df)

    # 8.4. Validar columnas necesarias.
    validar_columnas_valoracion(df)

    # Se guarda el porcentaje del origen antes de recalcular todo
    porcentaje_origen = (
        df["PORCENTAJE"].copy() if "PORCENTAJE" in df.columns else None
    )
    # 8.5. Limpiar identificación.
    df = limpiar_identificacion(df)

    # 8.6. Filtrar únicamente primaria.
    df = df[
        df["NIVEL"] == NIVEL_OBJETIVO
        ].copy()

    # 8.7. Filtrar grados 2.º a 6.º.
    df = df[
        df["GRADO"].isin(GRADOS_OBJETIVO)
    ].copy()

    if df.empty:
        print(
            f"No hay registros de {NIVEL_OBJETIVO} "
            f"entre los grados {GRADOS_OBJETIVO}."
        )

        return pd.DataFrame()

    # 8.8. Eliminar resultados anteriores.
    df = eliminar_columnas_calculadas(df)

    # 8.9. Calcular resultados.
    df = calcular_resultados(df)
    df = agregar_porcentajes(df, porcentaje_origen)

    # 8.10. Ordenar columnas.
    df = ordenar_columnas(df)

    print(
        f"Registros procesados en {nombre_hoja}: {len(df):,}"
    )

    return df


# ================================================================
# 9. PROCESAR TODAS LAS HOJAS Y EXPORTAR
# ================================================================

def procesar_datos():
    # ------------------------------------------------------------
    # 9.1. Crear carpetas y validar configuración.
    # ------------------------------------------------------------

    crear_carpetas()

    validar_configuracion()

    # ------------------------------------------------------------
    # 9.2. Verificar archivo de origen.
    # ------------------------------------------------------------

    if not ARCHIVO_ORIGEN.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de origen: "
            f"{ARCHIVO_ORIGEN}"
        )

    # ------------------------------------------------------------
    # 9.3. Identificar hoja disponible.
    # ------------------------------------------------------------

    xls = pd.ExcelFile(ARCHIVO_ORIGEN)

    hojas_disponibles = xls.sheet_names

    if not hojas_disponibles:
        raise ValueError(
            "El archivo Excel no contiene hojas."
        )

    hoja_preferida = next(
        (
            nombre
            for nombre in HOJAS_ORIGEN
            if nombre in hojas_disponibles
        ),
        hojas_disponibles[0]
    )

    print("\n" + "=" * 60)
    print("TUKUYRICUY 2026 - PERSONAL SOCIAL")
    print("=" * 60)

    print(
        f"Archivo de origen: {ARCHIVO_ORIGEN.name}"
    )

    print(
        f"Hoja utilizada: {hoja_preferida}"
    )

    # ------------------------------------------------------------
    # 9.4. Procesar hoja seleccionada.
    # ------------------------------------------------------------

    df = procesar_hoja(hoja_preferida)

    if df.empty:
        raise ValueError(
            "No se encontraron registros válidos de "
            "Personal Social de primaria, grados 2.º a 6.º."
        )

    # ------------------------------------------------------------
    # 9.5. Validar registros duplicados.
    # ------------------------------------------------------------

    if "ID_REGISTRO" in df.columns:

        ids = (
            df["ID_REGISTRO"]
            .fillna("")
            .astype(str)
            .str.strip()
        )

        ids_validos = ids.ne("")

        duplicados = (
                ids_validos
                & ids.duplicated(keep=False)
        )

        if duplicados.any():
            cantidad = ids[duplicados].nunique()

            print(
                f"ADVERTENCIA: Se encontraron {cantidad} "
                "ID_REGISTRO duplicados."
            )

    # ------------------------------------------------------------
    # 9.6. Exportar base procesada.
    # ------------------------------------------------------------

    salida = (
            CARPETA_RESULTADOS
            / NOMBRE_ARCHIVO_PROCESADO
    )

    with pd.ExcelWriter(
            salida,
            engine="openpyxl"
    ) as writer:

        df.to_excel(
            writer,
            sheet_name="DETALLE_ESTUDIANTES",
            index=False
        )

    # ------------------------------------------------------------
    # 9.7. Mostrar resumen.
    # ------------------------------------------------------------

    print("\n" + "=" * 60)
    print("PROCESAMIENTO COMPLETADO")
    print("=" * 60)

    print("Área: PERSONAL SOCIAL")

    print(
        f"Nivel: {NIVEL_OBJETIVO}"
    )

    print(
        f"Grados: {GRADOS_OBJETIVO}"
    )

    print(
        f"Estudiantes procesados: {len(df):,}"
    )

    print("\nDistribución global:")

    print(
        df["NIVEL_LOGRO_GLOBAL"]
        .value_counts()
        .reindex(
            ["INICIO", "PROCESO", "LOGRADO"],
            fill_value=0
        )
        .to_string()
    )

    print(
        f"\nArchivo generado: {salida}"
    )

    return df


# ================================================================
# 10. EJECUCIÓN PRINCIPAL
# ================================================================

if __name__ == "__main__":
    procesar_datos()
