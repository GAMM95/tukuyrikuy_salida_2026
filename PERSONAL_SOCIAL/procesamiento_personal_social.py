import pandas as pd

from config_personal_social import (
    ARCHIVO_ORIGEN, HOJA_ORIGEN, NIVEL_OBJETIVO, GRADOS_OBJETIVO,
    CARPETA_RESULTADOS, NOMBRE_ARCHIVO_PROCESADO, COMPETENCIAS,
    TOTAL_PREGUNTAS, COLUMNAS_IDENTIFICACION, COLUMNAS_PREGUNTAS,
    COLUMNAS_VALORACION, COLUMNAS_RESULTADOS, COLUMNAS_COMPETENCIAS,
    COLUMNAS_RESULTADO_GLOBAL, clasificar_nivel_competencia,
    clasificar_nivel_global, crear_carpetas, validar_configuracion,
)


def normalizar_columnas(df):
    df = df.copy()
    df.columns = [
        str(col).strip().upper().replace(" ", "_").replace("-", "_")
        for col in df.columns
    ]
    return df


def cargar_origen():
    if not ARCHIVO_ORIGEN.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de origen: {ARCHIVO_ORIGEN}"
        )

    xls = pd.ExcelFile(ARCHIVO_ORIGEN)
    hoja = HOJA_ORIGEN if HOJA_ORIGEN in xls.sheet_names else xls.sheet_names[0]
    print(f"Procesando archivo: {ARCHIVO_ORIGEN.name}")
    print(f"Hoja utilizada: {hoja}")
    df = pd.read_excel(ARCHIVO_ORIGEN, sheet_name=hoja, dtype=object)
    return normalizar_columnas(df)


def validar_columnas(df):
    requeridas = ["NIVEL", "GRADO"] + COLUMNAS_VALORACION
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError(
            "Faltan columnas necesarias en el Excel: " + ", ".join(faltantes)
        )


def normalizar_identificacion(df):
    for columna in COLUMNAS_IDENTIFICACION:
        if columna not in df.columns:
            df[columna] = ""

    for columna in ["NIVEL", "INSTITUCION", "SECCION", "APELLIDOS", "NOMBRES", "SEXO"]:
        df[columna] = df[columna].fillna("").astype(str).str.strip()

    df["NIVEL"] = df["NIVEL"].str.upper()
    df["INSTITUCION"] = df["INSTITUCION"].replace("", "SIN DATO")
    df["SECCION"] = df["SECCION"].replace("", "SIN DATO")
    df["GRADO"] = pd.to_numeric(df["GRADO"], errors="coerce")
    df["DNI"] = df["DNI"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
    return df


def convertir_valoracion(serie, nombre_columna):
    original = serie.copy()
    numerica = pd.to_numeric(serie, errors="coerce")
    invalidos = numerica.notna() & ~numerica.isin([0, 1])
    if invalidos.any():
        ejemplos = original[invalidos].astype(str).unique()[:5]
        raise ValueError(
            f"{nombre_columna} contiene valores distintos de 0/1: {list(ejemplos)}"
        )
    # Los vacíos se consideran 0 (respuesta no acertada/no registrada).
    return numerica.fillna(0).astype(int)


def calcular_resultados(df):
    for columna in COLUMNAS_VALORACION:
        df[columna] = convertir_valoracion(df[columna], columna)

    df["CORRECTAS"] = df[COLUMNAS_VALORACION].sum(axis=1)
    df["TOTAL_PREGUNTAS"] = TOTAL_PREGUNTAS
    df["INCORRECTAS"] = TOTAL_PREGUNTAS - df["CORRECTAS"]
    df["PORCENTAJE"] = df["CORRECTAS"] / TOTAL_PREGUNTAS

    for competencia in COMPETENCIAS:
        codigo = competencia["codigo"]
        preguntas = [f"V{i}" for i in competencia["preguntas"]]
        aciertos = df[preguntas].sum(axis=1)
        total = len(preguntas)
        porcentaje = aciertos / total

        df[f"{codigo}_NOMBRE"] = competencia["nombre"]
        df[f"{codigo}_PREGUNTAS"] = ", ".join(
            f"P{i}" for i in competencia["preguntas"]
        )
        df[f"{codigo}_ACIERTOS"] = aciertos
        df[f"{codigo}_TOTAL"] = total
        df[f"{codigo}_PORCENTAJE"] = porcentaje
        df[f"{codigo}_NIVEL"] = aciertos.apply(
            lambda x: clasificar_nivel_competencia(x, competencia)
        )
        df[f"{codigo}_PESO"] = competencia["peso"]
        df[f"{codigo}_PUNTAJE_PONDERADO"] = porcentaje * competencia["peso"] * 100

    puntaje_global = sum(
        df[f"{c['codigo']}_PUNTAJE_PONDERADO"] for c in COMPETENCIAS
    )
    df["PUNTAJE_GLOBAL"] = puntaje_global
    df["NIVEL_LOGRO_GLOBAL"] = puntaje_global.apply(clasificar_nivel_global)
    return df


def ordenar_columnas(df):
    orden = (
        COLUMNAS_IDENTIFICACION + COLUMNAS_PREGUNTAS + COLUMNAS_VALORACION
        + COLUMNAS_RESULTADOS + COLUMNAS_COMPETENCIAS
        + COLUMNAS_RESULTADO_GLOBAL
    )
    existentes = [c for c in orden if c in df.columns]
    adicionales = [c for c in df.columns if c not in existentes]
    return df[existentes + adicionales]


def procesar_datos():
    crear_carpetas()
    validar_configuracion()

    df = cargar_origen()
    validar_columnas(df)

    # Mantener únicamente Personal Social de primaria.
    df = df[df["NIVEL"].astype(str).str.strip().str.upper() == NIVEL_OBJETIVO].copy()
    df = normalizar_identificacion(df)

    df = df[df["GRADO"].isin(GRADOS_OBJETIVO)].copy()
    if df.empty:
        raise ValueError(
            "No se encontraron registros de PRIMARIA entre 2.º y 6.º grado."
        )

    # Evitar utilizar columnas calculadas antiguas como fuente de resultados.
    columnas_calculadas = [
        c for c in df.columns
        if c.startswith("COMP_")
        or c in COLUMNAS_RESULTADOS + COLUMNAS_RESULTADO_GLOBAL
    ]
    df = df.drop(columns=columnas_calculadas, errors="ignore")

    df = calcular_resultados(df)
    df = ordenar_columnas(df)

    salida = CARPETA_RESULTADOS / NOMBRE_ARCHIVO_PROCESADO
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="DETALLE_ESTUDIANTES", index=False)

    print("\n" + "=" * 48)
    print("PROCESAMIENTO DE PERSONAL SOCIAL COMPLETADO")
    print("=" * 48)
    print(f"Estudiantes de primaria procesados: {len(df):,}")
    print("\nDistribución global:")
    print(df["NIVEL_LOGRO_GLOBAL"].value_counts().to_string())
    print(f"\nArchivo generado: {salida}")
    return df


if __name__ == "__main__":
    procesar_datos()
