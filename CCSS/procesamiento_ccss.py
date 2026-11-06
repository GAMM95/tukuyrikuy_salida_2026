import pandas as pd

from config_ccss import (
    ARCHIVO_ORIGEN, HOJA_ORIGEN, NIVEL_OBJETIVO, GRADOS_OBJETIVO,
    CARPETA_RESULTADOS, NOMBRE_ARCHIVO_PROCESADO, COMPETENCIAS,
    TOTAL_PREGUNTAS, COLUMNAS_IDENTIFICACION, COLUMNAS_PREGUNTAS,
    COLUMNAS_VALORACION, COLUMNAS_RESULTADOS, COLUMNAS_RESULTADO_GLOBAL,
    clasificar_nivel_competencia, clasificar_nivel_global,
    crear_carpetas, validar_configuracion,
)


def normalizar_columnas(df):
    df = df.copy()
    df.columns = [
        str(c).strip().upper().replace(" ", "_").replace("-", "_")
        for c in df.columns
    ]
    return df


def cargar_origen():
    if not ARCHIVO_ORIGEN.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {ARCHIVO_ORIGEN}")
    libro = pd.ExcelFile(ARCHIVO_ORIGEN)
    hoja = HOJA_ORIGEN if HOJA_ORIGEN in libro.sheet_names else libro.sheet_names[0]
    print(f"Procesando hoja: {hoja}")
    return normalizar_columnas(
        pd.read_excel(ARCHIVO_ORIGEN, sheet_name=hoja, dtype=object)
    )


def validar_columnas(df):
    requeridas = ["NIVEL", "GRADO"] + COLUMNAS_VALORACION
    faltantes = [c for c in requeridas if c not in df.columns]
    if faltantes:
        raise ValueError("Faltan columnas requeridas: " + ", ".join(faltantes))


def normalizar_identificacion(df):
    for c in COLUMNAS_IDENTIFICACION:
        if c not in df.columns:
            df[c] = ""
    for c in ["NIVEL", "INSTITUCION", "SECCION", "APELLIDOS", "NOMBRES", "SEXO"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    df["NIVEL"] = df["NIVEL"].str.upper()
    df["INSTITUCION"] = df["INSTITUCION"].replace("", "SIN DATO")
    df["SECCION"] = df["SECCION"].replace("", "SIN DATO")
    df["GRADO"] = pd.to_numeric(df["GRADO"], errors="coerce")
    df["DNI"] = df["DNI"].fillna("").astype(str).str.replace(r"\\.0$", "", regex=True)
    return df


def convertir_valoracion(serie, nombre):
    valores = pd.to_numeric(serie, errors="coerce")
    invalidos = valores.notna() & ~valores.isin([0, 1])
    if invalidos.any():
        muestra = serie[invalidos].astype(str).unique()[:5]
        raise ValueError(f"{nombre} contiene valores diferentes de 0/1: {list(muestra)}")
    return valores.fillna(0).astype(int)


def calcular_resultados(df):
    for c in COLUMNAS_VALORACION:
        df[c] = convertir_valoracion(df[c], c)

    df["CORRECTAS"] = df[COLUMNAS_VALORACION].sum(axis=1)
    df["TOTAL_PREGUNTAS"] = TOTAL_PREGUNTAS
    df["INCORRECTAS"] = TOTAL_PREGUNTAS - df["CORRECTAS"]
    df["PORCENTAJE"] = df["CORRECTAS"] / TOTAL_PREGUNTAS

    for comp in COMPETENCIAS:
        cod = comp["codigo"]
        cols = [f"V{i}" for i in comp["preguntas"]]
        aciertos = df[cols].sum(axis=1)
        total = len(cols)
        df[f"{cod}_NOMBRE"] = comp["nombre"]
        df[f"{cod}_PREGUNTAS"] = ", ".join(f"P{i}" for i in comp["preguntas"])
        df[f"{cod}_ACIERTOS"] = aciertos
        df[f"{cod}_TOTAL"] = total
        df[f"{cod}_PORCENTAJE"] = aciertos / total
        df[f"{cod}_NIVEL"] = aciertos.apply(
            lambda x: clasificar_nivel_competencia(x, comp)
        )
        df[f"{cod}_PESO"] = comp["peso"]
        df[f"{cod}_PUNTAJE_PONDERADO"] = (aciertos / total) * comp["peso"] * 100

    df["PUNTAJE_GLOBAL"] = sum(
        df[f"{c['codigo']}_PUNTAJE_PONDERADO"] for c in COMPETENCIAS
    )
    df["NIVEL_LOGRO_GLOBAL"] = df["PUNTAJE_GLOBAL"].apply(clasificar_nivel_global)
    return df


def ordenar_columnas(df):
    competencias = [
        f"{c['codigo']}_{campo}"
        for c in COMPETENCIAS
        for campo in ["NOMBRE", "PREGUNTAS", "ACIERTOS", "TOTAL", "PORCENTAJE",
                      "NIVEL", "PESO", "PUNTAJE_PONDERADO"]
    ]
    orden = (
        COLUMNAS_IDENTIFICACION + COLUMNAS_PREGUNTAS + COLUMNAS_VALORACION
        + COLUMNAS_RESULTADOS + competencias + COLUMNAS_RESULTADO_GLOBAL
    )
    existentes = [c for c in orden if c in df.columns]
    return df[existentes + [c for c in df.columns if c not in existentes]]


def procesar_datos():
    crear_carpetas()
    validar_configuracion()
    df = cargar_origen()
    validar_columnas(df)

    df = df[df["NIVEL"].astype(str).str.strip().str.upper() == NIVEL_OBJETIVO].copy()
    df = normalizar_identificacion(df)
    df = df[df["GRADO"].isin(GRADOS_OBJETIVO)].copy()
    if df.empty:
        raise ValueError("No hay registros de SECUNDARIA de 1.º a 5.º grado.")

    calculadas = [
        c for c in df.columns
        if c.startswith("COMP_") or c in COLUMNAS_RESULTADOS + COLUMNAS_RESULTADO_GLOBAL
    ]
    df = df.drop(columns=calculadas, errors="ignore")
    df = ordenar_columnas(calcular_resultados(df))

    salida = CARPETA_RESULTADOS / NOMBRE_ARCHIVO_PROCESADO
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="DETALLE_ESTUDIANTES", index=False)

    print("\\nPROCESAMIENTO DE CCSS COMPLETADO")
    print(f"Estudiantes de secundaria: {len(df):,}")
    print(df["NIVEL_LOGRO_GLOBAL"].value_counts().to_string())
    print(f"Archivo generado: {salida}")
    return df


if __name__ == "__main__":
    procesar_datos()
