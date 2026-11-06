import numpy as np
import pandas as pd

from config_comunicacion import (
    AREA, TOTAL_PREGUNTAS_ESPERADAS, COLUMNAS_OBLIGATORIAS,
    clasificar_nivel_logro,
)


def _normalizar_texto(serie):
    return serie.fillna("").astype(str).str.strip()


def _normalizar_porcentaje(valor):
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


def cargar_y_procesar(archivo, pestana="RESULTADOS"):
    """
    Lee la base exportada desde Google Sheets.
    Conserva todas las columnas originales y agrega campos calculados.
    No elimina registros ni modifica el archivo de origen.
    """
    try:
        df = pd.read_excel(archivo, sheet_name=pestana, engine="openpyxl")
    except ValueError as exc:
        raise ValueError(
            f"No se encontró la pestaña '{pestana}' en el archivo. "
            "Verifica PESTANA_ORIGEN en config_comunicacion.py."
        ) from exc

    df.columns = [str(c).strip().upper() for c in df.columns]

    faltantes = [c for c in COLUMNAS_OBLIGATORIAS if c not in df.columns]
    if faltantes:
        raise ValueError(
            "Faltan columnas obligatorias en el Excel: " + ", ".join(faltantes)
        )

    # Normaliza campos de texto sin descartar las demás columnas.
    for col in [
        "NIVEL", "INSTITUCION", "GRADO", "SECCION",
        "APELLIDOS", "NOMBRES", "COMP_1_NOMBRE"
    ]:
        if col in df.columns:
            df[col] = _normalizar_texto(df[col])

    df["NIVEL"] = df["NIVEL"].str.upper()
    df["INSTITUCION"] = df["INSTITUCION"].str.replace(r"\s+", " ", regex=True)
    df["GRADO"] = df["GRADO"].str.replace(r"\.0$", "", regex=True)

    for col in [
        "CORRECTAS", "INCORRECTAS", "TOTAL_PREGUNTAS",
        "COMP_1_ACIERTOS", "COMP_1_TOTAL"
    ]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Puntaje de logro específico de Comunicación.
    df["NIVEL_LOGRO"] = df["CORRECTAS"].apply(clasificar_nivel_logro)

    # Conserva el porcentaje recibido y genera uno consistente con los aciertos.
    if "PORCENTAJE" in df.columns:
        df["PORCENTAJE_FUENTE_NORMALIZADO"] = df["PORCENTAJE"].apply(
            _normalizar_porcentaje
        )
    else:
        df["PORCENTAJE_FUENTE_NORMALIZADO"] = np.nan

    df["PORCENTAJE_CALCULADO"] = (
        df["CORRECTAS"] / df["TOTAL_PREGUNTAS"].replace(0, np.nan)
    )
    df["PORCENTAJE_REPORTE"] = df["PORCENTAJE_CALCULADO"]

    df["ESTUDIANTE"] = (
        df.get("APELLIDOS", "").fillna("").astype(str).str.strip()
        + ", "
        + df.get("NOMBRES", "").fillna("").astype(str).str.strip()
    ).str.strip(", ")

    # Observaciones de calidad: se informa, pero no se borra ni altera la fila.
    df["OBSERVACION_DATOS"] = ""
    df.loc[df["NIVEL_LOGRO"].eq("REVISAR"), "OBSERVACION_DATOS"] = (
        "Puntaje ausente o fuera del rango 0-16"
    )
    df.loc[df["INSTITUCION"].eq(""), "OBSERVACION_DATOS"] += " | Institución vacía"
    df.loc[
        df["TOTAL_PREGUNTAS"].ne(TOTAL_PREGUNTAS_ESPERADAS),
        "OBSERVACION_DATOS"
    ] += " | Total de preguntas distinto de 16"

    if "ID_REGISTRO" in df.columns:
        duplicados = df["ID_REGISTRO"].duplicated(keep=False) & df["ID_REGISTRO"].notna()
        df.loc[duplicados, "OBSERVACION_DATOS"] += " | ID_REGISTRO duplicado"

    return df


def resumen_niveles(df, filtro_nivel="TODOS"):
    datos = df if filtro_nivel == "TODOS" else df[df["NIVEL"] == filtro_nivel]
    orden = ["INICIO", "PROCESO", "LOGRADO", "REVISAR"]
    conteo = datos["NIVEL_LOGRO"].value_counts()
    total = len(datos)
    return pd.DataFrame([
        {
            "NIVEL_LOGRO": nivel,
            "ESTUDIANTES": int(conteo.get(nivel, 0)),
            "PORCENTAJE": int(conteo.get(nivel, 0)) / total if total else 0,
        }
        for nivel in orden
    ])


def resumen_por_grado(df, filtro_nivel="TODOS"):
    datos = df if filtro_nivel == "TODOS" else df[df["NIVEL"] == filtro_nivel]
    columnas = [
        "NIVEL", "GRADO", "ESTUDIANTES",
        "INICIO", "PROCESO", "LOGRADO", "REVISAR",
        "PROMEDIO_ACIERTOS", "PROMEDIO_PORCENTAJE"
    ]
    if datos.empty:
        return pd.DataFrame(columns=columnas)

    tabla = pd.crosstab(
        [datos["NIVEL"], datos["GRADO"]], datos["NIVEL_LOGRO"]
    ).reset_index()

    for nivel in ["INICIO", "PROCESO", "LOGRADO", "REVISAR"]:
        if nivel not in tabla.columns:
            tabla[nivel] = 0

    promedios = (
        datos.groupby(["NIVEL", "GRADO"], dropna=False)
        .agg(
            ESTUDIANTES=("ID_REGISTRO", "size"),
            PROMEDIO_ACIERTOS=("CORRECTAS", "mean"),
            PROMEDIO_PORCENTAJE=("PORCENTAJE_REPORTE", "mean"),
        )
        .reset_index()
    )
    tabla = tabla.merge(promedios, on=["NIVEL", "GRADO"], how="left")
    return tabla[columnas].sort_values(["NIVEL", "GRADO"]).reset_index(drop=True)


def resumen_competencia(df):
    """Comunicación se resume usando los campos COMP_1_* existentes."""
    if df.empty:
        return pd.DataFrame(columns=[
            "NIVEL", "COMPETENCIA", "ESTUDIANTES",
            "ACIERTOS_ACUMULADOS", "PREGUNTAS_POSIBLES_ACUMULADAS",
            "PORCENTAJE_GLOBAL"
        ])

    datos = df.copy()
    if "COMP_1_NOMBRE" not in datos.columns:
        datos["COMP_1_NOMBRE"] = AREA

    datos["COMP_1_NOMBRE"] = datos["COMP_1_NOMBRE"].fillna("").astype(str).str.strip()
    datos.loc[datos["COMP_1_NOMBRE"].eq(""), "COMP_1_NOMBRE"] = AREA

    return (
        datos.groupby(["NIVEL", "COMP_1_NOMBRE"], dropna=False)
        .agg(
            ESTUDIANTES=("ID_REGISTRO", "size"),
            ACIERTOS_ACUMULADOS=("COMP_1_ACIERTOS", "sum"),
            PREGUNTAS_POSIBLES_ACUMULADAS=("COMP_1_TOTAL", "sum"),
        )
        .reset_index()
        .rename(columns={"COMP_1_NOMBRE": "COMPETENCIA"})
        .assign(
            PORCENTAJE_GLOBAL=lambda x:
                x["ACIERTOS_ACUMULADOS"] /
                x["PREGUNTAS_POSIBLES_ACUMULADAS"].replace(0, np.nan)
        )
    )
