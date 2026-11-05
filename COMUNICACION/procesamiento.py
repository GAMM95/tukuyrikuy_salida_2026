import pandas as pd
import numpy as np

from config_comunicacion import (
    AREA, TOTAL_PREGUNTAS_ESPERADAS, COLUMNAS_OBLIGATORIAS,
    COLUMNAS_IDENTIFICACION, COLUMNAS_PREGUNTAS, COLUMNAS_VALIDACION,
    clasificar_nivel_logro,
)


def _normalizar_texto(serie):
    return serie.fillna("").astype(str).str.strip()


def cargar_y_procesar(archivo, pestana="RESULTADOS"):
    """Lee el Excel fuente, valida su estructura y añade NIVEL_LOGRO."""
    try:
        df = pd.read_excel(archivo, sheet_name=pestana, engine="openpyxl")
    except ValueError as exc:
        raise ValueError(
            f"No se encontró la pestaña '{pestana}' en el archivo. "
            "Verifica el nombre de la hoja."
        ) from exc

    df.columns = [str(c).strip().upper() for c in df.columns]

    faltantes = [c for c in COLUMNAS_OBLIGATORIAS if c not in df.columns]
    if faltantes:
        raise ValueError(
            "Faltan columnas obligatorias en el Excel: " + ", ".join(faltantes)
        )

    # Limpieza de campos categóricos sin alterar el archivo original.
    for col in ["NIVEL", "INSTITUCION", "SECCION", "APELLIDOS", "NOMBRES",
                "COMP_1_NOMBRE"]:
        if col in df.columns:
            df[col] = _normalizar_texto(df[col])

    df["NIVEL"] = df["NIVEL"].str.upper()
    df["INSTITUCION"] = df["INSTITUCION"].str.replace(r"\s+", " ", regex=True)

    # Convierte a numérico y verifica el rango del puntaje.
    for col in ["CORRECTAS", "INCORRECTAS", "TOTAL_PREGUNTAS",
                "COMP_1_ACIERTOS", "COMP_1_TOTAL"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df["NIVEL_LOGRO"] = df["CORRECTAS"].apply(clasificar_nivel_logro)

    # El porcentaje fuente puede venir como 0.75 o como 75%.
    def normalizar_porcentaje(valor):
        if pd.isna(valor):
            return np.nan
        if isinstance(valor, str):
            valor = valor.strip().replace("%", "").replace(",", ".")
        try:
            numero = float(valor)
        except (TypeError, ValueError):
            return np.nan
        return numero / 100 if numero > 1 else numero

    df["PORCENTAJE_CALCULADO"] = (
        df["CORRECTAS"] / df["TOTAL_PREGUNTAS"].replace(0, np.nan)
    )
    df["PORCENTAJE_FUENTE_NORMALIZADO"] = df["PORCENTAJE"].apply(
        normalizar_porcentaje
    )

    # Se usa el porcentaje calculado desde CORRECTAS/TOTAL_PREGUNTAS
    # para que el reporte sea consistente con la clasificación.
    df["PORCENTAJE_REPORTE"] = df["PORCENTAJE_CALCULADO"]

    # Campos derivados útiles para tablas y filtros.
    df["GRADO"] = df["GRADO"].fillna("").astype(str).str.replace(r"\.0$", "", regex=True)
    df["ESTUDIANTE"] = (
        df["APELLIDOS"].fillna("").astype(str).str.strip()
        + ", "
        + df["NOMBRES"].fillna("").astype(str).str.strip()
    ).str.strip(", ")

    # Controles de calidad. No se eliminan registros automáticamente.
    df["OBSERVACION_DATOS"] = ""
    df.loc[df["NIVEL_LOGRO"] == "REVISAR", "OBSERVACION_DATOS"] = (
        "Puntaje ausente o fuera del rango 0-16"
    )
    df.loc[df["INSTITUCION"].eq(""), "OBSERVACION_DATOS"] += (
        " | Institución vacía"
    )
    df.loc[df["TOTAL_PREGUNTAS"].ne(TOTAL_PREGUNTAS_ESPERADAS),
           "OBSERVACION_DATOS"] += " | Total de preguntas distinto de 16"

    # Verifica posibles IDs repetidos, sin borrar filas.
    if "ID_REGISTRO" in df.columns:
        duplicados = df["ID_REGISTRO"].duplicated(keep=False) & df["ID_REGISTRO"].notna()
        df.loc[duplicados, "OBSERVACION_DATOS"] += " | ID_REGISTRO duplicado"

    return df


def resumen_niveles(df, filtro_nivel="TODOS"):
    datos = df if filtro_nivel == "TODOS" else df[df["NIVEL"] == filtro_nivel]
    orden = ["INICIO", "PROCESO", "LOGRADO", "REVISAR"]
    conteo = datos["NIVEL_LOGRO"].value_counts()
    total = len(datos)
    filas = []
    for nivel in orden:
        cantidad = int(conteo.get(nivel, 0))
        filas.append({
            "NIVEL_LOGRO": nivel,
            "ESTUDIANTES": cantidad,
            "PORCENTAJE": cantidad / total if total else 0,
        })
    return pd.DataFrame(filas)


def resumen_por_grado(df, filtro_nivel="TODOS"):
    datos = df if filtro_nivel == "TODOS" else df[df["NIVEL"] == filtro_nivel]
    if datos.empty:
        return pd.DataFrame(columns=["NIVEL", "GRADO", "ESTUDIANTES",
                                     "INICIO", "PROCESO", "LOGRADO"])
    tabla = pd.crosstab(
        [datos["NIVEL"], datos["GRADO"]],
        datos["NIVEL_LOGRO"]
    ).reset_index()
    for nivel in ["INICIO", "PROCESO", "LOGRADO"]:
        if nivel not in tabla.columns:
            tabla[nivel] = 0
    tabla["ESTUDIANTES"] = tabla[["INICIO", "PROCESO", "LOGRADO"]].sum(axis=1)
    return tabla[["NIVEL", "GRADO", "ESTUDIANTES",
                  "INICIO", "PROCESO", "LOGRADO"]]


def resumen_competencia(df):
    nombre = (
        df["COMP_1_NOMBRE"].dropna().astype(str).str.strip()
    )
    nombre = nombre[nombre.ne("")]
    competencia = nombre.mode().iloc[0] if not nombre.empty else (
        "Lee diversos tipos de textos escritos en su lengua materna"
    )
    aciertos = pd.to_numeric(df["COMP_1_ACIERTOS"], errors="coerce")
    total = pd.to_numeric(df["COMP_1_TOTAL"], errors="coerce")
    return pd.DataFrame([{
        "COMPETENCIA": competencia,
        "ESTUDIANTES": len(df),
        "ACIERTOS_ACUMULADOS": int(aciertos.fillna(0).sum()),
        "PREGUNTAS_POSIBLES_ACUMULADAS": int(total.fillna(0).sum()),
        "PORCENTAJE_GLOBAL": (
            aciertos.sum() / total.sum() if total.sum() else 0
        ),
    }])
