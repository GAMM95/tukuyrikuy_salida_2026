from pathlib import Path

# ============================================================
# TUKUYRICUY 2026 - COMUNICACIÓN
# Configuración general
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
CARPETA_DATOS = BASE_DIR / "resources"
CARPETA_RESULTADOS = BASE_DIR / "resultados"

ARCHIVO_ENTRADA = CARPETA_DATOS / "RESULTADOS_COMUNICACION.xlsx"
PESTANA_ORIGEN = "RESULTADOS"

AREA = "COMUNICACIÓN"
TOTAL_PREGUNTAS_ESPERADAS = 16

# Comunicación: Inicio 0-6; Proceso 7-11; Logrado 12-16.
def clasificar_nivel_logro(correctas):
    if correctas is None:
        return "REVISAR"
    try:
        puntaje = float(correctas)
    except (TypeError, ValueError):
        return "REVISAR"

    if puntaje < 0 or puntaje > TOTAL_PREGUNTAS_ESPERADAS:
        return "REVISAR"
    if puntaje <= 6:
        return "INICIO"
    if puntaje <= 11:
        return "PROCESO"
    return "LOGRADO"


COLUMNAS_OBLIGATORIAS = [
    "ID_REGISTRO", "NIVEL", "INSTITUCION", "GRADO", "SECCION",
    "APELLIDOS", "NOMBRES", "CORRECTAS", "INCORRECTAS",
    "TOTAL_PREGUNTAS", "PORCENTAJE",
    "COMP_1_NOMBRE", "COMP_1_ACIERTOS", "COMP_1_TOTAL",
    "COMP_1_PORCENTAJE",
]

COLUMNAS_IDENTIFICACION = [
    "ID_REGISTRO", "FECHA_HORA", "NIVEL", "INSTITUCION", "GRADO",
    "SECCION", "APELLIDOS", "NOMBRES", "DNI", "SEXO",
]

COLUMNAS_PREGUNTAS = [f"P{i}" for i in range(1, 17)]
COLUMNAS_VALIDACION = [f"V{i}" for i in range(1, 17)]
