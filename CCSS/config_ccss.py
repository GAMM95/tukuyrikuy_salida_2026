from pathlib import Path

CARPETA_PROYECTO = Path(__file__).resolve().parent
CARPETA_RESOURCES = CARPETA_PROYECTO / "resources"
CARPETA_RESULTADOS = CARPETA_PROYECTO / "resultados"
CARPETA_REPORTES_IE = CARPETA_RESULTADOS / "REPORTES_INSTITUCIONES"

ARCHIVO_ORIGEN = CARPETA_RESOURCES / "RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx"
HOJA_ORIGEN = "RESULTADOS"  # Si no existe, se toma la primera hoja.

AREA = "CIENCIAS SOCIALES"
NIVEL_OBJETIVO = "SECUNDARIA"
GRADOS_OBJETIVO = [1, 2, 3, 4, 5]

NOMBRE_ARCHIVO_PROCESADO = "BASE_PROCESADA_CCSS_SECUNDARIA.xlsx"
NOMBRE_ARCHIVO_GENERAL = "REPORTE_GENERAL_CCSS_SECUNDARIA.xlsx"
NOMBRE_ARCHIVO_ZIP = "REPORTES_INSTITUCIONES_CCSS_SECUNDARIA.zip"

TOTAL_PREGUNTAS = 16
NIVELES = ["INICIO", "PROCESO", "LOGRADO"]

COMPETENCIAS = [
    {
        "numero": 1, "codigo": "COMP_1",
        "nombre": "Construye interpretaciones históricas",
        "corto": "Históricas", "preguntas": [1, 2, 3, 4, 5, 6],
        "peso": 1/3, "limites": {"INICIO": 2, "PROCESO": 4, "LOGRADO": 6},
    },
    {
        "numero": 2, "codigo": "COMP_2",
        "nombre": "Gestiona responsablemente el espacio y el ambiente",
        "corto": "Espacio", "preguntas": [7, 8, 9, 10, 11],
        "peso": 1/3,
        # La matriz compartida tiene 5 preguntas en esta competencia.
        "limites": {"INICIO": 2, "PROCESO": 4, "LOGRADO": 5},
    },
    {
        "numero": 3, "codigo": "COMP_3",
        "nombre": "Gestiona responsablemente los recursos económicos",
        "corto": "Recursos económicos", "preguntas": [12, 13, 14, 15, 16],
        "peso": 1/3, "limites": {"INICIO": 2, "PROCESO": 3, "LOGRADO": 5},
    },
]
PESO_TOTAL = sum(c["peso"] for c in COMPETENCIAS)

# Umbrales globales provisionales; se pueden cambiar aquí.
LIMITE_GLOBAL_INICIO = 50
LIMITE_GLOBAL_PROCESO = 75

COLORES = {
    "CABECERA": "404040", "BLANCO": "FFFFFF", "TEXTO": "262626",
    "INICIO": "F4CCCC", "PROCESO": "FFF2CC", "LOGRADO": "D9EAD3",
    "GRIS_CLARO": "F2F2F2", "GRIS_MEDIO": "D9D9D9", "BORDE": "BFBFBF",
}

COLUMNAS_IDENTIFICACION = [
    "ID_REGISTRO", "FECHA_HORA", "NIVEL", "INSTITUCION", "GRADO",
    "SECCION", "APELLIDOS", "NOMBRES", "DNI", "SEXO"
]
COLUMNAS_PREGUNTAS = [f"P{i}" for i in range(1, 17)]
COLUMNAS_VALORACION = [f"V{i}" for i in range(1, 17)]
COLUMNAS_RESULTADOS = ["CORRECTAS", "INCORRECTAS", "TOTAL_PREGUNTAS", "PORCENTAJE"]
COLUMNAS_COMPETENCIAS = [
    f"{c['codigo']}_{campo}"
    for c in COMPETENCIAS
    for campo in ["NOMBRE", "PREGUNTAS", "ACIERTOS", "TOTAL", "PORCENTAJE",
                  "NIVEL", "PESO", "PUNTAJE_PONDERADO"]
]
COLUMNAS_RESULTADO_GLOBAL = ["PUNTAJE_GLOBAL", "NIVEL_LOGRO_GLOBAL"]


def clasificar_nivel_competencia(aciertos, competencia):
    try:
        valor = float(aciertos)
    except (TypeError, ValueError):
        return "REVISAR"
    if valor < 0 or valor > len(competencia["preguntas"]):
        return "REVISAR"
    if valor <= competencia["limites"]["INICIO"]:
        return "INICIO"
    if valor <= competencia["limites"]["PROCESO"]:
        return "PROCESO"
    return "LOGRADO"


def clasificar_nivel_global(puntaje):
    try:
        valor = float(puntaje)
    except (TypeError, ValueError):
        return "REVISAR"
    if not 0 <= valor <= 100:
        return "REVISAR"
    if valor <= LIMITE_GLOBAL_INICIO:
        return "INICIO"
    if valor <= LIMITE_GLOBAL_PROCESO:
        return "PROCESO"
    return "LOGRADO"


def validar_configuracion():
    if len(COMPETENCIAS) != 3 or abs(PESO_TOTAL - 1) > 0.000001:
        raise ValueError("CCSS requiere tres competencias con peso total de 100%.")
    preguntas = [p for c in COMPETENCIAS for p in c["preguntas"]]
    if sorted(preguntas) != list(range(1, 17)):
        raise ValueError("La distribución debe cubrir P1 a P16 sin repetir.")
    for c in COMPETENCIAS:
        if c["limites"]["LOGRADO"] != len(c["preguntas"]):
            raise ValueError(f"El máximo de {c['codigo']} no coincide con sus preguntas.")
    return True


def crear_carpetas():
    for carpeta in [CARPETA_RESOURCES, CARPETA_RESULTADOS, CARPETA_REPORTES_IE]:
        carpeta.mkdir(parents=True, exist_ok=True)
