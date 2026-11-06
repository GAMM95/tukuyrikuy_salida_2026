from pathlib import Path

# ============================================================
# TUKUYRICUY 2026 - PERSONAL SOCIAL (PRIMARIA)
# ============================================================

CARPETA_PROYECTO = Path(__file__).resolve().parent
CARPETA_RESOURCES = CARPETA_PROYECTO / "resources"
CARPETA_RESULTADOS = CARPETA_PROYECTO / "resultados"
CARPETA_REPORTES_IE = CARPETA_RESULTADOS / "REPORTES_INSTITUCIONES"

ARCHIVO_ORIGEN = CARPETA_RESOURCES / "RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx"
HOJA_ORIGEN = "RESULTADOS"  # Si no existe, se utilizará la primera hoja del libro.

AREA = "PERSONAL SOCIAL"
NIVEL_OBJETIVO = "PRIMARIA"
GRADOS_OBJETIVO = [2, 3, 4, 5, 6]

NOMBRE_ARCHIVO_PROCESADO = "BASE_PROCESADA_PERSONAL_SOCIAL_PRIMARIA.xlsx"
NOMBRE_ARCHIVO_GENERAL = "REPORTE_GENERAL_PERSONAL_SOCIAL_PRIMARIA.xlsx"
NOMBRE_ARCHIVO_ZIP = "REPORTES_INSTITUCIONES_PERSONAL_SOCIAL_PRIMARIA.zip"

TOTAL_PREGUNTAS = 16
NIVELES = ["INICIO", "PROCESO", "LOGRADO"]

# Las preguntas siguen la distribución que compartiste.
# Las cinco competencias tienen el mismo peso: 20% cada una.
COMPETENCIAS = [
    {
        "numero": 1,
        "codigo": "COMP_1",
        "nombre": "Construye su identidad",
        "corto": "Identidad",
        "preguntas": [1, 2, 3],
        "peso": 0.20,
        "limites": {"INICIO": 1, "PROCESO": 2, "LOGRADO": 3},
    },
    {
        "numero": 2,
        "codigo": "COMP_2",
        "nombre": "Convive y participa democráticamente en la búsqueda del bien común",
        "corto": "Convive",
        "preguntas": [4, 5, 6],
        "peso": 0.20,
        "limites": {"INICIO": 1, "PROCESO": 2, "LOGRADO": 3},
    },
    {
        "numero": 3,
        "codigo": "COMP_3",
        "nombre": "Construye interpretaciones históricas",
        "corto": "Históricas",
        "preguntas": [7, 8, 9, 10],
        "peso": 0.20,
        "limites": {"INICIO": 2, "PROCESO": 3, "LOGRADO": 4},
    },
    {
        "numero": 4,
        "codigo": "COMP_4",
        "nombre": "Gestiona responsablemente el espacio y el ambiente",
        "corto": "Espacio",
        "preguntas": [11, 12, 13],
        "peso": 0.20,
        "limites": {"INICIO": 1, "PROCESO": 2, "LOGRADO": 3},
    },
    {
        "numero": 5,
        "codigo": "COMP_5",
        "nombre": "Gestiona responsablemente los recursos económicos",
        "corto": "Recursos económicos",
        "preguntas": [14, 15, 16],
        "peso": 0.20,
        "limites": {"INICIO": 1, "PROCESO": 2, "LOGRADO": 3},
    },
]

PREGUNTAS_POR_COMPETENCIA = {
    c["numero"]: c["preguntas"] for c in COMPETENCIAS
}
PESO_TOTAL = sum(c["peso"] for c in COMPETENCIAS)

# Criterios globales provisionales, configurables.
# Se aplican al puntaje ponderado de 0 a 100.
LIMITE_GLOBAL_INICIO = 50
LIMITE_GLOBAL_PROCESO = 75

COLORES = {
    "CABECERA": "404040",
    "BLANCO": "FFFFFF",
    "TEXTO": "262626",
    "INICIO": "F4CCCC",
    "PROCESO": "FFF2CC",
    "LOGRADO": "D9EAD3",
    "GRIS_CLARO": "F2F2F2",
    "GRIS_MEDIO": "D9D9D9",
    "BORDE": "BFBFBF",
}

COLUMNAS_IDENTIFICACION = [
    "ID_REGISTRO", "FECHA_HORA", "NIVEL", "INSTITUCION", "GRADO",
    "SECCION", "APELLIDOS", "NOMBRES", "DNI", "SEXO"
]
COLUMNAS_PREGUNTAS = [f"P{i}" for i in range(1, TOTAL_PREGUNTAS + 1)]
COLUMNAS_VALORACION = [f"V{i}" for i in range(1, TOTAL_PREGUNTAS + 1)]
COLUMNAS_RESULTADOS = [
    "CORRECTAS", "INCORRECTAS", "TOTAL_PREGUNTAS", "PORCENTAJE"
]
COLUMNAS_COMPETENCIAS = []
for competencia in COMPETENCIAS:
    codigo = competencia["codigo"]
    for campo in [
        "NOMBRE", "PREGUNTAS", "ACIERTOS", "TOTAL",
        "PORCENTAJE", "NIVEL", "PESO", "PUNTAJE_PONDERADO"
    ]:
        COLUMNAS_COMPETENCIAS.append(f"{codigo}_{campo}")

COLUMNAS_RESULTADO_GLOBAL = ["PUNTAJE_GLOBAL", "NIVEL_LOGRO_GLOBAL"]


def clasificar_nivel_competencia(aciertos, competencia):
    try:
        valor = float(aciertos)
    except (TypeError, ValueError):
        return "REVISAR"

    limites = competencia["limites"]
    maximo = len(competencia["preguntas"])
    if valor < 0 or valor > maximo:
        return "REVISAR"
    if valor <= limites["INICIO"]:
        return "INICIO"
    if valor <= limites["PROCESO"]:
        return "PROCESO"
    return "LOGRADO"


def clasificar_nivel_global(puntaje):
    try:
        valor = float(puntaje)
    except (TypeError, ValueError):
        return "REVISAR"
    if valor < 0 or valor > 100:
        return "REVISAR"
    if valor <= LIMITE_GLOBAL_INICIO:
        return "INICIO"
    if valor <= LIMITE_GLOBAL_PROCESO:
        return "PROCESO"
    return "LOGRADO"


def validar_configuracion():
    if len(COMPETENCIAS) != 5:
        raise ValueError("Personal Social debe tener cinco competencias.")
    if abs(PESO_TOTAL - 1.0) > 0.000001:
        raise ValueError("La suma de los pesos debe ser 100%.")
    preguntas = []
    for competencia in COMPETENCIAS:
        preguntas.extend(competencia["preguntas"])
        if len(competencia["preguntas"]) != competencia["limites"]["LOGRADO"]:
            raise ValueError(
                f"El máximo de {competencia['codigo']} no coincide con sus preguntas."
            )
    if sorted(preguntas) != list(range(1, TOTAL_PREGUNTAS + 1)):
        raise ValueError("La distribución debe cubrir P1 a P16 sin repetir.")
    return True


def crear_carpetas():
    CARPETA_RESOURCES.mkdir(parents=True, exist_ok=True)
    CARPETA_RESULTADOS.mkdir(parents=True, exist_ok=True)
    CARPETA_REPORTES_IE.mkdir(parents=True, exist_ok=True)


if __name__ == "__main__":
    crear_carpetas()
    validar_configuracion()
    print("CONFIGURACIÓN DE PERSONAL SOCIAL VALIDADA.")
