# -*- coding: utf-8 -*-
"""
TUKUYRICUY 2026 - CONFIGURACIÓN DE CIENCIA Y TECNOLOGÍA
"""
from pathlib import Path

CARPETA_PROYECTO = Path(__file__).resolve().parent
CARPETA_RESOURCES = CARPETA_PROYECTO / "resources"
CARPETA_RESULTADOS = CARPETA_PROYECTO / "resultados"

ARCHIVO_ENTRADA = CARPETA_RESOURCES / "RESULTADOS_CIENCIA_TECNOLOGIA.xlsx"
HOJA_ENTRADA = "RESULTADOS"
ARCHIVO_BASE_PROCESADA = CARPETA_RESULTADOS / "BASE_PROCESADA_CYT.xlsx"

AREA = "CIENCIA Y TECNOLOGÍA"
NOMBRE_CARPETA_REPORTES = "REPORTES_CYT"

# Las tres competencias tienen el mismo peso (1/3).
# Clasificación global provisional: INICIO <= 50%, PROCESO <= 75%, LOGRADO > 75%.
PESO_COMPETENCIA = 1 / 3
UMBRAL_GLOBAL_INICIO = 50
UMBRAL_GLOBAL_PROCESO = 75

NIVELES = {
    "PRIMARIA": {"grados": [2, 3, 4, 5, 6]},
    "SECUNDARIA": {"grados": [1, 2, 3, 4, 5]},
}

COMPETENCIAS = [
    {
        "id": "COMP_1",
        "nombre": "Indaga mediante métodos científicos",
        "corto": "INDAGA",
        "preguntas": list(range(1, 7)),
        "inicio_max": 2,
        "proceso_max": 4,
        "logrado_max": 6,
        "peso": PESO_COMPETENCIA,
    },
    {
        "id": "COMP_2",
        "nombre": "Explica el mundo físico",
        "corto": "EXPLICA",
        "preguntas": list(range(7, 13)),
        "inicio_max": 2,
        "proceso_max": 4,
        "logrado_max": 6,
        "peso": PESO_COMPETENCIA,
    },
    {
        "id": "COMP_3",
        "nombre": "Diseña y construye soluciones tecnológicas",
        "corto": "DISEÑA",
        "preguntas": list(range(13, 17)),
        "inicio_max": 2,
        "proceso_max": 3,
        "logrado_max": 4,
        "peso": PESO_COMPETENCIA,
    },
]

COLUMNAS_IDENTIFICACION = [
    "ID_REGISTRO", "FECHA_HORA", "NIVEL", "INSTITUCION", "GRADO",
    "SECCION", "APELLIDOS", "NOMBRES", "DNI", "SEXO"
]
COLUMNAS_RESPUESTAS = [f"P{i}" for i in range(1, 17)]
COLUMNAS_VALORACIONES = [f"V{i}" for i in range(1, 17)]

COLOR_CABECERA = "404040"
COLOR_INICIO = "F4CCCC"
COLOR_PROCESO = "FFF2CC"
COLOR_LOGRADO = "D9EAD3"
COLOR_SECUNDARIO = "D9EAF7"

def clasificar_competencia(aciertos, competencia):
    if aciertos <= competencia["inicio_max"]:
        return "INICIO"
    if aciertos <= competencia["proceso_max"]:
        return "PROCESO"
    return "LOGRADO"

def clasificar_global(porcentaje):
    if porcentaje <= UMBRAL_GLOBAL_INICIO:
        return "INICIO"
    if porcentaje <= UMBRAL_GLOBAL_PROCESO:
        return "PROCESO"
    return "LOGRADO"

def validar_configuracion():
    assert abs(sum(c["peso"] for c in COMPETENCIAS) - 1) < 0.000001
    for c in COMPETENCIAS:
        assert c["preguntas"]
        assert c["logrado_max"] == len(c["preguntas"])
        assert c["inicio_max"] < c["proceso_max"] < c["logrado_max"]
    if not ARCHIVO_ENTRADA.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de entrada: {ARCHIVO_ENTRADA}\\n"
            "Coloca RESULTADOS_CIENCIA_TECNOLOGIA.xlsx dentro de CYT/resources."
        )
    return True
