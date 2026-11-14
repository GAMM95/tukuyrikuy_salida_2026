# -*- coding: utf-8 -*-
"""
============================================================
TUKUYRICUY 2026
CONFIGURACIÓN DE CIENCIA Y TECNOLOGÍA
============================================================

Módulo de configuración para el procesamiento y generación
de reportes de la Evaluación de Salida 2026.

Área: Ciencia y Tecnología
Niveles: Primaria y Secundaria
Preguntas: 16
Competencias: 3
============================================================
"""

from pathlib import Path
from typing import Any

# ============================================================
# 1. RUTAS DEL PROYECTO
# ============================================================

CARPETA_PROYECTO = Path(__file__).resolve().parent
CARPETA_RESOURCES = CARPETA_PROYECTO / "resources"
CARPETA_RESULTADOS = CARPETA_PROYECTO / "resultados"
CARPETA_REPORTES = CARPETA_RESULTADOS / "REPORTES_CYT"
CARPETA_REPORTES_IE = CARPETA_REPORTES / "REPORTES_POR_INSTITUCION"

# ============================================================
# 2. ARCHIVO EXCEL DE ORIGEN
# ============================================================

ARCHIVO_ORIGEN = CARPETA_RESOURCES / "RESULTADOS_CIENCIA_TECNOLOGIA.xlsx"
ARCHIVO_ENTRADA = ARCHIVO_ORIGEN
HOJAS_ORIGEN = ["RESULTADOS"]
HOJA_ENTRADA = "RESULTADOS"

# ============================================================
# 3. ARCHIVOS DE SALIDA
# ============================================================

ARCHIVO_BASE_PROCESADA = CARPETA_RESULTADOS / "BASE_PROCESADA_CYT.xlsx"
ARCHIVO_REPORTE_GENERAL = CARPETA_REPORTES / "REPORTE_GENERAL_CYT.xlsx"
ARCHIVO_ZIP = CARPETA_REPORTES / "REPORTES_CYT_POR_INSTITUCION.zip"

NOMBRE_ARCHIVO_GENERAL = "REPORTE_GENERAL_CYT.xlsx"
NOMBRE_ARCHIVO_PROCESADO = "BASE_PROCESADA_CYT.xlsx"
NOMBRE_ARCHIVO_ZIP = "REPORTES_CYT_POR_INSTITUCION.zip"

# ============================================================
# 4. INFORMACIÓN GENERAL
# ============================================================

AREA = "CIENCIA Y TECNOLOGÍA"
NIVELES = ["INICIO", "PROCESO", "LOGRADO"]
TOTAL_PREGUNTAS = 16
TOTAL_COMPETENCIAS = 3

# ============================================================
# 5. GRADOS Y NIVELES EDUCATIVOS
# ============================================================

GRADOS_POR_NIVEL = {
    "PRIMARIA": [2, 3, 4, 5, 6],
    "SECUNDARIA": [1, 2, 3, 4, 5],
}

NIVELES_EDUCATIVOS = list(GRADOS_POR_NIVEL.keys())

# ============================================================
# 6. COMPETENCIAS Y DISTRIBUCIÓN DE PREGUNTAS
# ============================================================

# Las tres competencias tienen el mismo peso.
PESO_COMPETENCIA = 1 / 3

COMPETENCIAS: list[dict[str, Any]] = [
    {
        "numero": 1,
        "codigo": "COMP_1",
        "nombre": (
            "Indaga mediante métodos científicos "
            "para construir sus conocimientos"
        ),
        "corto": "Indaga",
        "preguntas": [1, 2, 3, 4, 5, 6],
        "inicio_max": 2,
        "proceso_max": 4,
        "logrado_max": 6,
        "peso": PESO_COMPETENCIA,
    },
    {
        "numero": 2,
        "codigo": "COMP_2",
        "nombre": (
            "Explica el mundo físico basándose en "
            "conocimientos sobre los seres vivos, "
            "materia y energía, biodiversidad, "
            "Tierra y universo"
        ),
        "corto": "Explica",
        "preguntas": [7, 8, 9, 10, 11, 12],
        "inicio_max": 2,
        "proceso_max": 4,
        "logrado_max": 6,
        "peso": PESO_COMPETENCIA,
    },
    {
        "numero": 3,
        "codigo": "COMP_3",
        "nombre": (
            "Diseña y construye soluciones tecnológicas "
            "para resolver problemas de su entorno"
        ),
        "corto": "Diseña",
        "preguntas": [13, 14, 15, 16],
        "inicio_max": 2,
        "proceso_max": 3,
        "logrado_max": 4,
        "peso": PESO_COMPETENCIA,
    },
]

# ============================================================
# 7. DISTRIBUCIÓN ESPECIAL DE 1.° DE SECUNDARIA
# ============================================================

# Distribución establecida en el referente de UGEL.
COMPETENCIAS_SECUNDARIA_1: list[dict[str, Any]] = [
    {
        "numero": 1,
        "codigo": "COMP_1",
        "nombre": COMPETENCIAS[0]["nombre"],
        "corto": "Indaga",
        "preguntas": [1, 2, 3, 5, 7, 9],
        "inicio_max": 2,
        "proceso_max": 4,
        "logrado_max": 6,
        "peso": PESO_COMPETENCIA,
    },
    {
        "numero": 2,
        "codigo": "COMP_2",
        "nombre": COMPETENCIAS[1]["nombre"],
        "corto": "Explica",
        "preguntas": [4, 6, 8, 14, 15, 16],
        "inicio_max": 2,
        "proceso_max": 4,
        "logrado_max": 6,
        "peso": PESO_COMPETENCIA,
    },
    {
        "numero": 3,
        "codigo": "COMP_3",
        "nombre": COMPETENCIAS[2]["nombre"],
        "corto": "Diseña",
        "preguntas": [10, 11, 12, 13],
        "inicio_max": 2,
        "proceso_max": 3,
        "logrado_max": 4,
        "peso": PESO_COMPETENCIA,
    },
]


# ============================================================
# 8. OBTENER COMPETENCIAS SEGÚN NIVEL Y GRADO
# ============================================================

def obtener_competencias(nivel, grado):
    """
    Devuelve la distribución de competencias correspondiente
    al nivel y grado del estudiante.

    Primaria:
        Distribución general.

    Secundaria:
        1.° utiliza distribución especial.
        2.° a 5.° utilizan distribución general.

    La información se devuelve mediante una copia para evitar
    modificaciones accidentales de la configuración original.
    """

    nivel = str(nivel).strip().upper()

    try:
        grado = int(grado)
    except (TypeError, ValueError):
        raise ValueError(f"Grado no válido: {grado}")

    if nivel not in GRADOS_POR_NIVEL:
        raise ValueError(f"Nivel educativo no válido: {nivel}")

    if grado not in GRADOS_POR_NIVEL[nivel]:
        raise ValueError(f"El grado {grado} no corresponde al nivel {nivel}.")

    if nivel == "SECUNDARIA" and grado == 1:
        competencias = COMPETENCIAS_SECUNDARIA_1
    else:
        competencias = COMPETENCIAS

    return [
        {
            **competencia,
            "preguntas": list(competencia["preguntas"]),
        }
        for competencia in competencias
    ]


# ============================================================
# 9. PESOS DE LAS COMPETENCIAS
# ============================================================

PESOS_COMPETENCIAS: dict[str, float] = {
    "COMP_1": 1 / 3,
    "COMP_2": 1 / 3,
    "COMP_3": 1 / 3,
}

PESO_TOTAL = sum(PESOS_COMPETENCIAS.values())

# ============================================================
# 10. NIVELES DE LOGRO POR COMPETENCIA
# ============================================================

NIVELES_LOGRO = {
    "COMP_1": {
        "INICIO": {"minimo": 0, "maximo": 2},
        "PROCESO": {"minimo": 3, "maximo": 4},
        "LOGRADO": {"minimo": 5, "maximo": 6},
    },
    "COMP_2": {
        "INICIO": {"minimo": 0, "maximo": 2},
        "PROCESO": {"minimo": 3, "maximo": 4},
        "LOGRADO": {"minimo": 5, "maximo": 6},
    },
    "COMP_3": {
        "INICIO": {"minimo": 0, "maximo": 2},
        "PROCESO": {"minimo": 3, "maximo": 3},
        "LOGRADO": {"minimo": 4, "maximo": 4},
    },
}

# ============================================================
# 11. NIVELES DE LOGRO GLOBAL
# ============================================================

UMBRAL_GLOBAL_INICIO = 50
UMBRAL_GLOBAL_PROCESO = 75

NIVELES_GLOBAL = {
    "INICIO": {"minimo": 0, "maximo": 50},
    "PROCESO": {"minimo": 50, "maximo": 75},
    "LOGRADO": {"minimo": 75, "maximo": 100},
}

# ============================================================
# 12. COLORES DEL PROYECTO
# ============================================================

COLORES = {
    # Cabeceras
    "CABECERA": "404040",
    "BLANCO": "FFFFFF",
    "TEXTO": "262626",

    # Niveles de logro
    "INICIO": "F4CCCC",
    "PROCESO": "FFF2CC",
    "LOGRADO": "D9EAD3",

    # Colores de gráficos
    "GRAFICO_INICIO": "F4CCCC",
    "GRAFICO_PROCESO": "FFF2CC",
    "GRAFICO_LOGRADO": "D9EAD3",

    # Complementarios
    "GRIS_CLARO": "F2F2F2",
    "GRIS_MEDIO": "D9D9D9",
    "BORDE": "BFBFBF",
    "SECUNDARIO": "D9EAF7",
}

# ============================================================
# 13. COLUMNAS DE IDENTIFICACIÓN
# ============================================================

COLUMNAS_IDENTIFICACION = [
    "ID_REGISTRO",
    "FECHA_HORA",
    "NIVEL",
    "INSTITUCION",
    "GRADO",
    "SECCION",
    "APELLIDOS",
    "NOMBRES",
    "DNI",
    "SEXO",
]

# ============================================================
# 14. COLUMNAS DE RESPUESTAS P1-P16
# ============================================================

COLUMNAS_PREGUNTAS = [f"P{i}" for i in range(1, TOTAL_PREGUNTAS + 1)]
COLUMNAS_RESPUESTAS = COLUMNAS_PREGUNTAS

# ============================================================
# 15. COLUMNAS DE VALORACIÓN V1-V16
# ============================================================

COLUMNAS_VALORACION = [f"V{i}" for i in range(1, TOTAL_PREGUNTAS + 1)]
COLUMNAS_VALORACIONES = COLUMNAS_VALORACION

# ============================================================
# 16. COLUMNAS DE RESULTADOS GENERALES
# ============================================================

COLUMNAS_RESULTADOS = [
    "CORRECTAS",
    "INCORRECTAS",
    "TOTAL_PREGUNTAS",
    "PORCENTAJE",
]

# ============================================================
# 17. COLUMNAS DERIVADAS POR COMPETENCIA
# ============================================================

COLUMNAS_COMPETENCIAS = []

for _competencia in COMPETENCIAS:
    _codigo = _competencia["codigo"]

    COLUMNAS_COMPETENCIAS.extend([
        f"{_codigo}_NOMBRE",
        f"{_codigo}_PREGUNTAS",
        f"{_codigo}_ACIERTOS",
        f"{_codigo}_TOTAL",
        f"{_codigo}_PORCENTAJE",
        f"{_codigo}_NIVEL",
        f"{_codigo}_PESO",
        f"{_codigo}_PUNTAJE_PONDERADO",
    ])

# ============================================================
# 18. COLUMNAS DE RESULTADO GLOBAL
# ============================================================

COLUMNAS_RESULTADO_GLOBAL = [
    "PUNTAJE_GLOBAL_PONDERADO",
    "PORCENTAJE_GLOBAL",
    "NIVEL_LOGRO_GLOBAL",
]

# Porcentajes de control (fuente vs. calculado)
COLUMNAS_PORCENTAJES = [
    "PORCENTAJE_FUENTE_NORMALIZADO",
    "PORCENTAJE_CALCULADO",
]

# Columnas de DETALLE_ESTUDIANTES en los reportes por institución
COLUMNAS_DETALLE_INSTITUCION = (
    COLUMNAS_IDENTIFICACION
    + ["CORRECTAS", "INCORRECTAS"]
    + [f"{c['codigo']}_{campo}"
       for c in COMPETENCIAS
       for campo in ("ACIERTOS", "PORCENTAJE", "NIVEL")]
    + COLUMNAS_PREGUNTAS
    + COLUMNAS_VALORACION
    + COLUMNAS_PORCENTAJES
    + ["PUNTAJE_GLOBAL_PONDERADO", "NIVEL_LOGRO_GLOBAL"]
)


# ============================================================
# 19. CLASIFICAR NIVEL DE UNA COMPETENCIA
# ============================================================

def clasificar_nivel_competencia(aciertos, competencia):
    """Clasifica los aciertos según los umbrales de la competencia."""

    try:
        valor = float(aciertos)
    except (TypeError, ValueError):
        return "REVISAR"

    if valor < 0 or valor > competencia["logrado_max"]:
        return "REVISAR"
    if valor <= competencia["inicio_max"]:
        return "INICIO"
    if valor <= competencia["proceso_max"]:
        return "PROCESO"
    return "LOGRADO"


# ============================================================
# 20. CALCULAR PUNTAJE PONDERADO
# ============================================================

def calcular_puntaje_ponderado(aciertos, total_preguntas, peso=PESO_COMPETENCIA):
    """
    Calcula el aporte ponderado de una competencia.
    Cada competencia aporta como máximo 33.33 puntos.
    El cálculo considera la cantidad real de preguntas
    de cada competencia.
    """

    try:
        valor = float(aciertos)
        total = int(total_preguntas)
        peso = float(peso)
    except (TypeError, ValueError):
        return None

    if total <= 0:
        return None
    if valor < 0 or valor > total:
        return None
    if peso < 0 or peso > 1:
        return None

    return (valor / total) * (peso * 100)


# ============================================================
# 21. CALCULAR RESULTADO GLOBAL
# ============================================================

def calcular_resultado_global(puntajes_competencias):
    """
    Calcula el promedio ponderado de las tres competencias.
    Cada competencia tiene un peso de 1/3.
    Recibe los porcentajes de logro de las competencias.
    """

    if len(puntajes_competencias) != TOTAL_COMPETENCIAS:
        return None

    try:
        valores = [float(valor) for valor in puntajes_competencias]
    except (TypeError, ValueError):
        return None

    if any(valor < 0 or valor > 100 for valor in valores):
        return None

    return sum(
        valores[i] * COMPETENCIAS[i]["peso"]
        for i in range(TOTAL_COMPETENCIAS)
    )


# ============================================================
# 22. CLASIFICAR RESULTADO GLOBAL
# ============================================================

def clasificar_nivel_global(puntaje):
    """
    Inicio: menor o igual al 50 %.
    Proceso: mayor al 50 % y menor o igual al 75 %.
    Logrado: mayor al 75 %.
    """

    try:
        valor = float(puntaje)
    except (TypeError, ValueError):
        return "REVISAR"

    if valor < 0 or valor > 100:
        return "REVISAR"
    if valor <= UMBRAL_GLOBAL_INICIO:
        return "INICIO"
    if valor <= UMBRAL_GLOBAL_PROCESO:
        return "PROCESO"
    return "LOGRADO"


# ============================================================
# 23. VALIDAR DISTRIBUCIÓN DE COMPETENCIAS
# ============================================================

def validar_distribucion(competencias, descripcion):
    preguntas = []

    for competencia in competencias:
        codigo = competencia["codigo"]
        lista = competencia["preguntas"]

        if not lista:
            raise ValueError(f"{descripcion}: {codigo} no tiene preguntas.")

        if len(lista) != competencia["logrado_max"]:
            raise ValueError(
                f"{descripcion}: cantidad incorrecta de preguntas en {codigo}."
            )

        if not (
            0 <= competencia["inicio_max"]
            < competencia["proceso_max"]
            < competencia["logrado_max"]
        ):
            raise ValueError(f"{descripcion}: umbrales incorrectos en {codigo}.")

        preguntas.extend(lista)

    if sorted(preguntas) != list(range(1, TOTAL_PREGUNTAS + 1)):
        raise ValueError(
            f"{descripcion}: la distribución debe cubrir "
            "P1 a P16 sin duplicaciones ni omisiones."
        )


# ============================================================
# 24. VALIDAR CONFIGURACIÓN GENERAL
# ============================================================

def validar_configuracion():
    # Validar cantidad de competencias
    if len(COMPETENCIAS) != TOTAL_COMPETENCIAS:
        raise ValueError("CYT debe contener tres competencias.")

    # Validar ponderaciones
    if abs(PESO_TOTAL - 1.0) > 0.000001:
        raise ValueError("La suma de los pesos debe ser exactamente 100 %.")

    # Validar distribución general
    validar_distribucion(COMPETENCIAS, "Distribución general")

    # Validar distribución especial de 1.° secundaria
    validar_distribucion(
        COMPETENCIAS_SECUNDARIA_1,
        "Distribución especial de 1.° secundaria"
    )

    # Validar claves de ponderación
    codigos = [str(c["codigo"]) for c in COMPETENCIAS]

    if sorted(PESOS_COMPETENCIAS.keys()) != sorted(codigos):
        raise ValueError(
            "Los códigos de las ponderaciones no coinciden "
            "con las competencias."
        )

    # Validar pesos individuales
    for competencia in COMPETENCIAS:
        codigo = str(competencia["codigo"])

        if abs(float(competencia["peso"]) - PESOS_COMPETENCIAS[codigo]) > 0.000001:
            raise ValueError(f"El peso de {codigo} no coincide.")

    # Validar grados
    for nivel, grados in GRADOS_POR_NIVEL.items():
        if not grados:
            raise ValueError(f"No hay grados configurados para {nivel}.")

    # Validar umbrales globales
    if not (0 <= UMBRAL_GLOBAL_INICIO < UMBRAL_GLOBAL_PROCESO <= 100):
        raise ValueError("Los umbrales globales son incorrectos.")

    # Validar archivo de entrada
    if not ARCHIVO_ENTRADA.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo de entrada: "
            f"{ARCHIVO_ENTRADA}\n"
            "Coloca RESULTADOS_CIENCIA_TECNOLOGIA.xlsx "
            "dentro de CYT/resources."
        )

    return True


# ============================================================
# 25. CREAR CARPETAS DEL PROYECTO
# ============================================================

def crear_carpetas():
    CARPETA_RESOURCES.mkdir(parents=True, exist_ok=True)
    CARPETA_RESULTADOS.mkdir(parents=True, exist_ok=True)
    CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)
    CARPETA_REPORTES_IE.mkdir(parents=True, exist_ok=True)


# ============================================================
# 26. EJECUCIÓN DE VALIDACIÓN
# ============================================================

if __name__ == "__main__":
    crear_carpetas()
    validar_configuracion()

    print("=" * 64)
    print("TUKUYRICUY 2026")
    print("CONFIGURACIÓN DE CIENCIA Y TECNOLOGÍA")
    print("=" * 64)
    print(f"Área: {AREA}")
    print(f"Preguntas: {TOTAL_PREGUNTAS}")
    print(f"Competencias: {TOTAL_COMPETENCIAS}")
    print(f"Peso total: {PESO_TOTAL:.0%}")

    print("\nDISTRIBUCIÓN GENERAL:")
    for competencia in COMPETENCIAS:
        print(
            f"{competencia['codigo']} | "
            f"{competencia['nombre']} | "
            f"{competencia['preguntas']} | "
            f"{competencia['peso']:.2%}"
        )

    print("\nDISTRIBUCIÓN ESPECIAL - 1.° DE SECUNDARIA:")
    for competencia in COMPETENCIAS_SECUNDARIA_1:
        print(
            f"{competencia['codigo']} | "
            f"{competencia['corto']} | "
            f"{competencia['preguntas']} | "
            f"{competencia['peso']:.2%}"
        )

    print("\nRUTAS:")
    print(f"Excel: {ARCHIVO_ENTRADA}")
    print(f"Base procesada: {ARCHIVO_BASE_PROCESADA}")
    print(f"Reportes: {CARPETA_REPORTES}")
    print(f"Reportes IE: {CARPETA_REPORTES_IE}")
    print(f"ZIP: {ARCHIVO_ZIP}")
    print("\nCONFIGURACIÓN VALIDADA CORRECTAMENTE.")