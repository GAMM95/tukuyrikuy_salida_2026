from pathlib import Path

# 1. RUTAS DEL PROYECTO
CARPETA_PROYECTO = Path(__file__).resolve().parent
CARPETA_RESOURCES = CARPETA_PROYECTO / "resources"
CARPETA_RESULTADOS = CARPETA_PROYECTO / "resultados"
CARPETA_REPORTES_IE = (CARPETA_RESULTADOS / "REPORTES_INSTITUCIONES")

# 2. ARCHIVO EXCEL DE ORIGEN
ARCHIVO_ORIGEN = (CARPETA_RESOURCES / "RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx")

# Hojas del Excel de origen
HOJAS_ORIGEN = ["RESULTADOS"]

# 3. INFORMACIÓN GENERAL
AREA = "CIENCIAS SOCIALES"
NIVEL_OBJETIVO = "SECUNDARIA"
GRADOS_OBJETIVO = [1, 2, 3, 4, 5]

NIVELES = ["INICIO", "PROCESO", "LOGRADO"]

NOMBRE_ARCHIVO_GENERAL = "REPORTE_GENERAL_CCSS_SECUNDARIA.xlsx"
NOMBRE_ARCHIVO_PROCESADO = "BASE_PROCESADA_CCSS_SECUNDARIA.xlsx"
NOMBRE_ARCHIVO_ZIP = "REPORTES_INSTITUCIONES_CCSS_SECUNDARIA.zip"

TOTAL_PREGUNTAS = 16
TOTAL_COMPETENCIAS = 3

# 4. COMPETENCIAS Y DISTRIBUCIÓN DE PREGUNTAS
COMPETENCIAS = [
    {
        "numero": 1,
        "codigo": "COMP_1",
        "nombre": "Construye interpretaciones históricas",
        "corto": "Históricas",
        "preguntas": [1, 2, 3, 4, 5, 6],
        "peso": 1 / 3
    },
    {
        "numero": 2,
        "codigo": "COMP_2",
        "nombre": "Gestiona responsablemente el espacio y el ambiente",
        "corto": "Espacio",
        "preguntas": [7, 8, 9, 10, 11],
        "peso": 1 / 3
    },
    {
        "numero": 3,
        "codigo": "COMP_3",
        "nombre": "Gestiona responsablemente los recursos económicos",
        "corto": "Recursos económicos",
        "preguntas": [12, 13, 14, 15, 16],
        "peso": 1 / 3
    }
]

# 5. PESOS DE LAS COMPETENCIAS
PESOS_COMPETENCIAS = {
    "COMP_1": 1 / 3,
    "COMP_2": 1 / 3,
    "COMP_3": 1 / 3,
}

PESO_TOTAL = sum(PESOS_COMPETENCIAS.values())

# 6. NIVELES DE LOGRO POR COMPETENCIA
# Cada competencia tiene una cantidad diferente de preguntas.
#
# COMP_1: 6 preguntas
# INICIO  : 0, 1 o 2 aciertos
# PROCESO : 3 o 4 aciertos
# LOGRADO : 5 o 6 aciertos
#
# COMP_2: 5 preguntas
# INICIO  : 0, 1 o 2 aciertos
# PROCESO : 3 o 4 aciertos
# LOGRADO : 5 aciertos
#
# COMP_3: 5 preguntas
# INICIO  : 0, 1 o 2 aciertos
# PROCESO : 3 aciertos
# LOGRADO : 4 o 5 aciertos

NIVELES_LOGRO = {
    "COMP_1": {
        "INICIO": {
            "minimo": 0,
            "maximo": 2
        },
        "PROCESO": {
            "minimo": 3,
            "maximo": 4
        },
        "LOGRADO": {
            "minimo": 5,
            "maximo": 6
        },
        "REVISAR": {
            "minimo": None,
            "maximo": None
        }
    },
    "COMP_2": {
        "INICIO": {
            "minimo": 0,
            "maximo": 2
        },
        "PROCESO": {
            "minimo": 3,
            "maximo": 4
        },
        "LOGRADO": {
            "minimo": 5,
            "maximo": 5
        },
        "REVISAR": {
            "minimo": None,
            "maximo": None
        }
    },
    "COMP_3": {
        "INICIO": {
            "minimo": 0,
            "maximo": 2
        },
        "PROCESO": {
            "minimo": 3,
            "maximo": 3
        },
        "LOGRADO": {
            "minimo": 4,
            "maximo": 5
        },
        "REVISAR": {
            "minimo": None,
            "maximo": None
        }
    }
}

# 7. UMBRALES DEL RESULTADO GLOBAL
LIMITE_GLOBAL_INICIO = 50
LIMITE_GLOBAL_PROCESO = 75

# 8. COLORES DEL PROYECTO
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
}

# 9. COLUMNAS DE IDENTIFICACIÓN
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

# 10. COLUMNAS DE RESPUESTAS P1-P16
COLUMNAS_PREGUNTAS = [
    f"P{i}"
    for i in range(1, TOTAL_PREGUNTAS + 1)
]

# 11. COLUMNAS DE VALORACIÓN V1-V16
COLUMNAS_VALORACION = [
    f"V{i}"
    for i in range(1, TOTAL_PREGUNTAS + 1)
]

# 12. COLUMNAS DE RESULTADOS GENERALES
COLUMNAS_RESULTADOS = ["CORRECTAS", "INCORRECTAS", "TOTAL_PREGUNTAS", "PORCENTAJE", ]

# 13. COLUMNAS DERIVADAS POR COMPETENCIA
COLUMNAS_COMPETENCIAS = []

for competencia in COMPETENCIAS:
    codigo = competencia["codigo"]

    COLUMNAS_COMPETENCIAS.extend([
        f"{codigo}_NOMBRE",
        f"{codigo}_PREGUNTAS",
        f"{codigo}_ACIERTOS",
        f"{codigo}_TOTAL",
        f"{codigo}_PORCENTAJE",
        f"{codigo}_NIVEL",
        f"{codigo}_PESO",
        f"{codigo}_PUNTAJE_PONDERADO",
    ])

# 14. COLUMNAS DE RESULTADO GLOBAL
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

# 15. CLASIFICAR NIVEL DE UNA COMPETENCIA
def clasificar_nivel_competencia(aciertos, competencia):
    try:
        valor = float(aciertos)
    except (TypeError, ValueError):
        return "REVISAR"

    codigo = competencia["codigo"]
    total = len(competencia["preguntas"])

    if valor < 0 or valor > total or not valor.is_integer():
        return "REVISAR"

    limites = NIVELES_LOGRO[codigo]

    if limites["INICIO"]["minimo"] <= valor <= limites["INICIO"]["maximo"]:
        return "INICIO"

    if limites["PROCESO"]["minimo"] <= valor <= limites["PROCESO"]["maximo"]:
        return "PROCESO"

    if limites["LOGRADO"]["minimo"] <= valor <= limites["LOGRADO"]["maximo"]:
        return "LOGRADO"

    return "REVISAR"


# 16. CALCULAR PUNTAJE PONDERADO
def calcular_puntaje_ponderado(aciertos, competencia):
    try:
        valor = float(aciertos)
        peso = float(competencia["peso"])
    except (TypeError, ValueError):
        return None

    total = len(competencia["preguntas"])

    if valor < 0 or valor > total or not valor.is_integer():
        return None

    if peso < 0 or peso > 1:
        return None

    # Cada competencia aporta como máximo 33.33 puntos.
    return (valor / total) * (peso * 100)


# 17. CLASIFICAR RESULTADO GLOBAL
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


# 18. VALIDAR DISTRIBUCIÓN DE COMPETENCIAS
def validar_configuracion():
    if len(COMPETENCIAS) != 3:
        raise ValueError(
            "La configuración debe contener tres competencias."
        )

    if abs(PESO_TOTAL - 1.0) > 0.000001:
        raise ValueError(
            "La suma de los pesos debe ser exactamente 100 %."
        )

    preguntas = []

    for competencia in COMPETENCIAS:
        codigo = competencia["codigo"]
        total = len(competencia["preguntas"])

        preguntas.extend(competencia["preguntas"])

        if NIVELES_LOGRO[codigo]["LOGRADO"]["maximo"] != total:
            raise ValueError(
                "El máximo de " + codigo +
                " no coincide con sus preguntas."
            )

    if sorted(preguntas) != list(range(1, TOTAL_PREGUNTAS + 1)):
        raise ValueError(
            "La distribución debe cubrir P1 a P16 "
            "sin duplicaciones ni omisiones."
        )

    if len(preguntas) != len(set(preguntas)):
        raise ValueError(
            "Existen preguntas duplicadas entre competencias."
        )

    return True


# ============================================================
# 19. CREAR CARPETAS DEL PROYECTO
# ============================================================

def crear_carpetas():
    CARPETA_RESOURCES.mkdir(parents=True, exist_ok=True)
    CARPETA_RESULTADOS.mkdir(parents=True, exist_ok=True)
    CARPETA_REPORTES_IE.mkdir(parents=True, exist_ok=True)


# ============================================================
# 20. EJECUCIÓN DE VALIDACIÓN
# ============================================================

if __name__ == "__main__":
    crear_carpetas()
    validar_configuracion()

    print("==========================================")
    print("TUKUYRICUY 2026")
    print("CONFIGURACIÓN DE CIENCIAS SOCIALES")
    print("==========================================")

    print(f"Área: {AREA}")
    print(f"Nivel: {NIVEL_OBJETIVO}")
    print(f"Grados: {GRADOS_OBJETIVO}")
    print(f"Preguntas: {TOTAL_PREGUNTAS}")
    print(f"Competencias: {TOTAL_COMPETENCIAS}")
    print(f"Peso total: {PESO_TOTAL:.0%}")

    print("\nDISTRIBUCIÓN:")

    for competencia in COMPETENCIAS:
        print(
            f"{competencia['codigo']} | "
            f"{competencia['nombre']} | "
            f"{competencia['preguntas']} | "
            f"{competencia['peso']:.2%}"
        )

    print("\nRUTAS:")
    print(f"Excel: {ARCHIVO_ORIGEN}")
    print(f"Resultados: {CARPETA_RESULTADOS}")

    print("\nCONFIGURACIÓN VALIDADA CORRECTAMENTE.")
