# ============================================================
# TUKUYRICUY 2026
# CONFIGURACIÓN DE PERSONAL SOCIAL - PRIMARIA
# ============================================================

from pathlib import Path

# ============================================================
# 1. RUTAS DEL PROYECTO
# ============================================================
CARPETA_PROYECTO = Path(__file__).resolve().parent
CARPETA_RESOURCES = CARPETA_PROYECTO / "resources"
CARPETA_RESULTADOS = CARPETA_PROYECTO / "resultados"
CARPETA_REPORTES_IE = (CARPETA_RESULTADOS / "REPORTES_INSTITUCIONES")

# ============================================================
# 2. ARCHIVO EXCEL DE ORIGEN
# ============================================================

ARCHIVO_ORIGEN = (CARPETA_RESOURCES / "RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx")

ARCHIVO_ENTRADA = ARCHIVO_ORIGEN

# Hoja preferida del Excel de origen.
# Si no existe, el procesamiento podrá utilizar
# la primera hoja disponible.

HOJAS_ORIGEN = ["RESULTADOS"]

# ============================================================
# 3. INFORMACIÓN GENERAL
# ============================================================

AREA = "PERSONAL SOCIAL"
NIVEL_OBJETIVO = "PRIMARIA"
GRADOS_OBJETIVO = [2, 3, 4, 5, 6]
NIVELES = ["INICIO", "PROCESO", "LOGRADO"]
NOMBRE_ARCHIVO_GENERAL = ("REPORTE_GENERAL_PERSONAL_SOCIAL_PRIMARIA.xlsx")
NOMBRE_ARCHIVO_PROCESADO = ("BASE_PROCESADA_PERSONAL_SOCIAL_PRIMARIA.xlsx")
NOMBRE_ARCHIVO_ZIP = ("REPORTES_INSTITUCIONES_PERSONAL_SOCIAL_PRIMARIA.zip")

TOTAL_PREGUNTAS = 16

TOTAL_COMPETENCIAS = 5

# ============================================================
# 4. COMPETENCIAS Y DISTRIBUCIÓN DE PREGUNTAS
# ============================================================

COMPETENCIAS = [

    {
        "numero": 1,
        "codigo": "COMP_1",
        "nombre": "Construye su identidad",
        "corto": "Identidad",
        "preguntas": [1, 2, 3],
        "peso": 0.20
    },

    {
        "numero": 2,
        "codigo": "COMP_2",
        "nombre": (
            "Convive y participa democráticamente "
            "en la búsqueda del bien común"
        ),
        "corto": "Convive",
        "preguntas": [4, 5, 6],
        "peso": 0.20
    },

    {
        "numero": 3,
        "codigo": "COMP_3",
        "nombre": "Construye interpretaciones históricas",
        "corto": "Históricas",
        "preguntas": [7, 8, 9, 10],
        "peso": 0.20
    },

    {
        "numero": 4,
        "codigo": "COMP_4",
        "nombre": (
            "Gestiona responsablemente "
            "el espacio y el ambiente"
        ),
        "corto": "Espacio",
        "preguntas": [11, 12, 13],
        "peso": 0.20
    },

    {
        "numero": 5,
        "codigo": "COMP_5",
        "nombre": (
            "Gestiona responsablemente "
            "los recursos económicos"
        ),
        "corto": "Recursos económicos",
        "preguntas": [14, 15, 16],
        "peso": 0.20
    }

]

# ============================================================
# 5. PESOS DE LAS COMPETENCIAS
# ============================================================

PESOS_COMPETENCIAS = {
    "COMP_1": 0.20,
    "COMP_2": 0.20,
    "COMP_3": 0.20,
    "COMP_4": 0.20,
    "COMP_5": 0.20,
}

PESO_TOTAL = sum(PESOS_COMPETENCIAS.values())

# ============================================================
# 6. NIVELES DE LOGRO POR COMPETENCIA
# ============================================================

# COMP_1: 3 preguntas
# INICIO  : 0-1 aciertos
# PROCESO : 2 aciertos
# LOGRADO : 3 aciertos
#
# COMP_2: 3 preguntas
# INICIO  : 0-1 aciertos
# PROCESO : 2 aciertos
# LOGRADO : 3 aciertos
#
# COMP_3: 4 preguntas
# INICIO  : 0-2 aciertos
# PROCESO : 3 aciertos
# LOGRADO : 4 aciertos
#
# COMP_4: 3 preguntas
# INICIO  : 0-1 aciertos
# PROCESO : 2 aciertos
# LOGRADO : 3 aciertos
#
# COMP_5: 3 preguntas
# INICIO  : 0-1 aciertos
# PROCESO : 2 aciertos
# LOGRADO : 3 aciertos

NIVELES_LOGRO = {

    "COMP_1": {
        "INICIO": {
            "minimo": 0,
            "maximo": 1
        },
        "PROCESO": {
            "minimo": 2,
            "maximo": 2
        },
        "LOGRADO": {
            "minimo": 3,
            "maximo": 3
        },
        "REVISAR": {
            "minimo": None,
            "maximo": None
        }
    },

    "COMP_2": {
        "INICIO": {
            "minimo": 0,
            "maximo": 1
        },
        "PROCESO": {
            "minimo": 2,
            "maximo": 2
        },
        "LOGRADO": {
            "minimo": 3,
            "maximo": 3
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
            "maximo": 4
        },
        "REVISAR": {
            "minimo": None,
            "maximo": None
        }
    },

    "COMP_4": {
        "INICIO": {
            "minimo": 0,
            "maximo": 1
        },
        "PROCESO": {
            "minimo": 2,
            "maximo": 2
        },
        "LOGRADO": {
            "minimo": 3,
            "maximo": 3
        },
        "REVISAR": {
            "minimo": None,
            "maximo": None
        }
    },

    "COMP_5": {
        "INICIO": {
            "minimo": 0,
            "maximo": 1
        },
        "PROCESO": {
            "minimo": 2,
            "maximo": 2
        },
        "LOGRADO": {
            "minimo": 3,
            "maximo": 3
        },
        "REVISAR": {
            "minimo": None,
            "maximo": None
        }
    }

}

# ============================================================
# 7. UMBRALES DEL RESULTADO GLOBAL
# ============================================================

LIMITE_GLOBAL_INICIO = 50

LIMITE_GLOBAL_PROCESO = 75

# ============================================================
# 8. COLORES DEL PROYECTO
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

}

# ============================================================
# 9. COLUMNAS DE IDENTIFICACIÓN
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
# 10. COLUMNAS DE RESPUESTAS P1-P16
# ============================================================

COLUMNAS_PREGUNTAS = [

    f"P{i}"
    for i in range(1, TOTAL_PREGUNTAS + 1)

]

# ============================================================
# 11. COLUMNAS DE VALORACIÓN V1-V16
# ============================================================

COLUMNAS_VALORACION = [

    f"V{i}"
    for i in range(1, TOTAL_PREGUNTAS + 1)

]

# ============================================================
# 12. COLUMNAS DE RESULTADOS GENERALES
# ============================================================

COLUMNAS_RESULTADOS = [

    "CORRECTAS",
    "INCORRECTAS",
    "TOTAL_PREGUNTAS",
    "PORCENTAJE",

]

# ============================================================
# 13. COLUMNAS DERIVADAS POR COMPETENCIA
# ============================================================

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

# ============================================================
# 14. COLUMNAS DE RESULTADO GLOBAL
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
# 15. CLASIFICAR NIVEL DE UNA COMPETENCIA
# ============================================================

def clasificar_nivel_competencia(aciertos, competencia):
    try:
        valor = float(aciertos)

    except (TypeError, ValueError):
        return "REVISAR"

    codigo = competencia["codigo"]

    total = len(competencia["preguntas"])

    if not valor.is_integer():
        return "REVISAR"

    if valor < 0 or valor > total:
        return "REVISAR"

    limites = NIVELES_LOGRO[codigo]

    if (
            limites["INICIO"]["minimo"]
            <= valor
            <= limites["INICIO"]["maximo"]
    ):
        return "INICIO"

    if (
            limites["PROCESO"]["minimo"]
            <= valor
            <= limites["PROCESO"]["maximo"]
    ):
        return "PROCESO"

    if (
            limites["LOGRADO"]["minimo"]
            <= valor
            <= limites["LOGRADO"]["maximo"]
    ):
        return "LOGRADO"

    return "REVISAR"


# ============================================================
# 16. CALCULAR PUNTAJE PONDERADO
# ============================================================

def calcular_puntaje_ponderado(aciertos, competencia):
    try:
        valor = float(aciertos)
        peso = float(competencia["peso"])

    except (TypeError, ValueError):
        return None

    total = len(competencia["preguntas"])

    if total == 0:
        return None

    if not valor.is_integer():
        return None

    if valor < 0 or valor > total:
        return None

    if peso < 0 or peso > 1:
        return None

    # Normalización según el número de preguntas.
    # Cada competencia aporta como máximo 20 puntos.
    #
    # Competencias de 3 preguntas:
    # 0 = 0.00
    # 1 = 6.67
    # 2 = 13.33
    # 3 = 20.00
    #
    # Competencia de 4 preguntas:
    # 0 = 0.00
    # 1 = 5.00
    # 2 = 10.00
    # 3 = 15.00
    # 4 = 20.00

    return (valor / total) * (peso * 100)


# ============================================================
# 17. CLASIFICAR RESULTADO GLOBAL
# ============================================================

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


# ============================================================
# 18. VALIDAR DISTRIBUCIÓN DE COMPETENCIAS
# ============================================================


def validar_configuracion():
    if len(COMPETENCIAS) != TOTAL_COMPETENCIAS:
        raise ValueError(
            "Personal Social debe contener cinco competencias."
        )

    if abs(PESO_TOTAL - 1.0) > 0.000001:
        raise ValueError(
            "La suma de los pesos debe ser exactamente 100 %."
        )

    if (
            LIMITE_GLOBAL_INICIO < 0
            or LIMITE_GLOBAL_PROCESO > 100
            or LIMITE_GLOBAL_INICIO >= LIMITE_GLOBAL_PROCESO
    ):
        raise ValueError(
            "Los umbrales globales no son válidos."
        )

    preguntas = []
    codigos = []

    for competencia in COMPETENCIAS:

        codigo = competencia["codigo"]

        total = len(competencia["preguntas"])

        preguntas.extend(competencia["preguntas"])

        # Validar códigos duplicados
        if codigo in codigos:
            raise ValueError(
                "Código de competencia duplicado: " + str(codigo)
            )

        codigos.append(codigo)

        # Validar configuración de niveles
        if codigo not in NIVELES_LOGRO:
            raise ValueError(
                "No existen niveles de logro definidos para "
                + str(codigo)
            )

        # Validar ponderaciones
        if codigo not in PESOS_COMPETENCIAS:
            raise ValueError(
                "No existe ponderación definida para "
                + str(codigo)
            )

        if competencia["peso"] != PESOS_COMPETENCIAS[codigo]:
            raise ValueError(
                "El peso de " + str(codigo)
                + " no coincide con PESOS_COMPETENCIAS."
            )

        # Validar cantidad de preguntas
        if total == 0:
            raise ValueError(
                "La competencia no tiene preguntas: "
                + str(codigo)
            )

        # Validar distribución de niveles de logro
        limites = NIVELES_LOGRO[codigo]

        if (
                limites["INICIO"]["minimo"] != 0
                or limites["INICIO"]["maximo"] + 1
                != limites["PROCESO"]["minimo"]
                or limites["PROCESO"]["maximo"] + 1
                != limites["LOGRADO"]["minimo"]
                or limites["LOGRADO"]["maximo"] != total
        ):
            raise ValueError(
                "Los niveles de logro tienen una distribución "
                "incorrecta en " + str(codigo)
            )

    # Validar que no existan ponderaciones adicionales
    if len(codigos) != len(PESOS_COMPETENCIAS):
        raise ValueError(
            "Existen pesos definidos para competencias inexistentes."
        )

    # Validar que las preguntas cubran P1 hasta P16
    if sorted(preguntas) != list(
            range(1, TOTAL_PREGUNTAS + 1)
    ):
        raise ValueError(
            "La distribución debe cubrir P1 a P16 "
            "sin duplicaciones ni omisiones."
        )

    # Validar preguntas duplicadas
    if len(preguntas) != len(set(preguntas)):
        raise ValueError(
            "Existen preguntas duplicadas entre competencias."
        )

    # Validar columnas de preguntas
    if TOTAL_PREGUNTAS != len(COLUMNAS_PREGUNTAS):
        raise ValueError(
            "La cantidad de columnas de preguntas no coincide."
        )

    # Validar columnas de valoración
    if TOTAL_PREGUNTAS != len(COLUMNAS_VALORACION):
        raise ValueError(
            "La cantidad de columnas de valoración no coincide."
        )

    return True


# ============================================================
# 19. CREAR CARPETAS DEL PROYECTO
# ============================================================

def crear_carpetas():
    CARPETA_RESOURCES.mkdir(
        parents=True,
        exist_ok=True
    )

    CARPETA_RESULTADOS.mkdir(
        parents=True,
        exist_ok=True
    )

    CARPETA_REPORTES_IE.mkdir(
        parents=True,
        exist_ok=True
    )


# ============================================================
# 20. EJECUCIÓN DE VALIDACIÓN
# ============================================================

if __name__ == "__main__":

    crear_carpetas()

    validar_configuracion()

    print("=" * 60)
    print("TUKUYRICUY 2026")
    print("CONFIGURACIÓN DE PERSONAL SOCIAL")
    print("=" * 60)

    print(f"Área: {AREA}")
    print(f"Nivel: {NIVEL_OBJETIVO}")
    print(f"Grados: {GRADOS_OBJETIVO}")
    print(f"Preguntas: {TOTAL_PREGUNTAS}")
    print(f"Competencias: {TOTAL_COMPETENCIAS}")
    print(f"Peso total: {PESO_TOTAL:.0%}")

    print("\nDISTRIBUCIÓN DE COMPETENCIAS:")

    for competencia in COMPETENCIAS:
        print(
            f"{competencia['codigo']} | "
            f"{competencia['nombre']} | "
            f"{competencia['preguntas']} | "
            f"{competencia['peso']:.0%}"
        )

    print("\nUMBRALES GLOBALES:")

    print(
        f"INICIO: 0 - {LIMITE_GLOBAL_INICIO}"
    )

    print(
        f"PROCESO: "
        f">{LIMITE_GLOBAL_INICIO} - "
        f"{LIMITE_GLOBAL_PROCESO}"
    )

    print(
        f"LOGRADO: >{LIMITE_GLOBAL_PROCESO} - 100"
    )

    print("\nRUTAS:")

    print(f"Proyecto: {CARPETA_PROYECTO}")

    print(f"Excel de origen: {ARCHIVO_ENTRADA}")

    print(f"Resultados: {CARPETA_RESULTADOS}")

    print("\n" + "=" * 60)

    print("CONFIGURACIÓN VALIDADA CORRECTAMENTE.")

    print("=" * 60)
