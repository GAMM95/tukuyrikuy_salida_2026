from pathlib import Path

# RUTAS DEL PROYECTO
CARPETA_PROYECTO = Path(__file__).resolve().parent
CARPETA_RESOURCES = CARPETA_PROYECTO / "resources"
CARPETA_RESULTADOS = CARPETA_PROYECTO / "resultados"
CARPETA_REPORTES_IE = (CARPETA_RESULTADOS / "REPORTES_INSTITUCIONES")

# Archivo Excel de origen
ARCHIVO_ORIGEN = (CARPETA_RESOURCES / "RESULTADOS_MATEMATICA.xlsx")
# Hojas del Excel de origen
HOJAS_ORIGEN = ["RESULTADOS"]


def crear_carpetas():
    CARPETA_RESOURCES.mkdir(parents=True, exist_ok=True)
    CARPETA_RESULTADOS.mkdir(parents=True, exist_ok=True)
    CARPETA_REPORTES_IE.mkdir(parents=True, exist_ok=True)

# 2. ARCHIVO EXCEL DE ORIGEN
ARCHIVO_ENTRADA = (CARPETA_RESOURCES / "RESULTADOS_MATEMATICA.xlsx")

# El archivo de resultados puede contener dos pestañas:
# PRIMARIA y SECUNDARIA.
HOJAS_ORIGEN = ["RESULTADOS"]

# 3. INFORMACIÓN GENERAL
AREA = "MATEMÁTICA"
NIVELES = ["INICIO", "PROCESO", "LOGRADO"]
NOMBRE_ARCHIVO_GENERAL = ("REPORTE_GENERAL_MATEMATICA.xlsx")
NOMBRE_ARCHIVO_PROCESADO = ("BASE_PROCESADA_MATEMATICA.xlsx")
NOMBRE_ARCHIVO_ZIP = ("REPORTES_INSTITUCIONES_MATEMATICA.zip")

TOTAL_PREGUNTAS = 16
TOTAL_COMPETENCIAS = 4

# 4. COMPETENCIAS Y DISTRIBUCIÓN DE PREGUNTAS
COMPETENCIAS = [
    {
        "numero": 1,
        "codigo": "COMP_1",
        "nombre": "Resuelve problemas de cantidad",
        "corto": "Cantidad",
        "preguntas": [1, 2, 3, 4],
        "peso": 0.25
    },
    {
        "numero": 2,
        "codigo": "COMP_2",
        "nombre": "Resuelve problemas de regularidad, equivalencia y cambio",
        "corto": "Regularidad",
        "preguntas": [5, 6, 7, 8],
        "peso": 0.25
    },
    {
        "numero": 3,
        "codigo": "COMP_3",
        "nombre": "Resuelve problemas de forma, movimiento y localización",
        "corto": "Forma y movimiento",
        "preguntas": [9, 10, 11, 12],
        "peso": 0.25
    },
    {
        "numero": 4,
        "codigo": "COMP_4",
        "nombre": "Resuelve problemas de gestión de datos e incertidumbre",
        "corto": "Datos e incertidumbre",
        "preguntas": [13, 14, 15, 16],
        "peso": 0.25
    }
]

# 5. PESOS DE LAS COMPETENCIAS
PESOS_COMPETENCIAS = {
    "COMP_1": 0.25,
    "COMP_2": 0.25,
    "COMP_3": 0.25,
    "COMP_4": 0.25,
}
PESO_TOTAL = sum(PESOS_COMPETENCIAS.values())

# 6. NIVELES DE LOGRO POR COMPETENCIA
# Cada competencia contiene cuatro preguntas.
# INICIO  : 0, 1 o 2 aciertos
# PROCESO : 3 aciertos
# LOGRADO : 4 aciertos
NIVELES_LOGRO = {
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
}

# 7. COLORES DEL PROYECTO
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

# 8. COLUMNAS DE IDENTIFICACIÓN
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

# 9. COLUMNAS DE RESPUESTAS P1-P16
COLUMNAS_PREGUNTAS = [
    f"P{i}"
    for i in range(1, TOTAL_PREGUNTAS + 1)
]

# 10. COLUMNAS DE VALORACIÓN V1-V16
COLUMNAS_VALORACION = [
    f"V{i}"
    for i in range(1, TOTAL_PREGUNTAS + 1)
]

# 11. COLUMNAS DE RESULTADOS GENERALES
COLUMNAS_RESULTADOS = ["CORRECTAS", "INCORRECTAS", "TOTAL_PREGUNTAS", "PORCENTAJE", ]

# 12. COLUMNAS DERIVADAS POR COMPETENCIA
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

# 13. COLUMNAS DE RESULTADO GLOBAL
COLUMNAS_RESULTADO_GLOBAL = [
    "PUNTAJE_GLOBAL",
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
    + ["PUNTAJE_GLOBAL", "NIVEL_LOGRO_GLOBAL"]
)

# 14. CLASIFICAR NIVEL DE UNA COMPETENCIA
def clasificar_nivel_competencia(aciertos):
    try:
        valor = float(aciertos)
    except (TypeError, ValueError):
        return "REVISAR"
    if valor < 0 or valor > 4:        return "REVISAR"
    if valor <= 2:        return "INICIO"
    if valor <= 3:        return "PROCESO"
    return "LOGRADO"


# 15. CALCULAR PUNTAJE PONDERADO
def calcular_puntaje_ponderado(aciertos, peso=0.25):
    try:
        valor = float(aciertos)
        peso = float(peso)
    except (TypeError, ValueError):
        return None

    if valor < 0 or valor > 4:        return None
    if peso < 0 or peso > 1:        return None

    # Cada competencia aporta como máximo 25 puntos.
    #
    # 0 aciertos = 0 puntos
    # 1 acierto  = 6.25 puntos
    # 2 aciertos = 12.50 puntos
    # 3 aciertos = 18.75 puntos
    # 4 aciertos = 25 puntos
    return (valor / 4) * (peso * 100)


# 16. CLASIFICAR RESULTADO GLOBAL
def clasificar_nivel_global(puntaje):
    try:

        valor = float(puntaje)

    except (TypeError, ValueError):
        return "REVISAR"

    if valor < 0 or valor > 100:        return "REVISAR"
    if valor <= 50:        return "INICIO"
    if valor <= 75:        return "PROCESO"
    return "LOGRADO"


# ============================================================
# 17. VALIDAR DISTRIBUCIÓN DE COMPETENCIAS
def validar_configuracion():
    if len(COMPETENCIAS) != 4:
        raise ValueError(
            "La configuración debe contener cuatro competencias."
        )

    if abs(PESO_TOTAL - 1.0) > 0.000001:
        raise ValueError(
            "La suma de los pesos debe ser exactamente 100 %."
        )

    preguntas = []

    for competencia in COMPETENCIAS:
        preguntas.extend(
            competencia["preguntas"]
        )

    if sorted(preguntas) != list(range(1, TOTAL_PREGUNTAS + 1)):
        raise ValueError(
            "La distribución debe cubrir P1 a P16 "
            "sin duplicaciones ni omisiones."
        )
    for competencia in COMPETENCIAS:
        if len(competencia["preguntas"]) != 4:
            raise ValueError(
                "Cada competencia debe tener cuatro preguntas: "
                + competencia["nombre"]
            )
    return True


# 18. CREAR CARPETAS DEL PROYECTO
def crear_carpetas():
    CARPETA_RESOURCES.mkdir(parents=True, exist_ok=True)
    CARPETA_RESULTADOS.mkdir(parents=True, exist_ok=True)
    CARPETA_REPORTES_IE.mkdir(parents=True, exist_ok=True)


# 19. EJECUCIÓN DE VALIDACIÓN
if __name__ == "__main__":
    crear_carpetas()
    validar_configuracion()

    print("==========================================")
    print("TUKUYRICUY 2026")
    print("CONFIGURACIÓN DE MATEMÁTICA")
    print("==========================================")
    print(f"Área: {AREA}")
    print(f"Preguntas: {TOTAL_PREGUNTAS}")
    print(f"Competencias: {TOTAL_COMPETENCIAS}")
    print(f"Peso total: {PESO_TOTAL:.0%}")
    print("\nDISTRIBUCIÓN:")

    for competencia in COMPETENCIAS:
        print(
            f"{competencia['codigo']} | "
            f"{competencia['nombre']} | "
            f"{competencia['preguntas']} | "
            f"{competencia['peso']:.0%}"
        )
    print("\nRUTAS:")
    print(f"Excel: {ARCHIVO_ENTRADA}")
    print(f"Resultados: {CARPETA_RESULTADOS}")
    print("\nCONFIGURACIÓN VALIDADA CORRECTAMENTE.")
