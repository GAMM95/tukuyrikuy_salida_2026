# from pathlib import Path
#
# # ============================================================
# # TUKUYRICUY 2026 - COMUNICACIÓN
# # Configuración general
# # ============================================================
#
# BASE_DIR = Path(__file__).resolve().parent
# CARPETA_DATOS = BASE_DIR / "resources"
# CARPETA_RESULTADOS = BASE_DIR / "resultados"
#
# ARCHIVO_ENTRADA = CARPETA_DATOS / "RESULTADOS_COMUNICACION.xlsx"
# PESTANA_ORIGEN = "RESULTADOS"
#
# AREA = "COMUNICACIÓN"
# TOTAL_PREGUNTAS_ESPERADAS = 16
#
# # Comunicación: Inicio 0-6; Proceso 7-11; Logrado 12-16.
# def clasificar_nivel_logro(correctas):
#     if correctas is None:
#         return "REVISAR"
#     try:
#         puntaje = float(correctas)
#     except (TypeError, ValueError):
#         return "REVISAR"
#
#     if puntaje < 0 or puntaje > TOTAL_PREGUNTAS_ESPERADAS:
#         return "REVISAR"
#     if puntaje <= 6:
#         return "INICIO"
#     if puntaje <= 11:
#         return "PROCESO"
#     return "LOGRADO"
#
#
# COLUMNAS_OBLIGATORIAS = [
#     "ID_REGISTRO", "NIVEL", "INSTITUCION", "GRADO", "SECCION",
#     "APELLIDOS", "NOMBRES", "CORRECTAS", "INCORRECTAS",
#     "TOTAL_PREGUNTAS", "PORCENTAJE",
#     "COMP_1_NOMBRE", "COMP_1_ACIERTOS", "COMP_1_TOTAL",
#     "COMP_1_PORCENTAJE",
# ]
#
# COLUMNAS_IDENTIFICACION = [
#     "ID_REGISTRO", "FECHA_HORA", "NIVEL", "INSTITUCION", "GRADO",
#     "SECCION", "APELLIDOS", "NOMBRES", "DNI", "SEXO",
# ]
#
# COLUMNAS_PREGUNTAS = [f"P{i}" for i in range(1, 17)]
# COLUMNAS_VALIDACION = [f"V{i}" for i in range(1, 17)]
import math
from pathlib import Path

# RUTAS DEL PROYECTO
CARPETA_PROYECTO = Path(__file__).resolve().parent
CARPETA_RESOURCES = CARPETA_PROYECTO / "resources"
CARPETA_RESULTADOS = CARPETA_PROYECTO / "resultados"
CARPETA_REPORTES_IE = CARPETA_RESULTADOS / "REPORTES_INSTITUCIONES"

# Archivo Excel de origen
ARCHIVO_ORIGEN = CARPETA_RESOURCES / "RESULTADOS_COMUNICACION.xlsx"
ARCHIVO_ENTRADA = ARCHIVO_ORIGEN
HOJAS_ORIGEN = ["RESULTADOS"]

# INFORMACIÓN GENERAL
AREA = "COMUNICACIÓN"
NIVELES = ["INICIO", "PROCESO", "LOGRADO"]
NOMBRE_ARCHIVO_GENERAL = "REPORTE_GENERAL_COMUNICACION.xlsx"
NOMBRE_ARCHIVO_PROCESADO = "BASE_PROCESADA_COMUNICACION.xlsx"
NOMBRE_ARCHIVO_ZIP = "REPORTES_INSTITUCIONES_COMUNICACION.zip"

TOTAL_PREGUNTAS = 16
TOTAL_COMPETENCIAS = 1

# COMPETENCIAS
# Comunicación usa una sola competencia con las 16 preguntas.
COMPETENCIAS = [
    {
        "numero": 1,
        "codigo": "COMP_1",
        "nombre": "Comunicación",
        "corto": "Comunicación",
        "preguntas": list(range(1, TOTAL_PREGUNTAS + 1)),
        "peso": 1.0,
    }
]

PESOS_COMPETENCIAS = {str(c["codigo"]): float(c["peso"]) for c in COMPETENCIAS}
PESO_TOTAL = sum(PESOS_COMPETENCIAS.values())

# NIVELES DE LOGRO (según CORRECTAS)
# INICIO  : 0 a 6
# PROCESO : 7 a 11
# LOGRADO : 12 a 16
# Fuera de rango o sin puntaje: REVISAR
NIVELES_LOGRO = {
    "INICIO": {"minimo": 0, "maximo": 6},
    "PROCESO": {"minimo": 7, "maximo": 11},
    "LOGRADO": {"minimo": 12, "maximo": 16},
}

# COLORES DEL PROYECTO
COLORES = {
    # Cabeceras
    "CABECERA": "404040",
    "BLANCO": "FFFFFF",
    "TEXTO": "262626",
    # Niveles de logro
    "INICIO": "F4CCCC",
    "PROCESO": "FFF2CC",
    "LOGRADO": "D9EAD3",
    # Complementarios
    "GRIS_CLARO": "F2F2F2",
    "GRIS_MEDIO": "D9D9D9",
    "BORDE": "BFBFBF",
}

# COLUMNAS
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

COLUMNAS_PREGUNTAS = [f"P{i}" for i in range(1, TOTAL_PREGUNTAS + 1)]
COLUMNAS_VALORACION = [f"V{i}" for i in range(1, TOTAL_PREGUNTAS + 1)]

COLUMNAS_RESULTADOS = ["CORRECTAS", "INCORRECTAS", "TOTAL_PREGUNTAS", "PORCENTAJE"]

COLUMNAS_COMPETENCIAS = []

for _c in COMPETENCIAS:
    _codigo = _c["codigo"]
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
       for campo in ("ACIERTOS", "PORCENTAJE")]
    + COLUMNAS_PREGUNTAS
    + COLUMNAS_VALORACION
    + COLUMNAS_PORCENTAJES
    + ["PUNTAJE_GLOBAL", "NIVEL_LOGRO_GLOBAL", "OBSERVACION_DATOS"]
)

# Columnas obligatorias del Excel de origen
COLUMNAS_OBLIGATORIAS = ["INSTITUCION", "GRADO", "SECCION", "CORRECTAS"]


# CLASIFICAR NIVEL DE LOGRO
def clasificar_nivel_logro(aciertos, total=TOTAL_PREGUNTAS):
    try:
        valor = float(aciertos)
    except (TypeError, ValueError):
        return "REVISAR"

    if math.isnan(valor) or valor < 0 or valor > total or valor != int(valor):
        return "REVISAR"

    for nivel in NIVELES:
        rango = NIVELES_LOGRO[nivel]
        if rango["minimo"] <= valor <= rango["maximo"]:
            return nivel

    return "REVISAR"


# En Comunicación la competencia y el resultado global usan el mismo corte.
clasificar_nivel_competencia = clasificar_nivel_logro
clasificar_nivel_global = clasificar_nivel_logro


# CALCULAR PUNTAJE PONDERADO (sobre 100)
def calcular_puntaje_ponderado(aciertos, peso=1.0, total=TOTAL_PREGUNTAS):
    try:
        valor = float(aciertos)
        peso = float(peso)
    except (TypeError, ValueError):
        return float("nan")

    if math.isnan(valor) or valor < 0 or valor > total:
        return float("nan")
    if peso < 0 or peso > 1:
        return float("nan")

    return (valor / total) * (peso * 100)


# VALIDAR CONFIGURACIÓN
def validar_configuracion():
    if len(COMPETENCIAS) != TOTAL_COMPETENCIAS:
        raise ValueError("La cantidad de competencias no coincide con la configuración.")

    if abs(PESO_TOTAL - 1.0) > 0.000001:
        raise ValueError("La suma de los pesos debe ser exactamente 100 %.")

    preguntas = []
    for competencia in COMPETENCIAS:
        preguntas.extend(competencia["preguntas"])

    if sorted(preguntas) != list(range(1, TOTAL_PREGUNTAS + 1)):
        raise ValueError(
            "La distribución debe cubrir P1 a P16 sin duplicaciones ni omisiones."
        )

    # Los rangos de niveles deben cubrir 0-16 sin huecos ni cruces.
    esperado = 0
    for nivel in NIVELES:
        rango = NIVELES_LOGRO[nivel]
        if rango["minimo"] != esperado:
            raise ValueError(f"El rango de {nivel} no es continuo con el anterior.")
        esperado = rango["maximo"] + 1

    if esperado != TOTAL_PREGUNTAS + 1:
        raise ValueError("Los niveles de logro deben cubrir de 0 a 16 aciertos.")

    return True


# CREAR CARPETAS DEL PROYECTO
def crear_carpetas():
    CARPETA_RESOURCES.mkdir(parents=True, exist_ok=True)
    CARPETA_RESULTADOS.mkdir(parents=True, exist_ok=True)
    CARPETA_REPORTES_IE.mkdir(parents=True, exist_ok=True)


if __name__ == "__main__":
    crear_carpetas()
    validar_configuracion()

    print("==========================================")
    print("TUKUYRICUY 2026")
    print("CONFIGURACIÓN DE COMUNICACIÓN")
    print("==========================================")
    print(f"Área: {AREA}")
    print(f"Preguntas: {TOTAL_PREGUNTAS}")
    print(f"Competencias: {TOTAL_COMPETENCIAS}")
    print("\nNIVELES (por CORRECTAS):")

    for nivel in NIVELES:
        r = NIVELES_LOGRO[nivel]
        print(f"{nivel}: {r['minimo']} a {r['maximo']}")

    print("\nRUTAS:")
    print(f"Excel: {ARCHIVO_ORIGEN}")
    print(f"Resultados: {CARPETA_RESULTADOS}")
    print("\nCONFIGURACIÓN VALIDADA CORRECTAMENTE.")