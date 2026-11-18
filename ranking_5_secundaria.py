# ================================================================
# TUKUYRICUY 2026
# RANKING DE MEJORES ESTUDIANTES - 5TO SECUNDARIA
#
# CRITERIOS:
# 1. Mayor promedio de las cuatro áreas.
# 2. Desempate: mayor suma de porcentajes.
# 3. Solo estudiantes con resultados en las cuatro áreas.
# 4. No se utiliza tiempo de resolución.
#
# IDENTIFICACIÓN:
# 1. DNI válido.
# 2. Apellidos + nombres + institución.
# 3. Si existe ambigüedad, no unir automáticamente.
#
# DUPLICADOS:
# Se conserva el primer intento según FECHA_HORA.
#
# ================================================================

from pathlib import Path
import re
import unicodedata

import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# ================================================================
# 1. CONFIGURACIÓN
# ================================================================

CARPETA_PROYECTO = Path(__file__).resolve().parent

CARPETA_RESULTADOS = (
        CARPETA_PROYECTO / "RANKING_5_SECUNDARIA"
)

ARCHIVOS_AREAS = {
    "COMUNICACION": {
        "carpetas": ["COMUNICACION"],
        "archivos": [
            "BASE_PROCESADA_COMUNICACION.xlsx"
        ]
    },
    "MATEMATICA": {
        "carpetas": ["MATEMATICA"],
        "archivos": [
            "BASE_PROCESADA_MATEMATICA.xlsx"
        ]
    },
    "CCSS": {
        "carpetas": ["CCSS"],
        "archivos": [
            "BASE_PROCESADA_CCSS_SECUNDARIA.xlsx"
        ]
    },
    "CYT": {
        "carpetas": ["CYT"],
        "archivos": [
            "BASE_PROCESADA_CYT.xlsx"
        ]
    }
}

NOMBRE_ARCHIVO_SALIDA = (
    "RANKING_MEJORES_ESTUDIANTES_5_SECUNDARIA.xlsx"
)

HOJA_ORIGEN = "RESULTADOS"
GRADO_OBJETIVO = 5

AREAS = list(ARCHIVOS_AREAS.keys())

AREA_NOMBRES = {
    "COMUNICACION": "COMUNICACIÓN",
    "MATEMATICA": "MATEMÁTICA",
    "CCSS": "CIENCIAS SOCIALES",
    "CYT": "CIENCIA Y TECNOLOGÍA"
}

COLOR_CABECERA = "404040"
COLOR_BLANCO = "FFFFFF"
COLOR_VERDE = "C6E0B4"
COLOR_AMARILLO = "FFF2CC"
COLOR_ROJO = "F4CCCC"


# ================================================================
# 2. NORMALIZACIÓN
# ================================================================

def normalizar_texto(valor):
    if pd.isna(valor):
        return ""

    texto = str(valor).strip().upper()

    texto = unicodedata.normalize(
        "NFKD",
        texto
    )

    texto = "".join(
        caracter
        for caracter in texto
        if not unicodedata.combining(caracter)
    )

    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def normalizar_dni(valor):
    if pd.isna(valor):
        return None

    texto = str(valor).strip().upper()

    if texto.endswith(".0"):
        texto = texto[:-2]

    texto = re.sub(r"\s+", "", texto)

    # DNI peruano: exactamente 8 dígitos.
    if not re.fullmatch(r"\d{8}", texto):
        return None

    # Evitar valores evidentemente incorrectos.
    if texto == "00000000":
        return None

    return texto


def normalizar_nivel(valor):
    return normalizar_texto(valor)


def normalizar_grado(valor):
    if pd.isna(valor):
        return None

    texto = normalizar_texto(valor)

    # Captura 5, 5.0, 5°, 5TO, 5.º, 5 DE SECUNDARIA.
    coincidencia = re.match(
        r"^\s*([1-6])",
        texto
    )

    if coincidencia:
        return int(coincidencia.group(1))

    return None


def normalizar_porcentaje(valor):
    if pd.isna(valor):
        return None

    if isinstance(valor, str):
        valor = valor.strip().replace("%", "")
        valor = valor.replace(",", ".")

    numero = pd.to_numeric(
        valor,
        errors="coerce"
    )

    if pd.isna(numero):
        return None

    return float(numero)


def crear_clave_nombre(fila):
    apellidos = normalizar_texto(
        fila.get("APELLIDOS", "")
    )

    nombres = normalizar_texto(
        fila.get("NOMBRES", "")
    )

    institucion = normalizar_texto(
        fila.get("INSTITUCION", "")
    )

    # No realizar coincidencias por nombre si faltan
    # datos esenciales.
    if not apellidos or not nombres or not institucion:
        return None

    return (
        institucion,
        apellidos,
        nombres
    )


# ================================================================
# 3. LOCALIZAR ARCHIVOS
# ================================================================

def buscar_archivo_area(area, configuracion):
    nombres = configuracion["archivos"]
    carpetas = configuracion["carpetas"]

    for carpeta in carpetas:

        directorio = CARPETA_PROYECTO / carpeta

        for nombre in nombres:

            candidatos = [
                directorio / "resultados" / nombre,
                directorio / nombre
            ]

            for archivo in candidatos:

                if archivo.exists():
                    return archivo

    for nombre in nombres:

        encontrados = list(
            CARPETA_PROYECTO.rglob(nombre)
        )

        if len(encontrados) == 1:
            return encontrados[0]

        if len(encontrados) > 1:
            raise ValueError(
                f"Se encontraron varios archivos para {area}:\n"
                + "\n".join(
                    str(a) for a in encontrados
                )
            )

    raise FileNotFoundError(
        f"No se encontró la base procesada de {area}.\n"
        f"Nombres esperados: {nombres}"
    )


# ================================================================
# 4. OBTENER COLUMNA DE PORCENTAJE
# ================================================================

def obtener_columna_porcentaje(df):
    posibles = [
        "PORCENTAJE_GLOBAL",
        "PORCENTAJE",
        "PORCENTAJE_REPORTE"
    ]

    for columna in posibles:

        if columna in df.columns:
            return columna

    raise ValueError(
        "No se encontró una columna de porcentaje.\n"
        f"Columnas disponibles: {list(df.columns)}"
    )


# ================================================================
# 5. LEER Y PREPARAR UNA ÁREA
# ================================================================

def cargar_area(area, configuracion):
    archivo = buscar_archivo_area(
        area,
        configuracion
    )

    print(f"\nLeyendo {AREA_NOMBRES[area]}...")
    print(f"Archivo: {archivo}")

    libro = pd.ExcelFile(archivo)

    hoja = (
        HOJA_ORIGEN
        if HOJA_ORIGEN in libro.sheet_names
        else libro.sheet_names[0]
    )

    df = pd.read_excel(
        archivo,
        sheet_name=hoja,
        dtype={"DNI": str}
    )

    if df.empty:
        raise ValueError(
            f"La base de {area} no contiene registros."
        )

    # ------------------------------------------------------------
    # VALIDAR COLUMNAS
    # ------------------------------------------------------------

    if "GRADO" not in df.columns:
        raise ValueError(
            f"La base de {area} no contiene GRADO."
        )

    if "NIVEL" not in df.columns:
        raise ValueError(
            f"La base de {area} no contiene NIVEL."
        )

    columna_porcentaje = obtener_columna_porcentaje(df)

    # Crear columnas opcionales si no existen.
    for columna in [
        "DNI",
        "APELLIDOS",
        "NOMBRES",
        "INSTITUCION",
        "SECCION",
        "FECHA_HORA"
    ]:

        if columna not in df.columns:
            df[columna] = None

    # ------------------------------------------------------------
    # FILTRAR SECUNDARIA
    # ------------------------------------------------------------

    nivel = df["NIVEL"].apply(
        normalizar_nivel
    )

    df = df[
        nivel.str.contains(
            "SECUNDARIA",
            na=False
        )
    ].copy()

    # ------------------------------------------------------------
    # FILTRAR QUINTO GRADO
    # ------------------------------------------------------------

    df["_GRADO_NUM"] = df["GRADO"].apply(
        normalizar_grado
    )

    df = df[
        df["_GRADO_NUM"] == GRADO_OBJETIVO
        ].copy()

    # ------------------------------------------------------------
    # NORMALIZAR IDENTIFICADORES
    # ------------------------------------------------------------

    df["DNI_ORIGINAL"] = df["DNI"]

    df["DNI_VALIDO"] = df["DNI"].apply(
        normalizar_dni
    )

    df["CLAVE_NOMBRE"] = df.apply(
        crear_clave_nombre,
        axis=1
    )

    # ------------------------------------------------------------
    # NORMALIZAR PORCENTAJES
    # ------------------------------------------------------------

    df["PORCENTAJE_AREA"] = df[
        columna_porcentaje
    ].apply(
        normalizar_porcentaje
    )

    # ------------------------------------------------------------
    # PREPARAR FECHA
    # ------------------------------------------------------------

    df["_FECHA_ORDEN"] = pd.to_datetime(
        df["FECHA_HORA"],
        errors="coerce",
        dayfirst=True
    )

    # Mantener orden original para registros sin fecha.
    df["_ORDEN_ORIGINAL"] = range(len(df))

    # ------------------------------------------------------------
    # DESCARTAR REGISTROS SIN IDENTIFICACIÓN SUFICIENTE
    # ------------------------------------------------------------

    identificables = (
            df["DNI_VALIDO"].notna()
            | df["CLAVE_NOMBRE"].notna()
    )

    no_identificables = df[
        ~identificables
    ].copy()

    df = df[
        identificables
    ].copy()

    # ------------------------------------------------------------
    # ORDENAR INTENTOS
    # ------------------------------------------------------------

    df = df.sort_values(
        by=[
            "_FECHA_ORDEN",
            "_ORDEN_ORIGINAL"
        ],
        ascending=[
            True,
            True
        ],
        na_position="last"
    )

    # ------------------------------------------------------------
    # CONSERVAR PRIMER INTENTO POR ESTUDIANTE
    # ------------------------------------------------------------

    claves_vistas = set()
    filas_conservadas = []
    registros_duplicados = []

    for _, fila in df.iterrows():

        dni = fila["DNI_VALIDO"]
        clave_nombre = fila["CLAVE_NOMBRE"]

        if dni:

            clave = ("DNI", dni)

        elif clave_nombre:

            clave = ("NOMBRE", clave_nombre)

        else:
            continue

        if clave in claves_vistas:
            registros_duplicados.append({
                "AREA": area,
                "DNI": dni,
                "APELLIDOS": fila["APELLIDOS"],
                "NOMBRES": fila["NOMBRES"],
                "INSTITUCION": fila["INSTITUCION"],
                "FECHA_HORA": fila["FECHA_HORA"],
                "MOTIVO": "Intento duplicado. Se conserva el primero."
            })

            continue

        claves_vistas.add(clave)
        filas_conservadas.append(fila)

    df = pd.DataFrame(filas_conservadas)

    if df.empty:
        raise ValueError(
            f"No existen registros identificables para {area}."
        )

    # ------------------------------------------------------------
    # RESULTADO ESTANDARIZADO
    # ------------------------------------------------------------

    columnas = [
        "DNI_VALIDO",
        "CLAVE_NOMBRE",
        "APELLIDOS",
        "NOMBRES",
        "INSTITUCION",
        "SECCION",
        "FECHA_HORA",
        "PORCENTAJE_AREA"
    ]

    resultado = df[columnas].copy()

    resultado = resultado.rename(
        columns={
            "DNI_VALIDO": "DNI",
            "APELLIDOS": f"APELLIDOS_{area}",
            "NOMBRES": f"NOMBRES_{area}",
            "INSTITUCION": f"INSTITUCION_{area}",
            "SECCION": f"SECCION_{area}",
            "FECHA_HORA": f"FECHA_HORA_{area}",
            "PORCENTAJE_AREA": f"PORCENTAJE_{area}"
        }
    )

    resultado = resultado.reset_index(drop=True)

    print(
        f"Registros identificables: {len(resultado):,}"
    )

    print(
        f"Intentos duplicados descartados: "
        f"{len(registros_duplicados):,}"
    )

    print(
        f"Registros sin identificación suficiente: "
        f"{len(no_identificables):,}"
    )

    return (
        resultado,
        registros_duplicados,
        no_identificables
    )


# ================================================================
# 6. COMPARAR IDENTIDADES
# ================================================================

def buscar_coincidencia(registro, estudiante):
    dni_nuevo = registro.get("DNI")
    clave_nueva = registro.get("CLAVE_NOMBRE")

    dni_existente = estudiante.get("DNI")
    clave_existente = estudiante.get("CLAVE_NOMBRE")

    # No unir si ambos tienen DNI válido pero son diferentes.
    if (
            dni_nuevo
            and dni_existente
            and dni_nuevo != dni_existente
    ):
        return None

    # Primera prioridad: DNI idéntico.
    if (
            dni_nuevo
            and dni_existente
            and dni_nuevo == dni_existente
    ):
        return "DNI"

    # Segunda prioridad: nombre completo e institución.
    if (
            clave_nueva
            and clave_existente
            and clave_nueva == clave_existente
    ):
        return "NOMBRE_INSTITUCION"

    return None


# ================================================================
# 7. CONSTRUIR CRUCE PROGRESIVO
# ================================================================

def cruzar_resultados():
    estudiantes = []

    registros_duplicados = []
    registros_no_identificables = []
    coincidencias_revision = []

    for area, configuracion in ARCHIVOS_AREAS.items():

        resultado, duplicados, no_identificables = cargar_area(
            area,
            configuracion
        )

        registros_duplicados.extend(duplicados)

        for _, fila in no_identificables.iterrows():
            registros_no_identificables.append({
                "AREA": area,
                "DNI_ORIGINAL": fila.get("DNI_ORIGINAL"),
                "APELLIDOS": fila.get("APELLIDOS"),
                "NOMBRES": fila.get("NOMBRES"),
                "INSTITUCION": fila.get("INSTITUCION"),
                "MOTIVO": "No tiene DNI válido ni datos completos para identificarlo."
            })

        # --------------------------------------------------------
        # PROCESAR CADA ESTUDIANTE DE ESTA ÁREA
        # --------------------------------------------------------

        for _, fila in resultado.iterrows():

            registro = fila.to_dict()

            candidatos_dni = []
            candidatos_nombre = []

            for indice, estudiante in enumerate(estudiantes):

                # No permitir dos registros de la misma área
                # dentro del mismo estudiante consolidado.
                if area in estudiante["AREAS"]:
                    continue

                metodo = buscar_coincidencia(
                    registro,
                    estudiante
                )

                if metodo == "DNI":
                    candidatos_dni.append(indice)

                elif metodo == "NOMBRE_INSTITUCION":
                    candidatos_nombre.append(indice)

            # ----------------------------------------------------
            # PRIORIZAR DNI
            # ----------------------------------------------------

            candidatos = (
                candidatos_dni
                if candidatos_dni
                else candidatos_nombre
            )

            if len(candidatos) == 1:

                indice = candidatos[0]
                estudiante = estudiantes[indice]

                estudiante["AREAS"][area] = registro

                # Completar DNI si el registro nuevo lo aporta.
                if (
                        not estudiante.get("DNI")
                        and registro.get("DNI")
                ):
                    estudiante["DNI"] = registro["DNI"]

                metodo = (
                    "DNI"
                    if candidatos_dni
                    else "NOMBRE + INSTITUCION"
                )

                estudiante["METODOS"].append(
                    f"{area}: {metodo}"
                )

            elif len(candidatos) > 1:

                coincidencias_revision.append({
                    "AREA": area,
                    "DNI": registro.get("DNI"),
                    "APELLIDOS": registro.get(
                        f"APELLIDOS_{area}"
                    ),
                    "NOMBRES": registro.get(
                        f"NOMBRES_{area}"
                    ),
                    "INSTITUCION": registro.get(
                        f"INSTITUCION_{area}"
                    ),
                    "CANDIDATOS_ENCONTRADOS": len(candidatos),
                    "MOTIVO": "Coincidencia ambigua. No se unió automáticamente."
                })

                # Crear registro independiente para evitar
                # una unión incorrecta.
                estudiantes.append({
                    "DNI": registro.get("DNI"),
                    "CLAVE_NOMBRE": registro.get(
                        "CLAVE_NOMBRE"
                    ),
                    "AREAS": {
                        area: registro
                    },
                    "METODOS": [
                        f"{area}: REGISTRO INDEPENDIENTE"
                    ]
                })

            else:

                # No se encontró coincidencia.
                estudiantes.append({
                    "DNI": registro.get("DNI"),
                    "CLAVE_NOMBRE": registro.get(
                        "CLAVE_NOMBRE"
                    ),
                    "AREAS": {
                        area: registro
                    },
                    "METODOS": [
                        f"{area}: REGISTRO INICIAL"
                    ]
                })

    # ============================================================
    # 8. CONVERTIR ESTUDIANTES CONSOLIDADOS EN DATAFRAME
    # ============================================================

    filas = []

    for estudiante in estudiantes:

        fila = {
            "DNI": estudiante.get("DNI"),
            "METODO_IDENTIFICACION": " | ".join(
                estudiante["METODOS"]
            )
        }

        for area in AREAS:

            registro = estudiante["AREAS"].get(area)

            if registro:

                fila[f"PORCENTAJE_{area}"] = registro.get(
                    f"PORCENTAJE_{area}"
                )

                fila[f"APELLIDOS_{area}"] = registro.get(
                    f"APELLIDOS_{area}"
                )

                fila[f"NOMBRES_{area}"] = registro.get(
                    f"NOMBRES_{area}"
                )

                fila[f"INSTITUCION_{area}"] = registro.get(
                    f"INSTITUCION_{area}"
                )

                fila[f"SECCION_{area}"] = registro.get(
                    f"SECCION_{area}"
                )

                fila[f"FECHA_HORA_{area}"] = registro.get(
                    f"FECHA_HORA_{area}"
                )

            else:

                fila[f"PORCENTAJE_{area}"] = None
                fila[f"APELLIDOS_{area}"] = None
                fila[f"NOMBRES_{area}"] = None
                fila[f"INSTITUCION_{area}"] = None
                fila[f"SECCION_{area}"] = None
                fila[f"FECHA_HORA_{area}"] = None

        filas.append(fila)

    df = pd.DataFrame(filas)

    if df.empty:
        raise ValueError(
            "No se encontraron estudiantes para consolidar."
        )

    # ============================================================
    # 9. CALCULAR RESULTADOS
    # ============================================================

    columnas_porcentajes = [
        f"PORCENTAJE_{area}"
        for area in AREAS
    ]

    df["AREAS_CON_RESULTADO"] = (
        df[columnas_porcentajes]
        .notna()
        .sum(axis=1)
    )

    df["AREAS_FALTANTES"] = (
            4 - df["AREAS_CON_RESULTADO"]
    )

    df["RESULTADO_COMPLETO"] = (
            df["AREAS_CON_RESULTADO"] == 4
    )

    df["PROMEDIO_GENERAL"] = (
        df[columnas_porcentajes]
        .mean(axis=1)
    )

    df["TOTAL_PORCENTAJES"] = (
        df[columnas_porcentajes]
        .sum(axis=1, min_count=4)
    )

    # ============================================================
    # 10. CONSOLIDAR DATOS PERSONALES
    # ============================================================

    def primer_valor(fila, prefijo):

        for area in AREAS:

            columna = f"{prefijo}_{area}"

            if columna in fila.index:

                valor = fila[columna]

                if pd.notna(valor) and str(valor).strip():
                    return valor

        return ""

    df["APELLIDOS"] = df.apply(
        lambda fila: primer_valor(
            fila,
            "APELLIDOS"
        ),
        axis=1
    )

    df["NOMBRES"] = df.apply(
        lambda fila: primer_valor(
            fila,
            "NOMBRES"
        ),
        axis=1
    )

    df["INSTITUCION"] = df.apply(
        lambda fila: primer_valor(
            fila,
            "INSTITUCION"
        ),
        axis=1
    )

    df["SECCION"] = df.apply(
        lambda fila: primer_valor(
            fila,
            "SECCION"
        ),
        axis=1
    )

    df["ESTUDIANTE"] = (
            df["APELLIDOS"].fillna("").astype(str)
            + ", "
            + df["NOMBRES"].fillna("").astype(str)
    ).str.strip(" ,")

    # ============================================================
    # 11. SEPARAR COMPLETOS E INCOMPLETOS
    # ============================================================

    completos = df[
        df["RESULTADO_COMPLETO"]
    ].copy()

    incompletos = df[
        ~df["RESULTADO_COMPLETO"]
    ].copy()

    # ============================================================
    # 12. ORDENAR RANKING
    # ============================================================

    completos["PROMEDIO_GENERAL"] = (
        completos["PROMEDIO_GENERAL"].round(6)
    )

    completos["TOTAL_PORCENTAJES"] = (
        completos["TOTAL_PORCENTAJES"].round(6)
    )

    completos = completos.sort_values(
        by=[
            "PROMEDIO_GENERAL",
            "TOTAL_PORCENTAJES",
            "DNI"
        ],
        ascending=[
            False,
            False,
            True
        ],
        na_position="last"
    ).reset_index(drop=True)

    # Puestos: mismo puesto cuando coinciden promedio
    # y suma de porcentajes.
    pares = completos[
        [
            "PROMEDIO_GENERAL",
            "TOTAL_PORCENTAJES"
        ]
    ].drop_duplicates().copy()

    pares = pares.sort_values(
        by=[
            "PROMEDIO_GENERAL",
            "TOTAL_PORCENTAJES"
        ],
        ascending=[
            False,
            False
        ]
    ).reset_index(drop=True)

    pares["PUESTO"] = range(
        1,
        len(pares) + 1
    )

    completos = completos.merge(
        pares,
        on=[
            "PROMEDIO_GENERAL",
            "TOTAL_PORCENTAJES"
        ],
        how="left"
    )

    completos = completos.sort_values(
        by=[
            "PUESTO",
            "DNI"
        ],
        ascending=[
            True,
            True
        ],
        na_position="last"
    ).reset_index(drop=True)

    # ============================================================
    # 13. PREPARAR HOJA RANKING
    # ============================================================

    columnas_ranking = [
        "PUESTO",
        "DNI",
        "ESTUDIANTE",
        "INSTITUCION",
        "SECCION",
        "PORCENTAJE_COMUNICACION",
        "PORCENTAJE_MATEMATICA",
        "PORCENTAJE_CCSS",
        "PORCENTAJE_CYT",
        "PROMEDIO_GENERAL",
        "TOTAL_PORCENTAJES"
    ]

    ranking = completos[
        columnas_ranking
    ].copy()

    columnas_incompletos = [
        "DNI",
        "ESTUDIANTE",
        "INSTITUCION",
        "SECCION",
        "PORCENTAJE_COMUNICACION",
        "PORCENTAJE_MATEMATICA",
        "PORCENTAJE_CCSS",
        "PORCENTAJE_CYT",
        "AREAS_CON_RESULTADO",
        "AREAS_FALTANTES",
        "METODO_IDENTIFICACION"
    ]

    incompletos = incompletos[
        columnas_incompletos
    ].copy()

    duplicados_df = pd.DataFrame(
        registros_duplicados
    )

    no_identificables_df = pd.DataFrame(
        registros_no_identificables
    )

    revision_df = pd.DataFrame(
        coincidencias_revision
    )

    return (
        ranking,
        completos,
        incompletos,
        duplicados_df,
        no_identificables_df,
        revision_df
    )


# ================================================================
# 14. ESTILOS EXCEL
# ================================================================

def aplicar_estilo(ws):
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"

    if ws.max_row >= 2 and ws.max_column >= 1:
        ws.auto_filter.ref = ws.dimensions

    for celda in ws[1]:
        celda.fill = PatternFill(
            fill_type="solid",
            fgColor=COLOR_CABECERA
        )

        celda.font = Font(
            bold=True,
            color=COLOR_BLANCO,
            size=10
        )

        celda.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True
        )

        celda.border = Border(
            bottom=Side(
                style="thin",
                color="BFBFBF"
            )
        )

    ws.row_dimensions[1].height = 35

    for numero, columna in enumerate(
            ws.iter_cols(
                min_row=1,
                max_row=ws.max_row
            ),
            start=1
    ):
        letra = get_column_letter(numero)

        valores = [
            len(str(celda.value))
            for celda in columna
            if celda.value is not None
        ]

        ancho = min(
            max(max(valores, default=12) + 3, 12),
            45
        )

        ws.column_dimensions[letra].width = ancho


# ================================================================
# 15. ESCRIBIR DATAFRAME EN EXCEL
# ================================================================

def escribir_dataframe(ws, df):
    if df.empty:
        ws.append([
            "Sin registros"
        ])

        aplicar_estilo(ws)
        return

    ws.append(
        list(df.columns)
    )

    for fila in df.itertuples(
            index=False,
            name=None
    ):

        valores = []

        for valor in fila:

            if pd.isna(valor):
                valores.append(None)

            elif isinstance(valor, pd.Timestamp):
                valores.append(
                    valor.to_pydatetime()
                )

            else:
                valores.append(valor)

        ws.append(valores)

    aplicar_estilo(ws)


# ================================================================
# 16. GENERAR EXCEL
# ================================================================

def generar_excel(
        ranking,
        completos,
        incompletos,
        duplicados,
        no_identificables,
        revision
):
    CARPETA_RESULTADOS.mkdir(
        parents=True,
        exist_ok=True
    )

    ruta = (
            CARPETA_RESULTADOS
            / NOMBRE_ARCHIVO_SALIDA
    )

    wb = Workbook()

    # ------------------------------------------------------------
    # HOJA 1: RANKING
    # ------------------------------------------------------------

    ws = wb.active
    ws.title = "RANKING_5_SECUNDARIA"

    escribir_dataframe(
        ws,
        ranking
    )

    if ws.max_row >= 2:

        for fila in range(2, ws.max_row + 1):

            for columna in range(6, 12):
                ws.cell(
                    row=fila,
                    column=columna
                ).number_format = '0.00"%"'

        # Resaltar primer puesto.
        primer_puesto = ws[2][0].value

        for fila in range(2, ws.max_row + 1):

            if ws.cell(fila, 1).value == primer_puesto:

                for celda in ws[fila]:
                    celda.fill = PatternFill(
                        fill_type="solid",
                        fgColor=COLOR_VERDE
                    )

                    celda.font = Font(
                        bold=True
                    )

    # ------------------------------------------------------------
    # HOJA 2: DETALLE CRUCE
    # ------------------------------------------------------------

    ws = wb.create_sheet("DETALLE_CRUCE")

    columnas = [
        "PUESTO",
        "DNI",
        "ESTUDIANTE",
        "INSTITUCION",
        "SECCION",
        "PORCENTAJE_COMUNICACION",
        "PORCENTAJE_MATEMATICA",
        "PORCENTAJE_CCSS",
        "PORCENTAJE_CYT",
        "TOTAL_PORCENTAJES",
        "PROMEDIO_GENERAL",
        "AREAS_CON_RESULTADO",
        "METODO_IDENTIFICACION"
    ]

    detalle = completos[
        [
            columna
            for columna in columnas
            if columna in completos.columns
        ]
    ].copy()

    escribir_dataframe(
        ws,
        detalle
    )

    # ------------------------------------------------------------
    # HOJA 3: NO INCLUIDOS
    # ------------------------------------------------------------

    ws = wb.create_sheet("NO_INCLUIDOS")

    escribir_dataframe(
        ws,
        incompletos
    )

    # ------------------------------------------------------------
    # HOJA 4: REGISTROS DUPLICADOS
    # ------------------------------------------------------------

    ws = wb.create_sheet("REGISTROS_DUPLICADOS")

    escribir_dataframe(
        ws,
        duplicados
    )

    # ------------------------------------------------------------
    # HOJA 5: SIN IDENTIFICACIÓN
    # ------------------------------------------------------------

    ws = wb.create_sheet("SIN_IDENTIFICACION")

    escribir_dataframe(
        ws,
        no_identificables
    )

    # ------------------------------------------------------------
    # HOJA 6: COINCIDENCIAS PARA REVISAR
    # ------------------------------------------------------------

    ws = wb.create_sheet("COINCIDENCIAS_REVISION")

    escribir_dataframe(
        ws,
        revision
    )

    # ------------------------------------------------------------
    # CONFIGURACIÓN
    # ------------------------------------------------------------

    try:

        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
        wb.calculation.calcMode = "auto"

    except Exception:
        pass

    wb.save(ruta)

    return ruta


# ================================================================
# 17. FUNCIÓN PRINCIPAL
# ================================================================

def main():
    print("=" * 75)
    print("TUKUYRICUY 2026")
    print("RANKING DE MEJORES ESTUDIANTES - 5TO SECUNDARIA")
    print("=" * 75)

    print(
        "\nCRITERIO: PROMEDIO DE LAS CUATRO ÁREAS"
    )

    print(
        "Desempate: mayor suma de porcentajes."
    )

    print(
        "No se utiliza tiempo de resolución."
    )

    print("\n[1/3] Leyendo y cruzando resultados...")

    (
        ranking,
        completos,
        incompletos,
        duplicados,
        no_identificables,
        revision
    ) = cruzar_resultados()

    print("\n[2/3] RESUMEN DEL PROCESAMIENTO")

    print(
        f"Estudiantes con las cuatro áreas: "
        f"{len(completos):,}"
    )

    print(
        f"Estudiantes sin las cuatro áreas: "
        f"{len(incompletos):,}"
    )

    print(
        f"Intentos duplicados descartados: "
        f"{len(duplicados):,}"
    )

    print(
        f"Registros sin identificación suficiente: "
        f"{len(no_identificables):,}"
    )

    print(
        f"Coincidencias ambiguas para revisar: "
        f"{len(revision):,}"
    )

    print("\n[3/3] Generando archivo Excel...")

    ruta = generar_excel(
        ranking,
        completos,
        incompletos,
        duplicados,
        no_identificables,
        revision
    )

    print("\n" + "=" * 75)
    print("PROCESO FINALIZADO")
    print("=" * 75)

    print(
        f"\nEstudiantes incluidos en el ranking: "
        f"{len(ranking):,}"
    )

    if not ranking.empty:

        print(
            f"Primer puesto: "
            f"{ranking.iloc[0]['ESTUDIANTE']}"
        )

        print(
            f"Promedio: "
            f"{ranking.iloc[0]['PROMEDIO_GENERAL']:.2f}%"
        )

        print(
            f"Suma de porcentajes: "
            f"{ranking.iloc[0]['TOTAL_PORCENTAJES']:.2f}%"
        )

    else:

        print(
            "\nNo se encontraron estudiantes con "
            "resultados en las cuatro áreas."
        )

    print(
        f"\nArchivo generado:\n{ruta}"
    )


# ================================================================
# 18. EJECUCIÓN
# ================================================================

if __name__ == "__main__":
    main()
