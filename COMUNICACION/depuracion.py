
# ================================================================
# TUKUYRICUY 2026
# MODULO DE DEPURACION DE RESULTADOS
#
# FUNCIONES:
# 1. Normalizar nombres y apellidos.
# 2. Detectar registros claramente invalidos.
# 3. Conservar el primer intento de cada estudiante.
# 4. Registrar los intentos duplicados.
# 5. Registrar los excluidos.
# 6. Identificar casos que requieren revision manual.
#
# ================================================================

from pathlib import Path
import re
import unicodedata
import pandas as pd


# ================================================================
# 1. CONFIGURACION
# ================================================================

CARPETA_PROYECTO = Path(__file__).resolve().parent
CARPETA_RESOURCES = CARPETA_PROYECTO / "resources"

# CAMBIAR ESTOS DATOS PARA CADA AREA
AREA = "COMUNICACION"

ARCHIVO_ORIGINAL = (
    CARPETA_RESOURCES / "RESULTADOS_COMUNICACION.xlsx"
)

ARCHIVO_SALIDA = (
    CARPETA_RESOURCES /
    "RESULTADOS_COMUNICACION_DEPURADO.xlsx"
)

# Primera hoja del archivo original.
HOJA_ORIGEN = 0

COLUMNAS_IDENTIDAD = [
    "ID_REGISTRO",
    "FECHA_HORA",
    "NIVEL",
    "INSTITUCION",
    "GRADO",
    "SECCION",
    "APELLIDOS",
    "NOMBRES",
    "DNI",
    "SEXO"
]

COLUMNAS_OBLIGATORIAS = [
    "INSTITUCION",
    "NIVEL",
    "GRADO",
    "APELLIDOS",
    "NOMBRES"
]

# Palabras que indican claramente registros de prueba,
# bromas o datos intencionalmente incomprensibles.
# La deteccion es conservadora.
PALABRAS_INVALIDAS = {
    "prueba",
    "test",
    "asdf",
    "qwerty",
    "jajaja",
    "jejeje",
    "xd",
    "broma",
    "burla",
    "incomprensible",
    "incomprensibles",
    "noentiendo",
    "noseentiende",
    "anonimo",
    "anonima",
    "sinNombre",
    "xxx",
    "xxxxx"
}

PARTICULAS = {
    "de", "del", "la", "las", "los", "y"
}


# ================================================================
# 2. NORMALIZACION DE TEXTO
# ================================================================

def limpiar_espacios(valor):

    if pd.isna(valor):
        return ""

    texto = str(valor)

    texto = texto.replace("\xa0", " ")
    texto = texto.replace("\t", " ")

    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


def convertir_formato_nombre(valor):

    texto = limpiar_espacios(valor)

    if not texto:
        return ""

    palabras = texto.lower().split()

    resultado = []

    for palabra in palabras:

        partes = palabra.split("-")

        partes = [
            parte.capitalize()
            if parte.lower() not in PARTICULAS
            else parte.lower()
            for parte in partes
        ]

        resultado.append("-".join(partes))

    return " ".join(resultado)


def normalizar_clave(valor):

    texto = limpiar_espacios(valor).lower()

    texto = unicodedata.normalize("NFD", texto)

    texto = "".join(
        caracter
        for caracter in texto
        if unicodedata.category(caracter) != "Mn"
    )

    texto = re.sub(r"[^a-z0-9 ]", " ", texto)

    texto = re.sub(r"\s+", " ", texto)

    return texto.strip()


# ================================================================
# 3. VALIDACION DE DNI
# ================================================================

def dni_valido(valor):

    if pd.isna(valor):
        return False

    texto = str(valor).strip()

    # Excel puede guardar DNI como numero decimal.
    if texto.endswith(".0"):
        texto = texto[:-2]

    return (
        texto.isdigit()
        and len(texto) == 8
        and texto != "00000000"
    )


# ================================================================
# 4. DETECCION DE REGISTROS NO VALIDOS
# ================================================================


def detectar_nombre_invalido(apellidos, nombres):

    apellidos_clave = normalizar_clave(apellidos)
    nombres_clave = normalizar_clave(nombres)

    if not apellidos_clave and not nombres_clave:
        return "APELLIDOS Y NOMBRES VACIOS"

    if not apellidos_clave:
        return "APELLIDOS VACIOS"

    if not nombres_clave:
        return "NOMBRES VACIOS"

    # Una sola letra o inicial no permite identificar
    # adecuadamente al estudiante.
    if (
        len(apellidos_clave.replace(" ", "")) <= 1
        or len(nombres_clave.replace(" ", "")) <= 1
    ):
        return "DATOS INCOMPLETOS: NOMBRE O APELLIDO INSUFICIENTE"

    # Ambos campos exactamente iguales.
    if apellidos_clave == nombres_clave:
        return (
            "APELLIDOS Y NOMBRES IDENTICOS: "
            "IDENTIFICACION INCOMPLETA"
        )

    palabras_invalidas = {
        "asdf", "qwerty", "jajaja",
        "jejeje", "broma", "burla",
        "xxx", "xxxxx"
    }

    palabras = set(
        (apellidos_clave + " " + nombres_clave).split()
    )

    if palabras.intersection(palabras_invalidas):
        return "REGISTRO DE PRUEBA O BURLA"

    return ""



# ================================================================
# 5. CORRECCION DE NOMBRES Y APELLIDOS
# ================================================================

def distribuir_nombre_completo(apellidos, nombres):

    apellidos_originales = apellidos
    nombres_originales = nombres

    apellidos = convertir_formato_nombre(apellidos)
    nombres = convertir_formato_nombre(nombres)

    observacion = ""
    revision = False

    apellidos_clave = normalizar_clave(apellidos)
    nombres_clave = normalizar_clave(nombres)

    # ------------------------------------------------------------
    # CASO 1:
    # Ambos campos contienen exactamente el mismo nombre completo.
    # Se aplica una distribucion estimada y se marca para verificar.
    # ------------------------------------------------------------

    if (
        apellidos_clave
        and nombres_clave
        and apellidos_clave == nombres_clave
    ):

        palabras = apellidos.split()

        if len(palabras) >= 3:

            cantidad_nombres = 2 if len(palabras) >= 4 else 1

            nombres = " ".join(
                palabras[:cantidad_nombres]
            )

            apellidos = " ".join(
                palabras[cantidad_nombres:]
            )

            observacion = (
                "DISTRIBUCION ESTIMADA DE NOMBRE COMPLETO. "
                "VERIFICAR CON NOMINA OFICIAL."
            )

            revision = True

        else:

            observacion = (
                "APELLIDOS Y NOMBRES REPETIDOS. "
                "REQUIERE REVISION MANUAL."
            )

            revision = True

    # ------------------------------------------------------------
    # CASO 2:
    # APELLIDOS esta correctamente registrado y NOMBRES contiene
    # tambien los apellidos al final.
    #
    # Ejemplo:
    # APELLIDOS: Julca Laiza
    # NOMBRES: Alexander Rodrigo Julca Laiza
    # ------------------------------------------------------------

    elif (
        apellidos_clave
        and nombres_clave
        and nombres_clave.endswith(apellidos_clave)
    ):

        palabras_nombres = nombres.split()
        palabras_apellidos = apellidos.split()

        if len(palabras_nombres) > len(palabras_apellidos):

            nombres = " ".join(
                palabras_nombres[
                    :len(palabras_nombres) - len(palabras_apellidos)
                ]
            )

            observacion = (
                "SE RETIRARON APELLIDOS REPETIDOS "
                "DEL CAMPO NOMBRES."
            )

    # ------------------------------------------------------------
    # CASO 3:
    # Solo se registro el nombre completo en APELLIDOS.
    # ------------------------------------------------------------

    elif apellidos_clave and not nombres_clave:

        palabras = apellidos.split()

        if len(palabras) >= 3:

            cantidad_nombres = 2 if len(palabras) >= 4 else 1

            nombres = " ".join(
                palabras[:cantidad_nombres]
            )

            apellidos = " ".join(
                palabras[cantidad_nombres:]
            )

            observacion = (
                "DISTRIBUCION ESTIMADA DESDE APELLIDOS. "
                "VERIFICAR CON NOMINA OFICIAL."
            )

            revision = True

    # ------------------------------------------------------------
    # CASO 4:
    # Los dos campos tienen informacion, pero no es posible
    # identificar claramente la distribucion.
    # ------------------------------------------------------------

    # Se conserva la informacion original normalizada.

    return {
        "APELLIDOS": apellidos,
        "NOMBRES": nombres,
        "APELLIDOS_ORIGINAL": apellidos_originales,
        "NOMBRES_ORIGINAL": nombres_originales,
        "OBSERVACION_NOMBRES": observacion,
        "REQUIERE_REVISION_NOMBRES": revision
    }


# ================================================================
# 6. CREACION DE CLAVE DE IDENTIFICACION
# ================================================================

def crear_clave_estudiante(fila):

    dni = limpiar_espacios(fila.get("DNI", ""))

    if dni.endswith(".0"):
        dni = dni[:-2]

    if dni_valido(dni):

        return "DNI|" + dni

    # Cuando no existe DNI valido, se utiliza una combinacion
    # de datos del estudiante y su contexto escolar.

    campos = [
        "INSTITUCION",
        "NIVEL",
        "GRADO",
        "SECCION",
        "APELLIDOS",
        "NOMBRES"
    ]

    valores = []

    for columna in campos:

        valor = normalizar_clave(
            fila.get(columna, "")
        )

        valores.append(valor)

    if not valores[-1] or not valores[-2]:

        return ""

    if not valores[0]:

        return ""

    return "SIN_DNI|" + "|".join(valores)


# ================================================================
# 7. REGISTRO DE OBSERVACIONES
# ================================================================

def agregar_observacion(actual, nueva):

    actual = limpiar_espacios(actual)
    nueva = limpiar_espacios(nueva)

    if not nueva:
        return actual

    if not actual:
        return nueva

    return actual + " | " + nueva


# ================================================================
# 8. PROCESAMIENTO PRINCIPAL
# ================================================================

def procesar_depuracion():

    print("=" * 70)
    print("TUKUYRICUY 2026 - DEPURACION DE RESULTADOS")
    print("AREA:", AREA)
    print("=" * 70)

    # ------------------------------------------------------------
    # VALIDAR ARCHIVO
    # ------------------------------------------------------------

    if not ARCHIVO_ORIGINAL.exists():

        raise FileNotFoundError(
            f"No se encontro el archivo original:\n{ARCHIVO_ORIGINAL}"
        )

    CARPETA_RESOURCES.mkdir(
        parents=True,
        exist_ok=True
    )

    # ------------------------------------------------------------
    # LEER ORIGINAL
    # ------------------------------------------------------------

    print("\n[1/6] Leyendo archivo original...")

    df = pd.read_excel(
        ARCHIVO_ORIGINAL,
        sheet_name=HOJA_ORIGEN,
        dtype=object
    )

    df.columns = [
        str(columna).strip().upper()
        for columna in df.columns
    ]

    faltantes = [
        columna
        for columna in COLUMNAS_OBLIGATORIAS
        if columna not in df.columns
    ]

    if faltantes:

        raise ValueError(
            "Faltan columnas obligatorias: "
            + ", ".join(faltantes)
        )

    total_original = len(df)

    print("Registros originales:", total_original)

    # Guardar el orden original.
    df["_ORDEN_ORIGINAL"] = range(len(df))

    # ------------------------------------------------------------
    # NORMALIZAR COLUMNAS DE IDENTIDAD
    # ------------------------------------------------------------

    print("\n[2/6] Normalizando nombres y apellidos...")

    for columna in [
        "INSTITUCION",
        "NIVEL",
        "GRADO",
        "SECCION",
        "DNI",
        "SEXO"
    ]:

        if columna in df.columns:

            df[columna] = df[columna].apply(
                limpiar_espacios
            )

    # Normalizar fecha.
    df["_FECHA_ORDEN"] = pd.to_datetime(
        df["FECHA_HORA"],
        dayfirst=True,
        errors="coerce"
    ) if "FECHA_HORA" in df.columns else pd.NaT

    # Aplicar correccion de nombres.
    resultados_nombres = df.apply(
        lambda fila: distribuir_nombre_completo(
            fila.get("APELLIDOS", ""),
            fila.get("NOMBRES", "")
        ),
        axis=1
    )

    df["APELLIDOS_ORIGINAL"] = [
        resultado["APELLIDOS_ORIGINAL"]
        for resultado in resultados_nombres
    ]

    df["NOMBRES_ORIGINAL"] = [
        resultado["NOMBRES_ORIGINAL"]
        for resultado in resultados_nombres
    ]

    df["APELLIDOS"] = [
        resultado["APELLIDOS"]
        for resultado in resultados_nombres
    ]

    df["NOMBRES"] = [
        resultado["NOMBRES"]
        for resultado in resultados_nombres
    ]

    df["OBSERVACION_DEPURACION"] = [
        resultado["OBSERVACION_NOMBRES"]
        for resultado in resultados_nombres
    ]

    df["REQUIERE_REVISION_NOMBRES"] = [
        resultado["REQUIERE_REVISION_NOMBRES"]
        for resultado in resultados_nombres
    ]

    # ------------------------------------------------------------
    # IDENTIFICAR EXCLUIDOS
    # ------------------------------------------------------------

    print("\n[3/6] Identificando registros no validos...")

    motivos = []

    for _, fila in df.iterrows():

        motivo = detectar_nombre_invalido(
            fila.get("APELLIDOS", ""),
            fila.get("NOMBRES", "")
        )

        motivos.append(motivo)

    df["_MOTIVO_EXCLUSION"] = motivos

    excluidos = df[
        df["_MOTIVO_EXCLUSION"] != ""
    ].copy()

    validos = df[
        df["_MOTIVO_EXCLUSION"] == ""
    ].copy()

    excluidos["MOTIVO_EXCLUSION"] = (
        excluidos["_MOTIVO_EXCLUSION"]
    )

    excluidos["ESTADO_DEPURACION"] = "EXCLUIDO"

    # ------------------------------------------------------------
    # IDENTIFICAR DUPLICADOS
    # ------------------------------------------------------------

    print("\n[4/6] Identificando intentos duplicados...")

    validos["_CLAVE_ESTUDIANTE"] = validos.apply(
        crear_clave_estudiante,
        axis=1
    )

    # Los registros sin clave suficiente no se eliminan.
    # Se conservan, pero se envian a revision.

    sin_clave = validos[
        validos["_CLAVE_ESTUDIANTE"] == ""
    ].copy()

    validos.loc[
        validos["_CLAVE_ESTUDIANTE"] == "",
        "OBSERVACION_DEPURACION"
    ] = validos.loc[
        validos["_CLAVE_ESTUDIANTE"] == "",
        "OBSERVACION_DEPURACION"
    ].apply(
        lambda x: agregar_observacion(
            x,
            "NO SE PUDO GENERAR CLAVE DE IDENTIFICACION. REVISAR."
        )
    )

    validos.loc[
        validos["_CLAVE_ESTUDIANTE"] == "",
        "REQUIERE_REVISION_NOMBRES"
    ] = True

    con_clave = validos[
        validos["_CLAVE_ESTUDIANTE"] != ""
    ].copy()

    # Ordenar por estudiante, fecha y orden original.
    con_clave = con_clave.sort_values(
        by=[
            "_CLAVE_ESTUDIANTE",
            "_FECHA_ORDEN",
            "_ORDEN_ORIGINAL"
        ],
        na_position="last",
        kind="stable"
    )

    primeros = []
    duplicados = []

    for clave, grupo in con_clave.groupby(
        "_CLAVE_ESTUDIANTE",
        sort=False
    ):

        grupo = grupo.sort_values(
            by=[
                "_FECHA_ORDEN",
                "_ORDEN_ORIGINAL"
            ],
            na_position="last",
            kind="stable"
        )

        total_intentos = len(grupo)

        primer_registro = grupo.iloc[0].copy()

        id_conservado = primer_registro.get(
            "ID_REGISTRO",
            ""
        )

        fecha_conservada = primer_registro.get(
            "FECHA_HORA",
            ""
        )

        if total_intentos > 1:

            primer_registro[
                "OBSERVACION_DEPURACION"
            ] = agregar_observacion(
                primer_registro["OBSERVACION_DEPURACION"],
                f"TIENE {total_intentos} INTENTOS. "
                "SE CONSERVA EL PRIMERO."
            )

        primer_registro["TOTAL_INTENTOS"] = total_intentos

        primeros.append(primer_registro)

        # Los siguientes intentos pasan a la hoja DUPLICADOS.
        if total_intentos > 1:

            posteriores = grupo.iloc[1:].copy()

            posteriores["ID_REGISTRO_CONSERVADO"] = (
                id_conservado
            )

            posteriores["FECHA_HORA_CONSERVADA"] = (
                fecha_conservada
            )

            posteriores["TOTAL_INTENTOS"] = total_intentos

            posteriores["ESTADO_DEPURACION"] = (
                "DUPLICADO - NO CONSIDERADO"
            )

            posteriores["OBSERVACION_DEPURACION"] = (
                "INTENTO POSTERIOR. "
                "SE CONSERVA EL PRIMER REGISTRO."
            )

            posteriores["_NUMERO_INTENTO"] = range(
                2,
                total_intentos + 1
            )

            duplicados.append(posteriores)

    # Incorporar registros sin clave.
    for _, fila in sin_clave.iterrows():

        fila = fila.copy()

        fila["TOTAL_INTENTOS"] = 1

        primeros.append(fila)

    # ------------------------------------------------------------
    # CONSTRUIR BASE DEPURADA
    # ------------------------------------------------------------

    if primeros:

        depurados = pd.DataFrame(primeros)

    else:

        depurados = validos.iloc[0:0].copy()

    depurados["ESTADO_DEPURACION"] = "CONSERVADO"

    depurados = depurados.sort_values(
        "_ORDEN_ORIGINAL",
        kind="stable"
    )

    # ------------------------------------------------------------
    # CONSTRUIR BASE DE DUPLICADOS
    # ------------------------------------------------------------

    if duplicados:

        duplicados_df = pd.concat(
            duplicados,
            ignore_index=True
        )

        duplicados_df = duplicados_df.sort_values(
            "_ORDEN_ORIGINAL",
            kind="stable"
        )

    else:

        duplicados_df = pd.DataFrame(
            columns=list(df.columns) + [
                "ID_REGISTRO_CONSERVADO",
                "FECHA_HORA_CONSERVADA",
                "_NUMERO_INTENTO",
                "TOTAL_INTENTOS",
                "ESTADO_DEPURACION"
            ]
        )

    # ------------------------------------------------------------
    # CONSTRUIR REVISION MANUAL
    # ------------------------------------------------------------

    revision = depurados[
        depurados["REQUIERE_REVISION_NOMBRES"] == True
    ].copy()

    # ------------------------------------------------------------
    # LIMPIAR COLUMNAS INTERNAS
    # ------------------------------------------------------------

    columnas_internas = [
        "_ORDEN_ORIGINAL",
        "_FECHA_ORDEN",
        "_CLAVE_ESTUDIANTE",
        "_MOTIVO_EXCLUSION"
    ]

    def limpiar_salida(tabla):

        columnas_eliminar = [
            columna
            for columna in columnas_internas
            if columna in tabla.columns
        ]

        return tabla.drop(
            columns=columnas_eliminar,
            errors="ignore"
        )

    depurados = limpiar_salida(depurados)
    duplicados_df = limpiar_salida(duplicados_df)
    excluidos = limpiar_salida(excluidos)
    revision = limpiar_salida(revision)

    # ------------------------------------------------------------
    # EXPORTAR ARCHIVO
    # ------------------------------------------------------------

    print("\n[5/6] Generando archivo depurado...")

    with pd.ExcelWriter(
        ARCHIVO_SALIDA,
        engine="openpyxl"
    ) as writer:

        depurados.to_excel(
            writer,
            sheet_name="RESULTADOS_DEPURADOS",
            index=False
        )

        duplicados_df.to_excel(
            writer,
            sheet_name="DUPLICADOS",
            index=False
        )

        excluidos.to_excel(
            writer,
            sheet_name="EXCLUIDOS",
            index=False
        )

        revision.to_excel(
            writer,
            sheet_name="REVISION_MANUAL",
            index=False
        )

        # Formato basico de las hojas.
        for hoja in writer.book.worksheets:

            hoja.freeze_panes = "A2"

            hoja.auto_filter.ref = hoja.dimensions

            for columna in hoja.columns:

                letra = columna[0].column_letter

                longitud = max(
                    len(str(celda.value or ""))
                    for celda in list(columna)[:100]
                )

                hoja.column_dimensions[letra].width = min(
                    max(longitud + 2, 12),
                    40
                )

    # ------------------------------------------------------------
    # RESUMEN
    # ------------------------------------------------------------

    print("\n[6/6] Resumen de depuracion")

    total_validos = len(depurados)
    total_duplicados = len(duplicados_df)
    total_excluidos = len(excluidos)
    total_revision = len(revision)

    print("-" * 70)

    print("REGISTROS ORIGINALES:", total_original)
    print("REGISTROS CONSERVADOS:", total_validos)
    print("INTENTOS DUPLICADOS:", total_duplicados)
    print("REGISTROS EXCLUIDOS:", total_excluidos)
    print("CASOS PARA REVISION:", total_revision)

    print("-" * 70)

    print("\nArchivo generado correctamente:")
    print(ARCHIVO_SALIDA)

    print("\nDepuracion finalizada.")


# ================================================================
# 9. EJECUCION
# ================================================================

if __name__ == "__main__":

    procesar_depuracion()
