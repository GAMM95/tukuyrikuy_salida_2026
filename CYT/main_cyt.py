
# ============================================================
# TUKUYRICUY 2026
# GENERADOR DE REPORTES DE CIENCIA Y TECNOLOGÍA
# ============================================================

from config_cyt import (
    crear_carpetas,
    validar_configuracion
)

from procesamiento_cyt import procesar_datos

from reportes_cyt import generar_todos_los_reportes


def mostrar_diagnostico(df):

    print("\n" + "=" * 65)
    print("DIAGNÓSTICO DE ESTUDIANTES POR NIVEL Y GRADO")
    print("=" * 65)

    if df is None or df.empty:
        print("No existen registros para analizar.")
        return

    if "NIVEL" not in df.columns or "GRADO" not in df.columns:
        print("ERROR: No se encontraron las columnas NIVEL y GRADO.")
        print("Columnas disponibles:")
        print(list(df.columns))
        return

    # Normalizar temporalmente los datos para el diagnóstico
    datos = df.copy()

    datos["NIVEL"] = (
        datos["NIVEL"]
        .astype(str)
        .str.strip()
        .str.upper()
    )

    datos["GRADO"] = (
        datos["GRADO"]
        .astype(str)
        .str.strip()
    )

    # Cantidad de estudiantes por nivel y grado
    diagnostico = (
        datos.groupby(["NIVEL", "GRADO"], dropna=False)
        .size()
        .reset_index(name="ESTUDIANTES")
    )

    print("\nESTUDIANTES ENCONTRADOS:")
    print("-" * 65)
    print(diagnostico.to_string(index=False))

    # Verificar grados de primaria
    print("\nPRIMARIA:")
    print("-" * 65)

    for grado in range(2, 7):

        cantidad = len(
            datos[
                (datos["NIVEL"] == "PRIMARIA") &
                (datos["GRADO"] == str(grado))
            ]
        )

        print(f"{grado}° de primaria: {cantidad:,} estudiantes")

    # Verificar grados de secundaria
    print("\nSECUNDARIA:")
    print("-" * 65)

    for grado in range(1, 6):

        cantidad = len(
            datos[
                (datos["NIVEL"] == "SECUNDARIA") &
                (datos["GRADO"] == str(grado))
            ]
        )

        estado = "OK" if cantidad > 0 else "SIN REGISTROS"

        print(
            f"{grado}° de secundaria: "
            f"{cantidad:,} estudiantes | {estado}"
        )

    print("\n" + "=" * 65)
    print("FIN DEL DIAGNÓSTICO")
    print("=" * 65)


def main():

    print("=" * 65)
    print("TUKUYRICUY 2026 - GENERADOR DE REPORTES DE CIENCIA Y TECNOLOGÍA")
    print("=" * 65)

    try:

        # ====================================================
        # 1. CREAR CARPETAS
        # ====================================================

        print("\nPreparando carpetas del proyecto...")
        crear_carpetas()

        # ====================================================
        # 2. VALIDAR CONFIGURACIÓN
        # ====================================================

        print("\n[1/4] Validando configuración...")

        validar_configuracion()

        print("Configuración validada correctamente.")

        # ====================================================
        # 3. PROCESAR RESULTADOS
        # ====================================================

        print("\n[2/4] Procesando resultados...")

        df = procesar_datos()

        if df is None or df.empty:

            print("\nNo existen registros para generar reportes.")
            return

        print(f"\nEstudiantes procesados: {len(df):,}")

        # 4. DIAGNÓSTICO POR NIVEL Y GRADO
        mostrar_diagnostico(df)

        # 5. GENERAR REPORTES
        print("\n[3/4] Generando reporte general...")
        print("[4/4] Generando reportes por institución...")

        generar_todos_los_reportes(df)

        # FINALIZACIÓN
        print("\n" + "=" * 65)
        print("PROCESO FINALIZADO CORRECTAMENTE")
        print("=" * 65)

        print(f"\nTotal de estudiantes procesados: {len(df):,}")

    except Exception as e:

        print("\n" + "=" * 65)
        print("ERROR DURANTE LA GENERACIÓN DE REPORTES")
        print("=" * 65)

        print(f"\nDetalle: {e}")

        raise


if __name__ == "__main__":
    main()
