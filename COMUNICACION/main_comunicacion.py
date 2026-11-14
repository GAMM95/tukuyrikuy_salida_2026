# from config_comunicacion import ARCHIVO_ENTRADA, PESTANA_ORIGEN, CARPETA_RESULTADOS
# from procesamiento_comunicacion import cargar_y_procesar
# from reportes_comunicacion import generar_todos_los_reportes
#
#
# def main():
#     print("=" * 65)
#     print("TUKUYRICUY 2026 - GENERADOR DE REPORTES DE COMUNICACIÓN")
#     print("=" * 65)
#
#     if not ARCHIVO_ENTRADA.exists():
#         print("\nERROR: No se encontró el archivo de entrada:")
#         print(ARCHIVO_ENTRADA)
#         print("\nCopia el Excel descargado de Google Sheets a la carpeta 'resources'")
#         print("y nómbralo exactamente: RESULTADOS_COMUNICACION.xlsx")
#         return
#
#     print(f"\nLeyendo archivo: {ARCHIVO_ENTRADA.name}")
#     df = cargar_y_procesar(ARCHIVO_ENTRADA, PESTANA_ORIGEN)
#
#     print(f"Registros leídos: {len(df):,}")
#     print(f"Instituciones: {df['INSTITUCION'].nunique():,}")
#     print("\nDistribución por nivel de logro:")
#     print(df["NIVEL_LOGRO"].value_counts().reindex(
#         ["INICIO", "PROCESO", "LOGRADO", "REVISAR"], fill_value=0
#     ).to_string())
#
#     print("\nGenerando reportes Excel...")
#     general, carpeta_ie, zip_ie, base = generar_todos_los_reportes(df)
#
#     print("\nPROCESO FINALIZADO")
#     print(f"Reporte general: {general}")
#     print(f"Reportes por institución: {carpeta_ie}")
#     print(f"Archivo ZIP: {zip_ie}")
#     print(f"Base procesada: {base}")
#     print("\nAbre el reporte general en Excel y utiliza el filtro de nivel")
#     print("en la celda B3 del DASHBOARD: TODOS, PRIMARIA o SECUNDARIA.")
#
#
# if __name__ == "__main__":
#     main()



# ============================================================
# TUKUYRICUY 2026
# GENERADOR DE REPORTES DE COMUNICACIÓN
# ============================================================

from config_comunicacion import crear_carpetas, validar_configuracion
from procesamiento_comunicacion import procesar_datos
from reportes_comunicacion import generar_todos_los_reportes


def main():
    print("=" * 60)
    print("TUKUYRICUY 2026 - GENERADOR DE REPORTES DE COMUNICACIÓN")
    print("=" * 60)

    try:
        # 1. Crear carpetas y validar configuración
        print("\n[1/3] Validando configuración...")
        crear_carpetas()
        validar_configuracion()
        print("Configuración validada correctamente.")

        # 2. Procesar información
        print("\n[2/3] Procesando resultados...")
        df = procesar_datos()

        if df is None or df.empty:
            print("No existen registros para generar reportes.")
            return

        print(f"Estudiantes procesados: {len(df):,}")
        print(f"Instituciones: {df['INSTITUCION'].nunique():,}")

        revisar = int(df["NIVEL_LOGRO_GLOBAL"].eq("REVISAR").sum())
        if revisar:
            print(f"Registros a REVISAR (ver OBSERVACION_DATOS): {revisar}")

        # 3. Generar reportes
        print("\n[3/3] Generando reportes Excel (general e instituciones)...")
        generar_todos_los_reportes(df)

        print("\n" + "=" * 60)
        print("PROCESO FINALIZADO CORRECTAMENTE")
        print("=" * 60)

    except Exception as e:
        print("\n" + "=" * 60)
        print("ERROR DURANTE LA GENERACIÓN DE REPORTES")
        print("=" * 60)
        print(f"\nDetalle: {e}")
        raise


if __name__ == "__main__":
    main()