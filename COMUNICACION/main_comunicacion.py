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