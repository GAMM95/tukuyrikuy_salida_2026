from config_personal_social import crear_carpetas, validar_configuracion
from procesamiento_personal_social import procesar_datos
from reportes_personal_social import generar_todos_los_reportes

def main():
    print("=" * 64)
    print("TUKUYRICUY 2026 - GENERADOR DE PERSONAL SOCIAL (PRIMARIA)")
    print("=" * 64)

    try:
        crear_carpetas()
        print("\n[1/3] Validando configuración...")
        validar_configuracion()
        print("Configuración validada correctamente.")

        print("\n[2/3] Procesando resultados...")
        df = procesar_datos()
        print(f"Estudiantes procesados: {len(df):,}")

        print("\n[3/3] Generando reportes...")
        generar_todos_los_reportes(df)

        print("\n" + "=" * 64)
        print("PROCESO FINALIZADO CORRECTAMENTE")
        print("=" * 64)
    except Exception as error:
        print("\n" + "=" * 64)
        print("ERROR DURANTE LA GENERACIÓN DE REPORTES")
        print("=" * 64)
        print(f"Detalle: {error}")
        raise


if __name__ == "__main__":
    main()
