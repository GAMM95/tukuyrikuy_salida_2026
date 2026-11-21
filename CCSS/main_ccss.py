# ============================================================
# TUKUYRICUY 2026
# GENERADOR DE REPORTES DE CIENCIAS SOCIALES
# ============================================================

from config_ccss import (    crear_carpetas,    validar_configuracion)
from procesamiento_ccss import procesar_datos
from reportes_ccss import generar_todos_los_reportes

def main():
    print("=" * 64)
    print("TUKUYRICUY 2026 - GENERADOR DE REPORTES DE CCSS")
    print("=" * 64)

    try:

        # 1. Crear carpetas
        crear_carpetas()

        # 2. Validar configuración
        print("\n[1/4] Validando configuración...")
        validar_configuracion()
        print("Configuración validada correctamente.")

        # 3. Procesar información
        print("\n[2/4] Procesando resultados...")
        df = procesar_datos()

        if df is None or df.empty:
            print("No existen registros para generar reportes.")
            return

        print(f"Estudiantes procesados: {len(df):,}")

        # 4. Generar reportes
        print("\n[3/4] Generando reportes Excel...")
        print("\n[4/4] Generando reportes por institución...")

        generar_todos_los_reportes(df)

        print("\n" + "=" * 64)
        print("PROCESO FINALIZADO CORRECTAMENTE")
        print("=" * 64)

    except Exception as e:

        print("\n" + "=" * 64)
        print("ERROR DURANTE LA GENERACIÓN DE REPORTES")
        print("=" * 64)

        print(f"\nDetalle: {e}")

        raise


if __name__ == "__main__":
    main()
