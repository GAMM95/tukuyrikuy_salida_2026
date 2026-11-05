from config_comunicacion import ARCHIVO_ENTRADA, PESTANA_ORIGEN, CARPETA_RESULTADOS
from procesamiento import cargar_y_procesar
from reportes import generar_todos_los_reportes


def main():
    print("=" * 65)
    print("TUKUYRICUY 2026 - GENERADOR DE REPORTES DE COMUNICACIÓN")
    print("=" * 65)

    if not ARCHIVO_ENTRADA.exists():
        print("\nERROR: No se encontró el archivo de entrada:")
        print(ARCHIVO_ENTRADA)
        print("\nCopia el Excel descargado de Google Sheets a la carpeta 'resources'")
        print("y nómbralo exactamente: RESULTADOS_COMUNICACION.xlsx")
        return

    print(f"\nLeyendo archivo: {ARCHIVO_ENTRADA.name}")
    df = cargar_y_procesar(ARCHIVO_ENTRADA, PESTANA_ORIGEN)

    print(f"Registros leídos: {len(df):,}")
    print(f"Instituciones: {df['INSTITUCION'].nunique():,}")
    print("\nDistribución por nivel de logro:")
    print(df["NIVEL_LOGRO"].value_counts().reindex(
        ["INICIO", "PROCESO", "LOGRADO", "REVISAR"], fill_value=0
    ).to_string())

    print("\nGenerando reportes Excel...")
    general, carpeta_ie, zip_ie, base = generar_todos_los_reportes(df)

    print("\nPROCESO FINALIZADO")
    print(f"Reporte general: {general}")
    print(f"Reportes por institución: {carpeta_ie}")
    print(f"Archivo ZIP: {zip_ie}")
    print(f"Base procesada: {base}")
    print("\nAbre el reporte general en Excel y utiliza el filtro de nivel")
    print("en la celda B3 del DASHBOARD: TODOS, PRIMARIA o SECUNDARIA.")


if __name__ == "__main__":
    main()
