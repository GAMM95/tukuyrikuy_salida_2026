# -*- coding: utf-8 -*-
"""Punto de entrada del generador de reportes de Ciencia y Tecnología."""
from config_cyt import validar_configuracion
from procesamiento_cyt import procesar_resultados
from reportes_cyt import generar_todos_los_reportes

def main():
    print("=" * 64)
    print("TUKUYRICUY 2026 - GENERADOR DE CIENCIA Y TECNOLOGÍA")
    print("=" * 64)
    try:
        print("\\n[1/3] Validando configuración...")
        validar_configuracion()
        print("Configuración validada correctamente.")

        print("\\n[2/3] Procesando resultados...")
        df = procesar_resultados()
        print(f"Estudiantes procesados: {len(df):,}")

        print("\\n[3/3] Generando reportes...")
        generar_todos_los_reportes(df)
        print("\\nPROCESO FINALIZADO CORRECTAMENTE.")
    except Exception as e:
        print("\\nERROR DURANTE LA EJECUCIÓN")
        print(f"Detalle: {e}")
        raise

if __name__ == "__main__":
    main()
