TUKUYRICUY 2026 - COMUNICACIÓN
================================

REQUISITOS
- Windows
- Python 3.10 o superior
- PyCharm Community o Professional
- Microsoft Excel o compatible para ver los reportes

1. ABRIR EL PROYECTO
   Descomprime la carpeta y ábrela desde PyCharm.

2. INSTALAR LIBRERÍAS
   Abre la terminal de PyCharm y ejecuta:
       pip install -r requirements.txt

3. COLOCAR EL EXCEL DE ORIGEN
   Descarga el archivo de Google Sheets como Microsoft Excel (.xlsx).
   Coloca el archivo dentro de la carpeta:
       datos
   El nombre debe ser exactamente:
       RESULTADOS_COMUNICACION.xlsx

   La pestaña de origen debe llamarse:
       RESULTADOS

4. EJECUTAR
   En la terminal:
       python main_comunicacion.py

   También puedes abrir main_comunicacion.py y pulsar Run.

5. RESULTADOS
   Se crearán:
       resultados/REPORTE_GENERAL_COMUNICACION.xlsx
       resultados/REPORTES_INSTITUCIONES/
       resultados/REPORTES_INSTITUCIONES.zip
       resultados/BASE_PROCESADA_COMUNICACION.xlsx

CRITERIOS DE COMUNICACIÓN
- INICIO: 0 a 6 aciertos (<= 6)
- PROCESO: 7 a 11 aciertos (<= 11)
- LOGRADO: 12 a 16 aciertos (<= 16)
- Fuera de rango o sin puntaje: REVISAR

DASHBOARD
En el Excel general, hoja DASHBOARD, selecciona la celda B3:
TODOS, PRIMARIA o SECUNDARIA.
Las fórmulas se recalculan al abrir el archivo en Excel.

IMPORTANTE
- El programa no se conecta a Google ni modifica el documento original.
- Cada vez que tengas nuevas respuestas, vuelve a descargar el Excel,
  reemplaza datos/RESULTADOS_COMUNICACION.xlsx y ejecuta el programa.
- No compartas públicamente archivos con datos personales de estudiantes.
- El programa conserva registros con observaciones de calidad; no los elimina.
- La clasificación usa CORRECTAS. El porcentaje del reporte se calcula
  como CORRECTAS / TOTAL_PREGUNTAS para mantener coherencia.
