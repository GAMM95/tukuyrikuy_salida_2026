# TUKUYRICUY 2026 – PROCESAMIENTO Y REPORTES DE COMUNICACIÓN

## 1. Descripción del proyecto

Este módulo forma parte del proyecto **TUKUYRICUY 2026 – Evaluación de Salida** y tiene como finalidad procesar los resultados de las evaluaciones del área de Comunicación, calcular los niveles de logro de los estudiantes y generar reportes estadísticos en archivos Excel.

El programa permite sistematizar la información de los estudiantes, obtener resultados consolidados y generar reportes individuales por institución educativa, facilitando el análisis de los resultados de la evaluación.

## 2. Estructura de carpetas

La carpeta del módulo de Comunicación tiene la siguiente estructura:

```text
COMUNICACION/
│
├── resources/
│   └── RESULTADOS_COMUNICACION.xlsx
│
├── resultados/
│   ├── REPORTES_INSTITUCIONES/
│   ├── BASE_PROCESADA_COMUNICACION.xlsx
│   ├── REPORTE_GENERAL_COMUNICACION.xlsx
│   └── REPORTES_INSTITUCIONES.zip
│
├── config_comunicacion.py
├── main_comunicacion.py
├── procesamiento_comunicacion.py
├── reportes_comunicacion.py
└── README.md
```

### Descripción de los archivos Python

| Archivo | Función |
|---|---|
| `config_comunicacion.py` | Contiene la configuración general y los criterios de evaluación del área. |
| `main_comunicacion.py` | Ejecuta el proceso completo de procesamiento y generación de reportes. |
| `procesamiento_comunicacion.py` | Organiza y procesa los resultados de los estudiantes. |
| `reportes_comunicacion.py` | Genera el reporte general y los reportes individuales por institución educativa. |

## 3. Requisitos del sistema

Para ejecutar el programa se necesita:

- Sistema operativo Windows.
- Python 3.10 o superior.
- PyCharm Community o Professional, Visual Studio Code u otro editor compatible con Python.
- Microsoft Excel o un programa compatible para visualizar los reportes.

### Instalación de librerías

Abrir la terminal dentro de la carpeta `COMUNICACION` y ejecutar:

```bash
pip install -r requirements.txt
```

## 4. Archivo de entrada

El programa utiliza el archivo Excel de resultados del área de Comunicación:

**`resources/RESULTADOS_COMUNICACION.xlsx`**

La hoja de origen del archivo se denomina:

`RESULTADOS`

Este archivo contiene la información recopilada de las evaluaciones y constituye la fuente de datos para el procesamiento.

Antes de ejecutar el programa, se recomienda verificar que el archivo esté actualizado y que conserve la estructura de datos esperada por los scripts del módulo.

## 5. Criterios de evaluación

El procesamiento clasifica los resultados de los estudiantes de acuerdo con la cantidad de respuestas correctas obtenidas.

### 5.1. Niveles de logro

| Nivel de logro | Cantidad de respuestas correctas |
|---|---:|
| INICIO | De 0 a 6 |
| PROCESO | De 7 a 11 |
| LOGRADO | De 12 a 16 |
| REVISAR | Valores fuera del rango establecido o sin puntaje válido |

### 5.2. Cálculo del porcentaje

El porcentaje de respuestas correctas se calcula mediante la siguiente fórmula:

`PORCENTAJE = (CORRECTAS / TOTAL_PREGUNTAS) × 100`

Este cálculo permite conocer el rendimiento de cada estudiante en relación con el total de preguntas de la evaluación.

## 6. Ejecución del programa

Para procesar los resultados y generar los reportes, seguir los siguientes pasos:

**Paso 1.** Colocar el archivo `RESULTADOS_COMUNICACION.xlsx` dentro de la carpeta `resources`.

**Paso 2.** Abrir la terminal o consola de comandos en la carpeta `COMUNICACION`.

**Paso 3.** Ejecutar el siguiente comando:

```bash
python main_comunicacion.py
```

**Paso 4.** Esperar a que el programa realice el procesamiento de la información y genere los archivos Excel correspondientes.

Al finalizar, los resultados estarán disponibles en la carpeta `resultados/`.

## 7. Archivos de salida

Todos los resultados se almacenan dentro de la carpeta `resultados/`.

### 7.1. Base de datos procesada

**Archivo:** `BASE_PROCESADA_COMUNICACION.xlsx`

Contiene la información procesada de los estudiantes, incluyendo los resultados y cálculos generados por el programa.

### 7.2. Reporte general

**Archivo:** `REPORTE_GENERAL_COMUNICACION.xlsx`

Contiene información consolidada de los resultados del área de Comunicación.

El reporte incluye un Dashboard que permite visualizar los indicadores y resultados mediante tablas y gráficos estadísticos.

En el Dashboard, la celda **B3** permite seleccionar el ámbito de visualización:

- `TODOS`: muestra los resultados de primaria y secundaria.
- `PRIMARIA`: muestra los resultados correspondientes al nivel primaria.
- `SECUNDARIA`: muestra los resultados correspondientes al nivel secundaria.

Los indicadores y gráficos se actualizan de acuerdo con la selección realizada. En caso de que Excel no actualice inmediatamente los resultados, se recomienda habilitar el cálculo automático y actualizar las fórmulas del libro.

### 7.3. Reportes por institución educativa

**Carpeta:** `REPORTES_INSTITUCIONES/`

Contiene los archivos Excel generados individualmente para cada institución educativa, permitiendo revisar los resultados de sus estudiantes.

### 7.4. Archivo comprimido de reportes institucionales

**Archivo:** `REPORTES_INSTITUCIONES.zip`

Contiene los reportes individuales de las instituciones educativas en formato Excel, facilitando su almacenamiento y distribución.

## 8. Actualización de resultados

Cuando se incorporen nuevos registros de evaluación, se debe seguir el siguiente procedimiento:

1. Descargar los resultados actualizados en formato Excel.
2. Reemplazar el archivo `RESULTADOS_COMUNICACION.xlsx` ubicado en `resources/`.
3. Verificar que el archivo mantenga la hoja y estructura de datos esperadas.
4. Ejecutar nuevamente el programa mediante `python main_comunicacion.py`.
5. Revisar los archivos generados en la carpeta `resultados/`.

El programa trabaja con los archivos descargados y genera resultados independientes. No requiere conexión directa con Google Forms ni modifica el archivo original.

## 9. Consideraciones importantes

- Mantener los nombres de los archivos Python y sus respectivas dependencias.
- No modificar la estructura del archivo Excel de entrada sin verificar su compatibilidad con el programa.
- Revisar que los datos de los estudiantes estén correctamente registrados antes del procesamiento.
- Los registros con observaciones de calidad se conservan para su revisión; no se eliminan automáticamente.
- La clasificación de los niveles de logro se realiza utilizando la cantidad de respuestas correctas (`CORRECTAS`).
- Los porcentajes se calculan considerando el total de preguntas de la evaluación.
- Los archivos generados pueden actualizarse al ejecutar nuevamente el programa.
- Los reportes contienen información de estudiantes; deben almacenarse y compartirse únicamente con las personas autorizadas.

---

**Proyecto:** TUKUYRICUY 2026  
**Módulo:** Procesamiento y reportes de Comunicación  
**Finalidad:** Sistematización de resultados de la Evaluación de Salida 2026.