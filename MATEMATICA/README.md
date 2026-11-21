# TUKUYRICUY 2026 – PROCESAMIENTO Y REPORTES DE MATEMÁTICA

## 1. Descripción del proyecto

Este módulo forma parte del proyecto **TUKUYRICUY 2026 – Evaluación de Salida** y tiene como finalidad procesar los resultados de las evaluaciones del área de Matemática, calcular los niveles de logro por competencia y generar reportes estadísticos en archivos Excel.

El programa permite obtener información consolidada de los estudiantes y reportes individuales por institución educativa.

## 2. Estructura de carpetas

La carpeta del módulo de Matemática tiene la siguiente estructura:

```text
MATEMATICA/
│
├── resources/
│   └── RESULTADOS_MATEMATICA.xlsx
│
├── resultados/
│   ├── REPORTES_INSTITUCIONES/
│   ├── BASE_PROCESADA_MATEMATICA.xlsx
│   ├── REPORTE_GENERAL_MATEMATICA.xlsx
│   └── REPORTES_INSTITUCIONES_MATEMATICA.zip
│
├── config_matematica.py
├── main_matematica.py
├── procesamiento_matematica.py
├── reportes_matematica.py
└── README.md
```

**Descripción de los archivos Python:**

| Archivo | Función |
|---|---|
| `config_matematica.py` | Contiene la configuración general, las competencias, la distribución de preguntas, los pesos y los criterios de evaluación. |
| `main_matematica.py` | Ejecuta el proceso completo de validación, procesamiento y generación de reportes. |
| `procesamiento_matematica.py` | Limpia, organiza y calcula los resultados de los estudiantes. |
| `reportes_matematica.py` | Genera el reporte general y los reportes individuales por institución educativa. |

## 3. Requisitos del sistema

Para ejecutar el programa se necesita:

- Sistema operativo Windows.
- Python 3.10 o superior.
- Visual Studio Code, PyCharm u otro editor compatible con Python.
- Microsoft Excel o un programa compatible para visualizar los reportes.

### Instalación de librerías

Abrir la terminal dentro de la carpeta `MATEMATICA` y ejecutar:

```bash
pip install pandas numpy openpyxl
```

## 4. Archivo de entrada

El programa utiliza el siguiente archivo Excel:

**`resources/RESULTADOS_MATEMATICA.xlsx`**

El archivo debe contener una hoja denominada exactamente:

`RESULTADOS`

Esta hoja debe incluir, como mínimo, las siguientes columnas:

| Tipo de información | Columnas |
|---|---|
| Identificación | `NIVEL`, `INSTITUCION`, `GRADO` |
| Respuestas de los estudiantes | `P1` a `P16` |
| Valoración de respuestas | `V1` a `V16` |

También se pueden incluir datos de identificación como `ID_REGISTRO`, `FECHA_HORA`, `SECCION`, `APELLIDOS`, `NOMBRES`, `DNI` y `SEXO`.

**Importante:** Las columnas `V1` a `V16` contienen la valoración de cada pregunta: `1` representa una respuesta correcta y `0` una respuesta incorrecta. Los valores vacíos o no válidos se convierten en cero durante el procesamiento, por lo que se recomienda revisar la información de origen antes de ejecutar el programa.

La columna `NIVEL` debe identificar si el registro corresponde a `PRIMARIA` o `SECUNDARIA`.

## 5. Competencias y distribución de preguntas

La evaluación de Matemática comprende 16 preguntas distribuidas en cuatro competencias. Cada competencia tiene un peso del 25 % en el resultado global.

| N.º | Competencia | Preguntas | Peso |
|---|---|---|---:|
| 1 | Resuelve problemas de cantidad | P1 – P4 | 25 % |
| 2 | Resuelve problemas de regularidad, equivalencia y cambio | P5 – P8 | 25 % |
| 3 | Resuelve problemas de forma, movimiento y localización | P9 – P12 | 25 % |
| 4 | Resuelve problemas de gestión de datos e incertidumbre | P13 – P16 | 25 % |
| | **Total** | **16 preguntas** | **100 %** |

## 6. Criterios de evaluación

### 6.1. Niveles de logro por competencia

Cada competencia comprende cuatro preguntas.

| Nivel de logro | Cantidad de respuestas correctas |
|---|---:|
| INICIO | 0 a 2 |
| PROCESO | 3 |
| LOGRADO | 4 |

### 6.2. Cálculo del puntaje ponderado

Cada competencia aporta como máximo 25 puntos al resultado global.

La fórmula utilizada es:

`Puntaje ponderado = (Aciertos / 4) × 25`

De esta manera, el puntaje máximo global es de 100 puntos.

### 6.3. Niveles de logro global

| Nivel de logro | Puntaje global |
|---|---:|
| INICIO | De 0 a 50 |
| PROCESO | Mayor de 50 hasta 75 |
| LOGRADO | Mayor de 75 hasta 100 |

Los registros con valores fuera del rango válido o que no puedan clasificarse se identifican como `REVISAR`, según corresponda.

## 7. Ejecución del programa

Para procesar los resultados y generar los reportes, seguir estos pasos:

**Paso 1.** Colocar el archivo `RESULTADOS_MATEMATICA.xlsx` dentro de la carpeta `resources`.

**Paso 2.** Abrir la terminal o consola de comandos en la carpeta `MATEMATICA`.

**Paso 3.** Ejecutar el siguiente comando:

```bash
python main_matematica.py
```

**Paso 4.** Esperar a que el programa realice las siguientes operaciones:

1. Creación de las carpetas necesarias.
2. Validación de la configuración de Matemática.
3. Lectura y procesamiento de los resultados.
4. Cálculo de aciertos, porcentajes y niveles de logro por competencia.
5. Cálculo del puntaje y nivel de logro global.
6. Generación del reporte general.
7. Generación de reportes individuales por institución educativa.
8. Compresión de los reportes institucionales en un archivo ZIP.

Al finalizar, el programa mostrará un mensaje indicando que el proceso terminó correctamente.

## 8. Archivos de salida

Todos los resultados se almacenan dentro de la carpeta `resultados/`.

### 8.1. Base de datos procesada

**Archivo:** `BASE_PROCESADA_MATEMATICA.xlsx`

Contiene los registros de los estudiantes con los resultados calculados, incluyendo:

- Datos de identificación.
- Respuestas y valoraciones de las 16 preguntas.
- Cantidad de respuestas correctas e incorrectas.
- Porcentajes de rendimiento.
- Resultados por competencia.
- Niveles de logro por competencia.
- Puntaje y nivel de logro global.

### 8.2. Reporte general

**Archivo:** `REPORTE_GENERAL_MATEMATICA.xlsx`

Contiene información consolidada del área de Matemática, organizada en las siguientes hojas:

| Hoja | Contenido |
|---|---|
| `DASHBOARD` | Indicadores generales y gráficos estadísticos. |
| `RESUMEN_COMPETENCIAS` | Resultados y niveles de logro por competencia. |
| `RESUMEN_GRADOS` | Resultados agrupados por nivel y grado. |
| `RESUMEN_INSTITUCIONES` | Resultados consolidados por institución educativa y nivel. |
| `DETALLE_ESTUDIANTES` | Información detallada de los estudiantes evaluados. |

### 8.3. Reportes por institución educativa

**Carpeta:** `REPORTES_INSTITUCIONES/`

Contiene un archivo Excel independiente para cada institución educativa, con información organizada en:

- Dashboard institucional.
- Resumen de competencias.
- Resumen por grados.
- Resumen por secciones.
- Detalle de estudiantes.

### 8.4. Archivo comprimido de reportes institucionales

**Archivo:** `REPORTES_INSTITUCIONES_MATEMATICA.zip`

Contiene los archivos Excel generados para las instituciones educativas, facilitando su distribución y almacenamiento.

## 9. Actualización de resultados

Cuando se incorporen nuevos registros de evaluación:

1. Descargar o actualizar la información de resultados en formato Excel.
2. Reemplazar el archivo `RESULTADOS_MATEMATICA.xlsx` ubicado en `resources/`.
3. Verificar que se mantengan los nombres de las columnas requeridas y la hoja `RESULTADOS`.
4. Ejecutar nuevamente `python main_matematica.py`.
5. Revisar los archivos actualizados en la carpeta `resultados/`.

El programa trabaja con el archivo Excel de entrada y genera archivos de salida independientes. No requiere conexión directa con Google Forms ni modifica el archivo original.

## 10. Consideraciones importantes

- No cambiar los nombres de los archivos Python sin actualizar las importaciones correspondientes.
- Mantener la hoja de origen con el nombre `RESULTADOS`.
- Verificar que las columnas `V1` a `V16` contengan valores válidos (`0` o `1`).
- Comprobar que los campos `NIVEL`, `INSTITUCION` y `GRADO` estén correctamente registrados.
- Revisar los datos antes de ejecutar el procesamiento para evitar resultados incorrectos.
- Los archivos de salida pueden reemplazarse al ejecutar nuevamente el programa.
- Los reportes contienen información de estudiantes; deben almacenarse y compartirse únicamente con las personas autorizadas.

---

**Proyecto:** TUKUYRICUY 2026  
**Módulo:** Procesamiento y reportes de Matemática  
**Finalidad:** Sistematización de resultados de la Evaluación de Salida 2026.