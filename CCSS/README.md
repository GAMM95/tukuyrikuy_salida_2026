# TUKUYRICUY 2026 – MÓDULO CCSS (SECUNDARIA)

## 1. Descripción general

El módulo **CCSS** forma parte del sistema de procesamiento de resultados de la Evaluación de Salida 2026 – TUKUYRICUY, correspondiente al área de Ciencias Sociales del nivel secundario.

Su finalidad es procesar y sistematizar las respuestas de los estudiantes de 1.º a 5.º de secundaria, comparándolas con las claves de respuestas configuradas para cada grado.

El procesamiento permite calcular los resultados por competencia y el nivel de logro global de cada estudiante, así como generar una base procesada, un reporte general consolidado y reportes individuales por institución educativa.

## 2. Estructura de archivos

El módulo está organizado de la siguiente manera:

```text
CCSS/
│
├── resources/
│   └── RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx
│
├── resultados/
│   ├── BASE_PROCESADA_CCSS_SECUNDARIA.xlsx
│   ├── REPORTE_GENERAL_CCSS_SECUNDARIA.xlsx
│   ├── REPORTES_INSTITUCIONES_CCSS_SECUNDARIA.zip
│   └── REPORTES_INSTITUCIONES/
│       └── Reportes individuales por institución
│
├── config_ccss.py
├── procesamiento_ccss.py
├── reportes_ccss.py
└── main_ccss.py
```

## 3. Descripción de los archivos Python

| Archivo | Función |
|---|---|
| `config_ccss.py` | Contiene las configuraciones del módulo, las competencias, la distribución de preguntas y los umbrales de evaluación. |
| `procesamiento_ccss.py` | Lee el archivo Excel, valida y procesa las respuestas, calcula los resultados por competencia y determina los niveles de logro. |
| `reportes_ccss.py` | Genera el reporte general consolidado y los reportes individuales por institución educativa. |
| `main_ccss.py` | Ejecuta el flujo completo de procesamiento y generación de reportes. |

## 4. Archivo de entrada

El archivo Excel de entrada es:

```text
CCSS/resources/RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx
```

**Hoja de lectura:**
- Se utiliza la hoja `RESULTADOS`, si existe.
- Si no existe, se utiliza la primera hoja del archivo.

**Columnas requeridas:**

| Columna | Descripción |
|---|---|
| `NIVEL` | Nivel educativo del estudiante. |
| `GRADO` | Grado de estudios. |
| `V1` a `V16` | Respuestas correspondientes a los 16 ítems de evaluación. |

El procesamiento considera únicamente los registros del nivel **SECUNDARIA**, correspondientes a los grados de 1.º a 5.º.

## 5. Distribución de preguntas por competencia

La evaluación comprende tres competencias del área de Ciencias Sociales:

| Competencia | Preguntas | Total de ítems |
|---|---|---:|
| Construye interpretaciones históricas | P1–P6 | 6 |
| Gestiona responsablemente el espacio y el ambiente | P7–P11 | 5 |
| Gestiona responsablemente los recursos económicos | P12–P16 | 5 |

### 5.1. Niveles de logro por competencia

| Competencia | Inicio | Proceso | Logrado |
|---|---:|---:|---:|
| Históricas | 0–2 | 3–4 | 5–6 |
| Espacio | 0–2 | 3–4 | 5 |
| Recursos económicos | 0–2 | 3 | 4–5 |

**Consideración importante sobre la competencia Espacio:**

En la indicación original se establecía un umbral de LOGRADO de hasta 6 puntos. Sin embargo, la matriz de respuestas contempla únicamente 5 preguntas para esta competencia (`COMP_2_TOTAL=5`).

Por ello, el código utiliza como máximo real 5 preguntas y establece el nivel LOGRADO en 5 puntos, evitando clasificaciones imposibles.

## 6. Ponderación y nivel de logro global

Las tres competencias tienen el mismo peso dentro del resultado global:

| Competencia | Ponderación |
|---|---:|
| Construye interpretaciones históricas | 1/3 |
| Gestiona responsablemente el espacio y el ambiente | 1/3 |
| Gestiona responsablemente los recursos económicos | 1/3 |

Los umbrales globales son provisionales y se encuentran configurados en `config_ccss.py` mediante las variables:

- `LIMITE_GLOBAL_INICIO`
- `LIMITE_GLOBAL_PROCESO`

Los valores de referencia son 50 y 75 puntos, respectivamente. Estos límites pueden modificarse desde el archivo de configuración.

## 7. Instalación de dependencias

Se requiere Python y las siguientes bibliotecas:

```bash
pip install pandas openpyxl
```

## 8. Ejecución del módulo

Ubicarse en la carpeta del módulo:

```text
D:\Python\tukuyrikuy_salida_2026\CCSS\
```

Ejecutar el archivo principal:

```bash
python main_ccss.py
```

El archivo `main_ccss.py` coordina el procesamiento de los resultados y la generación de los reportes correspondientes.

## 9. Archivos de salida

Al finalizar la ejecución se generan los siguientes productos:

| Archivo o carpeta | Descripción |
|---|---|
| `BASE_PROCESADA_CCSS_SECUNDARIA.xlsx` | Base de datos con las respuestas procesadas y los resultados obtenidos por los estudiantes. |
| `REPORTE_GENERAL_CCSS_SECUNDARIA.xlsx` | Reporte general consolidado de los resultados de secundaria. |
| `REPORTES_INSTITUCIONES_CCSS_SECUNDARIA.zip` | Archivo comprimido con los reportes individuales por institución educativa. |
| `REPORTES_INSTITUCIONES/` | Carpeta que contiene un archivo Excel por cada institución educativa. |

## 10. Consideraciones importantes

- El archivo Excel de entrada debe conservar las columnas requeridas para el procesamiento.
- Se procesan únicamente los registros correspondientes al nivel SECUNDARIA y grados del 1.º al 5.º.
- La distribución de preguntas por competencia debe coincidir con la matriz de respuestas configurada.
- Los umbrales globales pueden modificarse en `config_ccss.py`.
- Se recomienda conservar una copia de seguridad del archivo original antes de ejecutar el procesamiento.
- Al volver a ejecutar el módulo, los archivos de salida se actualizarán según los datos disponibles en el archivo de entrada.

---

**Módulo:** CCSS – Ciencias Sociales (Secundaria)  
**Proyecto:** TUKUYRICUY 2026 – Evaluación de Salida  
**Ámbito:** UGEL 02 – La Esperanza