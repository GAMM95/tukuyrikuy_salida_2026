# TUKUYRICUY 2026
## MÓDULO DE CIENCIA Y TECNOLOGÍA (CYT)
### Generador de reportes y resultados

## 1. DESCRIPCIÓN

Este módulo permite procesar y sistematizar los resultados de la **Evaluación de Salida 2026 – TUKUYRICUY**, correspondiente al área de Ciencia y Tecnología de los niveles de primaria y secundaria.

El sistema procesa las respuestas de los estudiantes, las compara con las claves de respuestas establecidas y calcula los niveles de logro por competencia y el resultado global.

Asimismo, genera reportes consolidados y reportes individuales por institución educativa, facilitando el análisis de los resultados obtenidos.

## 2. ESTRUCTURA DE ARCHIVOS

El módulo está organizado de la siguiente manera:

```text
CYT/
│
├── resources/
│   └── RESULTADOS_CIENCIA_TECNOLOGIA.xlsx
│
├── resultados/
│   ├── REPORTES_CYT/
│   │   ├── REPORTE_GENERAL_CYT.xlsx
│   │   └── REPORTES_CYT_POR_INSTITUCION.zip
│   │
│   └── BASE_PROCESADA_CYT.xlsx
│
├── config_cyt.py
├── procesamiento_cyt.py
├── reportes_cyt.py
└── main_cyt.py
```

**Descripción de los archivos Python:**

- `config_cyt.py`: configuración de rutas, grados, competencias, ponderaciones y niveles de logro.
- `procesamiento_cyt.py`: lectura del archivo Excel, validación de respuestas, procesamiento de datos y cálculo de resultados.
- `reportes_cyt.py`: generación del reporte general consolidado y reportes individuales por institución educativa.
- `main_cyt.py`: archivo principal que ejecuta todo el proceso.

## 3. REQUISITOS E INSTALACIÓN

### 3.1. Ubicación del módulo

La carpeta `CYT` debe encontrarse dentro del directorio principal del proyecto:

```text
D:\Python\tukuyrikuy_salida_2026\
```

### 3.2. Archivo de entrada

Colocar el archivo Excel:

```text
RESULTADOS_CIENCIA_TECNOLOGIA.xlsx
```

Dentro de la carpeta:

```text
CYT/resources/
```

### 3.3. Instalación de dependencias

Activar el entorno virtual del proyecto e instalar las librerías necesarias:

```bash
pip install pandas openpyxl numpy
```

### 3.4. Ejecución del sistema

Desde el directorio principal del proyecto, ejecutar:

```bash
python CYT/main_cyt.py
```

## 4. GRADOS Y NIVELES EVALUADOS

El módulo contempla los siguientes grados:

**Educación primaria**
- 2.º grado
- 3.º grado
- 4.º grado
- 5.º grado
- 6.º grado

**Educación secundaria**
- 1.º grado
- 2.º grado
- 3.º grado
- 4.º grado
- 5.º grado

Cada evaluación contempla 16 preguntas, con excepción de las adecuaciones que puedan establecerse en la configuración correspondiente.

## 5. COMPETENCIAS Y DISTRIBUCIÓN DE PREGUNTAS

La evaluación comprende tres competencias del área de Ciencia y Tecnología.

### Competencia 1: Indaga mediante métodos científicos para construir sus conocimientos

- Primaria: preguntas 1 al 6.
- Secundaria: distribución según la clave configurada.

**Niveles de logro:**

| Nivel | Cantidad de aciertos |
|---|---:|
| Inicio | 0 a 2 |
| Proceso | 3 a 4 |
| Logrado | 5 a 6 |

### Competencia 2: Explica el mundo físico basándose en conocimientos sobre los seres vivos, materia y energía, biodiversidad, Tierra y universo

- Primaria: preguntas 7 al 12.
- Secundaria: distribución según la clave configurada.

**Niveles de logro:**

| Nivel | Cantidad de aciertos |
|---|---:|
| Inicio | 0 a 2 |
| Proceso | 3 a 4 |
| Logrado | 5 a 6 |

### Competencia 3: Diseña y construye soluciones tecnológicas para resolver problemas de su entorno

- Primaria: preguntas 13 al 16.
- Secundaria: distribución según la clave configurada.

**Niveles de logro:**

| Nivel | Cantidad de aciertos |
|---|---:|
| Inicio | 0 a 2 |
| Proceso | 3 |
| Logrado | 4 |

### Distribución especial de 1.º de secundaria

Para 1.º de secundaria se aplica la distribución establecida en el referente de UGEL:

| Competencia | Preguntas | Cantidad |
|---|---|---:|
| Indaga | 1, 2, 3, 5, 7 y 9 | 6 |
| Explica | 4, 6, 8, 14, 15 y 16 | 6 |
| Diseña | 10, 11, 12 y 13 | 4 |
| **Total** | **1 al 16** | **16** |

Esta distribución especial se encuentra contemplada en la configuración del módulo.

## 6. PONDERACIÓN Y RESULTADO GLOBAL

Las tres competencias tienen la misma ponderación:

| Competencia | Ponderación |
|---|---:|
| Indaga | 33.33 % |
| Explica | 33.33 % |
| Diseña | 33.33 % |

El resultado global se determina considerando el promedio de los porcentajes de logro obtenidos en las tres competencias.

### Niveles de logro global

| Nivel | Rango porcentual |
|---|---|
| Inicio | Menor o igual al 50 % |
| Proceso | Mayor al 50 % y menor o igual al 75 % |
| Logrado | Mayor al 75 % |

Los umbrales pueden modificarse desde `config_cyt.py`, según los criterios establecidos para la evaluación.

## 7. ARCHIVOS GENERADOS

Al ejecutar el módulo se generan los siguientes archivos:

### 7.1. Base procesada

**Archivo:** `BASE_PROCESADA_CYT.xlsx`

**Ubicación:**

```text
CYT/resultados/BASE_PROCESADA_CYT.xlsx
```

Contiene los resultados individuales de los estudiantes, las respuestas procesadas, la cantidad de aciertos, los resultados por competencia y los niveles de logro.

### 7.2. Reporte general consolidado

**Archivo:** `REPORTE_GENERAL_CYT.xlsx`

**Ubicación:**

```text
CYT/resultados/REPORTES_CYT/REPORTE_GENERAL_CYT.xlsx
```

Contiene los resultados consolidados del área de Ciencia y Tecnología para facilitar el análisis de los resultados obtenidos.

### 7.3. Reportes por institución educativa

**Archivo comprimido:** `REPORTES_CYT_POR_INSTITUCION.zip`

**Ubicación:**

```text
CYT/resultados/REPORTES_CYT/REPORTES_CYT_POR_INSTITUCION.zip
```

Contiene los reportes individuales de las instituciones educativas, organizados en archivos Excel.

Además, se genera una carpeta con los archivos correspondientes a cada institución educativa.

## 8. CONSIDERACIONES IMPORTANTES

- El archivo de entrada debe conservar la estructura esperada por el módulo.
- Las claves de respuestas deben estar correctamente configuradas.
- Los nombres de las instituciones educativas deben estar correctamente registrados para evitar duplicidades.
- Se recomienda mantener una copia de seguridad de los archivos originales.
- Los archivos generados se actualizan al ejecutar nuevamente el proceso.
- No se deben modificar manualmente los archivos de resultados mientras se ejecuta el procesamiento.
- La distribución de preguntas y los criterios de evaluación deben mantenerse de acuerdo con la configuración establecida.

---

**TUKUYRICUY 2026**  
**Sistematización de resultados y generación de reportes**  
**ÁREA: CIENCIA Y TECNOLOGÍA**