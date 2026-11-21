# TUKUYRICUY 2026 – SISTEMA DE PROCESAMIENTO DE EVALUACIONES

## 1. Descripción general

**TUKUYRICUY 2026** es un sistema desarrollado en Python para el procesamiento, sistematización y generación de reportes de resultados de la Evaluación de Salida 2026, correspondiente a las instituciones educativas del ámbito de la **UGEL 02 – La Esperanza**.

El sistema permite procesar las respuestas de los estudiantes, contrastarlas con las claves de respuestas establecidas para cada grado y área, calcular los resultados por competencia y determinar los niveles de logro.

Asimismo, genera bases de datos procesadas, reportes generales consolidados y reportes individuales por institución educativa, facilitando el análisis de los resultados de aprendizaje.

El proyecto está organizado en módulos independientes, de acuerdo con las áreas evaluadas.

## 2. Módulos del sistema

El proyecto comprende cinco módulos:

| N.º | Módulo | Nivel educativo |
|---:|---|---|
| 1 | COMUNICACIÓN | Primaria y secundaria |
| 2 | MATEMÁTICA | Primaria y secundaria |
| 3 | CIENCIA Y TECNOLOGÍA (CYT) | Primaria y secundaria |
| 4 | PERSONAL SOCIAL | Primaria |
| 5 | CIENCIAS SOCIALES (CCSS) | Secundaria |

### 2.1. Grados comprendidos

- **Primaria:** de 2.º a 6.º grado.
- **Secundaria:** de 1.º a 5.º grado.

Cada módulo procesa los registros correspondientes a su área y nivel educativo, de acuerdo con las configuraciones establecidas.

## 3. Estructura general del proyecto

La organización del proyecto es la siguiente:

```text
tukuyrikuy_salida_2026/
│
├── COMUNICACION/
│   ├── resources/
│   ├── resultados/
│   ├── config_comunicacion.py
│   ├── procesamiento_comunicacion.py
│   ├── reportes_comunicacion.py
│   └── main_comunicacion.py
│
├── MATEMATICA/
│   ├── resources/
│   ├── resultados/
│   ├── config_matematica.py
│   ├── procesamiento_matematica.py
│   ├── reportes_matematica.py
│   └── main_matematica.py
│
├── CYT/
│   ├── resources/
│   ├── resultados/
│   ├── config_cyt.py
│   ├── procesamiento_cyt.py
│   ├── reportes_cyt.py
│   └── main_cyt.py
│
├── PERSONAL_SOCIAL/
│   ├── resources/
│   ├── resultados/
│   ├── config_personal_social.py
│   ├── procesamiento_personal_social.py
│   ├── reportes_personal_social.py
│   └── main_personal_social.py
│
├── CCSS/
│   ├── resources/
│   ├── resultados/
│   ├── config_ccss.py
│   ├── procesamiento_ccss.py
│   ├── reportes_ccss.py
│   └── main_ccss.py
│
└── README.md
```

**Nota:** cada módulo dispone de sus propios archivos de entrada y resultados. La ubicación y el nombre exacto del Excel de entrada se especifican en el README de cada área.

## 4. Descripción de los componentes

Cada módulo contiene cuatro archivos Python principales:

| Archivo | Responsabilidad |
|---|---|
| `config_*.py` | Define las configuraciones, claves de respuestas, competencias, distribución de preguntas, ponderaciones y umbrales de evaluación. |
| `procesamiento_*.py` | Lee los archivos Excel, valida los datos, compara las respuestas con las claves y calcula los resultados. |
| `reportes_*.py` | Organiza la información procesada y genera los reportes generales e institucionales. |
| `main_*.py` | Coordina y ejecuta el flujo completo del módulo. |

### 4.1. Flujo general de procesamiento

El funcionamiento de cada módulo sigue esta secuencia:

1. Lectura del archivo Excel de entrada.
2. Selección de la hoja de resultados correspondiente.
3. Validación y selección de los registros según el nivel y grado.
4. Normalización y procesamiento de las respuestas.
5. Comparación con las claves de respuestas configuradas.
6. Cálculo de resultados por pregunta y competencia.
7. Determinación de los niveles de logro.
8. Generación de la base procesada.
9. Elaboración del reporte general consolidado.
10. Generación de reportes individuales por institución educativa.
11. Organización de los reportes institucionales en un archivo ZIP.

## 5. Requisitos del sistema

Para ejecutar el proyecto se requiere:

- Sistema operativo Windows.
- Python instalado.
- Entorno virtual de Python configurado.
- Archivos Excel de entrada correspondientes a cada módulo.
- Bibliotecas necesarias para la lectura, procesamiento y generación de archivos Excel.

### 5.1. Instalación de dependencias

Las bibliotecas principales utilizadas son:

```bash
pip install pandas openpyxl
```

Para el módulo de Ciencia y Tecnología también se contempla la dependencia:

```bash
pip install numpy
```

Se recomienda instalar las dependencias dentro del entorno virtual que se utilizará para ejecutar los módulos.

## 6. Ubicación del proyecto

La ubicación de referencia del proyecto es:

```text
D:\Python\tukuyrikuy_salida_2026\
```

Dentro de esta carpeta se encuentran los cinco módulos independientes.

## 7. Ejecución del sistema

**Importante:** los módulos se ejecutan individualmente mediante su respectivo archivo `main_*.py`. No se requiere un programa centralizador.

Para ejecutar un módulo, primero se debe ingresar a su carpeta correspondiente y luego ejecutar el archivo principal.

### 7.1. Comandos de ejecución

**COMUNICACIÓN**

```bash
cd D:\Python\tukuyrikuy_salida_2026\COMUNICACION
python main_comunicacion.py
```

**MATEMÁTICA**

```bash
cd D:\Python\tukuyrikuy_salida_2026\MATEMATICA
python main_matematica.py
```

**CIENCIA Y TECNOLOGÍA**

```bash
cd D:\Python\tukuyrikuy_salida_2026\CYT
python main_cyt.py
```

**PERSONAL SOCIAL**

```bash
cd D:\Python\tukuyrikuy_salida_2026\PERSONAL_SOCIAL
python main_personal_social.py
```

**CIENCIAS SOCIALES**

```bash
cd D:\Python\tukuyrikuy_salida_2026\CCSS
python main_ccss.py
```

## 8. Archivos de entrada

Cada módulo utiliza su propio archivo Excel de resultados.

Los archivos deben colocarse en la carpeta de recursos configurada para cada módulo.

Entre los archivos de entrada identificados se encuentran:

| Módulo | Archivo Excel |
|---|---|
| COMUNICACIÓN | `RESULTADOS_COMUNICACION.xlsx` |
| MATEMÁTICA | Archivo de resultados configurado para el módulo. |
| CIENCIA Y TECNOLOGÍA | `RESULTADOS_CIENCIA_TECNOLOGIA.xlsx` |
| PERSONAL SOCIAL | `RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx` |
| CCSS | `RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx` |

La hoja de origen, las columnas requeridas y los criterios de selección pueden variar según el módulo. Por ello, se debe revisar el README específico de cada área antes de ejecutar su procesamiento.

## 9. Archivos de salida

Cada módulo genera sus resultados dentro de su propia carpeta `resultados/`.

Los principales productos son:

| Producto | Descripción |
|---|---|
| BASE PROCESADA | Contiene los registros de estudiantes y los resultados calculados durante el procesamiento. |
| REPORTE GENERAL | Presenta los resultados consolidados del área evaluada. |
| REPORTES INSTITUCIONALES | Contiene un archivo Excel por institución educativa. |
| ARCHIVO ZIP | Agrupa los reportes individuales de las instituciones educativas. |

Los nombres exactos de los archivos de salida dependen del módulo y se detallan en su documentación específica.

## 10. Criterios de procesamiento y evaluación

Cada módulo dispone de su propia configuración de evaluación.

Esta configuración puede incluir:

- Claves de respuestas por grado.
- Número de preguntas por evaluación.
- Distribución de preguntas por competencia.
- Puntajes obtenidos por los estudiantes.
- Niveles de logro por competencia.
- Ponderaciones para el resultado global.
- Umbrales para determinar los niveles de logro globales.

Los criterios no necesariamente son iguales entre las áreas. Por esta razón, cada módulo mantiene sus propias reglas en el archivo `config_*.py`.

## 11. Consideraciones importantes

1. Verificar que los archivos Excel de entrada estén ubicados en la carpeta correspondiente.
2. Comprobar que las hojas y columnas requeridas estén disponibles.
3. Revisar las claves de respuestas antes de ejecutar el procesamiento.
4. Verificar que la distribución de preguntas por competencia coincida con la configuración del área.
5. Mantener los nombres de las instituciones educativas de manera consistente para facilitar la consolidación de reportes.
6. Ejecutar cada módulo desde su propia carpeta.
7. Revisar los archivos generados al finalizar cada ejecución.
8. Conservar los archivos originales de entrada para mantener la trazabilidad de la información procesada.

## 12. Documentación específica

Para conocer los procedimientos, estructuras de entrada, criterios de evaluación y archivos de salida de cada módulo, consultar los README correspondientes:

- `COMUNICACION/README.md`
- `MATEMATICA/README.md`
- `CYT/README.md`
- `PERSONAL_SOCIAL/README.md`
- `CCSS/README.md`

## 13. Información del proyecto

**Proyecto:** TUKUYRICUY 2026 – Evaluación de Salida  
**Institución:** UGEL 02 – La Esperanza  
**Ámbito:** Instituciones educativas de primaria y secundaria  
**Tecnología principal:** Python  
**Procesamiento de datos:** Pandas y OpenPyXL  
**Formato de entrada y salida:** Microsoft Excel (.xlsx)