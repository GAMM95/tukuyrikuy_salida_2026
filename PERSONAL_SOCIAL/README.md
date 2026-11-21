# TUKUYRICUY 2026 – MÓDULO PERSONAL SOCIAL (PRIMARIA)

## 1. Descripción general

El módulo **PERSONAL SOCIAL** forma parte del sistema de procesamiento de resultados de la Evaluación de Salida 2026 – TUKUYRICUY, correspondiente al área de Personal Social del nivel primario.

Su finalidad es procesar y sistematizar las respuestas de los estudiantes de 2.º a 6.º de primaria, comparándolas con las claves de respuestas configuradas para cada grado.

El procesamiento permite calcular los resultados por competencia, determinar los niveles de logro y obtener el puntaje global de cada estudiante. Asimismo, genera una base procesada, un reporte general consolidado y reportes individuales por institución educativa.

## 2. Estructura de archivos

El módulo está organizado de la siguiente manera:

```text
PERSONAL_SOCIAL/
│
├── resources/
│   └── RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx
│
├── resultados/
│   ├── BASE_PROCESADA_PERSONAL_SOCIAL_PRIMARIA.xlsx
│   ├── REPORTE_GENERAL_PERSONAL_SOCIAL_PRIMARIA.xlsx
│   ├── REPORTES_INSTITUCIONES_PERSONAL_SOCIAL_PRIMARIA.zip
│   └── REPORTES_INSTITUCIONES/
│       └── Reportes individuales por institución
│
├── config_personal_social.py
├── procesamiento_personal_social.py
├── reportes_personal_social.py
└── main_personal_social.py
```

## 3. Descripción de los archivos Python

| Archivo | Función |
|---|---|
| `config_personal_social.py` | Contiene las configuraciones del módulo, las competencias, la distribución de preguntas, las ponderaciones y los umbrales de evaluación. |
| `procesamiento_personal_social.py` | Lee el archivo Excel, valida y procesa las respuestas, calcula los resultados por competencia y determina los niveles de logro. |
| `reportes_personal_social.py` | Genera el reporte general consolidado y los reportes individuales por institución educativa. |
| `main_personal_social.py` | Ejecuta el flujo completo de procesamiento y generación de reportes. |

## 4. Instalación y configuración

Para utilizar el módulo:

1. Colocar la carpeta dentro del proyecto o copiar los cuatro archivos Python en la siguiente ubicación:

   ```text
   D:\Python\tukuyrikuy_salida_2026\PERSONAL_SOCIAL\
   ```

2. Crear la carpeta `resources`.
3. Colocar dentro de `resources` el archivo Excel de entrada:

   ```text
   RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx
   ```

4. Utilizar el mismo entorno virtual de Python configurado para los demás módulos.

## 5. Archivo de entrada

El archivo Excel de entrada es:

```text
PERSONAL_SOCIAL/resources/RESULTADOS_PERSONAL_SOCIAL_CCSS.xlsx
```

**Hoja de lectura:**

- Se utiliza preferentemente la hoja `RESULTADOS`.
- Si esta hoja no existe, se utiliza la primera hoja del archivo.

**Estructura requerida:**

| Campo | Descripción |
|---|---|
| `NIVEL` | Nivel educativo del estudiante. |
| `GRADO` | Grado de estudios. |
| `V1` a `V16` | Respuestas correspondientes a los 16 ítems de evaluación. |
| Campos de identificación | Se conservan cuando están presentes. |
| `P1` a `P16` | Se conservan cuando están presentes. |

El procesamiento considera únicamente los registros que cumplen las siguientes condiciones:

- `NIVEL = PRIMARIA`.
- Grados comprendidos entre 2.º y 6.º de primaria.

## 6. Distribución de preguntas por competencia

La evaluación comprende cinco competencias, distribuidas de la siguiente manera:

| Competencia | Preguntas | Total de ítems |
|---|---|---:|
| COMP_1 | P1–P3 | 3 |
| COMP_2 | P4–P6 | 3 |
| COMP_3 | P7–P10 | 4 |
| COMP_4 | P11–P13 | 3 |
| COMP_5 | P14–P16 | 3 |

### 6.1. Niveles de logro por competencia

| Competencia | Inicio | Proceso | Logrado |
|---|---:|---:|---:|
| COMP_1 | 0–1 | 2 | 3 |
| COMP_2 | 0–1 | 2 | 3 |
| COMP_3 | 0–2 | 3 | 4 |
| COMP_4 | 0–1 | 2 | 3 |
| COMP_5 | 0–1 | 2 | 3 |

Los niveles de logro se determinan a partir de la cantidad de respuestas correctas obtenidas en cada competencia.

## 7. Ponderación y cálculo del puntaje global

Cada una de las cinco competencias tiene el mismo peso dentro del resultado global:

| Competencia | Ponderación |
|---|---:|
| COMP_1 | 20 % |
| COMP_2 | 20 % |
| COMP_3 | 20 % |
| COMP_4 | 20 % |
| COMP_5 | 20 % |
| **Total** | **100 %** |

El puntaje global se calcula normalizando el resultado de cada competencia según su número de preguntas y aplicando una ponderación del 20 %.

De esta manera, todas las competencias tienen la misma importancia, aunque tengan diferente cantidad de ítems.

### 7.1. Umbrales globales

Los umbrales globales de referencia son:

- **Inicio:** hasta 50 puntos.
- **Proceso:** más de 50 y hasta 75 puntos.
- **Logrado:** más de 75 puntos.

Estos valores son provisionales y pueden modificarse en `config_personal_social.py` mediante las variables:

```python
LIMITE_GLOBAL_INICIO
LIMITE_GLOBAL_PROCESO
```

## 8. Instalación de dependencias

Se requiere Python y las siguientes bibliotecas:

```bash
pip install pandas openpyxl
```

## 9. Ejecución del módulo

Ubicarse en la carpeta:

```text
D:\Python\tukuyrikuy_salida_2026\PERSONAL_SOCIAL\
```

Ejecutar el archivo principal:

```bash
python main_personal_social.py
```

El archivo `main_personal_social.py` coordina el procesamiento de las respuestas y la generación de los reportes correspondientes.

## 10. Archivos de salida

Los resultados se generan dentro de la carpeta `resultados/`:

| Archivo o carpeta | Descripción |
|---|---|
| `BASE_PROCESADA_PERSONAL_SOCIAL_PRIMARIA.xlsx` | Base de datos con las respuestas procesadas y los resultados obtenidos por los estudiantes. |
| `REPORTE_GENERAL_PERSONAL_SOCIAL_PRIMARIA.xlsx` | Reporte general consolidado de los resultados de primaria. |
| `REPORTES_INSTITUCIONES_PERSONAL_SOCIAL_PRIMARIA.zip` | Archivo comprimido con los reportes individuales por institución educativa. |
| `REPORTES_INSTITUCIONES/` | Carpeta que contiene un archivo Excel por cada institución educativa. |

## 11. Consideraciones importantes

- El archivo Excel de entrada debe conservar las columnas requeridas para el procesamiento.
- Se procesan únicamente los registros del nivel PRIMARIA, correspondientes a los grados de 2.º a 6.º.
- Los campos de identificación y las columnas `P1` a `P16` se conservan cuando están presentes.
- La distribución de preguntas por competencia debe coincidir con la configuración del módulo.
- Cada competencia representa el 20 % del puntaje global.
- Los umbrales globales pueden modificarse desde `config_personal_social.py`.
- Se recomienda conservar una copia de seguridad del archivo original antes de ejecutar el procesamiento.
- Al volver a ejecutar el módulo, los archivos de salida se actualizarán según los datos disponibles en el archivo de entrada.

---

**Módulo:** Personal Social – Primaria  
**Proyecto:** TUKUYRICUY 2026 – Evaluación de Salida  
**Ámbito:** UGEL 02 – La Esperanza