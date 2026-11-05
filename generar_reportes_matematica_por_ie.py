# -*- coding: utf-8 -*-
"""
TUKUYRICUY 2026 - GENERADOR DE REPORTES DE MATEMÁTICA POR INSTITUCIÓN

Genera:
1. Un reporte general con todos los estudiantes.
2. Un archivo Excel independiente por cada institución educativa.
3. Un ZIP con todos los reportes institucionales.

Uso:
    pip install pandas openpyxl
    python generar_reportes_matematica_por_ie.py
    python generar_reportes_matematica_por_ie.py "C:/ruta/RESULTADOS_MATEMATICA.xlsx"

El archivo de entrada debe contener la hoja RESULTADOS y las columnas de
identificación y COMP_1...COMP_4 indicadas en el código.
"""

from pathlib import Path
import sys, re, zipfile, unicodedata

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

# ---------------- CONFIGURACIÓN ----------------
ENTRADA = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("RESULTADOS_MATEMATICA.xlsx")
CARPETA_SALIDA = ENTRADA.parent / "REPORTES_MATEMATICA_2026"
CARPETA_IE = CARPETA_SALIDA / "INSTITUCIONES"
ARCHIVO_GENERAL = CARPETA_SALIDA / "REPORTE_GENERAL_MATEMATICA.xlsx"
ARCHIVO_ZIP = CARPETA_SALIDA / "REPORTES_INSTITUCIONES.zip"
HOJA_ORIGEN = "RESULTADOS"

COMPETENCIAS = [
    ("COMP_1_NOMBRE", "COMP_1_ACIERTOS", "COMP_1_TOTAL"),
    ("COMP_2_NOMBRE", "COMP_2_ACIERTOS", "COMP_2_TOTAL"),
    ("COMP_3_NOMBRE", "COMP_3_ACIERTOS", "COMP_3_TOTAL"),
    ("COMP_4_NOMBRE", "COMP_4_ACIERTOS", "COMP_4_TOTAL"),
]

# Umbrales configurables para cada competencia (máximo esperado: 4 puntos).
# Inicio: 0-2; Proceso: 3; Logrado: 4.
def nivel_logro(x):
    if pd.isna(x): return "SIN DATO"
    try: x = float(x)
    except (TypeError, ValueError): return "SIN DATO"
    if x <= 2: return "INICIO"
    if x == 3: return "PROCESO"
    if x >= 4: return "LOGRADO"
    return "SIN DATO"

AZUL, AZUL2, CELESTE = "17365D", "1F4E78", "D9EAF7"
VERDE, AMARILLO, ROJO = "E2F0D9", "FFF2CC", "FCE4D6"
BLANCO, GRIS = "FFFFFF", "E7E6E6"
thin = Side(style="thin", color="B7C9D6")
BORDE = Border(left=thin, right=thin, top=thin, bottom=thin)

def limpio(x, default="SIN DATO"):
    if pd.isna(x) or str(x).strip() == "": return default
    return re.sub(r"\s+", " ", str(x)).strip()

def slug(texto):
    s = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    s = re.sub(r"[^A-Za-z0-9]+", "_", s).strip("_").upper()
    return (s[:65] or "INSTITUCION")

def titulo(ws, texto, end=8):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end)
    c = ws.cell(1, 1, texto)
    c.fill = PatternFill("solid", fgColor=AZUL)
    c.font = Font(name="Aptos Display", size=16, bold=True, color=BLANCO)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

def encabezado(ws, row, c1, c2):
    for col in range(c1, c2+1):
        c = ws.cell(row, col)
        c.fill = PatternFill("solid", fgColor=AZUL2)
        c.font = Font(bold=True, color=BLANCO)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDE
    ws.row_dimensions[row].height = 32

def formato_tabla(ws, ref, nombre):
    tab = Table(displayName=nombre, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True,
                                        showColumnStripes=False)
    ws.add_table(tab)

def preparar_libro(datos, institucion=None, general=False):
    """Construye un libro de reporte con fórmulas visibles y gráficos."""
    wb = Workbook()
    ws = wb.active
    ws.title = "DETALLE ESTUDIANTES"
    ws.sheet_view.showGridLines = False

    # Datos completos de la institución (o todos los registros en el general).
    cols = list(datos.columns)
    ws.append(cols)
    for row in datos.itertuples(index=False, name=None):
        ws.append([None if pd.isna(v) else v for v in row])
    if len(datos):
        encabezado(ws, 1, 1, len(cols))
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(datos)+1}"
        formato_tabla(ws, f"A1:{get_column_letter(len(cols))}{len(datos)+1}", "TablaDetalle")
    for c in range(1, len(cols)+1):
        ws.column_dimensions[get_column_letter(c)].width = min(
            max(13, max(len(str(ws.cell(r,c).value or "")) for r in range(1, min(len(datos)+2, 80)))), 30)

    # Resumen
    wr = wb.create_sheet("RESUMEN")
    titulo(wr, "TUKUYRICUY 2026 | REPORTE DE MATEMÁTICA", 6)
    wr["A2"] = "Institución educativa"
    wr["B2"] = institucion if institucion else "TODAS LAS INSTITUCIONES"
    wr["A3"] = "Estudiantes evaluados"
    wr["B3"] = f"=COUNTA('DETALLE ESTUDIANTES'!A2:A{len(datos)+1})"
    wr["A4"] = "Promedio de aciertos"
    col_correctas = get_column_letter(cols.index("CORRECTAS")+1)
    wr["B4"] = f"=IFERROR(AVERAGE('DETALLE ESTUDIANTES'!{col_correctas}2:{col_correctas}{len(datos)+1}),0)"
    col_pct = get_column_letter(cols.index("PORCENTAJE")+1) if "PORCENTAJE" in cols else None
    wr["A5"] = "Promedio de rendimiento"
    wr["B5"] = f"=IFERROR(AVERAGE('DETALLE ESTUDIANTES'!{col_pct}2:{col_pct}{len(datos)+1}),0)" if col_pct else "=0"
    wr["B4"].number_format = "0.00"
    wr["B5"].number_format = "0.0%"
    for r in range(2,6):
        wr.cell(r,1).font = Font(bold=True, color=AZUL)
        wr.cell(r,1).fill = PatternFill("solid", fgColor=CELESTE)
        wr.cell(r,1).border = BORDE
        wr.cell(r,2).border = BORDE

    # Resumen por competencia
    wc = wb.create_sheet("COMPETENCIAS")
    titulo(wc, "RESULTADOS POR COMPETENCIA", 6)
    heads = ["COMPETENCIA", "INICIO", "PROCESO", "LOGRADO", "TOTAL", "% LOGRADO"]
    for j,h in enumerate(heads,1): wc.cell(3,j,h)
    encabezado(wc,3,1,6)
    for i,(name_col,acc_col,total_col) in enumerate(COMPETENCIAS,4):
        name = limpio(datos[name_col].dropna().iloc[0], f"Competencia {i-3}") if name_col in datos and datos[name_col].notna().any() else f"Competencia {i-3}"
        col = get_column_letter(cols.index(acc_col)+1)
        end = len(datos)+1
        wc.cell(i,1,name)
        wc.cell(i,2,f'=COUNTIF(\'DETALLE ESTUDIANTES\'!{col}2:{col}{end},"<=2")')
        wc.cell(i,3,f'=COUNTIF(\'DETALLE ESTUDIANTES\'!{col}2:{col}{end},3)')
        wc.cell(i,4,f'=COUNTIF(\'DETALLE ESTUDIANTES\'!{col}2:{col}{end},4)')
        wc.cell(i,5,f"=SUM(B{i}:D{i})")
        wc.cell(i,6,f"=IFERROR(D{i}/E{i},0)")
        wc.cell(i,6).number_format = "0.0%"
    for row in wc.iter_rows(min_row=4,max_row=7,min_col=1,max_col=6):
        for c in row: c.border=BORDE; c.alignment=Alignment(wrap_text=True,vertical="center")
    chart = BarChart()
    chart.type = "bar"; chart.style = 10; chart.grouping = "stacked"; chart.overlap = 100
    chart.title = "Distribución de estudiantes por nivel y competencia"
    chart.x_axis.title = "Estudiantes"; chart.y_axis.title = "Competencia"
    chart.add_data(Reference(wc,min_col=2,max_col=4,min_row=3,max_row=7),titles_from_data=True)
    chart.set_categories(Reference(wc,min_col=1,min_row=4,max_row=7))
    chart.height=8; chart.width=17; chart.legend.position="b"
    wc.add_chart(chart,"H3")

    # Niveles por grado y sección
    wn = wb.create_sheet("GRADO Y SECCIÓN")
    titulo(wn, "RESULTADOS POR GRADO Y SECCIÓN", 9)
    heads2 = ["NIVEL","GRADO","SECCIÓN","ESTUDIANTES","PROMEDIO ACIERTOS",
              "INICIO (COMP.)","PROCESO (COMP.)","LOGRADO (COMP.)"]
    for j,h in enumerate(heads2,1): wn.cell(3,j,h)
    encabezado(wn,3,1,8)
    grupos = datos[["NIVEL","GRADO","SECCION"]].drop_duplicates().sort_values(["NIVEL","GRADO","SECCION"])
    for r,(_,g) in enumerate(grupos.iterrows(),4):
        wn.cell(r,1,g["NIVEL"]); wn.cell(r,2,int(g["GRADO"])); wn.cell(r,3,g["SECCION"])
        wn.cell(r,4,f'=COUNTIFS(\'DETALLE ESTUDIANTES\'!$C:$C,A{r},\'DETALLE ESTUDIANTES\'!$E:$E,B{r},\'DETALLE ESTUDIANTES\'!$F:$F,C{r})')
        wn.cell(r,5,f'=IFERROR(AVERAGEIFS(\'DETALLE ESTUDIANTES\'!${col_correctas}:${col_correctas},\'DETALLE ESTUDIANTES\'!$C:$C,A{r},\'DETALLE ESTUDIANTES\'!$E:$E,B{r},\'DETALLE ESTUDIANTES\'!$F:$F,C{r}),0)')
        for target, level in [(6,"INICIO"),(7,"PROCESO"),(8,"LOGRADO")]:
            formulas=[]
            for _,acc_col,_ in COMPETENCIAS:
                cc=get_column_letter(cols.index(acc_col)+1)
                criterion = '"<=2"' if level=="INICIO" else ("3" if level=="PROCESO" else "4")
                formulas.append(f'COUNTIFS(\'DETALLE ESTUDIANTES\'!$C:$C,A{r},\'DETALLE ESTUDIANTES\'!$E:$E,B{r},\'DETALLE ESTUDIANTES\'!$F:$F,C{r},\'DETALLE ESTUDIANTES\'!${cc}:${cc},{criterion})')
            wn.cell(r,target,"="+"+".join(formulas))
    last=3+len(grupos)
    if last>=4:
        for row in wn.iter_rows(min_row=4,max_row=last,min_col=1,max_col=8):
            for c in row: c.border=BORDE
        wn.auto_filter.ref=f"A3:H{last}"
        ch=BarChart(); ch.type="col"; ch.style=11; ch.title="Estudiantes por grado y sección"
        ch.y_axis.title="Estudiantes"
        ch.add_data(Reference(wn,min_col=4,min_row=3,max_row=last),titles_from_data=True)
        wn["I3"]="GRADO / SECCIÓN"
        for r in range(4,last+1): wn.cell(r,9,f'=B{r}&"° "&C{r}')
        ch.set_categories(Reference(wn,min_col=9,min_row=4,max_row=last))
        ch.height=8; ch.width=17; wn.add_chart(ch,"J3")
    for sh in wb.worksheets:
        sh.sheet_view.showGridLines=False
        sh.freeze_panes = "A4" if sh.title in ("COMPETENCIAS","GRADO Y SECCIÓN") else sh.freeze_panes
        for c in range(1, min(sh.max_column,12)+1):
            sh.column_dimensions[get_column_letter(c)].width=max(sh.column_dimensions[get_column_letter(c)].width or 0,16)
    wr.column_dimensions["A"].width=30; wr.column_dimensions["B"].width=42
    return wb

def main():
    if not ENTRADA.exists():
        raise FileNotFoundError(f"No existe el archivo de entrada: {ENTRADA}")
    print("Leyendo hoja RESULTADOS...")
    df = pd.read_excel(ENTRADA, sheet_name=HOJA_ORIGEN, engine="openpyxl")
    if df.empty: raise ValueError("La hoja RESULTADOS está vacía.")
    obligatorias=["ID_REGISTRO","NIVEL","INSTITUCION","GRADO","SECCION","CORRECTAS","TOTAL_PREGUNTAS"]
    faltan=[c for c in obligatorias if c not in df.columns]
    if faltan: raise ValueError("Faltan columnas: "+", ".join(faltan))
    for name,acc,total in COMPETENCIAS:
        for c in (name,acc,total):
            if c not in df.columns: raise ValueError(f"Falta la columna {c}")
    for c in ["NIVEL","INSTITUCION","SECCION"]:
        df[c]=df[c].apply(limpio)
    df["GRADO"]=pd.to_numeric(df["GRADO"],errors="coerce").fillna(0).astype(int)
    for c in ["CORRECTAS","TOTAL_PREGUNTAS","PORCENTAJE"]+[x for _,a,t in COMPETENCIAS for x in (a,t)]:
        if c in df.columns: df[c]=pd.to_numeric(df[c],errors="coerce")
    df["GRADO_TEXTO"]=df.apply(lambda r:f'{r["GRADO"]}° {str(r["NIVEL"]).lower()}' if r["GRADO"] else "Grado no identificado",axis=1)

    CARPETA_IE.mkdir(parents=True,exist_ok=True)
    # Limpiar solo archivos previos generados en esta carpeta.
    for f in CARPETA_IE.glob("*.xlsx"): f.unlink()
    print("Generando reporte general...")
    preparar_libro(df, general=True).save(ARCHIVO_GENERAL)

    instituciones=sorted(df["INSTITUCION"].dropna().unique(),key=lambda x:str(x).casefold())
    nombres_usados=set()
    manifiesto=[]
    for i,inst in enumerate(instituciones,1):
        sub=df[df["INSTITUCION"]==inst].copy()
        base=slug(inst)
        nombre=base
        n=2
        while nombre.lower() in nombres_usados:
            nombre=f"{base[:58]}_{n}"; n+=1
        nombres_usados.add(nombre.lower())
        ruta=CARPETA_IE/f"{nombre}_REPORTE_MATEMATICA.xlsx"
        preparar_libro(sub,institucion=inst).save(ruta)
        manifiesto.append(ruta)
        if i%25==0 or i==len(instituciones):
            print(f"Instituciones procesadas: {i}/{len(instituciones)}")

    with zipfile.ZipFile(ARCHIVO_ZIP,"w",zipfile.ZIP_DEFLATED) as z:
        for ruta in manifiesto: z.write(ruta,arcname=ruta.name)

    print("\\nPROCESO FINALIZADO")
    print(f"Estudiantes: {len(df):,}")
    print(f"Instituciones: {len(instituciones):,}")
    print(f"Reporte general: {ARCHIVO_GENERAL}")
    print(f"Carpeta de reportes individuales: {CARPETA_IE}")
    print(f"Archivo ZIP: {ARCHIVO_ZIP}")

if __name__=="__main__":
    main()
