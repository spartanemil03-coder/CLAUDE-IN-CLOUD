"""Agrega las tablas transversales abreviadas AL ARCHIVO DE LA PROFESORA ('Fuentes de consulta.xlsx'), con fórmulas vivas.

Hojas que se agregan (la hoja 'Fuentes de datos' de la profesora queda como está):
  Cohorte ficticio      (es la 'Hoja1' vacía, renombrada): insumos, reparto de 'no especificado', separación del año 0 y agrupación.
  Tabla hombres / Tabla mujeres / Tabla ambos sexos: las tablas transversales abreviadas.
  Gráficos y resumen:   indicadores, gráficos nativos de Excel, comparación con CONAPO y notas.

Las fórmulas leen directo de las celdas de 'Fuentes de datos' (población C37:H52, defunciones B106:F127).
Este script escribe con openpyxl; luego 03_finalizar_libro.py restaura las capturas de pantalla de la profesora y guarda los valores calculados.

Uso:  python3 scripts/02_construir_libro.py [salida.xlsx]
"""
import sys
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference, Series
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parent.parent
ORIGEN = RAIZ / "fuentes" / "Fuentes_de_consulta.xlsx"
SALIDA = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "resultados" / "_sin_terminar.xlsx"

FS = "'Fuentes de datos'!"
POB0 = 37                       # fila de 00-04 años en 'Fuentes de datos' (C: grupo, E: total, F: mujeres, H: hombres)
DEF_ROW0 = 108                  # fila de 'Menores de 1 año' (C: total, D: hombre, E: mujer, F: no especificado); 85+ = 126; edad NE = 127
DEF_TOTAL = 107

# ---------- estilos (los colores del archivo de la profesora: acento verde del tema y turquesa de las tablas) ----------
VERDE_OSC = PatternFill("solid", fgColor=Color(theme=9, tint=-0.249977111117893))
VERDE_CLA = PatternFill("solid", fgColor=Color(theme=9, tint=0.7999816888943144))
GRIS = PatternFill("solid", fgColor=Color(theme=2, tint=-0.0999786370433668))
TURQ = PatternFill("solid", fgColor="FF00C7BB")
AMARILLO = PatternFill("solid", fgColor="FFFFFF00")
BLANCO = Color(theme=0)
f_tit = Font(name="Calibri", size=20, bold=True, color=BLANCO)
f_band = Font(name="Calibri", size=12, bold=True, color=BLANCO)
f_sub = Font(name="Calibri", size=11, color=BLANCO)
f_enc = Font(name="Calibri", size=10, bold=True, color="FFFFFFFF")
f_n = Font(name="Calibri", size=11)
f_b = Font(name="Calibri", size=11, bold=True)
f_nota = Font(name="Calibri", size=10, italic=True, color="FF404040")
f_in = Font(name="Calibri", size=11, color="FF0000FF")
borde = Border(*(Side(style="thin", color="FFA6A6A6"),) * 4)
centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
izq_aj = Alignment(horizontal="left", vertical="center", wrap_text=True)
der = Alignment(horizontal="right", vertical="center")


def fondo(ws, filas, cols):
    for r in range(1, filas + 1):
        for c in range(1, cols + 1):
            ws.cell(r, c).fill = VERDE_CLA


def pon(ws, ref, v, font=f_n, fmt=None, fill=None, alin=None, caja=True):
    c = ws[ref]
    c.value = v
    c.font = font
    if fmt:
        c.number_format = fmt
    if fill:
        c.fill = fill
    if alin:
        c.alignment = alin
    if caja:
        c.border = borde
    return c


def banda(ws, fila, texto, c0, c1, font=f_band, alto=None):
    for c in range(c0, c1 + 1):
        ws.cell(fila, c).fill = VERDE_OSC
    ws.cell(fila, c0).value = texto
    ws.cell(fila, c0).font = font
    ws.cell(fila, c0).alignment = Alignment(vertical="center")
    if alto:
        ws.row_dimensions[fila].height = alto


def encabezado(ws, fila, c0, titulos, alto=32):
    for i, t in enumerate(titulos):
        pon(ws, f"{get_column_letter(c0 + i)}{fila}", t, f_enc, fill=TURQ, alin=centro)
    ws.row_dimensions[fila].height = alto


def cuerpo(fila):
    return GRIS if fila % 2 == 0 else VERDE_CLA


def nota(ws, fila, texto, c0=2, c1=9, alto=None):
    ws.cell(fila, c0).value = texto
    ws.cell(fila, c0).font = f_nota
    ws.cell(fila, c0).alignment = izq_aj
    ws.merge_cells(start_row=fila, start_column=c0, end_row=fila, end_column=c1)
    if alto:
        ws.row_dimensions[fila].height = alto


def anchos(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


wb = load_workbook(ORIGEN)
cf = wb["Hoja1"]
cf.title = "Cohorte ficticio"
HOJAS = {"H": "Tabla hombres", "M": "Tabla mujeres", "A": "Tabla ambos sexos"}
NOMBRE = {"H": "hombres", "M": "mujeres", "A": "ambos sexos"}
th = {s: wb.create_sheet(HOJAS[s]) for s in "HMA"}
gr = wb.create_sheet("Gráficos y resumen")

# =====================================================================
# Cohorte ficticio
# =====================================================================
ws = cf
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 85
NF, NC = 90, 12
fondo(ws, NF, NC)
banda(ws, 2, "COHORTE FICTICIO", 2, 11, f_tit, 30)
banda(ws, 3, 'Insumos de las "Tablas Transversales Abreviadas" con datos de INEGI 2025: Encuesta Intercensal (EIC) y defunciones registradas', 2, 11, f_sub, 18)
pon(ws, "B5", "Radix de la cohorte ficticia (l0)", f_b, caja=False)
pon(ws, "D5", 100000, f_in, "#,##0", AMARILLO)
nota(ws, 5, "Supuesto: 100,000 nacimientos hipotéticos que viven toda su vida con las tasas de mortalidad de 2025 (cohorte ficticia). Se puede cambiar.", 5, 11, 30)
nota(ws, 6, "Letra azul sobre amarillo = supuesto que se puede cambiar. Todo lo demás son fórmulas que leen la hoja 'Fuentes de datos'.", 2, 11)

# ---- A. defunciones ----
banda(ws, 8, "A. Defunciones registradas 2025 (INEGI): se reparten el sexo y la edad no especificados", 2, 11)
nota(ws, 9, "Paso 1: el sexo no especificado se reparte entre hombres y mujeres en proporción a sus defunciones, edad por edad.  Paso 2: las defunciones de edad no "
     "especificada se reparten entre todas las edades en proporción a las ya ajustadas, por sexo. Las cifras de INEGI vienen como texto con comas; aquí se convierten a número.", 2, 11, 44)
encabezado(ws, 10, 2, ["Grupo de edad", "Hombres (INEGI)", "Mujeres (INEGI)", "Sexo no especificado", "Hombres (paso 1)", "Mujeres (paso 1)",
                       "Hombres (paso 2)", "Mujeres (paso 2)"])
A0 = 11
A1 = A0 + 18          # 85+
ANE = A1 + 1          # edad no especificada
ATOT = ANE + 1
FAC = ATOT + 1
CTRL = FAC + 1


def num(celda):
    return f'IF({FS}{celda}="",0,VALUE(SUBSTITUTE({FS}{celda},",","")))'


for k in range(20):
    f = A0 + k
    fr = DEF_ROW0 + k                      # 108 ... 127
    fl = cuerpo(f)
    pon(ws, f"B{f}", f"={FS}B{fr}", f_n, fill=fl)
    pon(ws, f"C{f}", "=" + num(f"D{fr}"), f_n, "#,##0", fl)
    pon(ws, f"D{f}", "=" + num(f"E{fr}"), f_n, "#,##0", fl)
    pon(ws, f"E{f}", "=" + num(f"F{fr}"), f_n, "#,##0", fl)
    pon(ws, f"F{f}", f"=C{f}+E{f}*C{f}/(C{f}+D{f})", f_n, "#,##0.0", fl)
    pon(ws, f"G{f}", f"=D{f}+E{f}*D{f}/(C{f}+D{f})", f_n, "#,##0.0", fl)
    if k < 19:
        pon(ws, f"H{f}", f"=F{f}*$F${FAC}", f_n, "#,##0.0", fl)
        pon(ws, f"I{f}", f"=G{f}*$G${FAC}", f_n, "#,##0.0", fl)
    else:
        pon(ws, f"H{f}", None, fill=fl)
        pon(ws, f"I{f}", None, fill=fl)
pon(ws, f"B{ATOT}", "Total", f_b, fill=GRIS)
for col in "CDEFGHI":
    pon(ws, f"{col}{ATOT}", f"=SUM({col}{A0}:{col}{ANE})", f_b, "#,##0" if col in "CDE" else "#,##0.0", GRIS)
pon(ws, f"B{FAC}", "Factor del paso 2", f_n, fill=VERDE_CLA)
for col in "CDE":
    pon(ws, f"{col}{FAC}", None, fill=VERDE_CLA)
pon(ws, f"F{FAC}", f"=SUM(F{A0}:F{ANE})/SUM(F{A0}:F{A1})", f_n, "0.000000", VERDE_CLA)
pon(ws, f"G{FAC}", f"=SUM(G{A0}:G{ANE})/SUM(G{A0}:G{A1})", f_n, "0.000000", VERDE_CLA)
pon(ws, f"B{CTRL}", "Control: H + M (paso 2) − total de INEGI", f_n, fill=VERDE_CLA)
pon(ws, f"H{CTRL}", f"=ROUND(SUM(H{A0}:I{A1})-{num(f'C{DEF_TOTAL}')},6)", f_b, "0.0", VERDE_CLA)
nota(ws, CTRL, "← debe ser 0", 9, 11)

# ---- B. población: separar el año 0 ----
B0 = CTRL + 3
banda(ws, B0 - 1, "B. Población (EIC 2025): se separan los menores de 1 año", 2, 11)
nota(ws, B0, "La EIC trae 0-4 años junto, y la tabla necesita el año 0 por separado. Supuesto: las cohortes de nacimiento crecen hacia atrás a tasa constante r, "
     "r = ln(P5-9 / P0-4) / 5; entonces la fracción de 0-4 con menos de 1 año es (e^r − 1) / (e^(5r) − 1).", 2, 11, 44)
encabezado(ws, B0 + 1, 2, ["", "Hombres", "Mujeres"], 20)
filas_b = ["Población 0-4", "Población 5-9", "r (crecimiento de las cohortes)", "Fracción de 0-4 con menos de 1 año",
           "Población menor de 1 año", "Población de 1 a 4 años"]
for j, t in enumerate(filas_b):
    pon(ws, f"B{B0 + 2 + j}", t, f_n, fill=cuerpo(B0 + 2 + j))
for col, fcol in (("C", "H"), ("D", "F")):               # hombres = H (absolutos); mujeres = F (absolutos)
    r0 = B0 + 2
    fl = lambda r: cuerpo(r)
    pon(ws, f"{col}{r0}", f"={FS}{fcol}{POB0}", f_n, "#,##0", fl(r0))
    pon(ws, f"{col}{r0 + 1}", f"={FS}{fcol}{POB0 + 1}", f_n, "#,##0", fl(r0 + 1))
    pon(ws, f"{col}{r0 + 2}", f"=LN({col}{r0 + 1}/{col}{r0})/5", f_n, "0.00000", fl(r0 + 2))
    pon(ws, f"{col}{r0 + 3}", f"=(EXP({col}{r0 + 2})-1)/(EXP(5*{col}{r0 + 2})-1)", f_n, "0.0000", fl(r0 + 3))
    pon(ws, f"{col}{r0 + 4}", f"={col}{r0}*{col}{r0 + 3}", f_n, "#,##0", fl(r0 + 4))
    pon(ws, f"{col}{r0 + 5}", f"={col}{r0}-{col}{r0 + 4}", f_n, "#,##0", fl(r0 + 5))
PB_P0, PB_P14 = B0 + 6, B0 + 7

# ---- C. insumos por grupo ----
C0 = B0 + 11
banda(ws, C0 - 2, "C. Población y defunciones por grupo de la tabla: 0, 1-4, 5-9, ..., 70-74 y 75+", 2, 11)
nota(ws, C0 - 1, "75-79, 80-84 y 85+ de defunciones se suman en 75+, porque la EIC llega solo a 75 años y más.", 2, 11)
encabezado(ws, C0, 2, ["Grupo de edad", "Población hombres", "Defunciones hombres", "Población mujeres", "Defunciones mujeres",
                       "Población ambos sexos", "Defunciones ambos sexos"])
ETIQ = ["0", "1-4"] + [f"{x}-{x + 4}" for x in range(5, 75, 5)] + ["75+"]
I0 = C0 + 1
for k, et in enumerate(ETIQ):
    f = I0 + k
    fl = cuerpo(f)
    pon(ws, f"B{f}", et, f_n, fill=fl, alin=Alignment(horizontal="center"))
    if k == 0:
        ph, pm = f"=C{PB_P0}", f"=D{PB_P0}"
    elif k == 1:
        ph, pm = f"=C{PB_P14}", f"=D{PB_P14}"
    else:                                              # 5-9 es la fila POB0+1 ... 75+ es POB0+15
        ph, pm = f"={FS}H{POB0 + k - 1}", f"={FS}F{POB0 + k - 1}"
    pon(ws, f"C{f}", ph, f_n, "#,##0", fl)
    pon(ws, f"E{f}", pm, f_n, "#,##0", fl)
    if k < 16:
        dh, dm = f"=H{A0 + k}", f"=I{A0 + k}"
    else:
        dh, dm = f"=SUM(H{A0 + 16}:H{A1})", f"=SUM(I{A0 + 16}:I{A1})"
    pon(ws, f"D{f}", dh, f_n, "#,##0.0", fl)
    pon(ws, f"F{f}", dm, f_n, "#,##0.0", fl)
    pon(ws, f"G{f}", f"=C{f}+E{f}", f_n, "#,##0", fl)
    pon(ws, f"H{f}", f"=D{f}+F{f}", f_n, "#,##0.0", fl)
I1 = I0 + 16
pon(ws, f"B{I1 + 1}", "Total", f_b, fill=GRIS)
for col in "CDEFGH":
    pon(ws, f"{col}{I1 + 1}", f"=SUM({col}{I0}:{col}{I1})", f_b, "#,##0", GRIS)
anchos(ws, {"A": 3, "B": 38, "C": 17, "D": 17, "E": 17, "F": 17, "G": 17, "H": 17, "I": 17, "J": 10, "K": 10, "L": 3})

# =====================================================================
# Tablas abreviadas
# =====================================================================
T0 = 9
T1 = T0 + 16
COL_INS = {"H": ("C", "D"), "M": ("E", "F"), "A": ("G", "H")}
CF = "'Cohorte ficticio'!"


for s in "HMA":
    ws = th[s]
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    fondo(ws, 32, 16)
    pc, dc = COL_INS[s]
    banda(ws, 2, f"TABLA TRANSVERSAL ABREVIADA: {NOMBRE[s].upper()}", 2, 15, f_tit, 30)
    banda(ws, 3, "México 2025. Cohorte ficticia de 100,000 nacimientos con las tasas de mortalidad por edad observadas (INEGI: EIC 2025 y defunciones registradas)", 2, 15, f_sub, 18)
    pon(ws, "B5", "Radix l0", f_b, caja=False)
    pon(ws, "D5", f"={CF}D5", f_n, "#,##0", VERDE_CLA)
    nota(ws, 6, "nmx = D/P;  nqx = (2·n·nmx) / (2 + n·nmx);  npx = 1 − nqx;  lx+n = lx·npx;  ndx = lx·nqx;  nLx = lx+n·n + ndx·nAx;  Tx = suma de nLx hacia abajo;  ex = Tx/lx.", 2, 15)
    nota(ws, 7, "nAx: 0.12 en la edad 0, 1.5 en 1-4 y 2.5 en los demás grupos (los valores del Cuadro 4.3 del ejemplo de la profesora). En el grupo abierto 75+: nqx = 1 y nAx = 1/nmx.", 2, 15)
    encabezado(ws, T0 - 1, 2, ["Edad", "x", "n", "P (población)", "D (defunciones)", "nmx", "nqx", "npx", "lx", "ndx", "nAx", "nLx", "Tx", "ex"])
    for k, et in enumerate(ETIQ):
        f = T0 + k
        fi = I0 + k
        ab = k == 16
        fl = cuerpo(f)
        pon(ws, f"B{f}", f"={CF}B{fi}", f_n, fill=fl, alin=Alignment(horizontal="center"))
        pon(ws, f"C{f}", ([0, 1] + list(range(5, 80, 5)))[k], f_n, "0", fl)
        pon(ws, f"D{f}", "abierto" if ab else (1 if k == 0 else 4 if k == 1 else 5), f_n, "0", fl, der)
        pon(ws, f"E{f}", f"={CF}{pc}{fi}", f_n, "#,##0", fl)
        pon(ws, f"F{f}", f"={CF}{dc}{fi}", f_n, "#,##0", fl)
        pon(ws, f"G{f}", f"=F{f}/E{f}", f_n, "0.00000", fl)                                   # nmx
        pon(ws, f"H{f}", 1 if ab else f"=(2*D{f}*G{f})/(2+D{f}*G{f})", f_n, "0.00000", fl)  # nqx
        pon(ws, f"I{f}", f"=1-H{f}", f_n, "0.00000", fl)                                      # npx
        pon(ws, f"J{f}", "=$D$5" if k == 0 else f"=J{f - 1}*I{f - 1}", f_n, "#,##0", fl)      # lx
        pon(ws, f"K{f}", f"=J{f}*H{f}", f_n, "#,##0", fl)                                     # ndx
        pon(ws, f"L{f}", f"=1/G{f}" if ab else [0.12, 1.5][k] if k < 2 else 2.5, f_n, "0.00", fl)   # nAx
        pon(ws, f"M{f}", f"=K{f}*L{f}" if ab else f"=J{f + 1}*D{f}+K{f}*L{f}", f_n, "#,##0", fl)    # nLx
        pon(ws, f"N{f}", f"=SUM(M{f}:$M${T1})", f_n, "#,##0", fl)                              # Tx
        pon(ws, f"O{f}", f"=N{f}/J{f}", f_b, "0.00", fl)                                       # ex
    pon(ws, f"B{T1 + 2}", "Control: suma de ndx − l0", f_n, fill=VERDE_CLA, caja=False)
    ws.merge_cells(f"B{T1 + 2}:D{T1 + 2}")
    pon(ws, f"E{T1 + 2}", f"=ROUND(SUM(K{T0}:K{T1})-D5,6)", f_b, "0.0", VERDE_CLA)
    nota(ws, T1 + 2, "← debe ser 0", 6, 8)
    pon(ws, f"B{T1 + 3}", "q0 (por mil)", f_n, fill=VERDE_CLA, caja=False)
    ws.merge_cells(f"B{T1 + 3}:D{T1 + 3}")
    pon(ws, f"E{T1 + 3}", f"=H{T0}*1000", f_b, "0.00", VERDE_CLA)
    anchos(ws, {"A": 3, "B": 8, "C": 6, "D": 9, "E": 15, "F": 15, "G": 10, "H": 10, "I": 10, "J": 11, "K": 11, "L": 8, "M": 12, "N": 13, "O": 9, "P": 3})
    ws.freeze_panes = f"C{T0}"

# =====================================================================
# Gráficos y resumen
# =====================================================================
ws = gr
ws.sheet_view.showGridLines = False
ws.sheet_view.zoomScale = 85
fondo(ws, 100, 20)
banda(ws, 2, "GRÁFICOS Y RESUMEN", 2, 19, f_tit, 30)
banda(ws, 3, "Tablas transversales abreviadas, México 2025: indicadores, gráficos y notas", 2, 19, f_sub, 18)
banda(ws, 5, "Indicadores principales (se recalculan desde las tablas)", 2, 5)
encabezado(ws, 6, 2, ["Indicador", "Hombres", "Mujeres", "Ambos sexos"], 20)
ind = [
    ("Esperanza de vida al nacer, e0 (años)", lambda h: f"'{h}'!O{T0}", "0.00"),
    ("Esperanza de vida a los 5 años, e5", lambda h: f"'{h}'!O{T0 + 2}", "0.00"),
    ("Esperanza de vida a los 65 años, e65", lambda h: f"'{h}'!O{T0 + 13}", "0.00"),
    ("Esperanza de vida a los 75 años, e75", lambda h: f"'{h}'!O{T1}", "0.00"),
    ("Mortalidad infantil, q0 por mil", lambda h: f"'{h}'!H{T0}*1000", "0.00"),
    ("Mortalidad de menores de 5 años, 5q0 por mil", lambda h: f"(1-'{h}'!I{T0}*'{h}'!I{T0 + 1})*1000", "0.00"),
    ("Sobrevivientes a los 65 años (de 100,000), l65", lambda h: f"'{h}'!J{T0 + 13}", "#,##0"),
    ("Sobrevivientes a los 75 años (de 100,000), l75", lambda h: f"'{h}'!J{T1}", "#,##0"),
]
for j, (t, fn, fmt) in enumerate(ind):
    f = 7 + j
    pon(ws, f"B{f}", t, f_n, fill=cuerpo(f))
    for col, s in zip("CDE", "HMA"):
        pon(ws, f"{col}{f}", "=" + fn(HOJAS[s]), f_n, fmt, cuerpo(f))
R = 7 + len(ind) + 1
pon(ws, f"B{R}", "e0 de CONAPO, proyecciones 2025 (referencia)", f_n, fill=VERDE_CLA)
for col, v in zip("CDE", (72.57, 79.04, 75.67)):
    pon(ws, f"{col}{R}", v, f_in, "0.00", AMARILLO)
pon(ws, f"B{R + 1}", "Diferencia: estas tablas − CONAPO (años)", f_n, fill=GRIS)
for col in "CDE":
    pon(ws, f"{col}{R + 1}", f"={col}7-{col}{R}", f_n, "0.00", GRIS)
nota(ws, R + 2, "CONAPO: base de indicadores demográficos del repositorio (EVH, EVM, EV, año 2025, total nacional), capturada a mano.", 2, 5, 30)

# gráficos nativos
cats = Reference(th["H"], min_col=2, min_row=T0, max_row=T1)
cats16 = Reference(th["H"], min_col=2, min_row=T0, max_row=T1 - 1)
COLORES = {"H": "1F77B4", "M": "D95F02", "A": "595959"}


def grafico(titulo, col, ytit, ancla, hasta, log=False, ymin=None):
    ch = LineChart()
    ch.title = titulo
    ch.height, ch.width = 7.5, 15
    ch.y_axis.title = ytit
    ch.x_axis.title = "Edad (grupo)"
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    if log:
        ch.y_axis.scaling.logBase = 10
    if ymin is not None:
        ch.y_axis.scaling.min = ymin
    for s in "HMA":
        ref = Reference(th[s], min_col=col, min_row=T0, max_row=hasta)
        ser = Series(ref, title=NOMBRE[s].capitalize())
        ser.graphicalProperties.line.solidFill = COLORES[s]
        ser.graphicalProperties.line.width = 22000
        ser.marker.symbol = "circle"
        ser.marker.size = 5
        ser.marker.graphicalProperties.solidFill = COLORES[s]
        ser.marker.graphicalProperties.line.solidFill = COLORES[s]
        ser.smooth = False
        ch.series.append(ser)
    ch.set_categories(cats if hasta == T1 else cats16)
    ch.legend.position = "b"
    ws.add_chart(ch, ancla)


grafico("Sobrevivientes lx de la cohorte ficticia (eje desde 50,000)", 10, "lx", "G5", T1, ymin=50000)
grafico("Esperanza de vida ex por edad", 15, "ex (años)", "G21", T1)
grafico("Probabilidad de morir nqx (escala log, sin 75+)", 8, "nqx", "G37", T1 - 1, log=True)

N0 = R + 5
banda(ws, N0, "Método, supuestos y alertas", 2, 5)
notas = [
    "1. Tabla transversal (de periodo): las tasas de mortalidad por edad observadas en 2025 se aplican a una cohorte ficticia de 100,000 nacimientos.",
    "2. Grupos 0, 1-4, 5-9, ..., 70-74 y 75+ (abierto): el máximo detalle que permite la población de la EIC (llega a 75 años y más).",
    "3. Las defunciones de sexo no especificado (529) y de edad no especificada (4,016) se reparten en proporción (hoja 'Cohorte ficticio', bloque A).",
    "4. La EIC junta 0-4 años: el año 0 se separa suponiendo cohortes de nacimiento con crecimiento constante (bloque B). Es un supuesto y afecta sobre todo a q0.",
    "5. nqx = (2·n·nmx) / (2 + n·nmx) y nAx = 0.12, 1.5 y 2.5, como en el ejemplo de la profesora (Cuadro 4.3). En 75+ se usa nAx = 1/nmx, muy sensible a las defunciones de adultos mayores.",
    "6. ALERTA: la e0 sale 4 a 5 años por encima de la de CONAPO. En la captura de INEGI de la hoja 'Fuentes de datos' se lee que las cifras de 2025 son "
    "preliminares y aún no concluyen los procesos de generación (son defunciones por año de registro). Las 762,494 defunciones del archivo quedan por debajo de las "
    "836,585 que proyecta CONAPO para 2025, y las 12,207 defunciones infantiles son pocas para ~2.0 millones de nacimientos. Con datos definitivos la e0 bajaría.",
]
for j, t in enumerate(notas):
    nota(ws, N0 + 1 + j, t, 2, 5, 46 if j < 5 else 118)
anchos(ws, {"A": 3, "B": 46, "C": 14, "D": 14, "E": 14, "F": 3})
for c in range(7, 20):
    ws.column_dimensions[get_column_letter(c)].width = 9

wb.calculation.fullCalcOnLoad = True
SALIDA.parent.mkdir(parents=True, exist_ok=True)
wb.save(SALIDA)
print("Libro (sin terminar) guardado en", SALIDA)
