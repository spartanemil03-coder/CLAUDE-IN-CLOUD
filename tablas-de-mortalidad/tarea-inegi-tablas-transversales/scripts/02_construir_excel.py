"""Construye el libro de Excel con las tablas transversales abreviadas (hombres, mujeres, ambos sexos) con fórmulas vivas.

Hojas: Resumen, Fuentes (datos de INEGI tal cual), Insumos (reparto de 'no especificado', separación del año 0 y agrupación),
Hombres, Mujeres, Ambos sexos (tablas abreviadas). Después hay que recalcular con LibreOffice (ver README).

Uso:  python3 scripts/02_construir_excel.py [salida.xlsx]
"""
import sys
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

RAIZ = Path(__file__).resolve().parent.parent
SALIDA = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "resultados" / "Tablas_transversales_abreviadas_INEGI_2025.xlsx"

pob = pd.read_csv(RAIZ / "datos" / "inegi_eic2025_poblacion_por_edad_sexo.csv")
dfn = pd.read_csv(RAIZ / "datos" / "inegi_mortalidad_defunciones_por_edad_sexo.csv")

# ---------- estilos ----------
FUENTE = "Arial"
AZUL, VERDE, NEGRO = "0000FF", "008000", "000000"
f_norm = Font(name=FUENTE, size=10)
f_neg = Font(name=FUENTE, size=10, bold=True)
f_in = Font(name=FUENTE, size=10, color=AZUL)
f_link = Font(name=FUENTE, size=10, color=VERDE)
f_tit = Font(name=FUENTE, size=14, bold=True, color="1F3864")
f_sub = Font(name=FUENTE, size=10, italic=True, color="595959")
f_enc = Font(name=FUENTE, size=10, bold=True, color="FFFFFF")
rel_enc = PatternFill("solid", fgColor="1F3864")
rel_sec = PatternFill("solid", fgColor="D9E2F3")
rel_amar = PatternFill("solid", fgColor="FFFF00")
rel_gris = PatternFill("solid", fgColor="F2F2F2")
borde = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
izq = Alignment(horizontal="left", vertical="center", wrap_text=True)


def pon(ws, ref, valor, fuente=f_norm, fmt=None, rel=None, alin=None, caja=True):
    c = ws[ref]
    c.value = valor
    c.font = fuente
    if fmt:
        c.number_format = fmt
    if rel:
        c.fill = rel
    if alin:
        c.alignment = alin
    if caja:
        c.border = borde
    return c


def encabezado(ws, fila, col0, titulos):
    for i, t in enumerate(titulos):
        pon(ws, f"{get_column_letter(col0 + i)}{fila}", t, f_enc, rel=rel_enc, alin=centro)


def anchos(ws, d):
    for k, v in d.items():
        ws.column_dimensions[k].width = v


wb = Workbook()

# =====================================================================
# Fuentes
# =====================================================================
fu = wb.active
fu.title = "Fuentes"
fu.sheet_view.showGridLines = False
pon(fu, "B1", "Fuentes de consulta (INEGI)", f_tit, caja=False)
pon(fu, "B2", "Datos tal cual vienen en 'Fuentes de consulta.xlsx' (la que entregó la maestra). Letra azul = dato de entrada.", f_sub, caja=False)
pon(fu, "B4", "1. Población por grupos quinquenales de edad y sexo: Encuesta Intercensal (EIC) 2025, total nacional, viviendas particulares habitadas", f_neg, caja=False)
pon(fu, "B5", "https://www.inegi.org.mx/programas/eic/2025/#tabulados", f_sub, caja=False)
encabezado(fu, 6, 2, ["Grupo de edad", "Población total", "Mujeres", "Hombres"])
F_POB0 = 7                                                    # primera fila de datos de población
for i, r in pob.iterrows():
    f = F_POB0 + i
    pon(fu, f"B{f}", r.grupo_edad.replace(" años", "").replace(" y más", "+").replace(" y más", "+"), f_in)
    pon(fu, f"C{f}", int(r.total), f_in, "#,##0")
    pon(fu, f"D{f}", int(r.mujeres), f_in, "#,##0")
    pon(fu, f"E{f}", int(r.hombres), f_in, "#,##0")
F_POBN = F_POB0 + len(pob) - 1                                # 75+
fila = F_POBN + 1
pon(fu, f"B{fila}", "Total", f_neg)
for col in "CDE":
    pon(fu, f"{col}{fila}", f"=SUM({col}{F_POB0}:{col}{F_POBN})", f_neg, "#,##0")
F_POBT = fila

fila += 3
pon(fu, f"B{fila - 1}", "2. Defunciones por grupos de edad y sexo: Mortalidad, INEGI (el archivo no indica el año de las defunciones; aquí se tratan como las de 2025)", f_neg, caja=False)
pon(fu, f"B{fila}", "https://www.inegi.org.mx/sistemas/olap/Proyectos/bd/continuas/mortalidad/MortalidadGeneral.asp?", f_sub, caja=False)
fila += 1
encabezado(fu, fila, 2, ["Grupo de edad", "Total", "Hombres", "Mujeres", "Sexo no especificado"])
F_DEF0 = fila + 1
for i, r in dfn.iterrows():
    f = F_DEF0 + i
    nombre = {"Menores de 1 año": "<1"}.get(r.grupo_edad, r.grupo_edad.replace(" años", "").replace(" y más", "+"))
    pon(fu, f"B{f}", nombre, f_in)
    for col, k in zip("CDEF", ("total", "hombres", "mujeres", "sexo_no_especificado")):
        pon(fu, f"{col}{f}", int(r[k]), f_in, "#,##0")
F_DEFTOT = F_DEF0                                             # fila "Total"
F_DEFA0, F_DEFA1 = F_DEF0 + 1, F_DEF0 + 19                    # <1 ... 85+  (19 grupos)
F_DEFNE = F_DEF0 + 20                                         # edad no especificada
fu.cell(row=F_DEFTOT, column=2).font = f_in
anchos(fu, {"A": 2, "B": 22, "C": 18, "D": 14, "E": 14, "F": 20})

# =====================================================================
# Insumos
# =====================================================================
ins = wb.create_sheet("Insumos")
ins.sheet_view.showGridLines = False
pon(ins, "B1", "Insumos: del dato de INEGI a los 17 grupos de edad de la tabla", f_tit, caja=False)
pon(ins, "B2", "Letra negra = fórmula; letra verde = vínculo a otra hoja; letra azul sobre amarillo = supuesto que se puede cambiar.", f_sub, caja=False)

pon(ins, "B4", "Radix de la cohorte ficticia (l0)", f_neg)
pon(ins, "C4", 100000, f_in, "#,##0", rel_amar)
pon(ins, "D4", "Supuesto: se parte de 100,000 nacimientos hipotéticos (cohorte ficticia), como en la plantilla de la maestra.", f_sub, caja=False)

# ---- A. defunciones ----
pon(ins, "B6", "A. Defunciones: reparto de 'sexo no especificado' y de 'edad no especificada'", f_neg, caja=False)
pon(ins, "B7", "Paso 1: el sexo no especificado se reparte entre hombres y mujeres en proporción a sus defunciones, edad por edad.  "
    "Paso 2: las defunciones de edad no especificada se reparten entre las edades en proporción a las ya ajustadas, por sexo.", f_sub, caja=False)
encabezado(ins, 8, 2, ["Grupo de edad", "Hombres (INEGI)", "Mujeres (INEGI)", "Sexo no esp.", "Hombres (paso 1)", "Mujeres (paso 1)",
                       "Hombres (paso 2)", "Mujeres (paso 2)"])
A0 = 9                                                        # primera fila de edades (<1)
A1 = A0 + 18                                                  # 85+
ANE = A1 + 1                                                  # fila edad no especificada
ATOT = ANE + 1
for k in range(20):                                           # 19 grupos + edad NE
    f = A0 + k
    fu_f = F_DEFA0 + k if k < 19 else F_DEFNE
    pon(ins, f"B{f}", f"=Fuentes!B{fu_f}" if k < 19 else "Edad no especificada", f_link if k < 19 else f_norm)
    pon(ins, f"C{f}", f"=Fuentes!D{fu_f}", f_link, "#,##0")
    pon(ins, f"D{f}", f"=Fuentes!E{fu_f}", f_link, "#,##0")
    pon(ins, f"E{f}", f"=Fuentes!F{fu_f}", f_link, "#,##0")
    pon(ins, f"F{f}", f"=C{f}+E{f}*C{f}/(C{f}+D{f})", f_norm, "#,##0.0")
    pon(ins, f"G{f}", f"=D{f}+E{f}*D{f}/(C{f}+D{f})", f_norm, "#,##0.0")
    if k < 19:
        pon(ins, f"H{f}", f"=F{f}*$F${ATOT + 1}", f_norm, "#,##0.0")
        pon(ins, f"I{f}", f"=G{f}*$G${ATOT + 1}", f_norm, "#,##0.0")
    else:
        pon(ins, f"H{f}", None)
        pon(ins, f"I{f}", None)
pon(ins, f"B{ATOT}", "Total", f_neg)
for col in "CDEFGHI":
    pon(ins, f"{col}{ATOT}", f"=SUM({col}{A0}:{col}{ANE})", f_neg, "#,##0.0" if col in "FGHI" else "#,##0")
pon(ins, f"B{ATOT + 1}", "Factor del paso 2 (todas las defunciones / las de edad conocida)", f_norm)
pon(ins, f"F{ATOT + 1}", f"=SUM(F{A0}:F{ANE})/SUM(F{A0}:F{A1})", f_norm, "0.000000")
pon(ins, f"G{ATOT + 1}", f"=SUM(G{A0}:G{ANE})/SUM(G{A0}:G{A1})", f_norm, "0.000000")
pon(ins, f"B{ATOT + 2}", "Control: hombres + mujeres (paso 2) menos el total de INEGI (debe ser 0)", f_norm)
pon(ins, f"H{ATOT + 2}", f"=ROUND(SUM(H{A0}:I{A1})-Fuentes!C{F_DEFTOT},6)", f_neg, "0.0")

# ---- B. población: separar el año 0 ----
B0 = ATOT + 5
pon(ins, f"B{B0 - 1}", "B. Población: la EIC trae 0-4 años junto; hay que separar los menores de 1 año", f_neg, caja=False)
pon(ins, f"B{B0}", "Supuesto: las cohortes de nacimiento crecen hacia atrás a tasa constante r (r = ln(P5-9 / P0-4) / 5). Entonces la fracción de 0-4 que "
    "tiene menos de 1 año es (e^r - 1) / (e^(5r) - 1).", f_sub, caja=False)
encabezado(ins, B0 + 1, 2, ["", "Hombres", "Mujeres"])
filas_b = ["Población 0-4", "Población 5-9", "r (tasa de crecimiento de las cohortes)", "Fracción de 0-4 con menos de 1 año",
           "Población de menores de 1 año", "Población de 1 a 4 años"]
for j, t in enumerate(filas_b):
    pon(ins, f"B{B0 + 2 + j}", t, f_norm)
for col, fu_col in (("C", "E"), ("D", "D")):                   # hombres = columna E de Fuentes; mujeres = D
    r0 = B0 + 2
    pon(ins, f"{col}{r0}", f"=Fuentes!{fu_col}{F_POB0}", f_link, "#,##0")
    pon(ins, f"{col}{r0 + 1}", f"=Fuentes!{fu_col}{F_POB0 + 1}", f_link, "#,##0")
    pon(ins, f"{col}{r0 + 2}", f"=LN({col}{r0 + 1}/{col}{r0})/5", f_norm, "0.00000")
    pon(ins, f"{col}{r0 + 3}", f"=(EXP({col}{r0 + 2})-1)/(EXP(5*{col}{r0 + 2})-1)", f_norm, "0.0000")
    pon(ins, f"{col}{r0 + 4}", f"={col}{r0}*{col}{r0 + 3}", f_norm, "#,##0")
    pon(ins, f"{col}{r0 + 5}", f"={col}{r0}-{col}{r0 + 4}", f_norm, "#,##0")
PB_P0, PB_P14 = B0 + 6, B0 + 7                                 # filas de P(<1) y P(1-4)

# ---- C. insumos por grupo de la tabla ----
C0 = B0 + 11
pon(ins, f"B{C0 - 2}", "C. Insumos de las tablas: 17 grupos (0, 1-4, 5-9, ..., 70-74, 75+)", f_neg, caja=False)
pon(ins, f"B{C0 - 1}", "En defunciones, 75-79, 80-84 y 85+ se suman en 75+ porque la EIC llega solo a 75 años y más.", f_sub, caja=False)
encabezado(ins, C0, 2, ["Grupo de edad", "Población hombres", "Defunciones hombres", "Población mujeres", "Defunciones mujeres",
                        "Población ambos sexos", "Defunciones ambos sexos"])
ETIQ = ["0", "1-4"] + [f"{x}-{x + 4}" for x in range(5, 75, 5)] + ["75+"]
I0 = C0 + 1
for k, et in enumerate(ETIQ):
    f = I0 + k
    pon(ins, f"B{f}", et, f_norm, alin=Alignment(horizontal="center"))
    # población
    if k == 0:
        ph, pm = f"=C{PB_P0}", f"=D{PB_P0}"
    elif k == 1:
        ph, pm = f"=C{PB_P14}", f"=D{PB_P14}"
    else:                                                      # fila de Fuentes: 5-9 es F_POB0+1 ... 75+ es F_POB0+15
        ph, pm = f"=Fuentes!E{F_POB0 + k - 1}", f"=Fuentes!D{F_POB0 + k - 1}"
    pon(ins, f"C{f}", ph, f_norm if k < 2 else f_link, "#,##0")
    pon(ins, f"E{f}", pm, f_norm if k < 2 else f_link, "#,##0")
    # defunciones
    if k < 16:
        dh, dm = f"=H{A0 + k}", f"=I{A0 + k}"
    else:
        dh, dm = f"=SUM(H{A0 + 16}:H{A1})", f"=SUM(I{A0 + 16}:I{A1})"
    pon(ins, f"D{f}", dh, f_norm, "#,##0.0")
    pon(ins, f"F{f}", dm, f_norm, "#,##0.0")
    pon(ins, f"G{f}", f"=C{f}+E{f}", f_norm, "#,##0")
    pon(ins, f"H{f}", f"=D{f}+F{f}", f_norm, "#,##0.0")
I1 = I0 + 16
pon(ins, f"B{I1 + 1}", "Total", f_neg)
for col in "CDEFGH":
    pon(ins, f"{col}{I1 + 1}", f"=SUM({col}{I0}:{col}{I1})", f_neg, "#,##0")
anchos(ins, {"A": 2, "B": 44, "C": 18, "D": 18, "E": 18, "F": 18, "G": 18, "H": 18, "I": 18})
ins.row_dimensions[8].height = 30
ins.row_dimensions[C0].height = 30

# =====================================================================
# Tablas abreviadas
# =====================================================================
T0 = 9                                                          # primera fila de la tabla
T1 = T0 + 16                                                    # última (75+)
COL_INS = {"Hombres": ("C", "D"), "Mujeres": ("E", "F"), "Ambos sexos": ("G", "H")}


def a_coale_demeny(sexo, fila, g):
    """nax de la fila 0 (a0) o 1-4 según Coale-Demeny; g = celda de nM0 (tasa infantil). Ambos sexos: promedio de hombres y mujeres."""
    if fila == 0:
        h = f"IF({g}>=0.107,0.33,0.045+2.684*{g})"
        m = f"IF({g}>=0.107,0.35,0.053+2.8*{g})"
    else:
        h = f"IF({g}>=0.107,1.352,1.651-2.816*{g})"
        m = f"IF({g}>=0.107,1.361,1.522-1.518*{g})"
    return {"Hombres": f"={h}", "Mujeres": f"={m}", "Ambos sexos": f"=({h}+{m})/2"}[sexo]


for hoja in ("Hombres", "Mujeres", "Ambos sexos"):
    ws = wb.create_sheet(hoja)
    ws.sheet_view.showGridLines = False
    pc, dc = COL_INS[hoja]
    pon(ws, "B1", f"Tabla de mortalidad transversal abreviada de México, 2025: {hoja.lower()}", f_tit, caja=False)
    pon(ws, "B2", "Cohorte ficticia de 100,000 nacimientos expuesta a las tasas de mortalidad por edad observadas en México (INEGI: EIC 2025 y defunciones).",
        f_sub, caja=False)
    pon(ws, "B4", "Radix l0", f_neg)
    pon(ws, "E4", "=Insumos!C4", f_link, "#,##0")
    pon(ws, "B5", "Fórmulas: nMx = D/P;  nqx = n·nMx / (1 + (n − nax)·nMx);  lx+n = lx·(1 − nqx);  ndx = lx − lx+n;  nLx = n·lx+n + nax·ndx;  Tx = Σ nLx;  ex = Tx/lx.  "
        "Grupo abierto 75+: q = 1, L = l/M.", f_sub, caja=False)
    pon(ws, "B6", "nax: 0 y 1-4 de Coale-Demeny (según nM0 y sexo; en ambos sexos, promedio de los dos); resto n/2. Valores sin redondear (se muestran sin decimales).",
        f_sub, caja=False)
    encabezado(ws, T0 - 1, 2, ["Edad", "x", "n", "P (población)", "D (defunciones)", "nMx", "nax", "nqx", "npx", "lx", "ndx", "nLx", "Tx", "ex"])
    ws.row_dimensions[T0 - 1].height = 30
    for k, et in enumerate(ETIQ):
        f = T0 + k
        fi = I0 + k
        abierto = k == 16
        pon(ws, f"B{f}", f"=Insumos!B{fi}", f_link, alin=Alignment(horizontal="center"))
        x = [0, 1] + list(range(5, 80, 5))[0:15]
        pon(ws, f"C{f}", x[k], f_in, "0")
        pon(ws, f"D{f}", "abierto" if abierto else (1 if k == 0 else 4 if k == 1 else 5), f_in, "0", alin=Alignment(horizontal="right"))
        pon(ws, f"E{f}", f"=Insumos!{pc}{fi}", f_link, "#,##0")
        pon(ws, f"F{f}", f"=Insumos!{dc}{fi}", f_link, "#,##0")
        pon(ws, f"G{f}", f"=F{f}/E{f}", f_norm, "0.00000")
        if k < 2:
            pon(ws, f"H{f}", a_coale_demeny(hoja, k, f"$G${T0}"), f_norm, "0.0000")   # ambas dependen de M0 (infantil)
        elif not abierto:
            pon(ws, f"H{f}", f"=D{f}/2", f_norm, "0.0000")
        else:
            pon(ws, f"H{f}", "-", f_norm, alin=Alignment(horizontal="right"))
        pon(ws, f"I{f}", 1 if abierto else f"=D{f}*G{f}/(1+(D{f}-H{f})*G{f})", f_in if abierto else f_norm, "0.00000")
        pon(ws, f"J{f}", f"=1-I{f}", f_norm, "0.00000")
        pon(ws, f"K{f}", "=$E$4" if k == 0 else f"=K{f - 1}-L{f - 1}", f_norm, "#,##0")
        pon(ws, f"L{f}", f"=K{f}*I{f}", f_norm, "#,##0")
        pon(ws, f"M{f}", f"=K{f}/G{f}" if abierto else f"=D{f}*K{f + 1}+H{f}*L{f}", f_norm, "#,##0")
        pon(ws, f"N{f}", f"=SUM(M{f}:$M${T1})", f_norm, "#,##0")
        pon(ws, f"O{f}", f"=N{f}/K{f}", f_norm, "0.00")
    pon(ws, f"B{T1 + 2}", "Control: suma de ndx menos l0 (debe ser 0)", f_norm, caja=False)
    pon(ws, f"L{T1 + 2}", f"=ROUND(SUM(L{T0}:L{T1})-E4,6)", f_neg, "0.0")
    pon(ws, f"B{T1 + 3}", "q0: mortalidad infantil de la tabla (por mil)", f_norm, caja=False)
    pon(ws, f"L{T1 + 3}", f"=I{T0}*1000", f_neg, "0.00")
    anchos(ws, {"A": 2, "B": 8, "C": 6, "D": 9, "E": 16, "F": 16, "G": 10, "H": 9, "I": 10, "J": 10, "K": 11, "L": 11, "M": 12, "N": 13, "O": 8})
    ws.freeze_panes = f"C{T0}"

# =====================================================================
# Resumen
# =====================================================================
rs = wb.create_sheet("Resumen", 0)
rs.sheet_view.showGridLines = False
pon(rs, "B1", "Tablas transversales abreviadas: INEGI, México 2025", f_tit, caja=False)
pon(rs, "B2", "U2TC: INEGI Tablas transversales. Cohorte ficticia de 100,000 nacimientos con las tasas de mortalidad de 2025 (población: EIC 2025; defunciones: INEGI).",
    f_sub, caja=False)

pon(rs, "B4", "Indicadores principales (todo se recalcula desde las tablas)", f_neg, caja=False)
encabezado(rs, 5, 2, ["Indicador", "Hombres", "Mujeres", "Ambos sexos"])
ind = [
    ("Esperanza de vida al nacer, e0 (años)", lambda h: f"'{h}'!O{T0}", "0.00"),
    ("Esperanza de vida a los 5 años, e5", lambda h: f"'{h}'!O{T0 + 2}", "0.00"),
    ("Esperanza de vida a los 65 años, e65", lambda h: f"'{h}'!O{T0 + 13}", "0.00"),
    ("Esperanza de vida a los 75 años, e75", lambda h: f"'{h}'!O{T1}", "0.00"),
    ("Mortalidad infantil: q0 por mil", lambda h: f"'{h}'!I{T0}*1000", "0.00"),
    ("Mortalidad de menores de 5 años: 5q0 por mil", lambda h: f"(1-'{h}'!J{T0}*'{h}'!J{T0 + 1})*1000", "0.00"),
    ("Sobrevivientes a los 65 años (de 100,000): l65", lambda h: f"'{h}'!K{T0 + 13}", "#,##0"),
    ("Sobrevivientes a los 75 años (de 100,000): l75", lambda h: f"'{h}'!K{T1}", "#,##0"),
]
for j, (t, f, fmt) in enumerate(ind):
    fila = 6 + j
    pon(rs, f"B{fila}", t, f_norm)
    for col, hoja in zip("CDE", ("Hombres", "Mujeres", "Ambos sexos")):
        pon(rs, f"{col}{fila}", "=" + f(hoja), f_link, fmt)

fila = 6 + len(ind) + 1
pon(rs, f"B{fila}", "Referencia externa: e0 de CONAPO, proyecciones 2025 (en el repositorio: tarea-conapo-generacion-1970/bases_de_datos_CONAPO)", f_neg, caja=False)
encabezado(rs, fila + 1, 2, ["", "Hombres", "Mujeres", "Ambos sexos"])
pon(rs, f"B{fila + 2}", "e0 CONAPO 2025 (años)", f_norm)
for col, v in zip("CDE", (72.57, 79.04, 75.67)):
    pon(rs, f"{col}{fila + 2}", v, f_in, "0.00")
pon(rs, f"B{fila + 3}", "Diferencia: tabla de este libro − CONAPO (años)", f_norm)
for col in "CDE":
    pon(rs, f"{col}{fila + 3}", f"={col}6-{col}{fila + 2}", f_norm, "0.00")
pon(rs, f"B{fila + 4}", "Dato de CONAPO capturado a mano de su base de indicadores demográficos (EVH, EVM, EV, año 2025, total nacional).", f_sub, caja=False)

nf = fila + 6
pon(rs, f"B{nf}", "Método y supuestos", f_neg, caja=False)
notas = [
    "1. Tabla transversal (de periodo): las tasas de mortalidad por edad observadas en un solo año se aplican a una cohorte ficticia de 100,000 nacimientos (l0 en Insumos!C4).",
    "2. Grupos de edad: 0, 1-4, 5-9, ..., 70-74 y 75+ (abierto). Es el máximo detalle que permite la población de la EIC, que termina en 75 años y más.",
    "3. Las defunciones de sexo no especificado (529) y de edad no especificada (4,016) se reparten proporcionalmente (hoja Insumos, bloque A).",
    "4. La EIC junta 0-4 años: el año 0 se separa suponiendo que las cohortes de nacimiento crecen a tasa constante (hoja Insumos, bloque B). Es un supuesto; afecta sobre todo a q0.",
    "5. nqx = n·nMx / (1 + (n − nax)·nMx), con nax = n/2 salvo 0 y 1-4 (Coale-Demeny). En 75+ se usa e75 = 1/M, que es sensible a la cobertura de defunciones de adultos mayores.",
    "6. Las defunciones del archivo no traen año ni indican si son ocurridas o registradas; se usan tal cual como defunciones de 2025.",
    "7. Alerta: la e0 que sale es 4 a 5 años mayor que la de CONAPO 2025. Las defunciones infantiles del archivo (12,207) son bajas frente a los ~2.0 millones de nacimientos de "
    "CONAPO, y las tasas de adultos mayores parecen bajas. Hay que confirmar con la maestra el año y la cobertura de las defunciones.",
]
for j, t in enumerate(notas):
    pon(rs, f"B{nf + 1 + j}", t, f_norm, alin=izq, caja=False)
    rs.merge_cells(f"B{nf + 1 + j}:E{nf + 1 + j}")
    rs.row_dimensions[nf + 1 + j].height = 42
lg = nf + len(notas) + 2
pon(rs, f"B{lg}", "Leyenda de colores: azul = dato de entrada o supuesto; negro = fórmula; verde = vínculo a otra hoja; amarillo = supuesto que se puede cambiar.", f_sub, caja=False)
anchos(rs, {"A": 2, "B": 62, "C": 16, "D": 16, "E": 16})

SALIDA.parent.mkdir(parents=True, exist_ok=True)
wb.save(SALIDA)
print("Libro guardado en", SALIDA)
