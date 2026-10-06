"""Construye el libro de Excel con las tablas de mortalidad por generación (México, ONU WPP 2024).

Todas las tablas de vida son fórmulas vivas: los únicos valores fijos son los datos de la ONU
(hoja Datos_ONU, ONU_qx, ONU_e0), el año de generación y el radix.

Uso:  python3 scripts/02_construir_excel.py
Luego recalcular:  python3 <skill xlsx>/scripts/recalc.py resultados/tablas_mortalidad_generaciones_mexico.xlsx
"""
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference, ScatterChart, Series, BarChart
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "datos"
SALIDA = RAIZ / "resultados" / "tablas_mortalidad_generaciones_mexico.xlsx"

GENERACIONES = [1950, 1960, 1970, 1980, 1990, 2000]
SEXOS = [("H", "Hombres"), ("M", "Mujeres"), ("A", "Ambos sexos")]
ANIOS = list(range(1950, 2101))
EDADES = list(range(0, 101))
NA = len(EDADES)          # 101 filas por bloque
C0, C1 = 2, 2 + len(ANIOS) - 1   # columnas B..EX = años 1950..2100
RAMPA_GEN = ["#8fb8ea", "#6a9ede", "#3f82d0", "#2a64b0", "#1d4a87", "#12315c"]   # azul secuencial
COLOR_SEXO = {"H": "2A78D6", "M": "EB6834", "A": "1BAF7A"}

FUENTE = "Arial"
F = lambda **k: Font(name=FUENTE, size=k.pop("size", 10), **k)
AZUL, VERDE, NEGRO = "0000FF", "008000", "000000"
GRIS = PatternFill("solid", fgColor="F2F2F2")
FILL_SEXO = {"H": PatternFill("solid", fgColor="DCE9F9"), "M": PatternFill("solid", fgColor="FBE3D8"),
             "A": PatternFill("solid", fgColor="D5F1E6")}
FILL_TIT = PatternFill("solid", fgColor="1F3864")
FINO = Side(style="thin", color="BFBFBF")
BORDE = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)

wb = Workbook()


# ----------------------------------------------------------------------------- utilidades
def titulo(ws, texto, sub=None):
    ws["A1"] = texto
    ws["A1"].font = F(size=14, bold=True)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = F(italic=True, color="595959")
    ws.sheet_view.showGridLines = False


def cab(c, texto, fill=GRIS):
    c.value = texto
    c.font = F(bold=True)
    c.fill = fill
    c.alignment = CENTRO
    c.border = BORDE


def bloque_matriz(ws, top, nombre, matriz, color_fuente=AZUL, fmt="#,##0.000"):
    """Escribe una matriz edad x año. Devuelve (fila_encabezado, fila_datos_0, fila_datos_100)."""
    ws.cell(top, 1, nombre).font = F(bold=True, size=11)
    h = top + 1
    cab(ws.cell(h, 1), "Edad")
    for j, a in enumerate(ANIOS):
        cab(ws.cell(h, C0 + j), a)
    for i, e in enumerate(EDADES):
        c = ws.cell(h + 1 + i, 1, e)
        c.font = F(bold=True)
        if e == 100:
            c.number_format = '0"+"'
        for j, a in enumerate(ANIOS):
            c = ws.cell(h + 1 + i, C0 + j, float(matriz.loc[e, a]))
            c.font = F(color=color_fuente)
            c.number_format = fmt
    return h, h + 1, h + NA


def rango(hoja, col0, col1, f0, f1):
    return f"{hoja}!${L(col0)}${f0}:${L(col1)}${f1}"


# ----------------------------------------------------------------------------- datos de entrada
pop = pd.read_csv(DATOS / "mexico_PopulationBySingleAgeSex.csv")
dea = pd.read_csv(DATOS / "mexico_DeathsBySingleAgeSex.csv")
ind = pd.read_csv(DATOS / "mexico_DemographicIndicators.csv").set_index("Time")
tv = pd.read_csv(DATOS / "mexico_ONU_TablaVida_Completa.csv")
for d in (pop, dea):
    d["edad"] = d.AgeGrpStart.astype(int)
COL_POP = {"H": "PopMale", "M": "PopFemale", "A": "PopTotal"}
COL_DEA = {"H": "DeathMale", "M": "DeathFemale", "A": "DeathTotal"}
SEXO_TV = {"H": "Male", "M": "Female", "A": "Total"}
LEX = {"H": "LExMale", "M": "LExFemale", "A": "LEx"}


def matriz(df, col):
    m = df.pivot(index="edad", columns="Time", values=col)
    assert m.shape == (NA, len(ANIOS)) and not m.isna().any().any()
    return m


# ----------------------------------------------------------------------------- Leeme (primera hoja)
ws_l = wb.active
ws_l.title = "Leeme"

# ----------------------------------------------------------------------------- Datos_ONU
ws_d = wb.create_sheet("Datos_ONU")
titulo(ws_d, "Datos de entrada: ONU, World Population Prospects 2024 (variante media) - México",
       "Miles de personas. Población a mitad de año (1 de julio) y defunciones por edad simple (0-100+), 1950-2100. "
       "1950-2023 estimaciones; 2024-2100 proyección. Edad 100 = grupo abierto 100+.")
BLQ = {}   # (variable, sexo) -> (hdr, f0, f1)
top = 4
for var, dfv, cols, nom in (("Pob", pop, COL_POP, "Población a mitad de año"),
                            ("Def", dea, COL_DEA, "Defunciones")):
    for s, ns in SEXOS:
        BLQ[(var, s)] = bloque_matriz(ws_d, top, f"{nom} - {ns} (miles)", matriz(dfv, cols[s]))
        top += NA + 4
ws_d.freeze_panes = "B4"
ws_d.column_dimensions["A"].width = 8

# ----------------------------------------------------------------------------- ONU_qx
ws_q = wb.create_sheet("ONU_qx")
titulo(ws_q, "qx oficial de las tablas de vida completas de la ONU (WPP 2024, variante media) - México",
       "Probabilidad de morir entre la edad x y x+1, por año calendario. Solo se usa para validar las generaciones "
       "(hoja Validacion). Fuente: WPP2024_Life_Table_Complete_Medium_{Male,Female,Both}.")
BLQ_Q = {}
top = 4
for s, ns in SEXOS:
    t = tv[tv.Sex == SEXO_TV[s]].copy()
    t["edad"] = t.AgeGrpStart.astype(int)
    BLQ_Q[s] = bloque_matriz(ws_q, top, f"qx ONU - {ns}", matriz(t, "qx"), fmt="0.000000")
    top += NA + 4
ws_q.freeze_panes = "B4"
ws_q.column_dimensions["A"].width = 8

# ----------------------------------------------------------------------------- ONU_e0
ws_e = wb.create_sheet("ONU_e0")
titulo(ws_e, "Esperanza de vida al nacer oficial de la ONU por año calendario - México",
       "Años. Fuente: WPP2024_Demographic_Indicators_Medium (LEx, LExMale, LExFemale). Es una medida de periodo.")
for j, t in enumerate(["Año", "Hombres", "Mujeres", "Ambos sexos"]):
    cab(ws_e.cell(4, 1 + j), t)
for i, a in enumerate(ANIOS):
    r = 5 + i
    ws_e.cell(r, 1, a).font = F(bold=True)
    for j, s in enumerate("HMA"):
        c = ws_e.cell(r, 2 + j, float(ind.loc[a, LEX[s]]))
        c.font = F(color=AZUL)
        c.number_format = "0.00"
E0_F0, E0_F1 = 5, 5 + len(ANIOS) - 1
ws_e.freeze_panes = "A5"

# ----------------------------------------------------------------------------- Gen_XXXX
R0, R1 = 6, 6 + NA - 1         # filas de datos 6..106
OFF = dict(pob=0, dea=1, qx=2, lx=3, dx=4, Lx=5, Tx=6, ex=7)
BASE = {"H": 3, "M": 11, "A": 19}


def col(s, k):
    return L(BASE[s] + OFF[k])


def idx(var, s, edad, anio):
    h, f0, f1 = BLQ[(var, s)]
    return (f"=INDEX(Datos_ONU!${L(C0)}${f0}:${L(C1)}${f1},MATCH({edad},Datos_ONU!$A${f0}:$A${f1},0),"
            f"MATCH({anio},Datos_ONU!${L(C0)}${h}:${L(C1)}${h},0))")


FMT = dict(pob="#,##0.000", dea="#,##0.000", qx="0.000000", lx="#,##0.0", dx="#,##0.0", Lx="#,##0.0",
           Tx="#,##0.0", ex="0.00")
ENC = dict(pob="Población a mitad de año (miles)", dea="Defunciones (miles)", qx="qx", lx="lx", dx="dx",
           Lx="Lx", Tx="Tx", ex="ex")
for g in GENERACIONES:
    ws = wb.create_sheet(f"Gen_{g}")
    titulo(ws, f"Tabla de mortalidad de la generación {g} - México (ONU WPP 2024)")
    ws["A2"] = "Generación (año de nacimiento):"
    ws["A2"].font = F(bold=True)
    ws["D2"] = g
    ws["D2"].font = F(bold=True, color=AZUL)
    ws["D2"].fill = PatternFill("solid", fgColor="FFFF00")
    ws["D2"].comment = Comment("Año de nacimiento de la generación. Cambiarlo recalcula toda la hoja "
                               "(debe estar entre 1950 y 2000 para que haya datos hasta la edad 100).", "Modelo")
    ws["F2"] = "Radix:"
    ws["F2"].font = F(bold=True)
    ws["G2"] = "=Radix"
    ws["G2"].font = F(color=VERDE)
    ws["G2"].number_format = "#,##0"
    for s, ns in SEXOS:
        b = BASE[s]
        ws.merge_cells(start_row=4, start_column=b, end_row=4, end_column=b + 7)
        cab(ws.cell(4, b), ns, FILL_SEXO[s])
        for k in OFF:
            cab(ws.cell(5, b + OFF[k]), ENC[k], FILL_SEXO[s])
        # control: e0 y suma de dx
        ws.cell(3, b + OFF["qx"], "e0:").font = F(bold=True)
        ws.cell(3, b + OFF["lx"], f"={col(s, 'ex')}{R0}").font = F(bold=True)
        ws.cell(3, b + OFF["lx"]).number_format = "0.00"
        ws.cell(3, b + OFF["dx"], "Control Σdx−radix:").font = F(bold=True, size=9)
        ws.cell(3, b + OFF["Tx"], f"=SUM({col(s, 'dx')}{R0}:{col(s, 'dx')}{R1})-Radix").font = F(size=9)
        ws.cell(3, b + OFF["Tx"]).number_format = "0.000;-0.000;0.000"
    cab(ws.cell(5, 1), "Edad x")
    cab(ws.cell(5, 2), "Año calendario")
    ws.cell(4, 1).fill = ws.cell(4, 2).fill = GRIS
    for i, e in enumerate(EDADES):
        r = R0 + i
        c = ws.cell(r, 1, e)
        c.font = F(bold=True)
        c.alignment = Alignment(horizontal="center")
        if e == 100:
            c.number_format = '0"+"'
        ws.cell(r, 2, f"=$D$2+$A{r}").font = F()
        for s, ns in SEXOS:
            c_ = {k: col(s, k) for k in OFF}
            ws[f"{c_['pob']}{r}"] = idx("Pob", s, f"$A{r}", f"$B{r}")
            ws[f"{c_['dea']}{r}"] = idx("Def", s, f"$A{r}", f"$B{r}")
            ws[f"{c_['qx']}{r}"] = f"=IF($A{r}=100,1,{c_['dea']}{r}/({c_['pob']}{r}+{c_['dea']}{r}/2))"
            ws[f"{c_['lx']}{r}"] = "=Radix" if r == R0 else f"={c_['lx']}{r-1}-{c_['dx']}{r-1}"
            ws[f"{c_['dx']}{r}"] = f"={c_['lx']}{r}*{c_['qx']}{r}"
            ws[f"{c_['Lx']}{r}"] = (f"={c_['dx']}{r}/2" if r == R1
                                    else f"={c_['lx']}{r+1}+{c_['dx']}{r}/2")
            ws[f"{c_['Tx']}{r}"] = f"=SUM({c_['Lx']}{r}:{c_['Lx']}${R1})"
            ws[f"{c_['ex']}{r}"] = f"={c_['Tx']}{r}/{c_['lx']}{r}"
            for k in OFF:
                cc = ws[f"{c_[k]}{r}"]
                cc.number_format = FMT[k]
                cc.font = F(color=VERDE if k in ("pob", "dea") else NEGRO)
    ws.freeze_panes = "C6"
    ws.column_dimensions["A"].width = 9
    ws.column_dimensions["B"].width = 11
    for c in range(3, 27):
        ws.column_dimensions[L(c)].width = 13
    ws.row_dimensions[5].height = 42
    ws.cell(R1 + 2, 1, "Notas: qx = defunciones / (población a mitad de año + defunciones/2); edad 100 = grupo abierto "
                       "(qx = 1, Lx = dx/2). lx(x+1) = lx − dx; Lx = l(x+1) + dx/2; Tx = suma de Lx desde x; "
                       "ex = Tx/lx. Población y defunciones se leen de la hoja Datos_ONU en el año g + x.").font = \
        F(italic=True, size=9, color="595959")

# ----------------------------------------------------------------------------- Periodo (motor de validación)
ws_p = wb.create_sheet("Periodo")
titulo(ws_p, "Tablas de vida de periodo (año calendario) con las mismas fórmulas - motor de la validación",
       "Para cada año se aplican qx = D/(P + D/2), lx(x+1) = lx(1−qx) y e0 = (Σlx − l0/2)/l0 "
       "(equivale a Σ Lx/l0 con Lx = l(x+1) + dx/2 y L100 = l100/2).")
ws_p.cell(4, 1, "e0 calculada").font = F(bold=True, size=11)
cab(ws_p.cell(5, 1), "Sexo")
for j, a in enumerate(ANIOS):
    cab(ws_p.cell(5, C0 + j), a)
top = 10
PER = {}
for k_, (pref, nom) in enumerate((("qx", "qx de periodo"), ("lx", "lx de periodo"))):
    for s, ns in SEXOS:
        PER[(pref, s)] = top
        ws_p.cell(top, 1, f"{nom} - {ns}").font = F(bold=True, size=11)
        cab(ws_p.cell(top + 1, 1), "Edad")
        for j, a in enumerate(ANIOS):
            cab(ws_p.cell(top + 1, C0 + j), a)
        top += NA + 4
for s, ns in SEXOS:
    tq, tl = PER[("qx", s)], PER[("lx", s)]
    _, p0, _ = BLQ[("Pob", s)]
    _, d0, _ = BLQ[("Def", s)]
    for i, e in enumerate(EDADES):
        for top_ in (tq, tl):
            c = ws_p.cell(top_ + 2 + i, 1, e)
            c.font = F(bold=True)
            if e == 100:
                c.number_format = '0"+"'
        for j in range(len(ANIOS)):
            cl = L(C0 + j)
            rq, rl = tq + 2 + i, tl + 2 + i
            ws_p[f"{cl}{rq}"] = (f"=IF($A{rq}=100,1,Datos_ONU!{cl}{d0+i}/(Datos_ONU!{cl}{p0+i}+Datos_ONU!{cl}{d0+i}/2))")
            ws_p[f"{cl}{rq}"].number_format = "0.000000"
            ws_p[f"{cl}{rl}"] = "=Radix" if i == 0 else f"={cl}{rl-1}*(1-{cl}{rq-1})"
            ws_p[f"{cl}{rl}"].number_format = "#,##0.0"
    fr = 6 + "HMA".index(s)
    ws_p.cell(fr, 1, ns).font = F(bold=True)
    for j in range(len(ANIOS)):
        cl = L(C0 + j)
        l0_, l1_ = tl + 2, tl + 2 + NA - 1
        ws_p[f"{cl}{fr}"] = f"=(SUM({cl}{l0_}:{cl}{l1_})-{cl}{l0_}/2)/{cl}{l0_}"
        ws_p[f"{cl}{fr}"].number_format = "0.00"
        ws_p[f"{cl}{fr}"].font = F(bold=True)
ws_p.freeze_panes = "B6"
ws_p.column_dimensions["A"].width = 14

# ----------------------------------------------------------------------------- Val_Diagonal_ONU
ws_v = wb.create_sheet("Val_Diagonal_ONU")
titulo(ws_v, "Generaciones construidas con la qx oficial de la ONU (diagonal de las tablas de periodo)",
       "Para la generación g y la edad x se toma la qx oficial ONU del año g + x. lx con el mismo radix; "
       "sirve de referencia para comparar con las tablas Gen_XXXX.")
VCOL = {}
c = 2
for s, ns in SEXOS:
    ws_v.merge_cells(start_row=4, start_column=c, end_row=4, end_column=c + 2 * len(GENERACIONES) - 1)
    cab(ws_v.cell(4, c), ns, FILL_SEXO[s])
    for g in GENERACIONES:
        ws_v.merge_cells(start_row=5, start_column=c, end_row=5, end_column=c + 1)
        cab(ws_v.cell(5, c), g, FILL_SEXO[s])
        cab(ws_v.cell(6, c), "qx ONU", FILL_SEXO[s])
        cab(ws_v.cell(6, c + 1), "lx", FILL_SEXO[s])
        VCOL[(s, g)] = (c, c + 1)
        c += 2
cab(ws_v.cell(6, 1), "Edad x")
VR0, VR1 = 7, 7 + NA - 1
for i, e in enumerate(EDADES):
    r = VR0 + i
    ws_v.cell(r, 1, e).font = F(bold=True)
    for s, ns in SEXOS:
        hq, q0, q1 = BLQ_Q[s]
        for g in GENERACIONES:
            cq, cl = VCOL[(s, g)]
            gref = f"{L(cq)}$5"
            ws_v.cell(r, cq).value = (f"=IF($A{r}=100,1,INDEX(ONU_qx!${L(C0)}${q0}:${L(C1)}${q1},MATCH($A{r},ONU_qx!$A${q0}:$A${q1},0),"
                                      f"MATCH({gref}+$A{r},ONU_qx!${L(C0)}${hq}:${L(C1)}${hq},0)))")
            ws_v.cell(r, cq).number_format = "0.000000"
            ws_v.cell(r, cl).value = "=Radix" if i == 0 else f"={L(cl)}{r-1}*(1-{L(cq)}{r-1})"
            ws_v.cell(r, cl).number_format = "#,##0.0"
            ws_v.cell(r, cq).font = F(color=VERDE)
            ws_v.cell(r, cl).font = F()
ws_v.freeze_panes = "B7"
ws_v.column_dimensions["A"].width = 9

# ----------------------------------------------------------------------------- Datos_Graficos
ws_g = wb.create_sheet("Datos_Graficos")
titulo(ws_g, "Datos de los gráficos de la hoja Comparacion (vínculos a las tablas Gen_XXXX)")
GX = {}     # (medida, sexo) -> columna inicial; filas de datos 6..106
c = 2
for medida, k in (("ex", "ex"), ("lx", "lx")):
    cab(ws_g.cell(5, 1 if medida == "ex" else c), "Edad x")
    if medida == "lx":
        for i, e in enumerate(EDADES):
            ws_g.cell(6 + i, c, f"=A{6+i}").font = F()
        c += 1
    for s, ns in SEXOS:
        ws_g.merge_cells(start_row=3, start_column=c, end_row=3, end_column=c + 5)
        cab(ws_g.cell(3, c), f"{medida} - {ns}", FILL_SEXO[s])
        GX[(medida, s)] = c
        for j, g in enumerate(GENERACIONES):
            cab(ws_g.cell(5, c + j), f"Gen. {g}", FILL_SEXO[s])
            for i, e in enumerate(EDADES):
                cc = ws_g.cell(6 + i, c + j, f"=Gen_{g}!{col(s, k)}{R0+i}")
                cc.font = F(color=VERDE)
                cc.number_format = "0.00" if medida == "ex" else "#,##0"
        c += 6
    c += 1
for i, e in enumerate(EDADES):
    ws_g.cell(6 + i, 1, e).font = F(bold=True)
# e0 oficial ONU por año
c_e0 = c + 1
cab(ws_g.cell(5, c_e0), "Año")
for j, (s, ns) in enumerate(SEXOS):
    cab(ws_g.cell(5, c_e0 + 1 + j), f"e0 ONU {ns}", FILL_SEXO[s])
ws_g.cell(4, c_e0, "e0 oficial ONU (periodo)").font = F(bold=True)
for i in range(len(ANIOS)):
    ws_g.cell(6 + i, c_e0, f"=ONU_e0!A{E0_F0+i}").font = F(color=VERDE)
    for j, s in enumerate("HMA"):
        cc = ws_g.cell(6 + i, c_e0 + 1 + j, f"=ONU_e0!{L(2+j)}{E0_F0+i}")
        cc.font = F(color=VERDE)
        cc.number_format = "0.00"
c_gen = c_e0 + 6
cab(ws_g.cell(5, c_gen), "Generación")
ws_g.cell(4, c_gen, "e0 de cada generación (tablas Gen)").font = F(bold=True)
for j, (s, ns) in enumerate(SEXOS):
    cab(ws_g.cell(5, c_gen + 1 + j), f"e0 gen. {ns}", FILL_SEXO[s])
for i, g in enumerate(GENERACIONES):
    ws_g.cell(6 + i, c_gen, g).font = F(bold=True, color=AZUL)
    for j, (s, ns) in enumerate(SEXOS):
        cc = ws_g.cell(6 + i, c_gen + 1 + j, f"=Gen_{g}!{col(s, 'ex')}{R0}")
        cc.font = F(color=VERDE)
        cc.number_format = "0.00"
ws_g.freeze_panes = "B6"

# ----------------------------------------------------------------------------- Comparacion
ws_c = wb.create_sheet("Comparacion")
titulo(ws_c, "Comparación entre generaciones - México (ONU WPP 2024, variante media)",
       "Indicadores y gráficos vinculados a las hojas Gen_XXXX. Las generaciones posteriores a 1970 incluyen "
       "años proyectados (2024-2100).")
IND = [("e0", "e0: esperanza de vida al nacer (años)", 0, "0.00", "ex"),
       ("e65", "e65: años que se esperan vivir a los 65", 65, "0.00", "ex"),
       ("l65", "l65/l0: % que sobrevive a los 65", 65, "0.0%", "lx"),
       ("l80", "l80/l0: % que sobrevive a los 80", 80, "0.0%", "lx")]
ws_c.cell(4, 1, "Indicadores por generación").font = F(bold=True, size=11)
cab(ws_c.cell(5, 1), "Generación")
ws_c.merge_cells("A5:A6")
cc_ = 2
for s, ns in SEXOS:
    ws_c.merge_cells(start_row=5, start_column=cc_, end_row=5, end_column=cc_ + 3)
    cab(ws_c.cell(5, cc_), ns, FILL_SEXO[s])
    for j, (_, et, _, _, _) in enumerate(IND):
        cab(ws_c.cell(6, cc_ + j), et.split(":")[0], FILL_SEXO[s])
    cc_ += 4
for i, g in enumerate(GENERACIONES):
    r = 7 + i
    ws_c.cell(r, 1, g).font = F(bold=True)
    ws_c.cell(r, 1).alignment = Alignment(horizontal="center")
    cc_ = 2
    for s, ns in SEXOS:
        for j, (cod, et, edad, fmt, k) in enumerate(IND):
            fila = R0 + edad
            f = f"=Gen_{g}!{col(s, k)}{fila}" if k == "ex" else f"=Gen_{g}!{col(s, 'lx')}{fila}/Gen_{g}!{col(s, 'lx')}{R0}"
            cell = ws_c.cell(r, cc_ + j, f)
            cell.number_format = fmt
            cell.font = F(color=VERDE)
            cell.border = BORDE
        cc_ += 4
ws_c.cell(14, 1, "e0 = ex a edad 0; e65 = ex a edad 65; l65/l0 y l80/l0 = supervivencia desde el nacimiento. "
                 "Hojas fuente: Gen_1950 … Gen_2000.").font = F(italic=True, size=9, color="595959")
ws_c.column_dimensions["A"].width = 12
for c in range(2, 14):
    ws_c.column_dimensions[L(c)].width = 9.5


def estilo_linea(serie, hexcolor, ancho=22000, punteada=False):
    serie.graphicalProperties.line.solidFill = hexcolor.lstrip("#")
    serie.graphicalProperties.line.width = ancho
    serie.smooth = False
    serie.marker.symbol = "none"
    if punteada:
        serie.graphicalProperties.line.dashStyle = "dash"


def grafico_gen(medida, s, ns, ancla, titulo_y):
    ch = LineChart()
    ch.title = f"{titulo_y} por generación - {ns}"
    ch.height, ch.width = 8.2, 14.0
    c0 = GX[(medida, s)]
    cats = Reference(ws_g, min_col=1, min_row=6, max_row=106)
    for j in range(len(GENERACIONES)):
        ref = Reference(ws_g, min_col=c0 + j, min_row=5, max_row=106)
        ch.add_data(ref, titles_from_data=True)
    ch.set_categories(cats)
    for j, serie in enumerate(ch.series):
        estilo_linea(serie, RAMPA_GEN[j])
    ch.x_axis.title = "Edad x"
    ch.y_axis.title = "ex (años)" if medida == "ex" else "lx (por 100,000 nacidos)"
    ch.x_axis.tickLblSkip = 10
    ch.x_axis.tickMarkSkip = 10
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.legend.position = "b"
    ws_c.add_chart(ch, ancla)


# filas 16 / 34 / 52
for k_, (s, ns) in enumerate(SEXOS):
    grafico_gen("ex", s, ns, f"{L(1 + 8*k_)}16", "ex")
    grafico_gen("lx", s, ns, f"{L(1 + 8*k_)}34", "lx")

sc = ScatterChart()
sc.title = "Esperanza de vida al nacer por año calendario: ONU (periodo) y generaciones"
sc.style = 2
sc.height, sc.width = 9.5, 24
sc.x_axis.title = "Año calendario / año de nacimiento"
sc.y_axis.title = "e0 (años)"
sc.x_axis.scaling.min, sc.x_axis.scaling.max = 1950, 2100
sc.y_axis.scaling.min = 40
sc.x_axis.delete = False
sc.y_axis.delete = False
xs = Reference(ws_g, min_col=c_e0, min_row=6, max_row=6 + len(ANIOS) - 1)
for j, (s, ns) in enumerate(SEXOS):
    ys = Reference(ws_g, min_col=c_e0 + 1 + j, min_row=5, max_row=6 + len(ANIOS) - 1)
    se = Series(ys, xs, title_from_data=True)
    estilo_linea(se, COLOR_SEXO[s])
    sc.series.append(se)
xg = Reference(ws_g, min_col=c_gen, min_row=6, max_row=11)
for j, (s, ns) in enumerate(SEXOS):
    yg = Reference(ws_g, min_col=c_gen + 1 + j, min_row=5, max_row=11)
    se = Series(yg, xg, title_from_data=True)
    se.marker.symbol = "circle"
    se.marker.size = 8
    se.marker.graphicalProperties = GraphicalProperties(solidFill=COLOR_SEXO[s])
    se.marker.graphicalProperties.line.solidFill = "FFFFFF"
    se.graphicalProperties.line.noFill = True
    sc.series.append(se)
sc.legend.position = "b"
ws_c.add_chart(sc, "A52")
ws_c.cell(51, 1, "Líneas: e0 oficial de la ONU por año calendario (periodo). Puntos: e0 de cada generación según las "
                 "tablas de este libro, ubicada en su año de nacimiento.").font = F(italic=True, size=9, color="595959")

# ----------------------------------------------------------------------------- Validacion
ws_x = wb.create_sheet("Validacion")
titulo(ws_x, "Validación contra la esperanza de vida oficial de la ONU",
       "A: la misma metodología aplicada por año calendario reproduce la e0 oficial. "
       "B: las tablas por generación se comparan con la diagonal oficial de la ONU.")
ws_x["A4"] = "A. e0 de periodo calculada con las fórmulas del libro vs. e0 oficial ONU (1950-2100)"
ws_x["A4"].font = F(bold=True, size=11)
YA = 40                       # fila del encabezado de la tabla anual
YF0, YF1 = YA + 2, YA + 1 + len(ANIOS)
for j, t in enumerate(["Sexo", "Diferencia media", "Error absoluto medio", "Máx. |dif.|", "Año del máx.",
                       "Dif. 1950", "Dif. 2023", "Dif. 2100"]):
    cab(ws_x.cell(5, 1 + j), t)
colan = {}
for j, (s, ns) in enumerate(SEXOS):
    colan[s] = (2 + 3 * j, 3 + 3 * j, 4 + 3 * j)   # ONU, calculada, dif
for j, (s, ns) in enumerate(SEXOS):
    r = 6 + j
    cd = L(colan[s][2])
    rng = f"${cd}${YF0}:${cd}${YF1}"
    yrs = f"$A${YF0}:$A${YF1}"
    ws_x.cell(r, 1, ns).font = F(bold=True)
    ws_x.cell(r, 2, f"=AVERAGE({rng})")
    ws_x.cell(r, 3, f"=SUMPRODUCT(ABS({rng}))/COUNT({rng})")
    ws_x.cell(r, 4, f"=MAX(MAX({rng}),-MIN({rng}))")
    ws_x.cell(r, 5, f"=IF(MAX({rng})>=-MIN({rng}),INDEX({yrs},MATCH(MAX({rng}),{rng},0)),INDEX({yrs},MATCH(MIN({rng}),{rng},0)))")
    ws_x.cell(r, 6, f"=INDEX({rng},MATCH(1950,{yrs},0))")
    ws_x.cell(r, 7, f"=INDEX({rng},MATCH(2023,{yrs},0))")
    ws_x.cell(r, 8, f"=INDEX({rng},MATCH(2100,{yrs},0))")
    for c in range(2, 9):
        ws_x.cell(r, c).number_format = "0.000" if c != 5 else "0"
        ws_x.cell(r, c).border = BORDE
        ws_x.cell(r, c).font = F()
ws_x["A9"] = "Diferencia = e0 calculada − e0 oficial (años). 1950-2023 son estimaciones y 2024-2100 proyecciones de la ONU."
ws_x["A9"].font = F(italic=True, size=9, color="595959")

ws_x["A11"] = "B. e0 de cada generación: tablas del libro (Gen_XXXX) vs. diagonal oficial de la ONU (Val_Diagonal_ONU)"
ws_x["A11"].font = F(bold=True, size=11)
for j, t in enumerate(["Sexo", "Generación", "e0 tabla (libro)", "e0 diagonal ONU", "Diferencia (años)",
                       "Diferencia (%)", "e0 oficial ONU de periodo en el año de nacimiento", "Brecha generación − periodo (años)",
                       "l100 (por 100,000)", "Ajuste por grupo abierto (años)", "e0 tabla + ajuste"]):
    cab(ws_x.cell(12, 1 + j), t)
ws_x.row_dimensions[12].height = 54
ws_x["M12"] = "e100 supuesto (años)"
cab(ws_x["M12"], "e100 supuesto (años)")
ws_x["M13"] = 2.0
ws_x["M13"].font = F(color=AZUL, bold=True)
ws_x["M13"].fill = PatternFill("solid", fgColor="FFFF00")
ws_x["M13"].number_format = "0.00"
ws_x["M13"].comment = Comment("Supuesto: años que en promedio le quedan por vivir a quien llega a los 100 años. "
                              "Las tablas de vida de la ONU dan e100 entre 1.3 y 3.0 para México, 1950-2100 (mediana 2.4, ambos "
                              "sexos); 2.0 es un valor prudente. En las tablas del libro, con qx = 1 en la edad 100, resulta e100 = 0.5.", "Modelo")
r = 13
FILA_B = {}
for s, ns in SEXOS:
    for g in GENERACIONES:
        cq, cl = VCOL[(s, g)]
        l0_, l1_ = f"{L(cl)}{VR0}", f"{L(cl)}{VR1}"
        FILA_B[(s, g)] = r
        ws_x.cell(r, 1, ns).font = F(bold=True)
        ws_x.cell(r, 2, g).font = F(bold=True)
        ws_x.cell(r, 3, f"=Gen_{g}!{col(s, 'ex')}{R0}")
        ws_x.cell(r, 4, f"=(SUM(Val_Diagonal_ONU!{l0_}:{l1_})-Val_Diagonal_ONU!{l0_}/2)/Val_Diagonal_ONU!{l0_}")
        ws_x.cell(r, 5, f"=C{r}-D{r}")
        ws_x.cell(r, 6, f"=E{r}/D{r}")
        ws_x.cell(r, 7, f"=INDEX(ONU_e0!${L(2+'HMA'.index(s))}${E0_F0}:${L(2+'HMA'.index(s))}${E0_F1},MATCH(B{r},ONU_e0!$A${E0_F0}:$A${E0_F1},0))")
        ws_x.cell(r, 8, f"=C{r}-G{r}")
        ws_x.cell(r, 9, f"=Gen_{g}!{col(s, 'lx')}{R1}")
        ws_x.cell(r, 10, f"=I{r}/Radix*($M$13-0.5)")
        ws_x.cell(r, 11, f"=C{r}+J{r}")
        for c, fmt in zip(range(3, 12), ["0.00", "0.00", "0.00;-0.00;0.00", "0.0%", "0.00", "0.00", "#,##0", "0.000", "0.00"]):
            ws_x.cell(r, c).number_format = fmt
            ws_x.cell(r, c).font = F(color=VERDE if c in (3, 7, 9) else NEGRO)
        for c in range(1, 12):
            ws_x.cell(r, c).border = BORDE
        r += 1
ws_x.merge_cells(start_row=r, start_column=1, end_row=r, end_column=11)
ws_x.row_dimensions[r].height = 40
ws_x.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
ws_x.cell(r, 1, "Lectura: (i) 'Diferencia' mide el error de la estimación qx = D/(P + D/2) frente a la qx oficial, sobre las mismas edades "
                "y años. (ii) La 'brecha generación − periodo' NO es error: la generación vive las mejoras de mortalidad futuras "
                "y la e0 de periodo no. (iii) El grupo abierto (qx = 1 a los 100) subestima e0 en 'Ajuste' años; afecta igual a la tabla y a la diagonal ONU, "
                "por lo que no altera la columna 'Diferencia'.").font = \
    F(italic=True, size=9, color="595959")

for j, t in enumerate(["Año", "ONU Hombres", "Calculada Hombres", "Dif. Hombres", "ONU Mujeres", "Calculada Mujeres",
                       "Dif. Mujeres", "ONU Ambos", "Calculada Ambos", "Dif. Ambos"]):
    cab(ws_x.cell(YA + 1, 1 + j), t, FILL_SEXO["HMA"[max(0, (j - 1)) // 3]] if j > 0 else GRIS)
ws_x.cell(YA, 1, "Detalle anual (e0 de periodo, años)").font = F(bold=True, size=11)
ws_x.row_dimensions[YA + 1].height = 30
for i, a in enumerate(ANIOS):
    r = YF0 + i
    ws_x.cell(r, 1, a).font = F(bold=True)
    for j, s in enumerate("HMA"):
        cu, cc_, cd = colan[s]
        ws_x.cell(r, cu, f"=ONU_e0!{L(2+j)}{E0_F0+i}").font = F(color=VERDE)
        ws_x.cell(r, cc_, f"=Periodo!{L(C0+i)}{6+j}").font = F(color=VERDE)
        ws_x.cell(r, cd, f"={L(cc_)}{r}-{L(cu)}{r}").font = F()
        for c in (cu, cc_):
            ws_x.cell(r, c).number_format = "0.00"
        ws_x.cell(r, cd).number_format = "0.000;-0.000;0.000"
ws_x.column_dimensions["A"].width = 13
for c in range(2, 14):
    ws_x.column_dimensions[L(c)].width = 14

# gráficos de validación
lc = LineChart()
lc.title = "Diferencia de e0 (calculada − ONU) por año calendario"
lc.height, lc.width = 7.5, 16
for s, ns in SEXOS:
    ref = Reference(ws_x, min_col=colan[s][2], min_row=YA + 1, max_row=YF1)
    lc.add_data(ref, titles_from_data=True)
lc.set_categories(Reference(ws_x, min_col=1, min_row=YF0, max_row=YF1))
for serie, s in zip(lc.series, "HMA"):
    estilo_linea(serie, COLOR_SEXO[s], 19000)
lc.x_axis.tickLblSkip = 10
lc.x_axis.tickMarkSkip = 10
lc.y_axis.title = "años"
lc.x_axis.tickLblPos = "low"
lc.x_axis.delete = False
lc.y_axis.delete = False
lc.legend.position = "b"
ws_x.add_chart(lc, "P3")

bc = BarChart()
bc.type = "col"
bc.title = "e0 por generación, ambos sexos: tabla del libro vs. diagonal ONU"
bc.height, bc.width = 7.5, 16
fa = FILA_B[("A", GENERACIONES[0])]
for c_, nom_, color in ((3, "Tabla del libro", "2A78D6"), (4, "Diagonal ONU", "9AA0A6")):
    se = Series(Reference(ws_x, min_col=c_, min_row=fa, max_row=fa + 5), title=nom_)
    se.graphicalProperties.solidFill = color
    bc.series.append(se)
bc.set_categories(Reference(ws_x, min_col=2, min_row=fa, max_row=fa + 5))
bc.y_axis.scaling.min = 40
bc.y_axis.title = "años"
bc.x_axis.delete = False
bc.y_axis.delete = False
bc.legend.position = "b"
ws_x.add_chart(bc, "P19")

# ----------------------------------------------------------------------------- Leeme (contenido)
titulo(ws_l, "Tablas de mortalidad de generaciones mexicanas (1950, 1960, 1970, 1980, 1990, 2000)")
filas = [
    ("Fuente", "ONU, División de Población, World Population Prospects 2024 (variante media): población a mitad de año y "
               "defunciones por edad simple, México, 1950-2100 (https://population.un.org/wpp/). 1950-2023 estimaciones; "
               "2024-2100 proyecciones."),
    ("Radix (l0)", None),
    ("", ""),
    ("Método", "Generación g = personas nacidas en el año g. La edad x de esa generación se observa en el año calendario g + x."),
    ("", "qx = defunciones(x, g+x) / (población a mitad de año(x, g+x) + defunciones(x, g+x)/2)"),
    ("", "Edad 100 = grupo abierto (100+): qx = 1."),
    ("", "lx: l0 = radix; l(x+1) = lx − dx.   dx = lx · qx."),
    ("", "Lx = l(x+1) + dx/2   (a la edad 100, l101 = 0, así que L100 = d100/2 = l100/2)."),
    ("", "Tx = suma de Lx desde x hasta 100.   ex = Tx / lx."),
    ("", ""),
    ("Hojas", "Gen_1950 … Gen_2000: tabla completa por generación, con Hombres, Mujeres y Ambos sexos (columnas C-J, K-R, S-Z)."),
    ("", "Comparacion: indicadores y gráficos de ex, lx y e0 por año calendario (datos en Datos_Graficos)."),
    ("", "Validacion: contraste con la esperanza de vida oficial de la ONU (apoyada en Periodo, ONU_e0, ONU_qx y Val_Diagonal_ONU)."),
    ("", "Datos_ONU, ONU_qx, ONU_e0: datos de entrada de la ONU (México)."),
    ("", ""),
    ("Colores", "Azul = dato fijo / entrada.  Negro = fórmula.  Verde = vínculo a otra hoja.  Fondo amarillo = supuesto editable."),
    ("Limitaciones", "(1) Defunciones y población son de periodo-edad, no de triángulos de Lexis: aproximan la generación con la "
                     "población de edad x a mitad del año g + x. (2) Con qx = 1 a los 100 años, e100 = 0.5 años (la ONU estima 1.3-3.0): "
                     "subestima e0 hasta ≈ 0.1 años en la generación 2000 (ver Validacion). (2b) q0 = D/(P + D/2) sobrestima algo la "
                     "mortalidad infantil de 1950 (0.163 vs 0.157 oficial). (3) Los años posteriores a 2023 son "
                     "proyección: las generaciones 1960-2000 son tablas parcialmente proyectadas."),
]
r = 3
for k, v in filas:
    ws_l.cell(r, 1, k).font = F(bold=True)
    ws_l.cell(r, 1).alignment = Alignment(vertical="top")
    if k == "Radix (l0)":
        ws_l.cell(r, 2, 100000).font = F(bold=True, color=AZUL)
        ws_l.cell(r, 2).fill = PatternFill("solid", fgColor="FFFF00")
        ws_l.cell(r, 2).number_format = "#,##0"
        ws_l.cell(r, 3, "<- nombre definido 'Radix'; se usa en todas las tablas").font = F(italic=True, color="595959")
        wb.defined_names["Radix"] = DefinedName("Radix", attr_text=f"Leeme!$B${r}")
    else:
        ws_l.cell(r, 2, v).font = F()
        ws_l.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
        ws_l.merge_cells(start_row=r, start_column=2, end_row=r, end_column=12)
        if k in ("Limitaciones", "Fuente"):
            ws_l.row_dimensions[r].height = 58 if k == "Limitaciones" else 30
    r += 1
ws_l.column_dimensions["A"].width = 16
for c in range(2, 13):
    ws_l.column_dimensions[L(c)].width = 12

# orden de hojas: Leeme, Gen_*, Comparacion, Validacion, apoyo
orden = (["Leeme"] + [f"Gen_{g}" for g in GENERACIONES] + ["Comparacion", "Validacion", "Datos_Graficos", "Periodo",
                                                              "Val_Diagonal_ONU", "Datos_ONU", "ONU_qx", "ONU_e0"])
wb._sheets = [wb[n] for n in orden]
for n in orden:
    wb[n].sheet_properties.tabColor = ("1F3864" if n.startswith("Gen_") else "2A78D6" if n in ("Comparacion", "Validacion")
                                       else "808080")

# impresión: apaisado y ajustado al ancho de una página
for n in orden:
    w = wb[n]
    w.page_setup.orientation = "landscape"
    w.sheet_properties.pageSetUpPr.fitToPage = True
    w.page_setup.fitToWidth = 1
    w.page_setup.fitToHeight = 1 if n in ("Comparacion", "Leeme") else 0
for g in GENERACIONES:
    wb[f"Gen_{g}"].print_title_rows = "4:5"

SALIDA.parent.mkdir(exist_ok=True)
wb.save(SALIDA)
print("guardado", SALIDA)
