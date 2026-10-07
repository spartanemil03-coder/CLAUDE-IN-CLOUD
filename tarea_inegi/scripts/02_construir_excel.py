"""Construye el libro de Excel: tabla de mortalidad de México 2019 con datos de INEGI (formato de la plantilla de clase).

La estructura de cada hoja de tabla copia la plantilla de la profesora (columnas x, qx, lx, dx, Lx, Tx, ex y nota
técnica con Nacimientos / Defunciones / q(x)), pero con fórmulas vivas en todas las columnas.

Uso:  python3 scripts/02_construir_excel.py     y luego recalcular con LibreOffice (recalc.py de la skill xlsx).
"""
import sys
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference, Series
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName

sys.path.insert(0, str(Path(__file__).resolve().parent))
import calculo_inegi as calc  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "datos"
SALIDA = RAIZ / "resultados" / "T_Mortalidad_INEGI_Mexico_2019.xlsx"

dfn, nac, pob = calc.cargar(DATOS)
onu = pd.read_csv(DATOS / "onu_wpp2024_e0_mexico_2019.csv")
REF = calc.tablas(DATOS, suaviza=True)                  # cálculo de referencia (para el texto del análisis)
IND = {s: calc.indicadores(REF[s]) for s in REF}

FONT = "Arial"
F = lambda **k: Font(name=FONT, size=k.pop("size", 10), **k)
AZUL, VERDE, NEGRO, GRISTX = "0000FF", "008000", "000000", "595959"
AMARILLO = PatternFill("solid", fgColor="FFFF00")
GRIS = PatternFill("solid", fgColor="F2F2F2")
FILL_S = {"H": PatternFill("solid", fgColor="DCE9F9"), "M": PatternFill("solid", fgColor="FBE3D8"),
          "A": PatternFill("solid", fgColor="D5F1E6")}
FINO = Side(style="thin", color="BFBFBF")
BORDE = Border(left=FINO, right=FINO, top=FINO, bottom=FINO)
CENTRO = Alignment(horizontal="center", vertical="center", wrap_text=True)
COLOR = {"H": "2A78D6", "M": "EB6834", "A": "1BAF7A"}
NOMBRE = {"H": "Hombres", "M": "Mujeres", "A": "Ambos sexos"}
HOJA = {"A": "Tabla ambos sexos", "H": "Tabla hombres", "M": "Tabla mujeres"}

wb = Workbook()
ws_l = wb.active
ws_l.title = "Leeme"


def cab(c, texto, fill=GRIS, size=10):
    c.value = texto
    c.font = F(bold=True, size=size)
    c.fill = fill
    c.alignment = CENTRO
    c.border = BORDE


def nota(ws, ref, texto, size=9, color=GRISTX, bold=False):
    ws[ref] = texto
    ws[ref].font = F(italic=not bold, size=size, color=color, bold=bold)


# ============================================================================ Datos INEGI
wd = wb.create_sheet("Datos INEGI")
wd.sheet_view.showGridLines = False
wd["A1"] = "Datos de INEGI para la tabla de mortalidad de México, 2019"
wd["A1"].font = F(size=14, bold=True)
nota(wd, "A2", "Azul = dato tomado de INEGI (conteos); negro = fórmula. Defunciones y nacimientos por año de OCURRENCIA 2019 "
               "(registrados 2019-2021). Población: Censo 2020.")

cab(wd["A3"], "Nacimientos ocurridos en 2019 (ENR)")
for j, t in enumerate(["Hombres", "Mujeres", "Sexo no esp.", "Total"]):
    cab(wd.cell(3, 2 + j), t)
wd["A4"] = "Nacimientos"
wd["A4"].font = F(bold=True)
for j, k in enumerate(["hombres", "mujeres", "sexo_no_especificado"]):
    c = wd.cell(4, 2 + j, int(nac[k]))
    c.font, c.number_format = F(color=AZUL), "#,##0"
wd["E4"] = "=SUM(B4:D4)"
wd["E4"].font, wd["E4"].number_format = F(bold=True), "#,##0"
nota(wd, "F4", "Fuente: INEGI, Estadística de Nacimientos Registrados (datos abiertos 2019, 2020 y 2021), ano_nac = 2019.")

cab(wd["A6"], "e0 de la ONU, 2019 (solo validación)")
for j, t in enumerate(["Hombres", "Mujeres", "Ambos sexos"]):
    cab(wd.cell(6, 2 + j), t)
wd["A7"] = "e0 (años)"
wd["A7"].font = F(bold=True)
for j, s in enumerate(["Hombres", "Mujeres", "Ambos sexos"]):
    c = wd.cell(7, 2 + j, round(float(onu.set_index("sexo").loc[s, "e0_2019"]), 4))
    c.font, c.number_format = F(color=AZUL), "0.00"
nota(wd, "F7", "Fuente: ONU, World Population Prospects 2024 (variante media), LEx 2019. No se usa en la tabla; solo para comparar.")

HR, R0, R1, RNE, RT, RF = 12, 13, 113, 114, 115, 116
wd.merge_cells(start_row=11, start_column=2, end_row=11, end_column=5)
cab(wd.cell(11, 2), "Defunciones ocurridas en 2019 (EDR)")
wd.merge_cells(start_row=11, start_column=6, end_row=11, end_column=8)
cab(wd.cell(11, 6), "Población, Censo 2020")
wd.merge_cells(start_row=11, start_column=9, end_row=11, end_column=11)
cab(wd.cell(11, 9), "Defunciones con edad no especificada redistribuida")
wd.merge_cells(start_row=11, start_column=12, end_row=11, end_column=14)
cab(wd.cell(11, 12), "Defunciones usadas en qx")
wd.merge_cells(start_row=11, start_column=15, end_row=11, end_column=17)
cab(wd.cell(11, 15), "Población usada en qx")
heads = ["Edad x", "Hombres", "Mujeres", "Sexo no esp.", "Total", "Hombres", "Mujeres", "Total"] + ["Hombres", "Mujeres", "Total"] * 3
for j, t in enumerate(heads):
    cab(wd.cell(HR, 1 + j), t)
for i in range(101):
    r = R0 + i
    c = wd.cell(r, 1, i)
    c.font, c.alignment = F(bold=True), Alignment(horizontal="center")
    if i == 100:
        c.number_format = '0"+"'
    for j, k in enumerate(["hombres", "mujeres", "sexo_no_especificado"]):
        c = wd.cell(r, 2 + j, int(dfn.loc[str(i), k]))
        c.font, c.number_format = F(color=AZUL), "#,##0"
    wd.cell(r, 5, f"=SUM(B{r}:D{r})")
    wd.cell(r, 6, int(pob.loc[str(i), "hombres"]))
    wd.cell(r, 7, int(pob.loc[str(i), "mujeres"]))
    wd.cell(r, 8, f"=F{r}+G{r}")
    for c_, (src) in zip((9, 10, 11), ("B", "C", "E")):
        wd.cell(r, c_, f"={src}{r}*{src}${RF}")
    for c_, src in zip((12, 13, 14), ("I", "J", "K")):
        wd.cell(r, c_, f"=IF(AND(Suaviza=1,$A{r}>=10,$A{r}<=97),AVERAGE({src}{r-2}:{src}{r+2}),{src}{r})")
    for c_, src in zip((15, 16, 17), ("F", "G", "H")):
        wd.cell(r, c_, f"=IF(AND(Suaviza=1,$A{r}>=10,$A{r}<=97),AVERAGE({src}{r-2}:{src}{r+2}),{src}{r})")
    for c_ in range(2, 18):
        cc = wd.cell(r, c_)
        cc.number_format = "#,##0" if c_ < 9 else "#,##0.0"
        if c_ in (6, 7):
            cc.font = F(color=AZUL)
        elif c_ not in (2, 3, 4):
            cc.font = F()
# fila edad no especificada
wd.cell(RNE, 1, "No especificada").font = F(bold=True)
for j, k in enumerate(["hombres", "mujeres", "sexo_no_especificado"]):
    c = wd.cell(RNE, 2 + j, int(dfn.loc["NE", k]))
    c.font, c.number_format = F(color=AZUL), "#,##0"
wd.cell(RNE, 5, f"=SUM(B{RNE}:D{RNE})")
wd.cell(RNE, 6, int(pob.loc["NE", "hombres"])).font = F(color=AZUL)
wd.cell(RNE, 7, int(pob.loc["NE", "mujeres"])).font = F(color=AZUL)
wd.cell(RNE, 8, f"=F{RNE}+G{RNE}")
wd.cell(RT, 1, "Total").font = F(bold=True)
for c_ in range(2, 9):
    wd.cell(RT, c_, f"=SUM({L(c_)}{R0}:{L(c_)}{RNE})").font = F(bold=True)
wd.cell(RF, 1, "Factor de redistribución").font = F(bold=True)
for c_ in (2, 3, 5):
    wd.cell(RF, c_, f"={L(c_)}{RT}/({L(c_)}{RT}-{L(c_)}{RNE})").number_format = "0.0000"
for r in (RNE, RT):
    for c_ in range(2, 9):
        wd.cell(r, c_).number_format = "#,##0"
nota(wd, f"A{RF+2}", "Edad no especificada: se reparte entre todas las edades en proporción a sus defunciones (factor = total / total con edad). "
                     "Población con edad no especificada (Censo): no se redistribuye (0.2 % del total).")
nota(wd, f"A{RF+3}", "Fuente defunciones: INEGI, Estadísticas de Defunciones Registradas, datos abiertos 2019, 2020 y 2021 (anio_ocur = 2019; "
                     "edad: códigos 1001-3999 = menores de 1 año, 4001-4120 = años cumplidos). Fuente población: INEGI, Censo de Población y "
                     "Vivienda 2020, tabulado 'Población 3', Estados Unidos Mexicanos.")
nota(wd, f"A{RF+4}", "Suavizado: media móvil de 5 edades (x-2 … x+2) en defunciones y población, de 10 a 97 años, para quitar la atracción de edades "
                     "redondas del Censo (p. ej. 30 años: 2.37 millones vs. 31 años: 1.54 millones). Se activa/desactiva en Leeme.")
wd.freeze_panes = "B13"
wd.column_dimensions["A"].width = 24
for c_ in range(2, 18):
    wd.column_dimensions[L(c_)].width = 13
wd.row_dimensions[HR].height = 30

# ============================================================================ Tablas (formato de la plantilla)
P0, P1 = 5, 105
SRC = {"H": dict(nac="B4", d="L", p="O"), "M": dict(nac="C4", d="M", p="P"), "A": dict(nac="E4", d="N", p="Q")}
for s in ("A", "H", "M"):
    ws = wb.create_sheet(HOJA[s])
    ws.sheet_view.showGridLines = False
    ws.merge_cells("B1:N1")
    ws["B1"] = "Tabla de mortalidad"
    ws["B1"].font = F(size=14, bold=True)
    ws.merge_cells("B2:H2")
    ws["B2"] = f"Cuadro: Tabla de mortalidad de México, 2019, {NOMBRE[s].lower()} (INEGI)"
    ws["B2"].font = F(bold=True)
    for j, t in enumerate(["x", "qx", "lx", "dx", "Lx", "Tx", "ex"]):
        cab(ws.cell(3, 2 + j), t, FILL_S[s], 11)
    desc = ["", "Probabilidad de muerte: defunciones xi / PI", "Supervivientes a una edad exacta", "Defunciones xi",
            "Años vividos entre aniversarios", "Años vividos acumulados: años que quedan por vivir",
            "La media de vida de una generación hasta su extinción"]
    for j, t in enumerate(desc):
        c = ws.cell(4, 2 + j, t)
        c.font, c.alignment, c.border = F(size=9, italic=True), CENTRO, BORDE
    ws.row_dimensions[4].height = 52
    sc = SRC[s]
    for i in range(101):
        r = P0 + i
        ws.cell(r, 2, i).font = F(bold=True)
        ws.cell(r, 2).alignment = Alignment(horizontal="center")
        if i == 100:
            ws.cell(r, 2).number_format = '0"+"'
        dref, pref = f"'Datos INEGI'!{sc['d']}{R0+i}", f"'Datos INEGI'!{sc['p']}{R0+i}"
        if i == 0:
            ws.cell(r, 3, "=L9")
        elif i == 100:
            ws.cell(r, 3, 1)
        else:
            ws.cell(r, 3, f"={dref}/({pref}+{dref}/2)")
        ws.cell(r, 4, "=Radix" if i == 0 else f"=D{r-1}-E{r-1}")
        ws.cell(r, 5, f"=D{r}*C{r}")
        ws.cell(r, 6, f"=E{r}/2" if i == 100 else f"=D{r+1}+(E{r}/2)")
        ws.cell(r, 7, f"=SUM(F{r}:F${P1})")
        ws.cell(r, 8, f"=G{r}/D{r}")
        for c_, fmt in zip(range(3, 9), ["0.000000", "#,##0", "#,##0", "#,##0", "#,##0", "0.00"]):
            cc = ws.cell(r, c_)
            cc.number_format = fmt
            cc.font = F(color=VERDE if (c_ == 3 and 0 < i < 100) else NEGRO)
            if i == 100 and c_ == 3:
                cc.font = F(color=AZUL)
        for c_ in range(2, 9):
            ws.cell(r, c_).border = BORDE
    # nota técnica (como en la plantilla)
    ws["K3"] = "Nota técnica (qx)"
    ws["K3"].font = F(bold=True)
    ws["K4"] = "q(x): Probabilidad de muerte"
    ws["K6"], ws["L6"] = "Nacimientos", f"='Datos INEGI'!{sc['nac']}"
    ws["K7"], ws["L7"] = "Defunciones (menores de 1 año)", f"='Datos INEGI'!{sc['d']}{R0}"
    ws["K9"], ws["L9"] = "q(0)", "=L7/L6"
    ws["L6"].number_format = ws["L7"].number_format = "#,##0"
    ws["L9"].number_format = "0.000000"
    ws["L6"].font = ws["L7"].font = F(color=VERDE)
    for ref in ("K4", "K6", "K7", "K9", "K12", "K14", "K16", "K17", "K19"):
        ws[ref].font = F(bold=True)
    ws["K10"] = "q(0) = defunciones de menores de 1 año / nacimientos"
    ws["K11"] = "q(x), x = 1…99 = defunciones / (población + defunciones/2)"
    ws["K12"] = "Lx"
    ws["L12"] = "Años vividos por los sobrevivientes entre la edad x y x+1"
    ws["K14"] = "Lx = l(x+1) + (dx/2)"
    ws["K16"] = "Tx = Lx + L(x+1) + … + L100 (suma de Lx)"
    ws["K17"] = "ex = Tx / lx"
    ws["K19"], ws["L19"] = "Control: suma de dx − radix", f"=SUM(E{P0}:E{P1})-Radix"
    ws["L19"].number_format = "0.000;-0.000;0.000"
    ws["K21"] = "Edad 100 = grupo abierto (100 y más): qx = 1, L100 = d100/2"
    for ref in ("K10", "K11", "L12", "K14", "K16", "K17", "K21"):
        ws[ref].font = F(size=9, italic=True, color=GRISTX)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 7
    for c_ in "CDEFGH":
        ws.column_dimensions[c_].width = 21
    ws.column_dimensions["I"].width = 3
    ws.column_dimensions["J"].width = 3
    ws.column_dimensions["K"].width = 30
    ws.column_dimensions["L"].width = 20
    ws.freeze_panes = "C5"

# ============================================================================ Datos graficos
wg = wb.create_sheet("Datos graficos")
wg.sheet_view.showGridLines = False
wg["A1"] = "Datos de los gráficos (vínculos a las tablas)"
wg["A1"].font = F(size=14, bold=True)
cab(wg["A4"], "x")
bloques = [("lx", "D", "#,##0"), ("qx", "C", "0.000000"), ("dx", "E", "#,##0"), ("ex", "H", "0.00")]
GCOL = {}
c_ = 2
for nombre, col, fmt in bloques:
    wg.merge_cells(start_row=3, start_column=c_, end_row=3, end_column=c_ + 2)
    cab(wg.cell(3, c_), nombre)
    for s in ("H", "M", "A"):
        cab(wg.cell(4, c_), NOMBRE[s], FILL_S[s])
        GCOL[(nombre, s)] = c_
        for i in range(101):
            cc = wg.cell(5 + i, c_, f"='{HOJA[s]}'!{col}{P0+i}")
            cc.font, cc.number_format = F(color=VERDE), fmt
        c_ += 1
cab(wg.cell(4, c_), "qx H / qx M")
wg.cell(3, c_, "Sobremortalidad masculina").font = F(bold=True)
GCOL["razon"] = c_
for i in range(101):
    wg.cell(5 + i, 1, i).font = F(bold=True)
    cc = wg.cell(5 + i, c_, f"=IF({L(GCOL[('qx','M')])}{5+i}>0,{L(GCOL[('qx','H')])}{5+i}/{L(GCOL[('qx','M')])}{5+i},\"\")")
    cc.number_format = "0.00"
    cc.font = F()
wg.freeze_panes = "B5"
for k in range(1, c_ + 1):
    wg.column_dimensions[L(k)].width = 13

# ============================================================================ Indicadores
wi = wb.create_sheet("Indicadores")
wi.sheet_view.showGridLines = False
wi["A1"] = "Indicadores de la tabla de mortalidad de México, 2019 (INEGI)"
wi["A1"].font = F(size=14, bold=True)
cab(wi["A3"], "Indicador")
for j, s in enumerate(("H", "M", "A")):
    cab(wi.cell(3, 2 + j), NOMBRE[s], FILL_S[s])
cab(wi["E3"], "Mujeres − Hombres")


def hs(s):
    return f"'{HOJA[s]}'!"


filas_ind = [
    ("e0: esperanza de vida al nacer (años)", lambda h: f"={h}H{P0}", "0.00"),
    ("e1: esperanza de vida a los 1 año (años)", lambda h: f"={h}H{P0+1}", "0.00"),
    ("e15: esperanza de vida a los 15 años (años)", lambda h: f"={h}H{P0+15}", "0.00"),
    ("e65: esperanza de vida a los 65 años (años)", lambda h: f"={h}H{P0+65}", "0.00"),
    ("e80: esperanza de vida a los 80 años (años)", lambda h: f"={h}H{P0+80}", "0.00"),
    ("q(0): mortalidad infantil (por mil nacimientos)", lambda h: f"={h}C{P0}*1000", "0.00"),
    ("5q0: probabilidad de morir antes de los 5 años (por mil)", lambda h: f"=(1-{h}D{P0+5}/{h}D{P0})*1000", "0.00"),
    ("Probabilidad de llegar con vida a los 15 años", lambda h: f"={h}D{P0+15}/{h}D{P0}", "0.0%"),
    ("Probabilidad de llegar con vida a los 65 años", lambda h: f"={h}D{P0+65}/{h}D{P0}", "0.0%"),
    ("Probabilidad de llegar con vida a los 80 años", lambda h: f"={h}D{P0+80}/{h}D{P0}", "0.0%"),
    ("Edad mediana a la muerte (años cumplidos)", lambda h: f"=INDEX({h}B{P0}:B{P1},MATCH(Radix/2,{h}D{P0}:D{P1},-1))", "0"),
    ("Edad modal de las defunciones adultas (15-99 años)", lambda h: f"=INDEX({h}B{P0+15}:B{P1-1},MATCH(MAX({h}E{P0+15}:E{P1-1}),{h}E{P0+15}:E{P1-1},0))", "0"),
]
for i, (nom, fn, fmt) in enumerate(filas_ind):
    r = 4 + i
    wi.cell(r, 1, nom).font = F(bold=True)
    for j, s in enumerate(("H", "M", "A")):
        c = wi.cell(r, 2 + j, fn(hs(s)))
        c.number_format, c.font, c.border = fmt, F(color=VERDE), BORDE
    if fmt not in ("0",):
        c = wi.cell(r, 5, f"=C{r}-B{r}")
        c.number_format = "0.00;-0.00;0.00" if "%" not in fmt else "0.0%;-0.0%;0.0%"
        c.font, c.border = F(), BORDE
    wi.cell(r, 1).border = BORDE
r_val = 4 + len(filas_ind) + 2
wi.cell(r_val, 1, "Validación: e0 calculada con datos de INEGI vs. e0 de la ONU (2019)").font = F(bold=True, size=11)
cab(wi.cell(r_val + 1, 1), "Concepto")
for j, s in enumerate(("H", "M", "A")):
    cab(wi.cell(r_val + 1, 2 + j), NOMBRE[s], FILL_S[s])
conceptos = [("e0 con datos de INEGI (esta tabla)", lambda j, s: f"=B4".replace("B", L(2 + j)), "0.00"),
             ("e0 ONU, World Population Prospects 2024", lambda j, s: f"='Datos INEGI'!{L(2+j)}7", "0.00"),
             ("Diferencia (años)", lambda j, s: f"={L(2+j)}{r_val+2}-{L(2+j)}{r_val+3}", "0.00;-0.00;0.00"),
             ("Diferencia (%)", lambda j, s: f"={L(2+j)}{r_val+4}/{L(2+j)}{r_val+3}", "0.0%")]
for i, (nom, fn, fmt) in enumerate(conceptos):
    r = r_val + 2 + i
    wi.cell(r, 1, nom).font = F(bold=True)
    wi.cell(r, 1).border = BORDE
    for j, s in enumerate(("H", "M", "A")):
        c = wi.cell(r, 2 + j, fn(j, s))
        c.number_format, c.border = fmt, BORDE
        c.font = F(color=VERDE if i == 1 else NEGRO)
R_VAL = r_val + 2          # fila de e0 INEGI en la tabla de validación
wi.column_dimensions["A"].width = 58
for c_ in "BCDE":
    wi.column_dimensions[c_].width = 17

# ---- texto del análisis (cifras de la tabla de referencia, que el script de verificación compara con el libro)
a, h, m = IND["A"], IND["H"], IND["M"]
q = {s: REF[s].set_index("x").qx for s in REF}
razon = (q["H"] / q["M"])
r_pico_edad = int(razon[15:40].idxmax())
r_media_15_39 = float(razon[15:40].mean())
onu_e0 = onu.set_index("sexo").e0_2019
dif_a = a["e0"] - float(onu_e0["Ambos sexos"])
qmin_edad = int(q["A"][1:30].idxmin())
p100 = REF["A"].lx[100] / 1e5
onu_q = pd.read_csv(DATOS / "onu_wpp2024_tabla_vida_2019_mexico.csv")
onu_qa = onu_q[onu_q.sexo == "Ambos sexos"].set_index("edad").qx
rat = (REF["A"].set_index("x").qx / onu_qa)
dec = {a: float(rat[a:a + 9].mean()) for a in range(0, 100, 10)}
BAJO_20_79 = (100 * (1 - max(dec[a] for a in (20, 30, 40, 50, 60, 70))), 100 * (1 - min(dec[a] for a in (20, 30, 40, 50, 60, 70))))
BAJO_80 = (100 * (1 - max(dec[80], dec[90])), 100 * (1 - min(dec[80], dec[90])))
ANALISIS = [
    ("1. Resultado principal",
     f"Con las defunciones y nacimientos de INEGI de 2019 y la población del Censo 2020, la esperanza de vida al nacer en México es de "
     f"{a['e0']:.1f} años (hombres {h['e0']:.1f}; mujeres {m['e0']:.1f}). Las mujeres viven en promedio {m['e0'] - h['e0']:.1f} años más. "
     f"De cada 100,000 nacidos vivos, {a['l65'] * 1e5:,.0f} llegan a los 65 años y {a['l80'] * 1e5:,.0f} a los 80; la mitad ya murió a los "
     f"{a['mediana_muerte']} años (hombres {h['mediana_muerte']}; mujeres {m['mediana_muerte']})."),
    ("2. Forma de la mortalidad por edad",
     f"La curva de qx tiene la forma clásica: alta en el primer año (q0 = {a['q0_mil']:.1f} por mil), mínima entre los 8 y los 12 años "
     f"(el valor más bajo ocurre a los {qmin_edad} años) y creciente después, casi de forma exponencial. La esperanza de vida a los 1 año "
     f"({a['e1']:.1f}) ya no supera a la del nacimiento ({a['e0']:.1f}), señal de que la mortalidad infantil es hoy relativamente baja; "
     f"la probabilidad de morir antes de los 5 años es {a['5q0_mil']:.1f} por mil."),
    ("3. Sobremortalidad masculina",
     f"Entre los 15 y los 39 años la probabilidad de morir de los hombres es en promedio {r_media_15_39:.1f} veces la de las mujeres, con un "
     f"máximo de {razon[15:40].max():.1f} veces a los {r_pico_edad} años: en esas edades pesan las muertes violentas y accidentales. "
     f"Esa diferencia explica por qué solo {h['l65'] * 100:.0f} de cada 100 hombres llegan a los 65 años, frente a {m['l65'] * 100:.0f} de cada 100 mujeres."),
    ("4. Comparación con la ONU",
     f"La e0 con cifras registradas de INEGI ({a['e0']:.1f}) queda {dif_a:.1f} años por encima de la estimación de la ONU para 2019 "
     f"({float(onu_e0['Ambos sexos']):.1f}). Las probabilidades de morir de esta tabla son entre {BAJO_20_79[0]:.0f} % y {BAJO_20_79[1]:.0f} % menores "
     f"que las de la ONU de los 20 a los 79 años y entre {BAJO_80[0]:.0f} % y {BAJO_80[1]:.0f} % menores de los 80 años en adelante (promedios por década). "
     f"Es la dirección esperada si hay subregistro de defunciones y si el Censo exagera la población de edades avanzadas."),
    ("5. Limitaciones",
     f"(i) El Censo tiene atracción de edades redondas; por eso se suavizó con media móvil de 5 edades (interruptor en Leeme). "
     f"(ii) A partir de los 85-90 años la población censal está sobrestimada: qx se aplana en ~0.17-0.19 y {p100 * 100:.1f} % de los nacidos "
     f"llegaría a los 100 años, lo cual no es creíble; las esperanzas de vida a edades avanzadas son optimistas. "
     f"(iii) Se usó 2019 por ser el último año completo antes de la pandemia; en 2020-2021 la mortalidad fue mucho mayor. "
     f"(iv) Es una tabla de periodo (generación ficticia), no el seguimiento de una generación real."),
]
r = R_VAL + 6
wi.cell(r, 1, "Análisis").font = F(bold=True, size=12)
for tit, txt in ANALISIS:
    r += 1
    wi.cell(r, 1, tit).font = F(bold=True)
    r += 1
    wi.merge_cells(start_row=r, start_column=1, end_row=r, end_column=5)
    wi.cell(r, 1, txt).font = F()
    wi.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
    wi.row_dimensions[r].height = 15 * (len(txt) // 105 + 1) + 6
nota(wi, f"A{r+2}", "El texto del análisis está redactado con los parámetros por defecto (radix 100,000; suavizado activado).")

# ============================================================================ Gráficos
wc = wb.create_sheet("Graficos")
wc.sheet_view.showGridLines = False
wc["A1"] = "Gráficos de la tabla de mortalidad de México, 2019 (INEGI)"
wc["A1"].font = F(size=14, bold=True)
nota(wc, "A2", "Se grafican las edades 0-99; la edad 100 es el grupo abierto (qx = 1) y distorsionaría la escala.")


def linea(titulo, nombre, ytitulo, ancla, log=False, series=("H", "M", "A")):
    ch = LineChart()
    ch.title = titulo
    ch.height, ch.width = 8.5, 15.5
    for s in series:
        ch.add_data(Reference(wg, min_col=GCOL[(nombre, s)], min_row=4, max_row=104), titles_from_data=True)
    ch.set_categories(Reference(wg, min_col=1, min_row=5, max_row=104))
    for se, s in zip(ch.series, series):
        se.graphicalProperties.line.solidFill = COLOR[s]
        se.graphicalProperties.line.width = 22000
        se.smooth = False
        se.marker.symbol = "none"
    ch.x_axis.title = "Edad x"
    ch.y_axis.title = ytitulo
    ch.x_axis.tickLblSkip = 10
    ch.x_axis.tickMarkSkip = 10
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.legend.position = "b"
    if log:
        ch.y_axis.scaling.logBase = 10
    wc.add_chart(ch, ancla)


linea("lx: sobrevivientes de 100,000 nacidos", "lx", "lx", "A3")
linea("qx: probabilidad de morir (escala logarítmica)", "qx", "qx (log)", "K3", log=True)
linea("dx: defunciones de la tabla por edad", "dx", "dx", "A21")
linea("ex: esperanza de vida por edad", "ex", "años", "K21")

rz = LineChart()
rz.title = "Sobremortalidad masculina: qx hombres / qx mujeres"
rz.height, rz.width = 8.5, 15.5
rz.add_data(Reference(wg, min_col=GCOL["razon"], min_row=4, max_row=104), titles_from_data=True)
rz.set_categories(Reference(wg, min_col=1, min_row=5, max_row=104))
rz.series[0].graphicalProperties.line.solidFill = "2A78D6"
rz.series[0].graphicalProperties.line.width = 22000
rz.series[0].smooth = False
rz.series[0].marker.symbol = "none"
rz.x_axis.title, rz.y_axis.title = "Edad x", "razón"
rz.x_axis.tickLblSkip = rz.x_axis.tickMarkSkip = 10
rz.x_axis.delete = rz.y_axis.delete = False
rz.legend = None
wc.add_chart(rz, "A39")

bc = BarChart()
bc.type = "col"
bc.title = "e0 2019: tabla con datos de INEGI vs. ONU"
bc.height, bc.width = 8.5, 15.5
for fila, nom, color in ((R_VAL, "Esta tabla (INEGI)", "2A78D6"), (R_VAL + 1, "ONU (WPP 2024)", "9AA0A6")):
    se = Series(Reference(wi, min_col=2, max_col=4, min_row=fila), title=nom)
    se.graphicalProperties.solidFill = color
    bc.series.append(se)
bc.set_categories(Reference(wi, min_col=2, max_col=4, min_row=R_VAL - 1))
bc.y_axis.scaling.min = 60
bc.y_axis.title = "años"
bc.x_axis.delete = bc.y_axis.delete = False
bc.legend.position = "b"
wc.add_chart(bc, "K39")

# ============================================================================ Leeme
ws_l.sheet_view.showGridLines = False
ws_l["A1"] = "U2_P1: Tabla de mortalidad INEGI - México, 2019"
ws_l["A1"].font = F(size=14, bold=True)
ws_l["A3"], ws_l["B3"] = "Radix (l0)", 100000
ws_l["A4"], ws_l["B4"] = "Suavizado (1 = sí, 0 = no)", 1
for ref in ("B3", "B4"):
    ws_l[ref].font, ws_l[ref].fill = F(bold=True, color=AZUL), AMARILLO
ws_l["B3"].number_format = "#,##0"
ws_l["C3"] = "Población inicial de la tabla (nombre definido 'Radix')."
ws_l["C4"] = "Media móvil de 5 edades en defunciones y población de 10 a 97 años (nombre definido 'Suaviza')."
for ref in ("C3", "C4"):
    ws_l[ref].font = F(italic=True, color=GRISTX, size=9)
wb.defined_names["Radix"] = DefinedName("Radix", attr_text="Leeme!$B$3")
wb.defined_names["Suaviza"] = DefinedName("Suaviza", attr_text="Leeme!$B$4")
bloques_txt = [
    ("Qué es", "Tabla de mortalidad de México para 2019 (ambos sexos, hombres y mujeres) con el formato de la plantilla de la profesora "
               "(x, qx, lx, dx, Lx, Tx, ex y nota técnica de nacimientos y defunciones), calculada con cifras de INEGI."),
    ("Datos (INEGI)", "Defunciones: Estadísticas de Defunciones Registradas, datos abiertos 2019-2021 (https://www.inegi.org.mx/programas/edr/). "
                      "Nacimientos: Estadística de Nacimientos Registrados, datos abiertos 2019-2021 (https://www.inegi.org.mx/programas/natalidad/). "
                      "Población por edad simple: Censo de Población y Vivienda 2020 (https://www.inegi.org.mx/programas/ccpv/2020/). "
                      "Año de referencia 2019 (último año completo antes de la pandemia)."),
    ("Método", "q(0) = defunciones de menores de 1 año / nacimientos.  q(x), x = 1…99 = defunciones / (población + defunciones/2).  q(100) = 1 (grupo abierto)."),
    ("", "lx: l0 = radix; l(x+1) = lx − dx.  dx = lx·qx.  Lx = l(x+1) + dx/2.  Tx = suma de Lx desde x hasta 100.  ex = Tx/lx."),
    ("", "Defunciones por año de ocurrencia; la edad no especificada se reparte proporcionalmente entre las edades (hoja Datos INEGI)."),
    ("Hojas", "Tabla ambos sexos / Tabla hombres / Tabla mujeres: formato de la plantilla.  Indicadores: e0, e65, supervivencia, validación con la ONU y análisis.  "
              "Graficos: lx, qx, dx, ex, sobremortalidad masculina y comparación con la ONU.  Datos INEGI y Datos graficos: apoyo."),
    ("Colores", "Azul = dato de INEGI o parámetro.  Negro = fórmula.  Verde = vínculo a otra hoja.  Amarillo = parámetro editable."),
    ("Respecto a la plantilla", "La plantilla (generación francesa de 1899) trae qx escritos a mano y varios Tx fijos; aquí todo se calcula con fórmulas: "
                                "Tx es la suma de Lx (no de lx) y no se redondea dx ni Lx."),
]
r = 6
for k, v in bloques_txt:
    ws_l.cell(r, 1, k).font = F(bold=True)
    ws_l.cell(r, 1).alignment = Alignment(vertical="top")
    ws_l.merge_cells(start_row=r, start_column=2, end_row=r, end_column=12)
    ws_l.cell(r, 2, v).font = F()
    ws_l.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
    ws_l.row_dimensions[r].height = 15 * (len(v) // 120 + 1) + 4
    r += 1
ws_l.column_dimensions["A"].width = 24
for c_ in range(2, 13):
    ws_l.column_dimensions[L(c_)].width = 11

orden = ["Leeme", "Tabla ambos sexos", "Tabla hombres", "Tabla mujeres", "Indicadores", "Graficos", "Datos INEGI", "Datos graficos"]
wb._sheets = [wb[n] for n in orden]
for n in orden:
    wb[n].sheet_properties.tabColor = "1F3864" if n.startswith("Tabla") else ("2A78D6" if n in ("Indicadores", "Graficos") else "808080")
    wb[n].page_setup.orientation = "landscape" if n in ("Indicadores", "Graficos", "Datos INEGI") else "portrait"
    wb[n].sheet_properties.pageSetUpPr.fitToPage = True
    wb[n].page_setup.fitToWidth = 1
    wb[n].page_setup.fitToHeight = 0

SALIDA.parent.mkdir(exist_ok=True)
wb.save(SALIDA)
print("guardado", SALIDA)
