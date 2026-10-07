"""Gráficos PNG de la tabla de mortalidad de México 2019 (INEGI), a partir de los valores calculados del libro de Excel.

Uso:  python3 scripts/03_graficos.py
"""
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from matplotlib.ticker import FuncFormatter
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import calculo_inegi as calc  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
LIBRO = RAIZ / "resultados" / "T_Mortalidad_INEGI_Mexico_2019.xlsx"
AZUL, NARANJA, GRIS = "#2a78d6", "#eb6834", "#9aa0a6"
INK, INK2, GRID, BG = "#0b0b0b", "#52514e", "#e4e3df", "#fcfcfb"
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": GRID, "axes.labelcolor": INK2,
                     "xtick.color": INK2, "ytick.color": INK2, "text.color": INK, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.grid": True, "grid.color": GRID, "grid.linewidth": 0.8,
                     "axes.axisbelow": True, "figure.facecolor": BG, "axes.facecolor": BG, "savefig.facecolor": BG})
ws = load_workbook(LIBRO, data_only=True)["Tabla de mortalidad"]
T = {k: np.array([ws[f"{c}{5 + i}"].value for i in range(101)], float) for c, k in zip("BCDEFGH", ["x", "qx", "lx", "dx", "Lx", "Tx", "ex"])}
x = T["x"]
SUB = "México, 2019, ambos sexos. Defunciones y nacimientos: INEGI (EDR, ENR); población: Censo 2020"


def base(titulo, ylabel, xlim=(0, 100)):
    fig, ax = plt.subplots(figsize=(9, 5))
    ax.set_title(titulo, loc="left", fontsize=13, fontweight="bold", pad=22)
    ax.text(0, 1.02, SUB, transform=ax.transAxes, fontsize=8.5, color=INK2)
    ax.set_xlabel("Edad x (años cumplidos)")
    ax.set_ylabel(ylabel)
    ax.set_xlim(*xlim)
    return fig, ax


def guardar(fig, nombre):
    fig.tight_layout()
    fig.savefig(RAIZ / "resultados" / nombre, dpi=160)
    plt.close(fig)
    print("png", nombre)


# 1. lx
fig, ax = base("Sobrevivientes lx de una generación ficticia de 100,000 nacidos", "lx")
ax.plot(x, T["lx"], color=AZUL, lw=2.4)
ax.axhline(50000, color=INK2, lw=0.9, ls=(0, (2, 3)))
ax.text(1, 51500, "mitad de la generación", fontsize=8.5, color=INK2)
for edad, dy in ((65, 7000), (80, 8000)):
    ax.scatter([edad], [T["lx"][edad]], color=AZUL, zorder=3, s=36)
    ax.annotate(f"{T['lx'][edad]:,.0f} llegan a los {edad}", (edad, T["lx"][edad]), xytext=(edad + 2, T["lx"][edad] + dy), fontsize=9, color=INK)
ax.set_ylim(0, 102000)
ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{int(v):,}"))
guardar(fig, "lx_sobrevivientes.png")

# 2. qx (escala log)
fig, ax = base("Probabilidad de morir qx por edad (escala logarítmica)", "qx (log)", xlim=(0, 99))
ax.plot(x[:100], T["qx"][:100], color=AZUL, lw=2.4)
ax.set_yscale("log")
ax.set_ylim(1e-4, 0.5)
ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:g}"))
ax.grid(True, which="minor", color=GRID, lw=0.4)
ax.annotate(f"q0 = {T['qx'][0] * 1000:.1f} por mil", (0, T["qx"][0]), xytext=(6, T["qx"][0] * 1.3), fontsize=9, color=INK, arrowprops=dict(arrowstyle="-", color=INK2, lw=0.8))
ax.text(58, 2.2e-4, "A partir de ~90 años qx se aplana: el Censo exagera\nla población de edades avanzadas (ver análisis)", fontsize=8.5, color=INK2)
guardar(fig, "qx_log.png")

# 3. dx
fig, ax = base("Defunciones dx de la tabla, por edad", "dx (de 100,000 nacidos)", xlim=(0, 99))
ax.plot(x[:100], T["dx"][:100], color=AZUL, lw=2.4)
ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{int(v):,}"))
ax.annotate(f"primer año: {T['dx'][0]:,.0f}", (0.5, T["dx"][0]), xytext=(7, T["dx"][0] * 0.95), fontsize=9, color=INK, arrowprops=dict(arrowstyle="-", color=INK2, lw=0.8))
moda = int(15 + np.argmax(T["dx"][15:100]))
ax.annotate(f"edad más frecuente de muerte adulta: {moda} años", (moda, T["dx"][moda]), xytext=(24, T["dx"][moda] - 330), fontsize=9, color=INK, arrowprops=dict(arrowstyle="-", color=INK2, lw=0.8))
guardar(fig, "dx_defunciones.png")

# 4. ex
fig, ax = base("Esperanza de vida ex por edad", "ex (años por vivir)")
ax.plot(x, T["ex"], color=AZUL, lw=2.4)
for edad in (0, 65):
    ax.scatter([edad], [T["ex"][edad]], color=AZUL, zorder=3, s=36)
    ax.annotate(f"e{edad} = {T['ex'][edad]:.1f}", (edad, T["ex"][edad]), xytext=(edad + 3, T["ex"][edad] + 4), fontsize=9.5, color=INK)
ax.set_ylim(0, 85)
guardar(fig, "ex_esperanza.png")

# 5. sobremortalidad masculina (mismas fórmulas aplicadas a hombres y a mujeres con los datos de las hojas Defunciones y Población)
qh, qm = calc.qx("H")[0], calc.qx("M")[0]
r = qh[1:100] / qm[1:100]
fig, ax = base("Sobremortalidad masculina: qx de hombres / qx de mujeres", "veces", xlim=(1, 99))
ax.plot(x[1:100], r, color=AZUL, lw=2.4)
ax.axhline(1, color=INK2, lw=1)
pico = int(np.argmax(r[14:39]) + 15)
ax.annotate(f"{r[pico - 1]:.1f} veces a los {pico} años", (pico, r[pico - 1]), xytext=(pico + 8, r[pico - 1] + 0.15), fontsize=9, color=INK, arrowprops=dict(arrowstyle="-", color=INK2, lw=0.8))
ax.set_ylim(0.5, max(r) * 1.1)
guardar(fig, "sobremortalidad_masculina.png")

# 6. e0 con datos de INEGI vs. e0 oficial de CONAPO 2019 (referencia externa)
ref = pd.read_csv(RAIZ / "datos" / "conapo_e0_2019_referencia.csv").set_index("sexo").e0_2019
inegi = [calc.e0_continua(calc.qx(s)[0]) for s in "HMA"]
conapo = [ref["Hombres"], ref["Mujeres"], ref["Ambos sexos"]]
fig, ax = base("e0 2019: tabla con datos de INEGI vs. e0 oficial de CONAPO", "e0 (años)", xlim=(-0.6, 2.6))
xs = np.arange(3)
ax.bar(xs - 0.2, inegi, 0.38, color=AZUL, label="Esta tabla (INEGI)")
ax.bar(xs + 0.2, conapo, 0.38, color=GRIS, label="CONAPO (referencia externa)")
for xi, a, b in zip(xs, inegi, conapo):
    ax.text(xi - 0.2, a + 0.4, f"{a:.1f}", ha="center", fontsize=9)
    ax.text(xi + 0.2, b + 0.4, f"{b:.1f}", ha="center", fontsize=9)
    ax.text(xi, max(a, b) + 2.3, f"INEGI − CONAPO: +{a - b:.1f} años", ha="center", fontsize=9, color=INK2)
ax.set_xticks(xs)
ax.set_xticklabels(["Hombres", "Mujeres", "Ambos sexos"])
ax.set_xlabel("")
ax.set_ylim(60, 85)
ax.legend(frameon=False, loc="upper left")
guardar(fig, "e0_comparacion_conapo.png")
