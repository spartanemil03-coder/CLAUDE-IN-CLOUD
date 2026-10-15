"""Gráficos de las tablas abreviadas (lx, ex y nqx en escala log) a partir de calculo_abreviada.py. Guarda PNG en resultados/."""
import sys
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

sys.path.insert(0, str(Path(__file__).resolve().parent))
import calculo_abreviada as ca  # noqa: E402

SALIDA = Path(__file__).resolve().parent.parent / "resultados"
COL = {"H": "#1f77b4", "M": "#d95f02", "A": "#555555"}
tablas = {s: ca.tabla(s) for s in "HMA"}


def figura(nombre, titulo, ylabel, campo, logy=False, hasta=16):
    fig, ax = plt.subplots(figsize=(7.5, 4.2), dpi=150)
    for s, t in tablas.items():
        ax.plot(t.x[:hasta], t[campo][:hasta], marker="o", ms=3.5, lw=1.8, color=COL[s], label=ca.SEXOS[s], ls="--" if s == "A" else "-")
    ax.set_title(titulo, fontsize=11, loc="left")
    ax.set_xlabel("Edad exacta x")
    ax.set_ylabel(ylabel)
    if logy:
        ax.set_yscale("log")
    ax.grid(alpha=0.3)
    ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(SALIDA / nombre)
    plt.close(fig)


figura("lx_sobrevivientes.png", "Sobrevivientes lx de una cohorte ficticia de 100,000 (México, 2025)", "lx", "lx", hasta=17)
figura("ex_esperanza.png", "Esperanza de vida ex por edad (México, 2025)", "ex (años)", "ex", hasta=17)
figura("qx_log.png", "Probabilidad de morir nqx por grupo de edad (escala log, sin el grupo abierto)", "nqx", "nqx", logy=True, hasta=16)
print("PNG guardados en", SALIDA)
