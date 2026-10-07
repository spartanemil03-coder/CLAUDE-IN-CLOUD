"""Verifica el libro recalculado contra el cálculo independiente (calculo_inegi.py), los controles de las hojas de datos
y que el interruptor de suavizado funcione.

Uso:  python3 scripts/04_verificar.py        (AssertionError si algo no coincide)
"""
import subprocess
import sys
import tempfile
from pathlib import Path

import numpy as np
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import calculo_inegi as calc  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
LIBRO = RAIZ / "resultados" / "T_Mortalidad_INEGI_Mexico_2019.xlsx"


def comparar(wb, suaviza):
    ws = wb["Tabla de mortalidad"]
    q, d, p, b = calc.qx("A", suaviza)
    t = calc.tabla(q)
    maxdif = 0.0
    for col, k in zip("BCDEFGH", ["x", "qx", "lx", "dx", "Lx", "Tx", "ex"]):
        xl = np.array([ws[f"{col}{5 + i}"].value for i in range(101)], float)
        dif = float(np.max(np.abs(xl - t[k].values) / np.maximum(1, np.abs(t[k].values))))
        maxdif = max(maxdif, dif)
        assert dif < 1e-9, f"{k}: dif relativa {dif}"
    assert ws["L25"].value == 0 and ws["D5"].value == 100000
    assert ws["L6"].value == b and abs(ws["L9"].value - q[0]) < 1e-12
    return maxdif, t


wb = load_workbook(LIBRO, data_only=True)
assert wb["Tabla de mortalidad"]["L18"].value == 1, "el libro debe entregarse con el suavizado activado"
maxdif, t = comparar(wb, True)
# controles de las hojas de datos
assert wb["Defunciones"]["I9"].value == 0, "total de defunciones por archivo ≠ total por edad"
dfn, nac, pob = calc.cargar()
assert wb["Defunciones"]["E107"].value == dfn.values.sum()
assert wb["Nacimientos"]["F8"].value == nac.sum()
assert wb["Población"]["D107"].value == pob.values.sum() == 126_014_024
assert [wb["Defunciones"].cell(5 + i, 5).value for i in range(101)] == list(dfn[["hombres", "mujeres", "sexo_no_especificado"]].iloc[:101].sum(axis=1))
print(f"OK: libro verificado contra el cálculo independiente (dif. relativa máx {maxdif:.2e}); controles de datos en 0")

# interruptor de suavizado: se apaga en una copia, se recalcula con LibreOffice y se compara con la referencia sin suavizar
with tempfile.TemporaryDirectory() as tmp:
    copia = Path(tmp) / "prueba.xlsx"
    w = load_workbook(LIBRO)
    w["Tabla de mortalidad"]["L18"] = 0
    w.save(copia)
    skill = next(Path("/root/.claude/skills/synced").glob("*/xlsx/scripts/recalc.py"))
    subprocess.run([sys.executable, str(skill), str(copia), "300"], check=True, capture_output=True)
    ws2 = load_workbook(copia, data_only=True)["Tabla de mortalidad"]
    q0, *_ = calc.qx("A", False)
    t0 = calc.tabla(q0)
    dif = max(float(np.max(np.abs(np.array([ws2[f"{c}{5 + i}"].value for i in range(101)], float) - t0[k].values) / np.maximum(1, np.abs(t0[k].values))))
              for c, k in zip("CDEFGH", ["qx", "lx", "dx", "Lx", "Tx", "ex"]))
    assert dif < 1e-9
    print(f"OK: con suavizado = 0 el libro coincide con la referencia sin suavizar (e0 = {ws2['H5'].value:.2f})")
print(f"México 2019, ambos sexos: e0 = {t.ex[0]:.2f}, e65 = {t.ex[65]:.2f}, q0 = {t.qx[0] * 1000:.2f} por mil, l65 = {t.lx[65] / 1e5:.3f}")
