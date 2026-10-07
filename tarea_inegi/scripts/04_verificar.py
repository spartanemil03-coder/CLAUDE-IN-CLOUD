"""Verifica el libro recalculado contra el cálculo independiente de calculo_inegi.py.

Uso:  python3 scripts/04_verificar.py        (AssertionError si algo no coincide)
"""
import sys
from pathlib import Path

import numpy as np
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parent))
import calculo_inegi as calc  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
wb = load_workbook(RAIZ / "resultados" / "T_Mortalidad_INEGI_Mexico_2019.xlsx", data_only=True)
assert wb["Leeme"]["B4"].value == 1, "el libro debe entregarse con el suavizado activado"
REF = calc.tablas(RAIZ / "datos", suaviza=True)
HOJA = {"A": "Tabla ambos sexos", "H": "Tabla hombres", "M": "Tabla mujeres"}
maxdif = 0.0
for s, hoja in HOJA.items():
    ws, t = wb[hoja], REF[s]
    for col, k in zip("BCDEFGH", ["x", "qx", "lx", "dx", "Lx", "Tx", "ex"]):
        xl = np.array([ws[f"{col}{5 + i}"].value for i in range(101)], float)
        ref = t[k].values.astype(float)
        dif = float(np.max(np.abs(xl - ref) / np.maximum(1, np.abs(ref))))
        maxdif = max(maxdif, dif)
        assert dif < 1e-9, f"{hoja} {k}: dif relativa {dif}"
    assert abs(ws["L19"].value) < 1e-6, "Σdx ≠ radix"
    assert abs(ws["L6"].value - t.attrs["nacimientos"]) < 1e-9
    # indicadores de la hoja Indicadores
    wi = wb["Indicadores"]
    col = {"H": "B", "M": "C", "A": "D"}[s]
    ind = calc.indicadores(t)
    for fila, k in zip(range(4, 16), ["e0", "e1", "e15", "e65", "e80", "q0_mil", "5q0_mil", "l15", "l65", "l80", "mediana_muerte", "moda_adulta"]):
        esperado = t.lx[15] / 1e5 if k == "l15" else ind[k]
        assert abs(wi[f"{col}{fila}"].value - esperado) < 1e-6, (s, k, wi[f"{col}{fila}"].value, esperado)
print(f"OK: libro verificado contra el cálculo independiente (diferencia relativa máx {maxdif:.2e})")
for s in "HMA":
    i = calc.indicadores(REF[s])
    print(f"  {calc.SEXOS[s]:12s} e0 = {i['e0']:.2f}  e65 = {i['e65']:.2f}  q0 = {i['q0_mil']:.2f}‰  l65 = {i['l65']:.3f}")
