"""Cálculo de referencia (numpy/pandas) de la tabla de mortalidad de la generación mexicana de 1970 con datos de CONAPO.

Se usa para verificar de forma independiente el libro de Excel y para redactar el análisis. Reglas (las mismas del libro):
  - Datos nacionales = suma de los 32 estados y de ambos sexos (CONAPO publica estos archivos solo por entidad).
  - q(0) = defunciones de menores de 1 año / nacimientos (nota técnica de la plantilla).
  - q(x) = D / (P + D/2) para x = 1..99, con D y P (población a mitad de año) de la edad x en el año g + x.
  - q(100) = 1 (grupo abierto, 100 y más).
  - l0 = 100,000; dx = ROUND(lx * qx); l(x+1) = lx - dx; Lx = ROUND(l(x+1) + dx/2); Tx = suma de Lx; ex = Tx / lx.
"""
from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
BASES = RAIZ / "bases_de_datos_CONAPO"
F_DEF = BASES / "01_Defunciones_1950_2070.csv"
F_POB = BASES / "00_Pob_Mitad_1950_2070.csv"
F_IND = BASES / "05_indicadores_demograficos_proyecciones.csv"
GENERACION = 1970
RADIX = 100_000


def redondea(x):
    """ROUND de Excel (mitad hacia arriba para positivos)."""
    return np.floor(np.asarray(x, float) + 0.5)


def cargar():
    d = pd.read_csv(F_DEF)
    p = pd.read_csv(F_POB)
    ind = pd.read_csv(F_IND)
    ind = ind[ind.CVE_GEO == 0].set_index("ANIO")
    D = d.groupby(["EDAD", "ANIO"]).DEFUNCIONES.sum().unstack()   # edad x año, ambos sexos, 32 estados
    P = p.groupby(["EDAD", "ANIO"]).POBLACION.sum().unstack()
    return D, P, ind


def qx_generacion(D, P, nac, g=GENERACION):
    q = np.zeros(101)
    for x in range(100):
        q[x] = D.loc[x, g + x] / (P.loc[x, g + x] + D.loc[x, g + x] / 2)
    q[0] = D.loc[0, g] / nac
    q[100] = 1.0
    return q


def tabla(q, radix=RADIX):
    l = np.zeros(101)
    d = np.zeros(101)
    L = np.zeros(101)
    l[0] = radix
    for x in range(101):
        d[x] = redondea(l[x] * q[x])
        if x < 100:
            l[x + 1] = l[x] - d[x]
    for x in range(101):
        L[x] = redondea((l[x + 1] if x < 100 else 0) + d[x] / 2)
    T = L[::-1].cumsum()[::-1]
    return pd.DataFrame({"x": np.arange(101), "qx": q, "lx": l, "dx": d, "Lx": L, "Tx": T, "ex": T / l})


def tabla_periodo_e0(D, P, nac, t):
    """e0 de periodo (año calendario t) con las mismas fórmulas; sirve para validar contra la e0 oficial de CONAPO."""
    q = np.zeros(101)
    for x in range(100):
        q[x] = D.loc[x, t] / (P.loc[x, t] + D.loc[x, t] / 2)
    q[0] = D.loc[0, t] / nac
    q[100] = 1.0
    l = np.zeros(102)
    l[0] = 1.0
    for x in range(101):
        l[x + 1] = l[x] * (1 - q[x])
    L = l[1:] + l[:101] * q / 2
    return float(L.sum())
