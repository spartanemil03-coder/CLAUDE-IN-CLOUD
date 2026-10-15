"""Cálculo de referencia (numpy/pandas) de las tablas transversales abreviadas de México con datos de INEGI (EIC 2025 y defunciones).

Se usa para verificar de forma independiente el libro de Excel. Reglas (las mismas del libro):
  1. Defunciones de sexo no especificado: se reparten entre hombres y mujeres en proporción a sus defunciones, edad por edad.
  2. Defunciones de edad no especificada: se reparten proporcionalmente entre las edades, por sexo.
  3. Grupos de edad: 0, 1-4, 5-9, ..., 70-74, 75+ (75-79, 80-84 y 85+ de defunciones se suman, porque la EIC llega a 75+).
  4. La EIC no separa el año 0 del 1-4: se supone que las cohortes de nacimiento crecen (hacia atrás) a tasa constante r,
     r = ln(P[5-9]/P[0-4])/5, y la fracción de 0-4 que tiene menos de 1 año es (e^r - 1)/(e^(5r) - 1).
  5. Método del ejemplo de la profesora (Cuadro 4.3): nmx = D/P;  nqx = 2*n*nmx/(2 + n*nmx);  nAx = 0.12 (edad 0), 1.5 (1-4), 2.5 (resto);
     grupo abierto: nqx = 1 y nAx = 1/nmx.
  6. Cohorte ficticia l0 = 100,000;  l(x+n) = lx*npx;  ndx = lx*nqx;  nLx = n*l(x+n) + ndx*nAx;  Tx = suma de nLx hacia arriba;  ex = Tx/lx.
"""
from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "datos"
RADIX = 100_000
SEXOS = {"H": "Hombres", "M": "Mujeres", "A": "Ambos sexos"}
EDADES = [0, 1] + list(range(5, 80, 5))                  # edad inicial de cada grupo
N = np.array([1, 4] + [5] * 14 + [np.nan])               # amplitud (último grupo abierto)
ETIQ = ["0", "1-4"] + [f"{x}-{x + 4}" for x in range(5, 75, 5)] + ["75+"]


def cargar(datos=DATOS):
    pob = pd.read_csv(Path(datos) / "inegi_eic2025_poblacion_por_edad_sexo.csv")
    dfn = pd.read_csv(Path(datos) / "inegi_mortalidad_defunciones_por_edad_sexo.csv")
    return pob, dfn


def defunciones_ajustadas(dfn):
    """Devuelve (hombres, mujeres) por los 19 grupos de edad, ya sin 'sexo NE' ni 'edad NE'."""
    t = dfn[dfn.grupo_edad != "Total"].reset_index(drop=True)
    h, m, s = (t[c].astype(float).values for c in ("hombres", "mujeres", "sexo_no_especificado"))
    h2, m2 = h + s * h / (h + m), m + s * m / (h + m)    # paso 1 (incluye la fila de edad NE)
    h2, m2 = h2[:-1] * h2.sum() / h2[:-1].sum(), m2[:-1] * m2.sum() / m2[:-1].sum()   # paso 2
    return h2, m2


def agrupa_defunciones(d):
    """19 grupos de INEGI -> 17 grupos de la tabla (0, 1-4, 5-9, ..., 70-74, 75+)."""
    return np.r_[d[:16], d[16:].sum()]


def agrupa_poblacion(p):
    """16 grupos de la EIC -> 17 grupos (se separa el año 0 del 1-4 con el supuesto del paso 4)."""
    r = np.log(p[1] / p[0]) / 5
    f0 = (np.exp(r) - 1) / (np.exp(5 * r) - 1)
    return np.r_[p[0] * f0, p[0] * (1 - f0), p[1:]], f0


def insumos(s, datos=DATOS):
    """Población (17 grupos) y defunciones (17 grupos) de un sexo ('H', 'M' o 'A')."""
    pob, dfn = cargar(datos)
    dh, dm = defunciones_ajustadas(dfn)
    ph, f0h = agrupa_poblacion(pob.hombres.astype(float).values)
    pm, f0m = agrupa_poblacion(pob.mujeres.astype(float).values)
    if s == "H":
        return ph, agrupa_defunciones(dh)
    if s == "M":
        return pm, agrupa_defunciones(dm)
    return ph + pm, agrupa_defunciones(dh) + agrupa_defunciones(dm)


def tabla(s, datos=DATOS, radix=RADIX):
    p, d = insumos(s, datos)
    M = d / p
    n = N.copy()
    a = np.r_[0.12, 1.5, np.full(14, 2.5), 1 / M[16]]
    q = np.ones(17)
    q[:16] = 2 * n[:16] * M[:16] / (2 + n[:16] * M[:16])
    l = np.zeros(18)
    l[0] = radix
    for i in range(17):
        l[i + 1] = l[i] * (1 - q[i])
    dx = l[:17] * q
    L = np.zeros(17)
    L[:16] = n[:16] * l[1:17] + dx[:16] * a[:16]
    L[16] = dx[16] * a[16]
    T = L[::-1].cumsum()[::-1]
    return pd.DataFrame({"edad": ETIQ, "x": EDADES, "n": n, "P": p, "D": d, "nMx": M, "nqx": q, "lx": l[:17], "ndx": dx, "nAx": a,
                         "nLx": L, "Tx": T, "ex": T / l[:17]})


if __name__ == "__main__":
    pd.set_option("display.width", 200, "display.float_format", "{:,.5f}".format)
    for s in "HMA":
        t = tabla(s)
        print(SEXOS[s]); print(t.to_string(index=False)); print()
