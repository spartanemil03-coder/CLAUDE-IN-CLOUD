"""Cálculo de referencia (numpy/pandas) de la tabla de mortalidad de México 2019 con datos de INEGI.

Se usa para verificar de forma independiente el libro de Excel y para redactar el análisis. Reglas (las mismas del libro):
  1. Las defunciones de edad no especificada se redistribuyen proporcionalmente entre las edades.
  2. (Opcional) media móvil de 5 edades en defunciones y población de 10 a 97 años (corrige la atracción de edades redondas del Censo).
  3. q0 = defunciones menores de 1 año / nacimientos;  qx = D / (P + D/2) para x = 1..99;  q100 = 1.
  4. l0 = 100,000; dx = ROUND(lx*qx); l(x+1) = lx - dx; Lx = ROUND(l(x+1) + dx/2); Tx = suma de Lx; ex = Tx/lx.
"""
from pathlib import Path

import numpy as np
import pandas as pd

RAIZ = Path(__file__).resolve().parent.parent
DATOS = RAIZ / "datos"
SEXOS = {"H": "Hombres", "M": "Mujeres", "A": "Ambos sexos"}
RADIX = 100_000


def redondea(x):
    """ROUND de Excel (mitad hacia arriba para positivos)."""
    return np.floor(np.asarray(x, float) + 0.5)


def cargar(datos=DATOS):
    dfn = pd.read_csv(Path(datos) / "inegi_defunciones_2019_por_edad_sexo.csv", index_col=0)
    nac = pd.read_csv(Path(datos) / "inegi_nacimientos_2019_por_sexo.csv").iloc[0]
    pob = pd.read_csv(Path(datos) / "inegi_censo2020_poblacion_por_edad_sexo.csv", index_col=0)
    return dfn, nac, pob


def serie(s, datos=DATOS):
    """Defunciones (con edad NE redistribuida), población y nacimientos de un sexo."""
    dfn, nac, pob = cargar(datos)
    if s == "H":
        d, p, b = dfn["hombres"], pob["hombres"], nac["hombres"]
    elif s == "M":
        d, p, b = dfn["mujeres"], pob["mujeres"], nac["mujeres"]
    else:
        d = dfn["hombres"] + dfn["mujeres"] + dfn["sexo_no_especificado"]
        p, b = pob["hombres"] + pob["mujeres"], nac["hombres"] + nac["mujeres"] + nac["sexo_no_especificado"]
    ne = d["NE"]
    d = d.iloc[:101].astype(float).values
    return d * (d.sum() + ne) / d.sum(), p.iloc[:101].astype(float).values, float(b)


def media_movil(v, ini=10, fin=97):
    o = v.copy()
    for x in range(ini, fin + 1):
        o[x] = v[x - 2:x + 3].mean()
    return o


def qx(s="A", suaviza=True, datos=DATOS):
    d, p, b = serie(s, datos)
    d, p = (media_movil(d), media_movil(p)) if suaviza else (d, p)
    q = d / (p + d / 2)
    q[0] = d[0] / b
    q[100] = 1.0
    return q, d, p, b


def tabla(q, radix=RADIX):
    l, d, L = np.zeros(101), np.zeros(101), np.zeros(101)
    l[0] = radix
    for x in range(101):
        d[x] = redondea(l[x] * q[x])
        if x < 100:
            l[x + 1] = l[x] - d[x]
    for x in range(101):
        L[x] = redondea((l[x + 1] if x < 100 else 0) + d[x] / 2)
    T = L[::-1].cumsum()[::-1]
    return pd.DataFrame({"x": np.arange(101), "qx": q, "lx": l, "dx": d, "Lx": L, "Tx": T, "ex": T / l})


def e0_continua(q, radix=1.0):
    """e0 sin redondeos (para comparar sexos)."""
    l = np.zeros(102)
    l[0] = radix
    for x in range(101):
        l[x + 1] = l[x] * (1 - q[x])
    return float((l[1:] + l[:101] * q / 2).sum() / radix)
