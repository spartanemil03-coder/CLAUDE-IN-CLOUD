"""Cálculo de referencia (numpy/pandas) de la tabla de mortalidad de México 2019 con datos de INEGI.

Se usa para (a) redactar el análisis con cifras exactas y (b) verificar de forma independiente el libro de Excel.
Las reglas son las mismas que las fórmulas del libro:
  1. Las defunciones de edad no especificada se redistribuyen proporcionalmente entre las edades.
  2. (Opcional) media móvil de 5 edades en defunciones y población de 10 a 97 años (corrige la atracción de edades redondas).
  3. q0 = defunciones menores de 1 año / nacimientos;  qx = D / (P + D/2) para x = 1..99;  q100 = 1.
  4. lx(0) = radix; l(x+1) = lx - dx; dx = lx*qx; Lx = l(x+1) + dx/2 (L100 = d100/2); Tx = suma de Lx; ex = Tx/lx.
"""
from pathlib import Path

import numpy as np
import pandas as pd

SEXOS = {"H": "Hombres", "M": "Mujeres", "A": "Ambos sexos"}
EDADES = np.arange(101)


def cargar(datos):
    datos = Path(datos)
    dfn = pd.read_csv(datos / "inegi_defunciones_2019_por_edad_sexo.csv", index_col=0)
    nac = pd.read_csv(datos / "inegi_nacimientos_2019_por_sexo.csv").iloc[0]
    pob = pd.read_csv(datos / "inegi_censo2020_poblacion_por_edad_sexo.csv", index_col=0)
    return dfn, nac, pob


def _serie(dfn, nac, pob, s):
    if s == "H":
        d, p, b = dfn["hombres"], pob["hombres"], nac["hombres"]
    elif s == "M":
        d, p, b = dfn["mujeres"], pob["mujeres"], nac["mujeres"]
    else:
        d = dfn["hombres"] + dfn["mujeres"] + dfn["sexo_no_especificado"]
        p, b = pob["hombres"] + pob["mujeres"], nac["hombres"] + nac["mujeres"] + nac["sexo_no_especificado"]
    ne = d["NE"]
    d = d.iloc[:101].astype(float).values
    d_aj = d * (d.sum() + ne) / d.sum()
    return d_aj, p.iloc[:101].astype(float).values, float(b), d, float(ne)


def media_movil(v, ini=10, fin=97):
    o = v.copy()
    for x in range(ini, fin + 1):
        o[x] = v[x - 2:x + 3].mean()
    return o


def tabla_vida(q, radix=100_000.0):
    l = np.zeros(101)
    l[0] = radix
    for x in range(100):
        l[x + 1] = l[x] * (1 - q[x])
    d = l * q
    L = np.append(l[1:], 0.0) + d / 2
    T = L[::-1].cumsum()[::-1]
    return pd.DataFrame({"x": EDADES, "qx": q, "lx": l, "dx": d, "Lx": L, "Tx": T, "ex": T / l})


def tablas(datos, suaviza=True, radix=100_000.0):
    dfn, nac, pob = cargar(datos)
    out = {}
    for s in SEXOS:
        d_aj, p, b, d_raw, ne = _serie(dfn, nac, pob, s)
        d_us, p_us = (media_movil(d_aj), media_movil(p)) if suaviza else (d_aj, p)
        q = d_us / (p_us + d_us / 2)
        q[0] = d_us[0] / b
        q[100] = 1.0
        t = tabla_vida(q, radix)
        t.attrs.update(nacimientos=b, defunciones_inf=d_us[0], d_us=d_us, p_us=p_us, d_raw=d_raw, d_ne=ne)
        out[s] = t
    return out


def indicadores(t, radix=100_000.0):
    """Indicadores clave de una tabla de vida."""
    l, e, d = t.lx.values, t.ex.values, t.dx.values
    return {
        "e0": e[0], "e1": e[1], "e15": e[15], "e65": e[65], "e80": e[80],
        "q0_mil": t.qx[0] * 1000, "5q0_mil": (1 - l[5] / l[0]) * 1000,
        "l65": l[65] / radix, "l80": l[80] / radix,
        "mediana_muerte": int(t.x[np.where(l >= radix / 2)[0][-1]]),
        "moda_adulta": int(15 + np.argmax(d[15:100])),  # sin el grupo abierto 100+
    }
