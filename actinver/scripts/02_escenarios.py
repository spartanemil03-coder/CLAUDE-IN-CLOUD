"""Reto Actinver 2026: escenarios, optimización de portafolios y ganador.

Modelo (todo descriptivo del pasado; NO hay alfa por emisora, no se predice qué acción sube más):
  - "Mercado" = canasta de las 144 emisoras con historial, mismo peso. Beta de cada emisora vs esa canasta
    (sale de la hoja Covarianza). Mov. 1σ de la canasta en 25 sesiones = sigma_m.
  - Tres escenarios para la canasta en las sesiones que faltan: bajista -sigma_m, base +1%, alcista +sigma_m,
    con probabilidad 25% / 50% / 25%.
  - Ganancia en cada escenario = capital * (beta del portafolio * rendimiento de la canasta - costo de operar).
  - Riesgo: simulación con la covarianza completa del último año y media = beta * rendimiento medio de la canasta.
  - Costo de operar = spread compra-venta completo + 2 comisiones (0.116%): supuesto conservador (ida y vuelta).
Optimización: máximo (ganancia esperada neta / riesgo) con tope por emisora, hasta 12 emisoras (las filas de la hoja
Mi portafolio), solo emisoras operables con al menos $40,000 en el mejor nivel del libro.
"""
import importlib.util
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.optimize import minimize

warnings.filterwarnings("ignore")
AQUI = Path(__file__).resolve().parent
_spec = importlib.util.spec_from_file_location("analisis", AQUI / "01_analisis.py")
base = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(base)

CAPITAL, COMISION, SESIONES = base.CAPITAL, base.COMISION, base.SESIONES
SALIDA = base.SALIDA
PROB = {"Bajista": 0.25, "Base": 0.50, "Alcista": 0.25}
R_BASE = 0.01
LIQ_MIN = 40_000
MAX_EMISORAS = 12
N_SIM = 40_000
SEMILLA = 2026


def modelo():
    df, cov = base.cargar()
    d = df.set_index("emisora").loc[cov.index]
    n = len(cov)
    S = cov.values * SESIONES / 252           # covarianza de rendimientos en el lapso
    wm = np.ones(n) / n
    var_m = wm @ S @ wm
    sigma_m = float(np.sqrt(var_m))
    beta = S @ wm / var_m
    costo = (d.spread.fillna(9) + 2 * COMISION).values
    liq = (np.minimum(d.vol_compra, d.vol_venta) * d.ultimo).values
    permitido = (d.operable == "Sí").values & (liq >= LIQ_MIN)
    niveles = {"Bajista": -sigma_m, "Base": R_BASE, "Alcista": sigma_m}
    return dict(df=df, d=d, cov=cov, S=S, beta=beta, costo=costo, permitido=permitido,
                sigma_m=sigma_m, niveles=niveles, nombres=list(cov.index))


def optimiza(m, objetivo, tope, extra=None):
    S, beta, costo = m["S"], m["beta"], m["costo"]
    r_medio = sum(PROB[k] * m["niveles"][k] for k in PROB)
    mu = beta * r_medio - costo
    sig = lambda w, i: float(np.sqrt(w @ S[np.ix_(i, i)] @ w))

    def resolver(i, w0):
        cons = [{"type": "eq", "fun": lambda w: w.sum() - 1}]
        if extra:
            cons.append({"type": "ineq", "fun": lambda w: extra - sig(w, i)})
        return minimize(lambda w: objetivo(w @ mu[i], sig(w, i)), w0, bounds=[(0, tope)] * len(i),
                        constraints=cons, method="SLSQP", options={"maxiter": 500, "ftol": 1e-12}).x

    i = np.where(m["permitido"])[0]
    w = resolver(i, np.ones(len(i)) / len(i))
    while True:
        k = w >= 1e-4
        i, w = i[k], w[k]
        if len(i) <= MAX_EMISORAS and w.min() >= 0.04 - 1e-9:
            break
        j = int(np.argmin(w))
        i, w = np.delete(i, j), np.delete(w, j)
        w = resolver(i, w / w.sum())
    w = np.round(w, 3)
    w[np.argmax(w)] += round(1 - w.sum(), 3)
    return {m["nombres"][j]: float(x) for j, x in sorted(zip(i, w), key=lambda t: -t[1])}


def simula(m, pesos, rng):
    """Ganancia neta en pesos con la covarianza completa del último año y media = beta * rend. medio de la canasta."""
    nom = m["nombres"]
    idx = [nom.index(e) for e in pesos]
    w = np.array(list(pesos.values()))
    costo = float(w @ m["costo"][idx])
    r_medio = sum(PROB[k] * m["niveles"][k] for k in PROB)
    chol = np.linalg.cholesky(m["S"][np.ix_(idx, idx)] + 1e-12 * np.eye(len(idx)))
    r = m["beta"][idx] * r_medio + rng.standard_normal((N_SIM, len(idx))) @ chol.T
    return CAPITAL * (r @ w - costo), costo


def resumen(m, nombre, pesos, rng):
    nom = m["nombres"]
    idx = [nom.index(e) for e in pesos]
    w = np.array(list(pesos.values()))
    assert len(pesos) >= 5 and max(pesos.values()) <= 0.5 and abs(w.sum() - 1) < 1e-6, nombre
    assert len(pesos) <= MAX_EMISORAS or nombre.startswith(("A", "B", "C")), nombre
    mezcla, costo = simula(m, pesos, rng)
    beta_p = float(w @ m["beta"][idx])
    sigma = float(np.sqrt(w @ m["S"][np.ix_(idx, idx)] @ w))
    fila = {"portafolio": nombre, "n_emisoras": len(pesos), "peso_max": max(pesos.values()),
            "beta_vs_canasta": beta_p, "vol_anual": sigma * np.sqrt(252 / SESIONES), "costo_mxn": costo * CAPITAL}
    for esc in PROB:
        fila[f"gan_{esc}"] = CAPITAL * (beta_p * m["niveles"][esc] - costo)
    fila["gan_esperada"] = sum(PROB[e] * fila[f"gan_{e}"] for e in PROB)
    fila["P5"], fila["P95"] = np.percentile(mezcla, [5, 95])
    fila["desv_ganancia"] = float(mezcla.std())
    fila["P_ganar"] = float((mezcla > 0).mean())
    fila["P_ganar_50k"] = float((mezcla > 50_000).mean())
    fila["P_perder_50k"] = float((mezcla < -50_000).mean())
    fila["ganancia_por_riesgo"] = fila["gan_esperada"] / fila["desv_ganancia"]
    d = m["d"]
    fila["estres_peor_ventana_mxn"] = CAPITAL * sum(p * d.loc[e, "peor_vent"] for e, p in pesos.items())
    fila["estres_mejor_ventana_mxn"] = CAPITAL * sum(p * d.loc[e, "mejor_vent"] for e, p in pesos.items())
    return fila


def main():
    m = modelo()
    rng = np.random.default_rng(SEMILLA)
    print(f"sigma_m (canasta, {SESIONES} ses.) = {m['sigma_m']:.4f}; escenarios: "
          + ", ".join(f"{k} {v:+.2%}" for k, v in m["niveles"].items()))
    print(f"universo permitido: {int(m['permitido'].sum())} emisoras")
    cartera = {n: dict(p) for n, p in base.PORTAFOLIOS.items()}
    cartera = {"A Prudente": cartera["A Prudente"], "B Tu idea + diversificadores": cartera["B Tu idea + diversificadores"],
               "C Ranking (alto riesgo)": cartera["C Ranking (alto riesgo)"]}
    cartera["D Optimizado: ganancia/riesgo"] = optimiza(m, lambda mu, s: -mu / s, tope=0.15)
    cartera["E Optimizado: ganancia máx."] = optimiza(
        m, lambda mu, s: -mu, tope=0.20, extra=0.30 * np.sqrt(SESIONES / 252))
    filas = [resumen(m, n, p, rng) for n, p in cartera.items()]
    res = pd.DataFrame(filas)
    SALIDA.mkdir(exist_ok=True)
    res.round(4).to_csv(SALIDA / "pronostico_escenarios.csv", index=False)

    pesos = pd.DataFrame([{"portafolio": n, "emisora": e, "peso": w} for n, p in cartera.items() for e, w in p.items()])
    pesos.to_csv(SALIDA / "pesos_portafolios.csv", index=False)

    d = m["d"]
    ordenes = []
    for n, p in cartera.items():
        for e, w in p.items():
            titulos = int(w * CAPITAL * 0.997 // d.loc[e, "p_venta"])
            ordenes.append({"portafolio": n, "emisora": e, "peso": w, "precio_venta": d.loc[e, "p_venta"],
                            "titulos": titulos, "importe": round(titulos * d.loc[e, "p_venta"], 2)})
    pd.DataFrame(ordenes).to_csv(SALIDA / "ordenes_sugeridas.csv", index=False)

    pd.set_option("display.width", 250, "display.max_columns", 40)
    print(res[["portafolio", "n_emisoras", "beta_vs_canasta", "vol_anual", "costo_mxn", "gan_Bajista", "gan_Base",
               "gan_Alcista", "gan_esperada", "P5", "P95", "P_ganar", "ganancia_por_riesgo"]].round(3).to_string(index=False))
    for n, p in cartera.items():
        print(n, {e: round(w, 3) for e, w in p.items()})
    for col in ["gan_Bajista", "gan_Base", "gan_Alcista", "gan_esperada", "ganancia_por_riesgo"]:
        print("Ganador en", col, "->", res.loc[res[col].idxmax(), "portafolio"])


if __name__ == "__main__":
    main()
