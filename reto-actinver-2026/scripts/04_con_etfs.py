"""Reto Actinver 2026: ¿cómo cambia el ganador si también se consideran los 40 ETFs permitidos?

Mismo modelo que 02_pronostico.py (pronóstico, riesgo, objetivo de superar +20%), con el universo ampliado:
  - 145 acciones del monitor + 40 ETFs no apalancados (datos/etfs_monitor_2026-10-08.csv, precios de Yahoo).
  - Los ETFs no tienen precio objetivo de analistas: su puntaje usa solo el momento 12-1.
  - Reglas: al menos 5 ACCIONES distintas (un ETF no cuenta), 50% máximo también para cada ETF.
Los 25 fondos de inversión no entran: la plataforma no muestra su historial y no hay precios públicos diarios.
"""
import importlib.util
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.optimize import minimize

warnings.filterwarnings("ignore")
AQUI = Path(__file__).resolve().parent
_spec = importlib.util.spec_from_file_location("pron", AQUI / "02_pronostico.py")
P = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(P)

RAIZ = AQUI.parent
CAP, COM, SES = P.CAPITAL, P.COMISION, P.SESIONES
GANADOR_ACTUAL = {"MU *": .50, "INTC *": .41, "PINFRA *": .03, "BOLSA A": .03, "GFNORTE O": .03}


def universo():
    df, mxn, ipc = P.datos()
    est, _, _ = P.estadisticas(df, mxn, ipc)
    est["clase"] = "Acción"

    etf = pd.read_csv(RAIZ / "datos" / "etfs_monitor_2026-10-08.csv")
    etf = etf[etf.apalancado == "No"].reset_index(drop=True)
    px = pd.read_csv(RAIZ / "datos" / "yahoo" / "etfs_precios_cierre.csv", index_col=0, parse_dates=True).sort_index()
    fx = pd.read_csv(RAIZ / "datos" / "yahoo" / "precios_cierre.csv", index_col=0, parse_dates=True)["USDMXN=X"]
    idx = mxn.index.union(px.index)
    fx = fx.reindex(idx).ffill()
    etf_mxn = pd.DataFrame({r.emisora: px[r.ticker].reindex(idx).ffill() * (1 if r.ticker.endswith(".MX") else fx)
                            for r in etf.itertuples()})
    todo = pd.concat([mxn.reindex(idx).ffill(), etf_mxn], axis=1)
    ret = todo.pct_change().iloc[-252:].fillna(0)
    cov = ret.cov() * 252

    mid = (etf.p_compra + etf.p_venta) / 2
    e2 = pd.DataFrame({
        "emisora": etf.emisora, "tipo": "ETF (SIC)", "clase": "ETF", "ultimo": etf.ultimo,
        "p_compra": etf.p_compra, "p_venta": etf.p_venta,
        "spread": np.where((etf.p_compra > 0) & (etf.p_venta > 0), (etf.p_venta - etf.p_compra).clip(lower=0) / mid, np.nan),
        "operable": np.where((etf.p_compra > 0) & (etf.p_venta > 0), "Sí", "Sin contraparte"),
        "liq_monitor_mxn": np.minimum(etf.vol_compra, etf.vol_venta) * etf.ultimo,
        "mom_12_1": (etf_mxn.iloc[-22] / etf_mxn.iloc[-253] - 1).values,
        "vol_anual": (ret[etf.emisora].std() * np.sqrt(252)).values,
        "objetivo_medio_mxn": np.nan, "n_analistas": np.nan, "prox_reporte": None,
    })
    e2["liquida"] = e2.liq_monitor_mxn >= P.LIQ_SIC
    e2["mov_1s"] = e2.vol_anual * np.sqrt(SES / 252)
    e2["costo"] = e2.spread.fillna(1.0) + 2 * COM

    e = pd.concat([est, e2], ignore_index=True)
    assert list(e.emisora) == list(cov.index)
    # Beta contra la misma canasta de antes: las 144 acciones con mismo peso.
    S = cov.values * SES / 252
    wm = (e.clase == "Acción").to_numpy(dtype=float, copy=True)
    wm /= wm.sum()
    e["beta_canasta"] = S @ wm / (wm @ S @ wm)
    sigma_m = float(np.sqrt(wm @ S @ wm))
    return e, S, sigma_m


def pronostico(e, sigma_m):
    e = e.copy()
    e["potencial"] = np.where(e.n_analistas >= 5, e.objetivo_medio_mxn / e.ultimo - 1, np.nan)
    z = pd.concat([P.zscore(e.potencial), P.zscore(e.mom_12_1)], axis=1).mean(axis=1)
    e["puntaje"] = (z - z.mean()) / z.std()
    e["alfa"] = P.IC * e.mov_1s * e.puntaje
    e["r_mercado"] = P.PRIMA_ANUAL * SES / 252
    r_m = float(e.r_mercado.iloc[0])
    niveles = {"Bajista": r_m - sigma_m, "Base": r_m, "Alcista": r_m + sigma_m}
    e["mu"] = e.beta_canasta * r_m + e.alfa - e.costo
    for k, v in niveles.items():
        e[f"esc_{k}"] = e.beta_canasta * v + e.alfa - e.costo
    return e, niveles


def resolver(mu, S, i, piso):
    def f(w):
        return -(w @ mu[i] - P.UMBRAL) / np.sqrt(w @ S[np.ix_(i, i)] @ w)
    cons = [{"type": "eq", "fun": lambda w: w.sum() - 1}, {"type": "ineq", "fun": lambda w: w @ mu[i]}]
    w = minimize(f, np.ones(len(i)) / len(i), bounds=[(piso, P.TOPE)] * len(i), constraints=cons, method="SLSQP",
                 options={"maxiter": 600, "ftol": 1e-12}).x
    w = np.round(w, 3)
    w[np.argmax(w)] += round(1 - w.sum(), 3)
    return w


def ganar_con_etfs(e, S, n):
    """Mejor portafolio para superar +20% con n emisoras en total, de las cuales al menos 5 son acciones."""
    mu = e.mu.values
    ok = (e.operable == "Sí") & (e.spread <= P.SPREAD_MAX) & e.liquida.astype(bool) & (e.mu > 0)
    i, _ = P.optimiza(mu, S, list(np.where(ok)[0]), n, "ganar")
    i = list(i)
    acciones = [j for j in i if e.clase.iloc[j] == "Acción"]
    if len(acciones) < 5:
        # Relleno: acciones líquidas con pronóstico positivo y la menor volatilidad, para cumplir el mínimo de 5.
        candidatas = e[ok & (e.clase == "Acción") & ~e.index.isin(i)].sort_values("vol_anual").index
        relleno = list(candidatas[: 5 - len(acciones)])
        grandes = sorted(i, key=lambda j: -mu[j])[: n - len(relleno)]
        i = grandes + relleno
    i = np.array(i)
    return i, resolver(mu, S, i, P.PISO)


def main():
    e, S, sigma_m = universo()
    e, niveles = pronostico(e, sigma_m)
    rng = np.random.default_rng(P.SEMILLA)
    sims = P.simula(e.mu.values, S, rng)

    filas, pesos = [], []
    ii = np.array([int(e.index[e.emisora == k][0]) for k in GANADOR_ACTUAL])
    filas.append(P.evalua("Ganador actual (solo acciones)", e, ii, np.array(list(GANADOR_ACTUAL.values())), sims, niveles)
                 | {"n": 5, "etfs": 0})
    for n in range(5, 9):
        i, w = ganar_con_etfs(e, S, n)
        nombre = f"Con ETFs · {n} emisoras"
        n_etf = int(sum(e.clase.iloc[j] == "ETF" for j in i))
        filas.append(P.evalua(nombre, e, i, w, sims, niveles) | {"n": n, "etfs": n_etf})
        pesos += [{"portafolio": nombre, "emisora": e.emisora.iloc[j], "clase": e.clase.iloc[j], "peso": x}
                  for j, x in zip(i, w)]
    res = pd.DataFrame(filas)

    out = RAIZ / "resultados"
    res.to_csv(out / "con_etfs_portafolios.csv", index=False)
    pd.DataFrame(pesos).to_csv(out / "con_etfs_pesos.csv", index=False)
    cols = ["emisora", "clase", "ultimo", "spread", "liq_monitor_mxn", "liquida", "vol_anual", "mom_12_1",
            "beta_canasta", "puntaje", "alfa", "costo", "mu"]
    e[e.clase == "ETF"][cols].sort_values("mu", ascending=False).to_csv(out / "con_etfs_pronostico_etfs.csv", index=False)

    pd.set_option("display.width", 250, "display.max_colwidth", 120)
    print(e[e.clase == "ETF"][cols].sort_values("mu", ascending=False).round(3).to_string(index=False))
    print(res[["portafolio", "etfs", "gan_esperada", "P_top3_20pct", "P_30pct", "P_ganar", "P_perder_100k", "P5", "P95",
               "emisoras"]].round(3).to_string(index=False))


if __name__ == "__main__":
    main()
