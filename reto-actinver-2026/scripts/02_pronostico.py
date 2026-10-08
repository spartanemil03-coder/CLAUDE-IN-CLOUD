"""Reto Actinver 2026: pronóstico por emisora, portafolios de 5 a 12 emisoras y ganador.

Datos: precios de Yahoo Finance (00_actualizar_datos.py) + precios de compra/venta y spread del monitor (Excel original).

Pronóstico a 25 sesiones de cada emisora (método estándar "alfa = IC × volatilidad × puntaje"):
  mu_i = beta_i × E[R canasta] + alfa_i − costo_i
  alfa_i = IC × sigma_i(25 ses.) × z_i, con z_i = promedio de dos señales estandarizadas entre las 144:
     - potencial al precio objetivo medio de analistas (solo con 5 o más analistas)
     - momento 12-1 (rendimiento de 12 meses sin el último mes)
  IC = 0.05 (capacidad de pronóstico típica de una buena señal; es un supuesto).
Riesgo: covarianza de rendimientos diarios en pesos del último año; colas gruesas con t de Student (4 g.l.).
Objetivo "ganar": máxima probabilidad de superar el umbral del top 3 (+20%), solo con emisoras de pronóstico positivo.
Universo: operables, spread <= 1.5% y líquidas (mexicanas: >= $20M operados al día; SIC: >= $40k en el mejor nivel del monitor).
"""
import importlib.util
import warnings
from pathlib import Path

import numpy as np
import pandas as pd
from scipy.optimize import minimize
from scipy.stats import norm

warnings.filterwarnings("ignore")
AQUI = Path(__file__).resolve().parent
_spec = importlib.util.spec_from_file_location("analisis", AQUI / "01_analisis.py")
base = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(base)

RAIZ = AQUI.parent
YAHOO = RAIZ / "datos" / "yahoo"
SALIDA = RAIZ / "resultados"
CAPITAL, COMISION = base.CAPITAL, base.COMISION
SESIONES = 25
FECHA_FIN = "2026-11-13"
IC = 0.05
PRIMA_ANUAL = 0.08                       # rendimiento esperado anual de la canasta
UMBRAL = 0.20                            # rendimiento que sacaron los 3 primeros en una edición pasada (>20%)
PROB = {"Bajista": 0.25, "Base": 0.50, "Alcista": 0.25}
TOPE, PISO = 0.50, 0.03
SPREAD_MAX = 0.015
LIQ_MX = 20_000_000                      # mexicanas: importe operado diario promedio (20 días, Yahoo) en pesos
LIQ_SIC = 40_000                         # SIC: importe en el mejor nivel de compra y de venta del monitor
N_SIM, GL, SEMILLA = 60_000, 4, 2026
LISTA_N = range(5, 13)


def datos():
    df, _ = base.cargar()
    info = pd.read_csv(YAHOO / "info_emisoras.csv")
    px = pd.read_csv(YAHOO / "precios_cierre.csv", index_col=0, parse_dates=True).sort_index()
    px = px.ffill()
    fx = px["USDMXN=X"]
    df = df.merge(info[["emisora", "ticker", "moneda", "precio_actual_mxn", "objetivo_medio_mxn", "n_analistas",
                        "recomendacion", "prox_reporte"]], on="emisora", how="left")
    df = df[df.precio_actual_mxn.notna()].reset_index(drop=True)
    mxn = pd.DataFrame({r.emisora: px[r.ticker] * (fx if r.moneda == "USD" else 1.0) for r in df.itertuples()})
    vol = pd.read_csv(YAHOO / "volumen.csv", index_col=0, parse_dates=True).sort_index()
    crudo = pd.read_csv(YAHOO / "precios_cierre.csv", index_col=0, parse_dates=True).sort_index()
    df["importe_diario_mxn"] = [float((vol[t] * crudo[t]).iloc[-20:].mean()) * (1 if m != "USD" else fx.iloc[-1])
                                for t, m in zip(df.ticker, df.moneda)]
    df["liq_monitor_mxn"] = np.minimum(df.vol_compra, df.vol_venta) * df.ultimo
    df["liquida"] = np.where(df.tipo == "Mexicana", df.importe_diario_mxn >= LIQ_MX, df.liq_monitor_mxn >= LIQ_SIC)
    return df, mxn, px["^MXX"]


def estadisticas(df, mxn, ipc):
    ret = mxn.pct_change()
    r1 = ret.iloc[-252:].fillna(0)
    out = df.copy()
    last = mxn.iloc[-1]
    def rend(d):
        return (last / mxn.iloc[-1 - d] - 1).values
    out["ultimo"] = out.precio_actual_mxn
    out["var_dia"] = ret.iloc[-1].values
    out["r1m"], out["r3m"], out["r6m"], out["r12m"] = rend(21), rend(63), rend(126), rend(252)
    out["mom_12_1"] = (mxn.iloc[-22] / mxn.iloc[-253] - 1).values
    out["vol_anual"] = (r1.std() * np.sqrt(252)).values
    out["mov_1s"] = out.vol_anual * np.sqrt(SESIONES / 252)
    p1 = mxn.iloc[-252:]
    out["peor_caida"] = (p1 / p1.cummax() - 1).min().values
    ri = ipc.pct_change().iloc[-252:].fillna(0)
    out["beta"] = (r1.apply(lambda c: np.cov(c, ri)[0, 1]) / ri.var()).values
    v25 = mxn / mxn.shift(SESIONES) - 1
    out["mejor_vent"], out["peor_vent"] = v25.max().values, v25.min().values
    cov = r1.cov() * 252
    w = np.ones(len(cov)) / len(cov)
    S = cov.values * SESIONES / 252
    out["beta_canasta"] = S @ w / (w @ S @ w)
    out["costo"] = out.spread.fillna(1.0) + 2 * COMISION
    terc = out.vol_anual.rank(pct=True)
    out["perfil"] = np.where(terc <= 1 / 3, "Menor movimiento", np.where(terc <= 2 / 3, "Movimiento medio", "Mayor movimiento"))
    return out, cov, float(np.sqrt(w @ S @ w))


def zscore(x):
    x = x.clip(x.quantile(0.05), x.quantile(0.95))
    return (x - x.mean()) / x.std()


def pronostico(est, sigma_m):
    e = est.copy()
    e["potencial"] = np.where(e.n_analistas >= 5, e.objetivo_medio_mxn / e.ultimo - 1, np.nan)
    z_pot, z_mom = zscore(e.potencial), zscore(e.mom_12_1)
    z = pd.concat([z_pot, z_mom], axis=1).mean(axis=1)
    e["puntaje"] = (z - z.mean()) / z.std()
    e["alfa"] = IC * e.mov_1s * e.puntaje
    e["r_mercado"] = PRIMA_ANUAL * SESIONES / 252
    e["mu"] = e.beta_canasta * e.r_mercado + e.alfa - e.costo
    r_m = float(e.r_mercado.iloc[0])
    niveles = {"Bajista": r_m - sigma_m, "Base": r_m, "Alcista": r_m + sigma_m}   # promedio ponderado = r_m
    for k, v in niveles.items():
        e[f"esc_{k}"] = e.beta_canasta * v + e.alfa - e.costo
    e["reporta_en_reto"] = e.prox_reporte.fillna("").between("2026-10-08", FECHA_FIN)
    return e, niveles


def optimiza(mu, S, idx, n, objetivo):
    def resolver(i, w0, piso):
        def f(w):
            m, s = w @ mu[i], np.sqrt(w @ S[np.ix_(i, i)] @ w)
            return -(m - UMBRAL) / s if objetivo == "ganar" else -m / s
        cons = [{"type": "eq", "fun": lambda w: w.sum() - 1},
                {"type": "ineq", "fun": lambda w: w @ mu[i]}]          # ganancia esperada >= 0
        return minimize(f, w0, bounds=[(piso, TOPE)] * len(i), constraints=cons, method="SLSQP",
                        options={"maxiter": 600, "ftol": 1e-12}).x
    i = np.array(idx)
    w = resolver(i, np.ones(len(i)) / len(i), 0.0)
    while len(i) > n:
        quita = max(1, (len(i) - n) // 3)
        orden = np.argsort(w)[quita:]
        i, w = i[orden], w[orden]
        w = resolver(i, w / w.sum(), 0.0)
    w = resolver(i, np.clip(w, PISO, TOPE) / np.clip(w, PISO, TOPE).sum(), PISO)
    w = np.round(w, 3)
    w[np.argmax(w)] += round(1 - w.sum(), 3)
    return i, w


def simula(mu, S, rng):
    chol = np.linalg.cholesky(S + 1e-10 * np.eye(len(S)))
    z = rng.standard_normal((N_SIM, len(S))) @ chol.T
    chi = rng.chisquare(GL, (N_SIM, 1))
    return mu + z * np.sqrt((GL - 2) / chi)


def evalua(nombre, e, i, w, sims, niveles):
    g = CAPITAL * (sims[:, i] @ w)
    sub = e.iloc[i]
    fila = {"portafolio": nombre, "n_emisoras": len(i), "emisoras": ", ".join(f"{a} {b:.0%}" for a, b in zip(sub.emisora, w)),
            "peso_max": w.max(), "gan_esperada": float(w @ sub.mu) * CAPITAL, "desv": g.std(),
            "P_top3_20pct": (g >= UMBRAL * CAPITAL).mean(), "P_30pct": (g >= 0.30 * CAPITAL).mean(), "P_ganar": (g > 0).mean(),
            "P_perder_100k": (g < -100_000).mean(), "P5": np.percentile(g, 5), "P50": np.percentile(g, 50),
            "P95": np.percentile(g, 95), "costo_mxn": float(w @ sub.costo) * CAPITAL,
            "beta_canasta": float(w @ sub.beta_canasta),
            "gan_esperada_sin_alfa": float(w @ (sub.beta_canasta * sub.r_mercado - sub.costo)) * CAPITAL,
            "P_top3_sin_alfa": float((g - CAPITAL * float(w @ sub.alfa) >= UMBRAL * CAPITAL).mean())}
    for k in niveles:
        fila[f"gan_{k}"] = float(w @ sub[f"esc_{k}"]) * CAPITAL
    return fila


def main():
    df, mxn, ipc = datos()
    est, cov, sigma_m = estadisticas(df, mxn, ipc)
    e, niveles = pronostico(est, sigma_m)
    S = cov.values * SESIONES / 252
    mu = e.mu.values
    base_ok = (e.operable == "Sí") & (e.spread <= SPREAD_MAX) & e.liquida
    idx_univ = {"ganar": list(np.where(base_ok & (e.mu > 0))[0]), "riesgo": list(np.where(base_ok)[0])}
    idx = idx_univ["riesgo"]
    rng = np.random.default_rng(SEMILLA)
    sims = simula(mu, S, rng)

    filas, pesos = [], []
    for objetivo, etiqueta in [("ganar", "Ganar (máx. prob. top 3)"), ("riesgo", "Balanceado (ganancia/riesgo)")]:
        for n in LISTA_N:
            i, w = optimiza(mu, S, idx_univ[objetivo], n, objetivo)
            nombre = f"{etiqueta} · {n} emisoras"
            filas.append(evalua(nombre, e, i, w, sims, niveles) | {"objetivo": etiqueta, "n": n})
            pesos += [{"portafolio": nombre, "objetivo": etiqueta, "n": n, "emisora": e.emisora.iloc[j], "peso": x}
                      for j, x in zip(i, w)]
    # Referencia: tu idea original (B)
    ref = {"NVDA *": .30, "GFNORTE O": .20, "GMEXICO B": .20, "MSFT *": .15, "GOOGL *": .15}
    ii = [int(e.index[e.emisora == k][0]) for k in ref]
    filas.append(evalua("Tu idea (NVDA + Banorte)", e, np.array(ii), np.array(list(ref.values())), sims, niveles)
                 | {"objetivo": "Referencia", "n": 5})
    pesos += [{"portafolio": "Tu idea (NVDA + Banorte)", "objetivo": "Referencia", "n": 5, "emisora": k, "peso": v}
              for k, v in ref.items()]

    res = pd.DataFrame(filas)
    gan = res[res.objetivo.str.startswith("Ganar")]
    ganador = gan.sort_values(["P_top3_20pct", "gan_esperada"], ascending=False).iloc[0].portafolio
    res["ganador"] = res.portafolio == ganador

    SALIDA.mkdir(exist_ok=True)
    cols = ["emisora", "tipo", "ticker", "ultimo", "var_dia", "p_compra", "p_venta", "spread", "operable", "r1m", "r3m", "r6m",
            "r12m", "mom_12_1", "vol_anual", "mov_1s", "peor_caida", "beta", "beta_canasta", "mejor_vent", "peor_vent",
            "perfil", "importe_diario_mxn", "liq_monitor_mxn", "liquida", "objetivo_medio_mxn", "potencial", "n_analistas", "recomendacion", "prox_reporte",
            "reporta_en_reto", "puntaje", "alfa", "costo", "mu", "esc_Bajista", "esc_Base", "esc_Alcista"]
    e[cols].to_csv(SALIDA / "pronostico_emisoras.csv", index=False)
    cov.to_csv(YAHOO / "covarianza_1a.csv")
    res.to_csv(SALIDA / "portafolios_por_n.csv", index=False)
    pw = pd.DataFrame(pesos)
    pw.to_csv(SALIDA / "pesos_portafolios.csv", index=False)

    d = e.set_index("emisora")
    ordenes = []
    for r in pw.itertuples():
        precio = d.loc[r.emisora, "p_venta"]
        titulos = int(r.peso * CAPITAL * 0.997 // precio)
        ordenes.append({"portafolio": r.portafolio, "emisora": r.emisora, "peso": r.peso, "precio_venta_monitor": precio,
                        "titulos": titulos, "importe": round(titulos * precio, 2),
                        "comision_iva": round(titulos * precio * COMISION, 2),
                        "prox_reporte": d.loc[r.emisora, "prox_reporte"]})
    pd.DataFrame(ordenes).to_csv(SALIDA / "ordenes_sugeridas.csv", index=False)

    pd.set_option("display.width", 260, "display.max_columns", 30, "display.max_colwidth", 90)
    print(f"sigma canasta 25 ses. {sigma_m:.3%}; universo {len(idx)} emisoras ({len(idx_univ['ganar'])} con pronóstico positivo); escenarios {niveles}")
    print(e.sort_values("mu", ascending=False)[["emisora", "mu", "alfa", "puntaje", "potencial", "mom_12_1", "vol_anual",
                                                 "beta_canasta", "costo", "prox_reporte"]].head(20).round(3).to_string(index=False))
    print(res[["portafolio", "gan_esperada", "desv", "P_top3_20pct", "P_30pct", "P_ganar", "P_perder_100k", "P5", "P95",
               "gan_Bajista", "gan_Base", "gan_Alcista"]].round(3).to_string(index=False))
    print("GANADOR:", ganador)
    print(res.loc[res.ganador, "emisoras"].iloc[0])


if __name__ == "__main__":
    main()
