"""Genera el documento Word 'Explicación paso a paso' con los números reales del cálculo (llama a scripts/06_explicacion.js).

Uso:  python3 scripts/06_explicacion.py
"""
import json
import subprocess
import sys
import tempfile
from pathlib import Path

import numpy as np
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))
import calculo_inegi as calc  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
dfn, nac, pob = calc.cargar()
raw = dfn["hombres"] + dfn["mujeres"] + dfn["sexo_no_especificado"]
ne = float(raw["NE"])
con_edad = float(raw.iloc[:101].sum())
d, p, b = calc.serie("A")
ds, ps = calc.media_movil(d), calc.media_movil(p)
q_s, q_c = calc.qx("A", True)[0], calc.qx("A", False)[0]
t = calc.tabla(q_s)
ref = pd.read_csv(RAIZ / "datos" / "conapo_e0_2019_referencia.csv").set_index("sexo").e0_2019
d_arch = pd.read_csv(RAIZ / "datos" / "inegi_defunciones_2019_por_archivo.csv")
n_arch = pd.read_csv(RAIZ / "datos" / "inegi_nacimientos_2019_por_archivo.csv")
datos = {
    "e0": t.ex[0], "e65": t.ex[65], "e0_crudo": calc.tabla(q_c).ex[0], "q0": q_s[0], "q1": q_s[1], "q65": q_s[65],
    "l1": t.lx[1], "l65": t.lx[65], "l80": t.lx[80], "dx0": t.dx[0], "Lx0": t.Lx[0], "Tx0": t.Tx[0],
    "mediana": int(t.x[np.where(t.lx.values >= 5e4)[0][-1]]),
    "nacimientos": b, "nac_arch": [int(x) for x in (n_arch.hombres + n_arch.mujeres + n_arch.sexo_no_especificado)],
    "def_total": int(d_arch.defunciones_ocurridas_2019.sum()), "def_arch": [int(x) for x in d_arch.defunciones_ocurridas_2019],
    "pob_total": int(pob.values.sum()), "ne": ne, "factor": (con_edad + ne) / con_edad,
    "d0_raw": float(raw["0"]), "d0_aj": d[0], "d1_aj": d[1], "p1": p[1], "d65_us": ds[65], "p65_us": ps[65],
    "p30": p[30], "p31": p[31], "q30_crudo": q_c[30], "q31_crudo": q_c[31], "q30_suav": q_s[30], "q31_suav": q_s[31],
    "ref_a": float(ref["Ambos sexos"]),
}
salida = RAIZ / "resultados" / "Explicacion_paso_a_paso_INEGI_2019.docx"
with tempfile.TemporaryDirectory() as tmp:
    ruta = Path(tmp) / "datos.json"
    ruta.write_text(json.dumps(datos, ensure_ascii=False))
    subprocess.run(["node", str(RAIZ / "scripts" / "06_explicacion.js"), str(ruta), str(salida)], check=True)
print("explicación", salida)
