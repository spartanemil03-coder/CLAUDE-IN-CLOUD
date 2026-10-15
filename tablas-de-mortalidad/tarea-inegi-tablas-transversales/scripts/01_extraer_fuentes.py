"""Extrae de 'Fuentes de consulta.xlsx' (la que entrega la maestra) las dos tablas que se usan y las guarda como CSV en datos/.

  - Población por grupos quinquenales de edad y sexo (Encuesta Intercensal 2025, INEGI), total nacional.
  - Defunciones por grupos de edad y sexo (Mortalidad, INEGI). Vienen como texto con comas ("12,207"); las celdas vacías son cero.
"""
import csv
import sys
from pathlib import Path

import openpyxl

RAIZ = Path(__file__).resolve().parent.parent
ORIGEN = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "fuentes" / "Fuentes_de_consulta.xlsx"


def numero(v):
    if v in (None, ""):
        return 0
    return v if isinstance(v, (int, float)) else int(str(v).replace(",", ""))


ws = openpyxl.load_workbook(ORIGEN, data_only=True).active
pob, dfn = [], []
for fila in ws.iter_rows(values_only=True):
    a, b, c, d, e, f, g, h = (list(fila) + [None] * 8)[1:9]     # columnas B..I
    if a == "Estados Unidos Mexicanos":                           # grupo, total, mujeres, hombres (absolutos)
        pob.append([b, d, e, g])
    elif isinstance(a, str) and a and a not in ("Tamaño de localidad",) and (a in ("Total", "Menores de 1 año", "No especificado") or a[0].isdigit()) and isinstance(b, str):
        dfn.append([a, numero(b), numero(c), numero(d), numero(e)])  # total, hombre, mujer, sexo NE
assert len(pob) == 16 and len(dfn) == 21, (len(pob), len(dfn))

with open(RAIZ / "datos" / "inegi_eic2025_poblacion_por_edad_sexo.csv", "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    w.writerow(["grupo_edad", "total", "mujeres", "hombres"])
    w.writerows(pob)
with open(RAIZ / "datos" / "inegi_mortalidad_defunciones_por_edad_sexo.csv", "w", newline="", encoding="utf-8") as fh:
    w = csv.writer(fh)
    w.writerow(["grupo_edad", "total", "hombres", "mujeres", "sexo_no_especificado"])
    w.writerows(dfn)
print("población:", len(pob), "filas; defunciones:", len(dfn), "filas")
