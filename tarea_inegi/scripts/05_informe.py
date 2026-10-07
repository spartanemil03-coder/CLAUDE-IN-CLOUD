"""Extrae cifras y texto del libro recalculado y genera el documento Word (llama a scripts/05_informe.js).

Uso:  python3 scripts/05_informe.py
"""
import json
import subprocess
import tempfile
from pathlib import Path

from openpyxl import load_workbook

RAIZ = Path(__file__).resolve().parent.parent
wb = load_workbook(RAIZ / "resultados" / "T_Mortalidad_INEGI_Mexico_2019.xlsx", data_only=True)
wi = wb["Indicadores"]

# indicadores: filas 4..15 (Hombres, Mujeres, Ambos sexos)
ind = {}
for r in range(4, 16):
    ind[wi.cell(r, 1).value] = [wi.cell(r, c).value for c in (2, 3, 4)]
# validación con la ONU
fila_val = next(r for r in range(1, 60) if wi.cell(r, 1).value == "e0 con datos de INEGI (esta tabla)")
val = {"inegi": [wi.cell(fila_val, c).value for c in (2, 3, 4)], "onu": [wi.cell(fila_val + 1, c).value for c in (2, 3, 4)]}
# texto del análisis
fila_an = next(r for r in range(1, 100) if wi.cell(r, 1).value == "Análisis")
analisis, r = [], fila_an + 1
while wi.cell(r, 1).value and wi.cell(r + 1, 1).value:
    analisis.append([wi.cell(r, 1).value, wi.cell(r + 1, 1).value])
    r += 2
# extracto de la tabla (ambos sexos)
ws = wb["Tabla ambos sexos"]
extracto = []
for x in (0, 1, 5, 10, 20, 30, 40, 50, 60, 65, 70, 80, 90, 99, 100):
    r = 5 + x
    extracto.append([x] + [ws.cell(r, c).value for c in range(3, 9)])
datos = {"ind": ind, "val": val, "analisis": analisis, "extracto": extracto,
         "nacimientos": ws["L6"].value, "def_inf": ws["L7"].value, "q0": ws["L9"].value}

salida = RAIZ / "resultados" / "Analisis_Tabla_Mortalidad_INEGI_2019.docx"
with tempfile.TemporaryDirectory() as tmp:
    ruta = Path(tmp) / "datos_informe.json"
    ruta.write_text(json.dumps(datos, ensure_ascii=False))
    subprocess.run(["node", str(RAIZ / "scripts" / "05_informe.js"), str(ruta), str(RAIZ / "resultados"), str(salida)], check=True)
print("informe", salida)
