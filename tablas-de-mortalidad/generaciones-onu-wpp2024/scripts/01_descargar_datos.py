"""Descarga los CSV de la ONU (World Population Prospects 2024, variante media) y deja en
datos/ solo las filas de México (LocID 484).

Los archivos originales pesan ~1,6 GB en total; se procesan en streaming (curl | gunzip | pandas)
y no se guardan completos en disco.

Uso:  python3 scripts/01_descargar_datos.py [carpeta_salida]      (por defecto: datos/)
"""
import subprocess
import sys
from pathlib import Path

import pandas as pd

BASE = ("https://population.un.org/wpp/assets/Excel%20Files/1_Indicator%20(Standard)/CSV_FILES/"
        "WPP2024_{name}.csv.gz")
MEXICO = 484
SALIDA = Path(sys.argv[1] if len(sys.argv) > 1 else Path(__file__).resolve().parent.parent / "datos")


def filas_mexico(name):
    """Lee el CSV remoto por streaming y devuelve solo las filas de México."""
    url = BASE.format(name=name)
    proc = subprocess.Popen(["curl", "-sS", "-f", "-L", url], stdout=subprocess.PIPE)
    partes = [c[c.LocID == MEXICO]
              for c in pd.read_csv(proc.stdout, compression="gzip", chunksize=500_000, low_memory=False)]
    if proc.wait() != 0:
        raise RuntimeError(f"Falló la descarga de {url}")
    return pd.concat(partes)


def main():
    SALIDA.mkdir(parents=True, exist_ok=True)

    # Población a mitad de año (1 de julio) y defunciones por edad simple y sexo; miles de personas.
    for tag in ("PopulationBySingleAgeSex", "DeathsBySingleAgeSex"):
        df = pd.concat([filas_mexico(f"{tag}_Medium_1950-2023"), filas_mexico(f"{tag}_Medium_2024-2100")])
        df.to_csv(SALIDA / f"mexico_{tag}.csv", index=False)
        print(tag, df.shape)

    # Indicadores demográficos: esperanza de vida al nacer oficial (LEx) por año calendario.
    di = filas_mexico("Demographic_Indicators_Medium")
    cols = ["LocID", "Location", "Time", "LEx", "LExMale", "LExFemale", "Births", "Deaths", "IMR", "Q5"]
    di[cols].to_csv(SALIDA / "mexico_DemographicIndicators.csv", index=False)
    print("DemographicIndicators", di.shape)

    # Tablas de vida completas de la ONU (qx por edad simple): solo para validar las generaciones.
    partes = []
    for sexo in ("Male", "Female", "Both"):
        for periodo in ("1950-2023", "2024-2100"):
            t = filas_mexico(f"Life_Table_Complete_Medium_{sexo}_{periodo}")
            partes.append(t[["Time", "Sex", "AgeGrpStart", "mx", "qx", "lx", "ex"]])
            print("LifeTable", sexo, periodo, t.shape)
    pd.concat(partes).to_csv(SALIDA / "mexico_ONU_TablaVida_Completa.csv", index=False)


if __name__ == "__main__":
    main()
