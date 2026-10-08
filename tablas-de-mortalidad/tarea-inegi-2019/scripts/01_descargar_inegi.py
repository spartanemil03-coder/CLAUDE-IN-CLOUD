"""Descarga las bases de datos abiertas de INEGI a bases_de_datos_INEGI/ (originales, sin modificar) y condensa en datos/
solo los conteos necesarios para la tabla de mortalidad.

Fuentes (todas INEGI):
  - Estadísticas de Defunciones Registradas (EDR), conjuntos de datos abiertos 2019, 2020 y 2021 (CSV):
    https://www.inegi.org.mx/programas/edr/#datos_abiertos
  - Estadística de Nacimientos Registrados (ENR / natalidad), conjuntos de datos abiertos 2019, 2020 y 2021 (CSV):
    https://www.inegi.org.mx/programas/natalidad/#datos_abiertos
  - Censo de Población y Vivienda 2020, tabulados del cuestionario básico, "Población 3" (edad desplegada y sexo):
    https://www.inegi.org.mx/programas/ccpv/2020/#tabulados

Año de referencia: 2019 (último año completo antes de la pandemia). Defunciones y nacimientos se cuentan por AÑO DE
OCURRENCIA = 2019, tomando los registrados en 2019 y en los dos años siguientes (registro tardío).
Si un archivo ya está en bases_de_datos_INEGI/, no se vuelve a descargar.

Uso:  python3 scripts/01_descargar_inegi.py
"""
import subprocess
import tempfile
import zipfile
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

RAIZ = Path(__file__).resolve().parent.parent
BASES = RAIZ / "bases_de_datos_INEGI"
DATOS = RAIZ / "datos"
ANIO = 2019
ANIOS_ARCHIVO = (ANIO, ANIO + 1, ANIO + 2)
INEGI = "https://www.inegi.org.mx/contenidos/programas"
ARCH = {
    "edr": (BASES / "1_Defunciones_EDR", INEGI + "/edr/datosabiertos/defunciones/{y}/conjunto_de_datos_defunciones_registradas_{y}_csv.zip"),
    "enr": (BASES / "2_Nacimientos_ENR", INEGI + "/natalidad/datosabiertos/{y}/conjunto_de_datos_natalidad_{y}_csv.zip"),
}
CENSO = BASES / "3_Censo_2020" / "cpv2020_b_eum_01_poblacion.xlsx"
URL_CENSO = INEGI + "/ccpv/2020/tabulados/cpv2020_b_eum_01_poblacion.xlsx"


def bajar(url, destino):
    destino.parent.mkdir(parents=True, exist_ok=True)
    if not destino.exists():
        subprocess.run(["curl", "-sS", "-f", "-L", "-m", "900", "-o", str(destino), url], check=True)
    return destino


def zip_programa(prog, y):
    carpeta, url = ARCH[prog]
    return bajar(url.format(y=y), carpeta / Path(url).name.format(y=y))


def leer_csv(zpath, cols, tmp):
    with zipfile.ZipFile(zpath) as z:
        nombres = [n for n in z.namelist() if "conjunto_de_datos/" in n and n.lower().endswith(".csv") and "bitacora" not in n.lower()]
        assert len(nombres) == 1, nombres
        z.extract(nombres[0], tmp)
    d = pd.read_csv(Path(tmp) / nombres[0], encoding="latin1", usecols=lambda c: c.lower() in cols, low_memory=False)
    d.columns = d.columns.str.lower()
    return d


def defunciones(tmp):
    partes, por_archivo = [], []
    for y in ANIOS_ARCHIVO:
        z = zip_programa("edr", y)
        d = leer_csv(z, {"sexo", "edad", "anio_ocur"}, tmp)
        d = d[d.anio_ocur == ANIO]
        partes.append(d)
        por_archivo.append({"archivo": z.name, "anio_registro": y, "defunciones_ocurridas_2019": len(d)})
        print("EDR", y, len(d))
    d = pd.concat(partes, ignore_index=True)
    # EDAD: 1001-3999 = menores de un año (horas, días, meses); 4001-4120 = años cumplidos; 4998 = no especificada
    edad = np.where(d.edad < 4000, 0, np.where(d.edad == 4998, -1, d.edad - 4000))
    d["edad"] = np.where(edad < 0, -1, np.minimum(edad, 100))
    tab = pd.crosstab(d.edad, d.sexo).reindex(columns=[1, 2, 9], fill_value=0)
    tab.columns = ["hombres", "mujeres", "sexo_no_especificado"]
    tab = tab.reindex(list(range(101)) + [-1], fill_value=0)
    tab.index = [str(i) if i >= 0 else "NE" for i in tab.index]
    tab.index.name = "edad"
    return tab, pd.DataFrame(por_archivo)


def nacimientos(tmp):
    filas = []
    for y in ANIOS_ARCHIVO:
        z = zip_programa("enr", y)
        d = leer_csv(z, {"sexo", "ano_nac"}, tmp)
        d = d[d.ano_nac == ANIO]
        c = d.sexo.value_counts().reindex([1, 2, 9], fill_value=0)
        filas.append({"archivo": z.name, "anio_registro": y, "hombres": c[1], "mujeres": c[2], "sexo_no_especificado": c[9]})
        print("ENR", y, len(d))
    return pd.DataFrame(filas)


def poblacion_censo():
    bajar(URL_CENSO, CENSO)
    ws = load_workbook(CENSO, read_only=True, data_only=True)["03"]
    filas = [r[1:5] for r in ws.iter_rows(min_row=1, values_only=True) if r[0] == "Estados Unidos Mexicanos" and r[1] != "Total"]
    assert len(filas) == 102 and filas[0][0].startswith("00") and filas[100][0].startswith("100") and filas[101][0].startswith("No esp")
    df = pd.DataFrame({"edad": [str(i) for i in range(101)] + ["NE"], "hombres": [f[2] for f in filas], "mujeres": [f[3] for f in filas]})
    return df.set_index("edad")


def main():
    DATOS.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        d, d_arch = defunciones(tmp)
        n_arch = nacimientos(tmp)
    d.to_csv(DATOS / f"inegi_defunciones_{ANIO}_por_edad_sexo.csv")
    d_arch.to_csv(DATOS / f"inegi_defunciones_{ANIO}_por_archivo.csv", index=False)
    n_arch.to_csv(DATOS / f"inegi_nacimientos_{ANIO}_por_archivo.csv", index=False)
    n_arch[["hombres", "mujeres", "sexo_no_especificado"]].sum().to_frame().T.to_csv(DATOS / f"inegi_nacimientos_{ANIO}_por_sexo.csv", index=False)
    poblacion_censo().to_csv(DATOS / "inegi_censo2020_poblacion_por_edad_sexo.csv")
    print("defunciones", int(d.values.sum()), "| nacimientos", int(n_arch[["hombres", "mujeres", "sexo_no_especificado"]].values.sum()))


if __name__ == "__main__":
    main()
