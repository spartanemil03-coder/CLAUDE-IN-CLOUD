# Bases de datos usadas (CONAPO)

Estos son los archivos originales, sin modificar, tal como se descargaron de datos.gob.mx.

| Archivo | Qué contiene | Para qué se usó |
|---|---|---|
| `01_Defunciones_1950_2070.csv` | Defunciones por año (1970-2070), entidad (32), sexo y edad simple (0-109) | Defunciones de cada edad en el año 1970 + x |
| `00_Pob_Mitad_1950_2070.csv` | Población a mitad de año por año (1970-2070), entidad, sexo y edad simple (0-109) | Población expuesta al riesgo (denominador de qx) |
| `05_indicadores_demograficos_proyecciones.csv` | Indicadores demográficos por año (1950-2070), nacional y por entidad | Nacimientos (q0), verificación de totales y e0 oficial |

- Fuente: Consejo Nacional de Población (CONAPO), *Conciliación demográfica de México 1950-2019 y Proyecciones de la población de México y de las entidades federativas 2020-2070*. Licencia CC BY 4.0.
- Página: https://www.datos.gob.mx/dataset/proyecciones-de-poblacion
- Los datos nacionales del libro de Excel son la suma de las 32 entidades y de hombres y mujeres (las sumas coinciden con los totales nacionales de `05_…`, ver la hoja `Nacimientos y e0`).
