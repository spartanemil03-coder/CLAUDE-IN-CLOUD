# U2TC: INEGI, tablas transversales abreviadas (México, 2025)

Instrucción de la tarea: *usar las fuentes de consulta de INEGI para hacer un cohorte ficticio y formar las tablas transversales abreviadas.*

- `resultados/Tablas_transversales_abreviadas_INEGI_2025.xlsx`: **la entrega**. Seis hojas con fórmulas vivas: `Resumen` (indicadores y supuestos), `Fuentes` (datos de INEGI tal cual), `Insumos` (ajustes y agrupación) y una tabla abreviada para `Hombres`, `Mujeres` y `Ambos sexos`.
- `resultados/*.png`: lx, ex y nqx por sexo.
- `fuentes/Fuentes_de_consulta.xlsx`: el archivo que dio la maestra (EIC 2025 y defunciones de INEGI).
- `datos/`: las dos tablas del archivo, ya en CSV.
- `scripts/`: cómo se construyó y se verificó todo.

## Qué es la tabla

Cohorte ficticia de 100,000 nacimientos (l0, celda `Insumos!C4`) que vive toda su vida con las tasas de mortalidad por edad de 2025. Grupos de edad: 0, 1-4, 5-9, ..., 70-74 y 75+ (abierto), el máximo detalle que permite la población de la EIC.

Por grupo: `nMx = D/P`, `nqx = n·nMx / (1 + (n − nax)·nMx)`, `lx+n = lx·(1 − nqx)`, `ndx = lx − lx+n`, `nLx = n·lx+n + nax·ndx`, `Tx = Σ nLx`, `ex = Tx/lx`. Con `nax = n/2` salvo 0 y 1-4 (Coale-Demeny) y, en 75+, `q = 1` y `L = l/M`.

## Resultados

| | Hombres | Mujeres | Ambos sexos |
|---|---|---|---|
| e0 (años) | 76.64 | 83.92 | 80.31 |
| e65 | 24.81 | 28.24 | 26.59 |
| q0 por mil | 9.93 | 7.82 | 8.89 |
| e0 CONAPO 2025 (referencia) | 72.57 | 79.04 | 75.67 |

## Supuestos y límites (importante)

1. Las defunciones de sexo no especificado (529) y de edad no especificada (4,016) se reparten proporcionalmente.
2. La EIC trae 0-4 años junto. El año 0 se separa suponiendo que las cohortes de nacimiento crecen a tasa constante, con r = ln(P5-9 / P0-4)/5 (fracción de 0-4 con menos de 1 año ≈ 0.18).
3. El archivo de defunciones no dice el año ni si son ocurridas o registradas; se usan como las de 2025.
4. **La e0 sale 4 a 5 años por encima de la de CONAPO 2025.** Las defunciones infantiles del archivo (12,207) son bajas frente a los ~2.0 millones de nacimientos de CONAPO, y las tasas de 70+ parecen bajas. No es un error de fórmulas (el libro coincide con un cálculo independiente), sino de los datos de entrada. Conviene confirmar con la maestra el año y la cobertura de las defunciones.

## Reproducir

Desde esta carpeta (requiere LibreOffice para recalcular):

```bash
python3 scripts/01_extraer_fuentes.py     # fuentes/Fuentes_de_consulta.xlsx -> datos/*.csv
python3 scripts/02_construir_excel.py     # construye el libro
python3 <skill xlsx>/scripts/recalc.py resultados/Tablas_transversales_abreviadas_INEGI_2025.xlsx   # recalcula con LibreOffice
python3 scripts/04_verificar.py           # compara el libro con calculo_abreviada.py (error relativo ~1e-13)
python3 scripts/03_graficos.py            # PNG (requiere matplotlib)
```
