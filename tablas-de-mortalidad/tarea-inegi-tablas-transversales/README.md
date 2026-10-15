# U2TC: INEGI, tablas transversales abreviadas (México, 2025)

Instrucción de la tarea: *utilizar las fuentes de consulta de INEGI para hacer una cohorte ficticia y formar las tablas transversales abreviadas.*

## Entrega

`resultados/Fuentes_de_consulta.xlsx` es **el mismo archivo de la profesora** con el trabajo agregado. Su hoja `Fuentes de datos` no se tocó (con sus capturas de pantalla) y las hojas nuevas siguen sus colores (verde y turquesa):

| Hoja | Qué tiene |
|---|---|
| `Fuentes de datos` | La de la profesora, igual (población EIC 2025 en `C37:H52`, defunciones de INEGI en `B106:F127`) |
| `Cohorte ficticio` | Radix de 100,000, reparto de sexo y edad no especificados, separación del año 0 y agrupación en 17 grupos |
| `Tabla hombres`, `Tabla mujeres`, `Tabla ambos sexos` | Las tablas transversales abreviadas (nmx, nqx, npx, lx, ndx, nAx, nLx, Tx, ex) |
| `Gráficos y resumen` | Indicadores, tres gráficos de Excel (lx, ex, nqx), comparación con CONAPO y notas |

Todo son fórmulas que leen la hoja de la profesora; si cambia un dato de `Fuentes de datos`, las tablas se recalculan.

Otros archivos: `fuentes/Fuentes_de_consulta.xlsx` (el original que dio la profesora), `fuentes/otras_bases/` (las bases que se descargaron de INEGI: EIC 2025 y la exportación de defunciones, con los mismos datos que ya traía la hoja de la profesora), `datos/` (las dos tablas en CSV) y `scripts/`.

## Qué es la tabla

Cohorte ficticia de 100,000 nacimientos (l0, `Cohorte ficticio!D5`) que vive toda su vida con las tasas de mortalidad por edad de 2025. Grupos: 0, 1-4, 5-9, ..., 70-74 y 75+ (abierto), lo máximo que permite la población de la EIC.

Mismas fórmulas y columnas del ejemplo de la profesora (`TC_TablaTransversal_edit.xlsx`, Cuadro 4.3): `nmx = D/P`, `nqx = (2·n·nmx)/(2 + n·nmx)`, `npx = 1 − nqx`, `lx+n = lx·npx`, `ndx = lx·nqx`, `nAx = 0.12` (edad 0), `1.5` (1-4) y `2.5` (resto), `nLx = lx+n·n + ndx·nAx`, `Tx = Σ nLx`, `ex = Tx/lx`. En 75+ (abierto): `nqx = 1` y `nAx = 1/nmx`. No se hizo la columna `nZx`.

## Resultados

| | Hombres | Mujeres | Ambos sexos |
|---|---|---|---|
| e0 (años) | 76.64 | 83.92 | 80.31 |
| e65 | 24.81 | 28.24 | 26.59 |
| q0 por mil | 9.97 | 7.85 | 8.92 |
| e0 CONAPO 2025 (referencia) | 72.57 | 79.04 | 75.67 |

## Supuestos y límites

1. Las defunciones de sexo no especificado (529) y de edad no especificada (4,016) se reparten proporcionalmente. La población de edad no especificada de la EIC (0.014 %) no se reparte.
2. La EIC trae 0-4 años junto. El año 0 se separa suponiendo que las cohortes de nacimiento crecen a tasa constante, con r = ln(P5-9 / P0-4)/5 (fracción de 0-4 con menos de 1 año ≈ 0.18).
3. **La e0 sale 4 a 5 años por encima de la de CONAPO 2025.** La captura de INEGI de la hoja de la profesora dice que las cifras de defunciones de 2025 son preliminares (por año de registro). Las 762,494 defunciones quedan por debajo de las 836,585 que proyecta CONAPO, y las 12,207 infantiles son pocas para ~2.0 millones de nacimientos. No es un error de fórmulas (el libro coincide con un cálculo independiente en Python); con cifras definitivas la e0 bajaría.

## Reproducir

Desde esta carpeta (requiere LibreOffice y la skill xlsx para recalcular):

```bash
python3 scripts/01_extraer_fuentes.py                 # fuentes/Fuentes_de_consulta.xlsx -> datos/*.csv (para la verificación)
python3 scripts/02_construir_libro.py                 # agrega las hojas al archivo de la profesora -> resultados/_sin_terminar.xlsx
python3 scripts/03_finalizar_libro.py <skill xlsx>/scripts/recalc.py   # guarda los valores calculados y restaura las capturas -> resultados/Fuentes_de_consulta.xlsx
python3 scripts/04_verificar.py                       # compara el libro con calculo_abreviada.py (error relativo ~1e-13)
```
