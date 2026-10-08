# Reto Actinver 2026

Trabajo del Reto Actinver (5 de octubre al 13 de noviembre de 2026). Está separado del trabajo de mortalidad de la raíz del repo.

- `datos/Reto_Actinver_145_acciones.xlsx`: libro con las 145 acciones del monitor (foto del 2026-10-08), correlaciones, covarianzas y la calculadora `Mi portafolio`. Original sin modificar.
- `scripts/01_analisis.py`: cribado de las 145 acciones y comparación de tres portafolios candidatos.
- `resultados/cribado_145_acciones.csv`: costo de operar, volatilidad, rendimientos y beta por emisora.
- `resultados/portafolios_candidatos.csv`: volatilidad, costo y probabilidades de cada portafolio.
- `resultados/ordenes_sugeridas.csv`: títulos enteros a comprar al precio de venta del Excel.
- `propuesta.md`: qué comprar, por qué y qué confirmar antes de operar.

```bash
python3 -I actinver/scripts/01_analisis.py
```

Requiere `pandas`, `numpy` y `openpyxl`. Todo es hipotético y descriptivo del pasado; no es una recomendación de inversión.
