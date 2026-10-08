# Reto Actinver 2026

Trabajo del Reto Actinver (5 de octubre al 13 de noviembre de 2026). Está separado del trabajo de mortalidad de la raíz del repo.

- `datos/Reto_Actinver_145_acciones.xlsx`: libro original con las 145 acciones del monitor (foto del 2026-10-08). Sin modificar.
- `resultados/Reto_Actinver_mejorado.xlsx`: el mismo libro con tres hojas nuevas (`Pronóstico`, `Portafolios`, `Beta`) y `Mi portafolio` ya llenado con el ganador.
- `propuesta.md`: qué comprar, cómo se obtuvo y qué confirmar antes de operar.
- `resultados/pronostico_escenarios.csv`, `pesos_portafolios.csv`, `ordenes_sugeridas.csv`, `cribado_145_acciones.csv`: salidas de los scripts.

```bash
python3 -I actinver/scripts/01_analisis.py        # cribado por costo de operar
python3 actinver/scripts/02_escenarios.py         # escenarios, optimización y ganador (requiere scipy)
python3 actinver/scripts/03_construir_excel.py    # agrega las hojas al Excel original
python3 <skill xlsx>/scripts/recalc.py actinver/resultados/Reto_Actinver_mejorado.xlsx   # recalcula con LibreOffice
```

Requiere `pandas`, `numpy`, `scipy`, `openpyxl` y LibreOffice. Todo es hipotético y descriptivo del pasado; no es una recomendación de inversión.
