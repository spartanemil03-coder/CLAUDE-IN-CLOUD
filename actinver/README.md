# Reto Actinver 2026

Trabajo del Reto Actinver (5 de octubre al 13 de noviembre de 2026). Está separado del trabajo de mortalidad de la raíz del repo.

- `datos/Reto_Actinver_145_acciones.xlsx`: libro original con las 145 acciones del monitor (foto del 2026-10-08). Sin modificar.
- `datos/yahoo/`: precios, volumen, precios objetivo de analistas y fechas de reporte descargados de Yahoo Finance; covarianza del último año.
- `resultados/Reto_Actinver_mejorado.xlsx`: el libro con precios actualizados y las hojas `Pronóstico`, `Órdenes` y `Portafolios`; `Mi portafolio` ya viene con el ganador.
- `propuesta.md`: ganador, por qué, riesgos y cómo operar.
- `bases_resumen.md`: reglas oficiales que afectan la operación y el Track del Inversionista.
- `resultados/*.csv`: pronóstico por emisora, portafolios de 5 a 12 emisoras, pesos y órdenes.

```bash
python3 actinver/scripts/00_actualizar_datos.py   # descarga de Yahoo Finance (requiere yfinance)
python3 -I actinver/scripts/01_analisis.py        # cribado por costo de operar
python3 actinver/scripts/02_pronostico.py         # pronóstico, optimización de 5 a 12 emisoras y ganador (requiere scipy)
python3 actinver/scripts/03_construir_excel.py    # actualiza el Excel
python3 actinver/scripts/04_con_etfs.py           # repite el ganador con los 40 ETFs permitidos (resultados/con_etfs_*.csv)
python3 <skill xlsx>/scripts/recalc.py actinver/resultados/Reto_Actinver_mejorado.xlsx   # recalcula con LibreOffice
```

Requiere `pandas`, `numpy`, `scipy`, `openpyxl`, `yfinance` y LibreOffice. Todo es hipotético; no es una recomendación de inversión.
