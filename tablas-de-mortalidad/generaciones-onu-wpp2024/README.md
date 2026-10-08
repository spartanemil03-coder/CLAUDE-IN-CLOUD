# Tablas de mortalidad de generaciones mexicanas (ONU WPP 2024)

Tablas de vida (qx, lx, dx, Lx, Tx, ex; radix 100,000) para las generaciones 1950, 1960, 1970, 1980, 1990 y 2000,
por sexo, con fórmulas vivas en Excel, comparación entre generaciones y validación contra la e0 oficial de la ONU.

- `resultados/tablas_mortalidad_generaciones_mexico.xlsx`: libro de Excel (hoja `Leeme` explica la estructura).
- `resultados/informe_tablas_mortalidad_generaciones.docx`: informe corto del método y la validación.
- `resultados/*.png`: gráficos de ex, lx, e0 por año calendario y validación.
- `datos/`: CSV de México extraídos de la ONU (World Population Prospects 2024, variante media).

## Reproducir

Desde esta carpeta (`tablas-de-mortalidad/generaciones-onu-wpp2024/`):

```bash
python3 scripts/01_descargar_datos.py          # descarga y filtra México (~1.6 GB en streaming)
python3 scripts/02_construir_excel.py          # construye el libro
python3 <skill xlsx>/scripts/recalc.py resultados/tablas_mortalidad_generaciones_mexico.xlsx   # recalcula con LibreOffice
python3 scripts/04_verificar.py                # compara el libro con un cálculo independiente en numpy
python3 scripts/03_graficos.py                 # PNG a partir de los valores del libro
python3 scripts/05_informe.py                  # informe Word (requiere node y el paquete npm `docx`)
```

Requiere `pandas`, `openpyxl`, `matplotlib` y LibreOffice (para recalcular las fórmulas).
