# U2_P1: Tabla de mortalidad INEGI (México, 2019)

- `resultados/2_Tabla_de_mortalidad_INEGI_2019.xlsx`: **la tabla de la entrega**. La hoja `Tabla de mortalidad` es la plantilla de la profesora (0 a 100 años, fórmulas simples); la hoja `Insumos` trae las defunciones y la población usadas.
- `resultados/1_Preparacion_datos_INEGI_2019.xlsx`: datos de INEGI ya contados (defunciones, población, nacimientos, fuentes) y su preparación (edad no especificada, suavizado del Censo en la celda J5).
- `resultados/Analisis_Tabla_Mortalidad_INEGI_2019.docx`: método, resultados, análisis y limitaciones, con gráficos.
- `resultados/Explicacion_paso_a_paso_INEGI_2019.docx`: explicación paso a paso de lo que se hizo y respuestas a posibles preguntas de la maestra.
- `resultados/*.png`: lx, qx (log), dx, ex, sobremortalidad masculina y comparación de e0 con CONAPO.
- `bases_de_datos_INEGI/`: las bases originales de INEGI usadas (ver el `LEEME.md` de esa carpeta). Los zips pesan ~234 MB; si faltan, el script 01 los descarga.
- `datos/`: conteos condensados de las bases (los que lee el Excel) y la e0 2019 de CONAPO como referencia externa.
- `plantilla_profesora/`: la plantilla original de la profesora.

## Reproducir

Desde esta carpeta:

```bash
python3 scripts/01_descargar_inegi.py     # descarga las bases a bases_de_datos_INEGI/ (si no están) y condensa los conteos
python3 scripts/02_construir_excel.py     # construye el libro a partir de la plantilla
python3 <skill xlsx>/scripts/recalc.py resultados/<cada libro>.xlsx   # recalcula con LibreOffice
python3 scripts/04_verificar.py           # compara el libro con un cálculo independiente
python3 scripts/03_graficos.py            # PNG
python3 scripts/05_informe.py             # documento Word (requiere node y el paquete npm `docx`)
python3 scripts/06_explicacion.py         # documento Word con la explicación paso a paso
```
