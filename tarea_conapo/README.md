# U2_P1: Tabla de mortalidad de la generación mexicana de 1970 (CONAPO)

- `resultados/T_Mortalidad_CONAPO_Generacion_1970.xlsx`: **la entrega**. La hoja `Tabla de mortalidad` es la plantilla de la profesora (de 0 a 100 años, fórmulas vivas); las demás hojas contienen los datos usados (`Datos usados`, `Defunciones`, `Población`, `Nacimientos y e0`, `Fuentes`).
- `resultados/Analisis_Tabla_Mortalidad_Generacion_1970.docx`: método, resultados, análisis y limitaciones (4 páginas, con gráficos).
- `resultados/*.png`: lx, qx (log), dx, ex y validación de e0.
- `bases_de_datos_CONAPO/`: las bases originales usadas (ver `LEEME.md` de esa carpeta).
- `plantilla_profesora/`: la plantilla original de la profesora.

## Reproducir

```bash
python3 scripts/02_construir_excel.py     # construye el libro a partir de la plantilla y de las bases
python3 <skill xlsx>/scripts/recalc.py resultados/T_Mortalidad_CONAPO_Generacion_1970.xlsx   # recalcula con LibreOffice
python3 scripts/04_verificar.py           # compara el libro con un cálculo independiente y valida contra la e0 de CONAPO
python3 scripts/03_graficos.py            # PNG
python3 scripts/05_informe.py             # documento Word (requiere node y el paquete npm `docx`)
```
