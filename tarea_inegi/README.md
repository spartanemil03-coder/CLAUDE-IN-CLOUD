# U2_P1: Tabla de mortalidad INEGI (México, 2019)

Tabla de mortalidad con el formato de la plantilla de clase (x, qx, lx, dx, Lx, Tx, ex), calculada con datos abiertos de INEGI.

- `resultados/T_Mortalidad_INEGI_Mexico_2019.xlsx`: libro con las tablas (ambos sexos, hombres y mujeres), indicadores, análisis y gráficos. Todo son fórmulas vivas; el radix y el suavizado se cambian en la hoja `Leeme`.
- `resultados/Analisis_Tabla_Mortalidad_INEGI_2019.docx`: método, resultados, análisis y limitaciones (4 páginas, con gráficos).
- `resultados/*.png`: lx, qx (log), dx, ex, sobremortalidad masculina y comparación con la ONU.
- `datos/`: conteos de INEGI ya condensados (defunciones por edad y sexo, nacimientos por sexo, población del Censo 2020 por edad y sexo) y dos archivos de referencia de la ONU solo para comparar.

## Fuentes

- Defunciones: INEGI, Estadísticas de Defunciones Registradas (datos abiertos 2019-2021). https://www.inegi.org.mx/programas/edr/
- Nacimientos: INEGI, Estadística de Nacimientos Registrados (datos abiertos 2019-2021). https://www.inegi.org.mx/programas/natalidad/
- Población: INEGI, Censo de Población y Vivienda 2020, tabulado "Población 3". https://www.inegi.org.mx/programas/ccpv/2020/

## Reproducir

```bash
python3 scripts/01_descargar_inegi.py          # descarga y condensa los datos de INEGI (~250 MB)
python3 scripts/02_construir_excel.py          # construye el libro
python3 <skill xlsx>/scripts/recalc.py resultados/T_Mortalidad_INEGI_Mexico_2019.xlsx   # recalcula con LibreOffice
python3 scripts/04_verificar.py                # compara el libro con un cálculo independiente
python3 scripts/03_graficos.py                 # PNG a partir de los valores del libro
python3 scripts/05_informe.py                  # documento Word (requiere node y el paquete npm `docx`)
```

Los archivos `datos/onu_wpp2024_*.csv` se derivaron de la tabla de vida completa de la ONU (WPP 2024) que está en `../datos/` (primera tarea del repositorio).
