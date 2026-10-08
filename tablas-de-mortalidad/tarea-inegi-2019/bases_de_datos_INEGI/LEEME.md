# Bases de datos usadas (INEGI)

Archivos originales, sin modificar, tal como se descargaron del sitio de INEGI.

| Carpeta | Archivos | Qué contiene | Para qué se usó |
|---|---|---|---|
| `1_Defunciones_EDR/` | `conjunto_de_datos_defunciones_registradas_2019_csv.zip`, `…_2020_csv.zip`, `…_2021_csv.zip` | Estadísticas de Defunciones Registradas (EDR), datos abiertos: una fila por defunción registrada en el año | Defunciones **ocurridas en 2019** por edad y sexo (se usan los tres años de registro por el registro tardío) |
| `2_Nacimientos_ENR/` | `conjunto_de_datos_natalidad_2019_csv.zip`, `…_2020_csv.zip`, `…_2021_csv.zip` | Estadística de Nacimientos Registrados (ENR), datos abiertos: una fila por nacimiento registrado en el año | Nacimientos **ocurridos en 2019** (q0 = defunciones de menores de 1 año / nacimientos) |
| `3_Censo_2020/` | `cpv2020_b_eum_01_poblacion.xlsx` | Censo de Población y Vivienda 2020, tabulados del cuestionario básico (Población), hoja 03: población por edad desplegada y sexo | Población por edad (denominador de qx) |

Enlaces:
- EDR: https://www.inegi.org.mx/programas/edr/#datos_abiertos
- ENR: https://www.inegi.org.mx/programas/natalidad/#datos_abiertos
- Censo 2020: https://www.inegi.org.mx/programas/ccpv/2020/#tabulados

Cada zip de EDR y ENR contiene el CSV (`conjunto_de_datos/`), el diccionario de datos, los catálogos y los metadatos de INEGI.
Campos usados: EDR → `sexo`, `edad`, `anio_ocur`; ENR → `sexo`, `ano_nac`; Censo → hoja 03, fila "Estados Unidos Mexicanos".
Los conteos condensados que usa el Excel están en `../datos/` y se generan con `../scripts/01_descargar_inegi.py` a partir de estos archivos.
