# CLAUDE-IN-CLOUD

Trabajos de actuaría y finanzas hechos con Claude Code. Cada tema vive en su propia carpeta y cada proyecto se puede reproducir con sus scripts.

## Resumen

| Tema | Carpeta | Qué se hizo | Fechas |
|---|---|---|---|
| Tablas de mortalidad | [`tablas-de-mortalidad/`](tablas-de-mortalidad/) | Tres tablas de mortalidad de México con fórmulas vivas en Excel, gráficos e informes en Word | 6 al 8 oct 2026 |
| Modelo de Markowitz | [`modelo-markowitz/`](modelo-markowitz/) | Portafolio de media-varianza con 8 acciones mexicanas, frontera eficiente y Sharpe en la plantilla del profesor | 7 oct 2026 |
| Reto Actinver 2026 | [`reto-actinver-2026/`](reto-actinver-2026/) | Análisis de las 145 acciones del simulador, pronóstico, elección del portafolio y operación en la plataforma | 8 oct 2026 |

### Tablas de mortalidad

- **Generaciones 1950 a 2000 (ONU WPP 2024):** tablas de vida (qx, lx, dx, Lx, Tx, ex) por generación y sexo, comparación entre generaciones y validación contra la e0 de la ONU. [Ver carpeta](tablas-de-mortalidad/generaciones-onu-wpp2024/).
- **Tarea CONAPO, generación 1970:** tabla en la plantilla de la profesora con las bases de CONAPO, análisis en Word. Esperanza de vida al nacer de la generación: 70.9 años. [Ver carpeta](tablas-de-mortalidad/tarea-conapo-generacion-1970/).
- **Tarea INEGI 2019:** tabla de México 2019 con defunciones y nacimientos registrados y el Censo 2020, libro de preparación de datos, explicación paso a paso y versión en R. Esperanza de vida al nacer en 2019: 76.0 años (ambos sexos). [Ver carpeta](tablas-de-mortalidad/tarea-inegi-2019/).

### Modelo de Markowitz

Precios de AC, ALSEA, AMXB, BIMBOA, CEMEXCPO, GFNORTEO, TLEVISACPO y WALMEX (oct 2025 a oct 2026), matriz de covarianzas, frontera eficiente con fórmulas o con Solver (solo largos) y Cetes 6.50% como tasa libre de riesgo. [Ver carpeta](modelo-markowitz/).

### Reto Actinver 2026

- **Datos:** Excel con las 145 acciones del simulador, actualizado con precios de Yahoo Finance, precios objetivo de analistas y fechas de reporte. También se revisaron los 56 ETFs y los 25 fondos de la plataforma.
- **Pronóstico:** cada acción combina su sensibilidad al mercado, el potencial según analistas y su momento, menos el costo de operar. Se probaron portafolios de 5 a 12 emisoras.
- **Portafolio elegido (8 oct 2026):** MU 50%, INTC 41%, PINFRA 3%, GFNORTE 3%, BOLSA 3%. Es el de mayor probabilidad de superar +20% (el nivel del top 3 en una edición pasada) y no mejora al incluir ETFs.
- **Pendiente:** compras mínimas del Track del Inversionista (fondo, ETF, acción y una operación, del 13 de octubre al 7 de noviembre).
- Detalle en [`propuesta.md`](reto-actinver-2026/propuesta.md), [`bases_resumen.md`](reto-actinver-2026/bases_resumen.md) y [`operaciones.md`](reto-actinver-2026/operaciones.md).

## Estructura

```
.
├── tablas-de-mortalidad/
│   ├── generaciones-onu-wpp2024/
│   ├── tarea-conapo-generacion-1970/
│   └── tarea-inegi-2019/
├── modelo-markowitz/
└── reto-actinver-2026/
```

Para recalcular los libros de Excel hace falta LibreOffice; cada carpeta explica sus requisitos en su propio README.
