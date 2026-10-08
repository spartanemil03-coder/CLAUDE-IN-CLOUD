# Modelo de Markowitz: portafolio de acciones mexicanas

Portafolio de media-varianza con 8 emisoras de la BMV (AC, ALSEA, AMXB, BIMBOA, CEMEXCPO, GFNORTEO, TLEVISACPO, WALMEX) y el IPC como referencia.

- `Portafolio_Markowitz_Profesor.xlsx`: **la entrega**, en el formato de la plantilla del profesor (hojas `PRECIOS`, `RENDIMIENTOS`, `Matriz`, `Portafolio`). Incluye la frontera eficiente, el índice de Sharpe y la beta del portafolio. En `Portafolio`, la celda del método elige la frontera: 1 = fórmulas (se actualiza sola y permite ventas en corto), 2 = Solver solo con posiciones largas (método por defecto).
- `Modelo_Markowitz.xlsx`: versión de trabajo con notas: precios, rendimientos logarítmicos, matriz de covarianzas y portafolios de mínima varianza y máximo Sharpe.

Datos: precios de cierre de Investing.com del 6 de octubre de 2025 al 6 de octubre de 2026 (252 cierres). Tasa libre de riesgo: Cetes 6.50%. Con solo posiciones largas, el portafolio de mayor Sharpe de la frontera queda 100% en GFNORTEO.
