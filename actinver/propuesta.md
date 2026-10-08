# Propuesta de portafolio: Reto Actinver 2026

Datos: Yahoo Finance al 8 de octubre de 2026 (~11:10 h CDMX) para precios, historial, precios objetivo de analistas y fechas de reporte; monitor de la plataforma para precio de compra/venta y profundidad. Hipotético. No es una recomendación de inversión.

## Ganador

| Emisora | Peso | Títulos* | Importe* | Próximo reporte |
|---|---|---|---|---|
| MU * (Micron) | 50% | 25 | $488,125 | 23 dic (fuera del reto) |
| INTC * (Intel) | 41% | 204 | $407,998 | 29 oct (dentro del reto) |
| PINFRA * | 3% | 109 | $29,880 | 22 oct |
| GFNORTE O | 3% | 157 | $29,759 | 3 nov |
| BOLSA A | 3% | 848 | $29,892 | 20 oct |

\* Al precio de venta del monitor. MU va -3.4% hoy en Yahoo, así que en la plataforma puede estar más barata; la hoja `Órdenes` recalcula los títulos al poner el precio que veas. Deja ~$13,000 en efectivo para comisiones.

| | Ganador | Balanceado (5) | Tu idea (NVDA + Banorte) |
|---|---|---|---|
| Prob. de superar +20% (top 3) | **15.6%** | 2.9% | 0.8% |
| Prob. de superar +30% | 6.9% | 0.7% | 0.2% |
| Ganancia esperada | +$31,600 | +$20,100 | +$5,300 |
| Si el mercado baja | -$68,700 | -$37,000 | -$42,000 |
| Si el mercado sube | +$131,900 | +$77,200 | +$52,600 |
| Prob. de perder más de $100k | 20.5% | 7.6% | 4.6% |
| Peor 5% | -$274,000 | -$124,500 | -$96,000 |

## ¿Cuántas emisoras?

**5.** Con el objetivo de quedar en el top 3, cada emisora extra baja la probabilidad de superar +20%: 15.6% con 5, 13.8% con 8, 12.5% con 10, 11.1% con 12. En el balanceado pasa lo mismo con la ganancia esperada ($20.1k con 5, $18.1k con 10). Todo está en la hoja `Pronóstico`.

## Por qué este

1. **Se gana quedando en el top 3**, no en promedio. Lo que importa es la probabilidad de un resultado muy alto (+20%, lo que lograron los 3 primeros en una edición pasada). Eso pide concentrar en lo que más se mueve y tiene mejor pronóstico.
2. **Pronóstico por acción** = beta × rendimiento esperado del mercado + alfa − costo de operar. El alfa combina dos señales reales: potencial al precio objetivo de analistas y momento de 12 meses. Se pondera con una capacidad de pronóstico baja (IC = 0.05), porque estas señales predicen poco.
3. **MU** tiene el mejor pronóstico de las 145: 46 analistas con objetivo 49% arriba, el mayor momento y volatilidad de 81%. **INTC** es la segunda entre las líquidas: 78% de volatilidad, momento fuerte y reporta el 29 de octubre (más movimiento). Las otras 3 son de bajo riesgo y solo cumplen el mínimo de 5.
4. **Se sostiene sin el pronóstico**: con IC = 0 sigue siendo el de mayor probabilidad de top 3 (12.6%), aunque su ganancia esperada baja a +$5,200.
5. **Universo filtrado**: operables, spread de 1.5% o menos y líquidas. Quedaron fuera MARA, LCID y MFRISCO (tenían buen pronóstico, pero menos de $3,000 en el mejor nivel del monitor) y BYND/MRNA (spread o pronóstico negativo).

## Riesgos

- MU e INTC son ambas de semiconductores: si el sector cae, caen juntas. 1 de cada 5 simulaciones pierde más de $100,000.
- Es dinero virtual; los 48 premios por avance (quizzes y Tracks) no dependen de esto.
- Si prefieres menos riesgo, el balanceado de 5 (MU 28%, LAB B 24%, CEMEX 23%, GCC 17%, BABA 9%) gana más en promedio que tu idea original con riesgo parecido.

## Cómo operar (hoy hasta las 14:00 h)

1. Revisa en la plataforma que las 5 estén operables y el precio de venta actual de cada una.
2. Primero MU y luego INTC, con orden limitada al precio de venta que ves.
3. Después PINFRA, GFNORTE y BOLSA.
4. No vuelvas a mover nada salvo que cambie algo importante: cada compra y venta cuesta.
5. Confirma en las bases si el 50% máximo se revisa solo al comprar o todo el tiempo: si MU sube mucho, podría pasar de 50%.
