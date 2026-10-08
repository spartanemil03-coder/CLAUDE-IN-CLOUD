# Propuesta de portafolio: Reto Actinver 2026

Datos: Excel del 2026-10-08. Hipotético y descriptivo del pasado. Nadie sabe qué va a pasar en 25 sesiones; el modelo no tiene un pronóstico por acción, solo escenarios de mercado.

## Ganador: portafolio E

Cinco acciones al 20% cada una, $1,000,000 en total:

| Emisora | Peso | Títulos (al precio de venta del Excel) | Importe |
|---|---|---|---|
| GOOGL * | 20% | 31 | $197,160 |
| CEMEX CPO | 20% | 11,492 | $199,386 |
| GMEXICO B | 20% | 860 | $199,176 |
| TSLA * | 20% | 29 | $196,909 |
| AVGO * | 20% | 29 | $196,620 |

Deja ~$11,000 en efectivo para comisiones. Ya viene cargado en la hoja `Mi portafolio` del Excel mejorado (`resultados/Reto_Actinver_mejorado.xlsx`).

## Cómo se obtuvo

1. **Mercado**: canasta de las 144 acciones con historial, mismo peso. Se mueve ±4.9% (1σ) en 25 sesiones. Cada acción tiene una beta contra esa canasta (sale de la hoja Covarianza).
2. **Tres escenarios** para la canasta: bajista -4.9% (25%), base +1.0% (50%), alcista +4.9% (25%). Se pueden cambiar en la hoja `Pronóstico`.
3. **Ganancia** = $1,000,000 × (beta del portafolio × rendimiento de la canasta) − costo de operar. El costo es la diferencia compra-venta completa más 2 comisiones.
4. **Optimización** sobre 93 emisoras operables con al menos $40,000 en el mejor nivel del libro: un optimizador buscó la mezcla con topes por emisora y hasta 12 emisoras (las filas de `Mi portafolio`). D busca la mejor ganancia por riesgo con tope de 15% por emisora; E busca la mayor ganancia con tope de 20% y riesgo anual hasta 30%. Se compararon contra tus ideas iniciales (A, B y C).

## Resultados (pesos, ya con costos)

| Portafolio | Emisoras | Beta | Bajista | Base | Alcista | Esperada | Vol. anual | Peor 5% |
|---|---|---|---|---|---|---|---|---|
| A Prudente | 5 | 0.66 | -$35,600 | +$3,700 | +$29,800 | +$400 | 17.7% | -$90,900 |
| B Tu idea (NVDA + Banorte + ...) | 5 | 0.98 | -$52,800 | +$5,600 | +$44,400 | +$700 | 21.4% | -$109,900 |
| C Ranking (NVDA 40%, TSLA, AVGO...) | 5 | 1.25 | -$67,100 | +$7,400 | +$56,900 | +$1,200 | 29.1% | -$150,200 |
| D Optimizado, 7 emisoras | 7 | 1.14 | -$59,800 | +$7,900 | +$52,800 | +$2,200 | 23.3% | -$118,500 |
| **E Optimizado, ganancia máx.** | 5 | 1.29 | -$67,800 | +$9,100 | +$60,200 | **+$2,700** | 27.0% | -$137,200 |

- **Ganador de ganancia esperada y de ganancia por riesgo: E.** También gana en el escenario base y en el alcista. En el bajista gana A, que pierde la mitad.
- **D casi empata** (ganancia por riesgo 0.030 contra 0.032) con 7 emisoras y menos riesgo (23% contra 27%). Si prefieres más acciones y menos susto, D es la alternativa: OMA, GOOGL, CEMEX, GMEXICO, TSLA, AVGO al 15% y GFNORTE al 10%.
- Más acciones sí bajan el riesgo (D con 7 tiene 23% de volatilidad contra 27% de E con casi el mismo rendimiento esperado). El número de emisoras lo fijan los topes que puse: 15% da al menos 7 y 20% da al menos 5. No probé topes más bajos.

## Lo que el modelo sí y no dice

- **No dice que E vaya a ganar $2,700.** La ganancia esperada es pequeña frente al riesgo: la desviación es de unos $85,000 y la probabilidad de terminar en positivo es de 51%. El resultado lo decide cómo se mueva el mercado.
- **E es una apuesta de beta alta**: más mercado por cada peso. Gana si el mercado sube, pierde más si baja. Es el efecto de no poder apalancarse y querer el máximo dinero con $1,000,000.
- **El optimizador dejó fuera NVIDIA**: su beta es 1.12, pero cuesta 0.68% ida y vuelta y se mueve mucho por razones propias. Banorte queda solo en D (10%). Si quieres una de las dos por convicción, B las incluye.
- **D y E se optimizaron con los mismos datos del último año**, así que su ventaja sobre A, B y C es optimista.
- **Liquidez**: GOOGL, CEMEX y AVGO muestran unos $42,000 a $47,000 en el mejor nivel del libro y E pide $200,000 en cada una. Revisa la profundidad en la plataforma antes de operar.

## La plataforma

La captura del monitor muestra pestañas de Acciones, ETFs y Fondos, con poder de compra de $1,000,000 y $0 invertidos. Seguimos solo con acciones, como decidiste. No analicé ETFs ni fondos: el Excel no trae sus datos.

## Antes de operar

1. Revisa cómo se valúa al 13 de noviembre. Si es a precio de mercado sin vender, no pagas el spread de salida y el costo real es menor.
2. Horario: 7:30 a 14:00 (8:30 a 15:00 desde el 3 de noviembre). Las órdenes asignadas no se cancelan.
3. Vuelve a ver precios y volumen del día: el Excel es del 8 de octubre y el reto empezó el 5.
4. Fechas de resultados trimestrales de Alphabet, Tesla y las demás: suelen caer a finales de octubre y mueven mucho el precio. El Excel no las trae.
