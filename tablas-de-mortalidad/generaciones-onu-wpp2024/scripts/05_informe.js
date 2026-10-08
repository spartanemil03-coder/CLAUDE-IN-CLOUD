// Informe Word: tablas de mortalidad de generaciones mexicanas.
// Uso: node scripts/05_informe.js datos.json carpeta_resultados salida.docx
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, BorderStyle, LevelFormat, Footer, PageNumber,
} = require("docx");

const [, , jsonPath, resDir, outPath] = process.argv;
const D = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
const f1 = (x) => x.toFixed(1);
const f2 = (x) => x.toFixed(2);
const pct = (x) => (100 * x).toFixed(1) + "%";
const sg = (x) => (x >= 0 ? "+" : "−") + Math.abs(x).toFixed(2);
const A = (s) => D.gen[s]; // series por sexo
const amb = A("Ambos sexos"), hom = A("Hombres"), muj = A("Mujeres");
const per = Object.fromEntries(D.periodo.map((p) => [p.sexo, p]));

const FONT = "Arial";
const W = 9360; // ancho útil, carta con márgenes de 1"
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const P = (children, o = {}) =>
  new Paragraph({ spacing: { after: 120, line: 276 }, ...o, children: Array.isArray(children) ? children : [run(children)] });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 26, color: "1F3864" })] });
const bullet = (children) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 60, line: 264 }, children: Array.isArray(children) ? children : [run(children)] });
const formula = (t) => new Paragraph({ indent: { left: 540 }, spacing: { after: 60 }, children: [new TextRun({ text: t, font: "Consolas", size: 20 })] });

const borde = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };
function celda(t, w, o = {}) {
  return new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: bordes,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    children: [new Paragraph({ keepNext: true, alignment: o.left ? AlignmentType.LEFT : AlignmentType.CENTER, children: [new TextRun({ text: String(t), font: FONT, size: 18, bold: !!o.bold })] })],
  });
}
function tabla(cols, anchos, filas) {
  return new Table({
    width: { size: W, type: WidthType.DXA }, columnWidths: anchos,
    rows: [
      new TableRow({ tableHeader: true, children: cols.map((c, i) => celda(c, anchos[i], { fill: "DCE3F0", bold: true })) }),
      ...filas.map((fila) => new TableRow({ children: fila.map((c, i) => celda(c, anchos[i], { left: i === 0, bold: i === 0 })) })),
    ],
  });
}
function figura(nombre, ancho, alto, pie) {
  return [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 40 }, keepNext: true,
      children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(resDir, nombre)), transformation: { width: ancho, height: alto },
        altText: { title: nombre, description: pie, name: nombre } })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 160 }, children: [run(pie, { size: 17, italics: true, color: "595959" })] }),
  ];
}

const g0 = amb[0], g5 = amb[5];
const proy = (g) => 101 - (2024 - g); // edades observadas en años >= 2024 (proyección)
const rangoDif = [Math.min(...D.sexos.flatMap((s) => A(s).map((x) => x.dif))), Math.max(...D.sexos.flatMap((s) => A(s).map((x) => x.dif)))];
const maxAdj = Math.max(...D.sexos.flatMap((s) => A(s).map((x) => x.ajuste)));

const hijos = [
  new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: "Tablas de mortalidad de generaciones mexicanas", font: FONT, bold: true, size: 36, color: "1F3864" })] }),
  new Paragraph({ spacing: { after: 200 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "1F3864", space: 4 } },
    children: [run("Generaciones 1950, 1960, 1970, 1980, 1990 y 2000. Método y validación. Datos: ONU, World Population Prospects 2024.", { color: "595959" })] }),

  H1("1. Qué se hizo"),
  P("Con los datos de la ONU se construyó una tabla de mortalidad para cada una de seis generaciones mexicanas (nacidas en 1950, 1960, 1970, 1980, 1990 y 2000), por separado para hombres, mujeres y ambos sexos, con radix de 100,000. Las tablas están en un libro de Excel donde todo el cálculo son fórmulas vivas: si se cambia el radix, el año de la generación o un dato de entrada, las tablas, los indicadores y los gráficos se recalculan. El libro incluye además una hoja de comparación entre generaciones y una de validación contra la esperanza de vida oficial de la ONU."),

  H1("2. Datos"),
  P("Se descargaron los archivos CSV de la ONU (World Population Prospects 2024, variante media) y se conservó solo México: población a mitad de año y defunciones, ambas por edad simple (0 a 100+) y sexo, de 1950 a 2100, en miles de personas. De 1950 a 2023 son estimaciones de la ONU y de 2024 a 2100 son proyecciones. Para validar se usaron también la esperanza de vida al nacer oficial (LEx) y la qx de las tablas de vida completas de la ONU. Los CSV filtrados están en la carpeta datos/."),

  H1("3. Método"),
  P("Una generación es el conjunto de personas nacidas en el año g, y se observa a la edad x en el año calendario g + x. Para cada edad se toman la población a mitad de año y las defunciones de esa edad en ese año, y se aplican estas fórmulas:"),
  formula("qx = defunciones / (población a mitad de año + defunciones / 2)"),
  formula("l0 = 100,000;   l(x+1) = lx − dx;   dx = lx · qx"),
  formula("Lx = l(x+1) + dx / 2"),
  formula("Tx = suma de Lx desde x hasta 100;   ex = Tx / lx"),
  P("La edad 100 es un grupo abierto (100 y más) con qx = 1, de modo que l101 = 0 y L100 = l100 / 2. Cada sexo se calcula con su propia población y sus propias defunciones; ambos sexos no es un promedio de las otras dos tablas. En el libro, cada hoja Gen_XXXX trae las ocho columnas (población, defunciones, qx, lx, dx, Lx, Tx, ex) para los tres sexos, y una celda de control comprueba que la suma de dx sea igual al radix.", { spacing: { before: 80, after: 120, line: 276 } }),

  H1("4. Resultados"),
  P(`La esperanza de vida al nacer (ambos sexos) pasa de ${f1(g0.e0)} años en la generación 1950 a ${f1(g5.e0)} en la generación 2000. En hombres sube de ${f1(hom[0].e0)} a ${f1(hom[5].e0)} y en mujeres de ${f1(muj[0].e0)} a ${f1(muj[5].e0)}; la ventaja femenina se mantiene en torno a ${f1(muj[0].e0 - hom[0].e0)}-${f1(muj[5].e0 - hom[5].e0)} años. La proporción que llega con vida a los 65 años (ambos sexos) sube de ${pct(g0.l65)} a ${pct(g5.l65)}.`),
  tabla(["Generación", "e0 H", "e0 M", "e0 Ambos", "e65 Ambos", "l65/l0 Ambos", "l80/l0 Ambos"],
    [1500, 1310, 1310, 1310, 1310, 1310, 1310],
    D.gens.map((g, i) => [g, f2(hom[i].e0), f2(muj[i].e0), f2(amb[i].e0), f2(amb[i].e65), pct(amb[i].l65), pct(amb[i].l80)])),
  P("", { spacing: { after: 60 } }),
  ...figura("lx_por_generacion.png", 624, 211, "Figura 1. Sobrevivientes lx por edad y generación (radix 100,000)."),
  ...figura("e0_por_anio_calendario.png", 520, 255, "Figura 2. e0 oficial de la ONU por año calendario (líneas) y e0 de cada generación (puntos) según estas tablas."),
  P(`La e0 de una generación es mayor que la e0 de periodo de su año de nacimiento (${f1(g0.e0)} frente a ${f1(g0.periodo)} en 1950, ambos sexos), porque la generación vive las mejoras de mortalidad que vienen después y la medida de periodo no. Esa brecha (entre ${f1(Math.min(...D.sexos.flatMap((s) => A(s).map((x) => x.brecha))))} y ${f1(Math.max(...D.sexos.flatMap((s) => A(s).map((x) => x.brecha))))} años) no es un error. También se ve el efecto de la pandemia de 2020-2021 como un escalón en lx de las generaciones 1950 y 1960 (edades 60-71) y como una caída en la e0 de periodo. En el libro, la hoja Comparacion reúne e0, e65, l65/l0 y l80/l0, y los gráficos de ex y lx separados por sexo (las demás figuras están en resultados/).`),

  H1("5. Validación contra la ONU"),
  P([run("A. Método. ", { bold: true }), run("Se aplicaron las mismas fórmulas a cada año calendario (tablas de periodo, hoja Periodo) y se compararon con la e0 oficial de la ONU, de 1950 a 2100:")]),
  tabla(["Sexo", "Diferencia media (años)", "Error absoluto medio", "Máx. |dif.| (año)", "Dif. 1950", "Dif. 2023", "Dif. 2100"],
    [1500, 1460, 1400, 1500, 1160, 1160, 1180],
    D.sexos.map((s) => [s, sg(per[s].media), f2(per[s].mae), `${f2(per[s].max)} (${per[s].anio_max})`, sg(per[s].d1950), sg(per[s].d2023), sg(per[s].d2100)])),
  P("", { spacing: { after: 60 } }),
  P(`El método reproduce la e0 oficial con un error absoluto medio de ${f2(per["Ambos sexos"].mae)} años y un máximo de ${f2(Math.max(...D.periodo.map((p) => p.max)))}. En promedio la diferencia es negativa (−0.07 a −0.08 años), con dos causas identificadas. En los años iniciales, la qx de la edad 0 calculada con población a mitad de año sobrestima algo la mortalidad infantil (en 1950, 0.163 frente a 0.157 oficial; desde la edad 1 la qx coincide con la oficial), y el efecto desaparece hacia 2023. En los años finales, el grupo abierto con qx = 1 asigna solo 0.5 años de vida a quien llega a los 100, mientras la ONU estima entre 1.3 y 3.0, y cada vez llega más gente a esa edad. Entre 1981 y 2008 la diferencia oscila alrededor de cero (hasta ±0.07) y no se investigó su origen.`),
  P([run("B. Generaciones. ", { bold: true }), run(`Como la e0 oficial es de periodo, no se puede comparar directamente con una generación. Por eso se construyó, para cada generación, una referencia con la qx oficial de la ONU a lo largo de la misma diagonal (edad x en el año g + x) y las mismas fórmulas de lx y Lx. Las tablas de este libro quedan entre ${f2(Math.abs(rangoDif[1]))} y ${f2(Math.abs(rangoDif[0]))} años por debajo de esa referencia, y la diferencia se reduce de ${f2(Math.abs(amb[0].dif))} años en la generación 1950 a ${f2(Math.abs(amb[5].dif))} en la 2000 (ambos sexos). Ningún caso supera 0.4 años, es decir, 0.7% de la e0.`)]),
  tabla(["Generación", "H: libro", "H: ONU", "M: libro", "M: ONU", "Ambos: libro", "Ambos: ONU", "Dif. ambos"],
    [1260, 1100, 1100, 1100, 1100, 1250, 1250, 1200],
    D.gens.map((g, i) => [g, f2(hom[i].libro), f2(hom[i].onu), f2(muj[i].libro), f2(muj[i].onu), f2(amb[i].libro), f2(amb[i].onu), sg(amb[i].dif)])),
  P("", { spacing: { after: 60 } }),
  ...figura("validacion_e0.png", 624, 213, "Figura 3. Diferencia entre la e0 calculada y la oficial: (A) por año calendario; (B) por generación."),
  P(`El grupo abierto reduce la e0 de las generaciones recientes hasta ${f2(maxAdj)} años (con e100 = 2.0, hoja Validacion), pero afecta por igual a las tablas y a la referencia de la ONU, así que no cambia las diferencias de la tabla anterior.`),

  H1("6. Limitaciones"),
  bullet("Las defunciones y la población son por edad cumplida y año calendario, no por triángulos de Lexis. La qx de la generación se aproxima con la población de edad x a mitad del año g + x, lo que explica la diferencia residual frente a la ONU."),
  bullet(`Toda generación incluye años proyectados (desde 2024): ${proy(1950)} de 101 edades en la generación 1950 y ${proy(2000)} de 101 en la 2000. La e0 de las generaciones 1980-2000 depende sobre todo de la variante media de la ONU.`),
  bullet("Los datos mexicanos de la ONU son estimaciones y modelos, no un registro completo de cada defunción; el libro no corrige ese origen."),
  bullet("La e0 de periodo oficial y la de generación miden cosas distintas; solo la comparación contra la diagonal ONU (sección 5B) valida las tablas de generación."),

  H1("7. Archivos"),
  bullet("resultados/tablas_mortalidad_generaciones_mexico.xlsx: libro con las 6 generaciones, comparación y validación (116,760 fórmulas, sin errores)."),
  bullet("resultados/*.png: ex y lx por generación, e0 por año calendario y dos gráficos de validación."),
  bullet("datos/: CSV de México extraídos de la ONU. scripts/: descarga, construcción del libro, gráficos, verificación independiente (coincide con un cálculo en numpy a 1e-14) e informe."),
];

const doc = new Document({
  creator: "Claude Code",
  title: "Tablas de mortalidad de generaciones mexicanas",
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1300, right: 1440, bottom: 1300, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Página ", PageNumber.CURRENT], font: FONT, size: 16, color: "808080" })] })] }) },
    children: hijos,
  }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync(outPath, b));
