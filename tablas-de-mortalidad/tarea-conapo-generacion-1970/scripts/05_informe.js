// Documento Word: tabla de mortalidad de la generación mexicana de 1970 (CONAPO): datos, método, resultados y análisis.
// Uso: node scripts/05_informe.js datos.json carpeta_resultados salida.docx
const fs = require("fs");
const path = require("path");
const { Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, HeadingLevel,
  BorderStyle, LevelFormat, Footer, PageNumber } = require("docx");

const [, , jsonPath, resDir, outPath] = process.argv;
const D = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
const f = (x, n = 1) => Number(x).toFixed(n);
const n0 = (x) => Math.round(x).toLocaleString("en-US");
const pct = (x) => (100 * x).toFixed(1) + " %";

const FONT = "Arial", W = 9360;
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const P = (c, o = {}) => new Paragraph({ spacing: { after: 120, line: 276 }, ...o, children: Array.isArray(c) ? c : [run(c)] });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 260, after: 110 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 25, color: "1F3864" })] });
const bullet = (c) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 60, line: 264 }, children: Array.isArray(c) ? c : [run(c)] });
const formula = (t) => new Paragraph({ indent: { left: 540 }, spacing: { after: 50 }, children: [new TextRun({ text: t, font: "Consolas", size: 20 })] });
const borde = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" }, bordes = { top: borde, bottom: borde, left: borde, right: borde };
const celda = (t, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: bordes,
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, margins: { top: 45, bottom: 45, left: 80, right: 80 },
  children: [new Paragraph({ keepNext: true, alignment: o.left ? AlignmentType.LEFT : AlignmentType.CENTER, children: [new TextRun({ text: String(t), font: FONT, size: 18, bold: !!o.bold })] })] });
const tabla = (cols, anchos, filas) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: anchos,
  rows: [new TableRow({ tableHeader: true, children: cols.map((c, i) => celda(c, anchos[i], { fill: "DCE3F0", bold: true })) }),
    ...filas.map((fila) => new TableRow({ children: fila.map((c, i) => celda(c, anchos[i], { left: i === 0, bold: i === 0 })) }))] });
const figura = (nombre, ancho, alto, pie) => [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 30 }, keepNext: true,
    children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(resDir, nombre)), transformation: { width: ancho, height: alto }, altText: { title: nombre, description: pie, name: nombre } })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 150 }, children: [run(pie, { size: 17, italics: true, color: "595959" })] })];

const hijos = [
  new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: `Tabla de mortalidad de la generación mexicana de ${D.G}`, font: FONT, bold: true, size: 34, color: "1F3864" })] }),
  new Paragraph({ spacing: { after: 180 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "1F3864", space: 4 } },
    children: [run("U2_P1 · Tabla de mortalidad. Ambos sexos, con datos de CONAPO.", { color: "595959" })] }),

  H1("1. Datos"),
  P("La tabla se calculó con las bases de datos abiertas de CONAPO (Conciliación demográfica de México 1950-2019 y Proyecciones de la población 2020-2070), descargadas de datos.gob.mx y guardadas sin modificar en la carpeta bases_de_datos_CONAPO/:"),
  bullet([run("Defunciones por edad simple y año ", { bold: true }), run("(01_Defunciones_1950_2070.csv).")]),
  bullet([run("Población a mitad de año por edad simple y año ", { bold: true }), run("(00_Pob_Mitad_1950_2070.csv).")]),
  bullet([run("Indicadores demográficos ", { bold: true }), run("(05_indicadores_demograficos_proyecciones.csv): nacimientos, defunciones totales y esperanza de vida oficial.")]),
  P("CONAPO publica defunciones y población por entidad, así que los datos nacionales son la suma de las 32 entidades y de hombres y mujeres. Esa suma coincide exactamente con los totales nacionales del archivo de indicadores. Las defunciones por edad empiezan en 1970, por eso la generación de 1970 es la más antigua que se puede seguir de los 0 a los 100 años (hasta 2070).", { spacing: { before: 80, after: 120, line: 276 } }),

  H1("2. Método"),
  P("Una generación es el conjunto de personas nacidas en el mismo año. Se sigue a la generación de 1970 edad por edad: a la edad x le corresponde el año calendario 1970 + x. Con esos datos se llena la plantilla de clase:"),
  formula("q(0) = defunciones de menores de 1 año / nacimientos"),
  formula("q(x) = defunciones / (población a mitad de año + defunciones / 2),  x = 1, …, 99"),
  formula("q(100) = 1   (100 y más es el grupo abierto)"),
  formula("l0 = 100,000;  dx = lx · qx;  l(x+1) = lx − dx"),
  formula("Lx = l(x+1) + dx / 2;  Tx = suma de Lx desde x;  ex = Tx / lx"),
  P("Como en la plantilla, dx y Lx se redondean a enteros. Se usa la población a mitad de año en lugar de los nacimientos acumulados porque así la migración internacional ya está incluida en el denominador: la población expuesta al riesgo es la que vive en México.", { spacing: { before: 80, after: 120, line: 276 } }),

  H1("3. Resultados"),
  P(`La esperanza de vida al nacer de la generación de ${D.G} es de ${f(D.e0)} años, y a los 65 años les quedan ${f(D.e65)} años más. De cada 100,000 nacidos, ${n0(D.l65 * 1e5)} llegan a los 65 años y ${n0(D.l80 * 1e5)} a los 80.`),
  tabla(["Indicador", "Valor"], [6360, 3000], [
    ["e0: esperanza de vida al nacer (años)", f(D.e0, 2)], ["e1: esperanza de vida a los 1 año (años)", f(D.e1, 2)],
    ["e65: esperanza de vida a los 65 años (años)", f(D.e65, 2)], ["q(0): mortalidad infantil (por mil)", f(D.q0_mil, 1)],
    ["5q0: probabilidad de morir antes de los 5 años (por mil)", f(D["5q0_mil"], 1)],
    ["Probabilidad de llegar con vida a los 65 años", pct(D.l65)], ["Probabilidad de llegar con vida a los 80 años", pct(D.l80)],
    ["Edad en la que muere la mitad de la generación (años cumplidos)", String(D.mediana)],
  ]),
  P("", { spacing: { after: 60 } }),
  P("Extracto de la tabla (la tabla completa de 0 a 100 años está en el Excel):", { keepNext: true }),
  tabla(["x", "qx", "lx", "dx", "Lx", "Tx", "ex"], [900, 1500, 1400, 1400, 1400, 1560, 1200],
    D.extracto.map((r) => [r[0] === 100 ? "100+" : r[0], r[1].toFixed(6), n0(r[2]), n0(r[3]), n0(r[4]), n0(r[5]), f(r[6], 2)])),
  P("", { spacing: { after: 80 } }),
  ...figura("lx_sobrevivientes.png", 470, 261, "Figura 1. Sobrevivientes lx de la generación de 1970."),
  ...figura("qx_log.png", 470, 261, "Figura 2. Probabilidad de morir qx por edad (escala logarítmica)."),

  H1("4. Análisis"),
  P([run("Primer año de vida. ", { bold: true }), run(`De 100,000 nacidos mueren ${n0(D.q0_mil * 100)} antes de cumplir un año (${f(D.q0_mil)} por mil). Es el riesgo más alto de la vida hasta los ${D.edad_supera_q0 - 1} años, y por eso la esperanza de vida a los 1 año (${f(D.e1)}) es mayor que al nacer (${f(D.e0)}): quien sobrevive el primer año gana casi cinco años de esperanza de vida. El riesgo baja rápido y llega a su mínimo a los ${D.edad_q_min} años (qx = ${D.q_min.toFixed(5)}).`)]),
  P([run("Vida adulta. ", { bold: true }), run(`A partir de la adolescencia qx crece de forma casi exponencial: ${D.q50.toFixed(4)} a los 50 años, ${D.q80.toFixed(3)} a los 80 y ${D.q99.toFixed(3)} a los 99. La mitad de la generación muere antes de los ${D.mediana + 1} años y las defunciones adultas se concentran alrededor de los ${D.moda} años.`)]),
  P([run("La pandemia. ", { bold: true }), run(`A los 50 y 51 años (años calendario 2020 y 2021) qx se dispara a ${D.q50.toFixed(4)} y vuelve a bajar a ${D.q52.toFixed(4)} a los 52. Es el efecto del COVID-19, que CONAPO incorpora en su reconstrucción, y se ve también en la esperanza de vida de periodo (Figura 3).`)]),
  P([run("Generación frente a periodo. ", { bold: true }), run(`La e0 de esta generación (${f(D.e0)}) es mayor que la e0 oficial de CONAPO para 1970 (${f(D.ev_1970)}) y menor que la de 2019 (${f(D.ev_2019)}): la generación nació con la mortalidad infantil de 1970, pero después vivió las mejoras de las décadas siguientes. Una e0 de periodo mide la mortalidad de un solo año; una de generación sigue a las mismas personas.`)]),
  ...figura("validacion_e0_periodo.png", 470, 261, "Figura 3. e0 de periodo con las fórmulas de la tabla frente a la e0 oficial de CONAPO."),
  P([run("Validación. ", { bold: true }), run(`Aplicando las mismas fórmulas a cada año calendario (1970-2070), la e0 que sale difiere de la oficial de CONAPO en ${f(D.mae, 2)} años en promedio y ${f(D.maxdif, 2)} años como máximo. Es una comprobación de que las fórmulas y los datos están bien armados.`)]),

  H1("5. Limitaciones"),
  bullet(`Solo las edades 0 a 49 de la generación (años 1970-2019) son cifras reconstruidas por CONAPO; de los 50 a los 100 años (2020-2070) son proyecciones, y la e0 depende de ellas.`),
  bullet(`Defunciones y población son por edad y año calendario, no por generación exacta: a cada edad x se le asignan los datos del año 1970 + x. Es la aproximación habitual cuando no hay datos por cohorte.`),
  bullet(`CONAPO proyecta una mortalidad baja en edades muy avanzadas: ${pct(D.l100)} de la generación llegaría a los 100 años. Con qx = 1 en el grupo abierto, e100 = 0.5 años, así que e0 queda ligeramente subestimada.`),
  bullet("Los datos de CONAPO son estimaciones demográficas conciliadas, no registros administrativos directos."),
  P("Archivos: T_Mortalidad_CONAPO_Generacion_1970.xlsx (tabla de la plantilla y hojas con los datos usados), los gráficos en PNG y las bases originales en bases_de_datos_CONAPO/.", { spacing: { before: 120 } }),
];

const doc = new Document({ creator: "Tarea U2_P1", title: `Tabla de mortalidad de la generación mexicana de ${D.G} (CONAPO)`,
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1300, right: 1440, bottom: 1300, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Página ", PageNumber.CURRENT], font: FONT, size: 16, color: "808080" })] })] }) },
    children: hijos }] });
Packer.toBuffer(doc).then((b) => fs.writeFileSync(outPath, b));
