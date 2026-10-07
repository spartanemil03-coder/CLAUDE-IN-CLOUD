// Documento Word: tabla de mortalidad de México 2019 con datos de INEGI (método, resultados y análisis).
// Uso: node scripts/05_informe.js datos.json carpeta_resultados salida.docx
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell, WidthType, ShadingType,
  AlignmentType, HeadingLevel, BorderStyle, LevelFormat, Footer, PageNumber,
} = require("docx");

const [, , jsonPath, resDir, outPath] = process.argv;
const D = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
const f = (x, n = 2) => Number(x).toFixed(n);
const n0 = (x) => Math.round(x).toLocaleString("en-US");
const pct = (x) => (100 * x).toFixed(1) + "%";
const I = (k) => D.ind[Object.keys(D.ind).find((n) => n.startsWith(k))]; // [H, M, A]

const FONT = "Arial";
const W = 9360;
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const P = (c, o = {}) => new Paragraph({ spacing: { after: 120, line: 276 }, ...o, children: Array.isArray(c) ? c : [run(c)] });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 260, after: 110 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 25, color: "1F3864" })] });
const bullet = (c) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 50, line: 264 }, children: Array.isArray(c) ? c : [run(c)] });
const formula = (t) => new Paragraph({ indent: { left: 540 }, spacing: { after: 50 }, children: [new TextRun({ text: t, font: "Consolas", size: 20 })] });

const borde = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" };
const bordes = { top: borde, bottom: borde, left: borde, right: borde };
const celda = (t, w, o = {}) => new TableCell({
  width: { size: w, type: WidthType.DXA }, borders: bordes,
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined,
  margins: { top: 45, bottom: 45, left: 80, right: 80 },
  children: [new Paragraph({ keepNext: true, alignment: o.left ? AlignmentType.LEFT : AlignmentType.CENTER, children: [new TextRun({ text: String(t), font: FONT, size: 18, bold: !!o.bold })] })],
});
const tabla = (cols, anchos, filas) => new Table({
  width: { size: W, type: WidthType.DXA }, columnWidths: anchos,
  rows: [new TableRow({ tableHeader: true, children: cols.map((c, i) => celda(c, anchos[i], { fill: "DCE3F0", bold: true })) }),
    ...filas.map((fila) => new TableRow({ children: fila.map((c, i) => celda(c, anchos[i], { left: i === 0, bold: i === 0 })) }))],
});
const figura = (nombre, ancho, alto, pie) => [
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 100, after: 30 }, keepNext: true,
    children: [new ImageRun({ type: "png", data: fs.readFileSync(path.join(resDir, nombre)), transformation: { width: ancho, height: alto },
      altText: { title: nombre, description: pie, name: nombre } })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 150 }, children: [run(pie, { size: 17, italics: true, color: "595959" })] }),
];

const e0 = I("e0"), e65 = I("e65"), q0 = I("q(0)"), l65 = I("Probabilidad de llegar con vida a los 65"), l80 = I("Probabilidad de llegar con vida a los 80");
const med = I("Edad mediana"), e5q0 = I("5q0");

const hijos = [
  new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: "Tabla de mortalidad de México, 2019", font: FONT, bold: true, size: 36, color: "1F3864" })] }),
  new Paragraph({ spacing: { after: 180 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "1F3864", space: 4 } },
    children: [run("U2_P1 · Tabla de mortalidad INEGI. Cálculo con cifras de defunciones y nacimientos de INEGI.", { color: "595959" })] }),

  H1("1. Datos"),
  P("La tabla se calculó con datos abiertos de INEGI. El año de referencia es 2019, el último año completo antes de la pandemia de COVID-19, que elevó mucho la mortalidad en 2020 y 2021."),
  bullet([run("Defunciones: ", { bold: true }), run("Estadísticas de Defunciones Registradas (EDR), conjuntos de datos abiertos 2019, 2020 y 2021. Se cuentan las defunciones ocurridas en 2019 (las registradas ese año y las registradas tarde en 2020 y 2021): 744,479, con la edad al morir.")]),
  bullet([run("Nacimientos: ", { bold: true }), run(`Estadística de Nacimientos Registrados, datos abiertos 2019, 2020 y 2021. Nacimientos ocurridos en 2019: ${n0(D.nacimientos)}.`)]),
  bullet([run("Población por edad: ", { bold: true }), run("Censo de Población y Vivienda 2020, tabulado de población por edad desplegada y sexo (INEGI). Es el denominador de qx a partir del año de edad.")]),

  H1("2. Método"),
  P("La tabla es de periodo (generación ficticia): se aplica a 100,000 recién nacidos la mortalidad por edad observada en 2019. Se sigue el formato de la plantilla de clase: x, qx, lx, dx, Lx, Tx y ex."),
  formula("q(0) = defunciones de menores de 1 año / nacimientos"),
  formula("q(x) = defunciones / (población + defunciones / 2),   x = 1, …, 99"),
  formula("q(100) = 1   (100 y más es el grupo abierto)"),
  formula("l0 = 100,000;   l(x+1) = lx − dx;   dx = lx · qx"),
  formula("Lx = l(x+1) + dx / 2;   Tx = suma de Lx desde x;   ex = Tx / lx"),
  P("Dos ajustes de los datos, ambos con interruptor o a la vista en el Excel: (a) las defunciones con edad no especificada (0.5 %) se reparten entre las edades en proporción; (b) el Censo concentra población en edades redondas (a los 30 años hay 2.37 millones de personas y a los 31 solo 1.54 millones), lo que haría zigzaguear qx, así que de 10 a 97 años se usó una media móvil de 5 edades en defunciones y población. El suavizado cambia e0 solo 0.3 años.", { spacing: { before: 80, after: 120, line: 276 } }),

  H1("3. Resultados"),
  P(`La esperanza de vida al nacer es de ${f(e0[2], 1)} años (hombres ${f(e0[0], 1)}; mujeres ${f(e0[1], 1)}). La mortalidad infantil es de ${f(q0[2], 1)} por mil nacimientos. De cada 100,000 nacidos, ${n0(l65[2] * 100000)} llegan a los 65 años y ${n0(l80[2] * 100000)} a los 80.`),
  tabla(["Indicador", "Hombres", "Mujeres", "Ambos sexos"], [4200, 1720, 1720, 1720], [
    ["e0: esperanza de vida al nacer (años)", ...e0.map((v) => f(v))],
    ["e65: esperanza de vida a los 65 años", ...e65.map((v) => f(v))],
    ["q(0): mortalidad infantil (por mil)", ...q0.map((v) => f(v))],
    ["5q0: morir antes de los 5 años (por mil)", ...e5q0.map((v) => f(v))],
    ["Llegar con vida a los 65 años", ...l65.map(pct)],
    ["Llegar con vida a los 80 años", ...l80.map(pct)],
    ["Edad mediana a la muerte (años)", ...med.map((v) => String(v))],
  ]),
  P("", { spacing: { after: 60 } }),
  P("Extracto de la tabla de mortalidad, ambos sexos (la tabla completa de 0 a 100 años, y las de hombres y mujeres, están en el Excel):", { keepNext: true }),
  tabla(["x", "qx", "lx", "dx", "Lx", "Tx", "ex"], [900, 1500, 1400, 1400, 1400, 1560, 1200],
    D.extracto.map((r) => [r[0] === 100 ? "100+" : r[0], r[1].toFixed(6), n0(r[2]), n0(r[3]), n0(r[4]), n0(r[5]), f(r[6])])),
  P("", { spacing: { after: 80 } }),
  ...figura("lx_sobrevivientes.png", 470, 261, "Figura 1. Sobrevivientes lx de 100,000 nacidos, por sexo."),
  ...figura("qx_log.png", 470, 261, "Figura 2. Probabilidad de morir qx por edad (escala logarítmica)."),

  H1("4. Análisis"),
  ...D.analisis.slice(0, 4).flatMap(([t, txt]) => [P([run(t + ". ", { bold: true }), run(txt)], { keepLines: true })]),
  ...figura("sobremortalidad_masculina.png", 470, 261, "Figura 3. Cociente entre la probabilidad de morir de hombres y de mujeres."),
  ...figura("e0_comparacion_onu.png", 470, 261, "Figura 4. e0 con datos de INEGI frente a la estimación de la ONU para 2019."),

];
hijos.push(H1("5. Limitaciones"));
// la sección de limitaciones se escribe como viñetas, separando (i) … (iv)
D.analisis[4][1].split(/\s*\((?:i|ii|iii|iv)\)\s*/).filter((t) => t.trim()).forEach((t) => hijos.push(bullet(t.trim())));
hijos.push(P("Archivos: T_Mortalidad_INEGI_Mexico_2019.xlsx (tablas con fórmulas, indicadores y gráficos) y los gráficos en PNG.", { spacing: { before: 120 } }));

const doc = new Document({
  creator: "Tarea U2_P1", title: "Tabla de mortalidad de México, 2019 (INEGI)",
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1300, right: 1440, bottom: 1300, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Página ", PageNumber.CURRENT], font: FONT, size: 16, color: "808080" })] })] }) },
    children: hijos,
  }],
});
Packer.toBuffer(doc).then((b) => fs.writeFileSync(outPath, b));
