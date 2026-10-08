// Documento Word: tabla de mortalidad de México 2019 con datos de INEGI (datos, método, resultados, análisis y limitaciones).
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
  new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: "Tabla de mortalidad de México, 2019", font: FONT, bold: true, size: 36, color: "1F3864" })] }),
  new Paragraph({ spacing: { after: 180 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "1F3864", space: 4 } },
    children: [run("U2_P1 · Tabla de mortalidad INEGI. Ambos sexos, con cifras de defunciones y nacimientos de INEGI.", { color: "595959" })] }),

  H1("1. Datos"),
  P("La tabla se calculó con bases de datos abiertas de INEGI, guardadas sin modificar en la carpeta bases_de_datos_INEGI/. El año de referencia es 2019, el último año completo antes de la pandemia de COVID-19, que elevó mucho la mortalidad en 2020 y 2021."),
  bullet([run("Defunciones: ", { bold: true }), run(`Estadísticas de Defunciones Registradas (EDR), datos abiertos 2019, 2020 y 2021. Se cuentan las defunciones ocurridas en 2019 (registradas ese año o después): ${n0(D.def_arch[0])} + ${n0(D.def_arch[1])} + ${n0(D.def_arch[2])} = ${n0(D.def_total)}.`)]),
  bullet([run("Nacimientos: ", { bold: true }), run(`Estadística de Nacimientos Registrados (ENR), datos abiertos 2019, 2020 y 2021. Nacimientos ocurridos en 2019: ${n0(D.nac_arch[0])} + ${n0(D.nac_arch[1])} + ${n0(D.nac_arch[2])} = ${n0(D.nacimientos)}. Se usan los tres años porque muchos nacimientos se registran con retraso.`)]),
  bullet([run("Población por edad: ", { bold: true }), run("Censo de Población y Vivienda 2020 (INEGI), población por edad simple. Es el denominador de qx a partir del año de edad, porque con solo nacimientos y defunciones no se puede calcular el riesgo de morir de las demás edades.")]),

  H1("2. Método"),
  P("La tabla es de periodo (generación ficticia): se aplica a 100,000 recién nacidos la mortalidad por edad observada en 2019. Sigue el formato de la plantilla de clase: x, qx, lx, dx, Lx, Tx y ex."),
  formula("q(0) = defunciones de menores de 1 año / nacimientos"),
  formula("q(x) = defunciones / (población + defunciones / 2),   x = 1, …, 99"),
  formula("q(100) = 1   (100 y más es el grupo abierto)"),
  formula("l0 = 100,000;   dx = lx · qx;   l(x+1) = lx − dx"),
  formula("Lx = l(x+1) + dx / 2;   Tx = suma de Lx desde x;   ex = Tx / lx"),
  P("Dos ajustes a los datos, visibles en el libro de preparación: (a) las defunciones con edad no especificada (0.5 %) se reparten entre las edades en proporción; (b) el Censo concentra población en edades redondas (a los 30 años hay 2.37 millones de personas y a los 31 solo 1.54 millones), lo que haría zigzaguear qx, así que de 10 a 97 años se usa una media móvil de 5 edades en defunciones y población (interruptor en la hoja Preparación). El suavizado cambia e0 en 0.3 años. Como en la plantilla, dx y Lx se redondean a enteros.", { spacing: { before: 80, after: 120, line: 276 } }),

  H1("3. Resultados"),
  P(`La esperanza de vida al nacer es de ${f(D.e0)} años (hombres ${f(D.e0_h)}; mujeres ${f(D.e0_m)}) y a los 65 años les quedan ${f(D.e65)} años más. La mortalidad infantil es de ${f(D.q0_mil)} por mil. De cada 100,000 nacidos, ${n0(D.l65 * 1e5)} llegan a los 65 años y ${n0(D.l80 * 1e5)} a los 80.`),
  tabla(["Indicador", "Valor"], [6360, 3000], [
    ["e0: esperanza de vida al nacer (años)", f(D.e0, 2)], ["e1: esperanza de vida a 1 año (años)", f(D.e1, 2)],
    ["e65: esperanza de vida a los 65 años (años)", f(D.e65, 2)], ["q(0): mortalidad infantil (por mil)", f(D.q0_mil, 1)],
    ["5q0: probabilidad de morir antes de los 5 años (por mil)", f(D["5q0_mil"], 1)],
    ["Probabilidad de llegar con vida a los 65 años", pct(D.l65)], ["Probabilidad de llegar con vida a los 80 años", pct(D.l80)],
    ["Edad hasta la que sobrevive al menos la mitad (años cumplidos)", String(D.mediana)],
  ]),
  P("", { spacing: { after: 60 } }),
  P("Extracto de la tabla (la tabla completa de 0 a 100 años está en el Excel):", { keepNext: true }),
  tabla(["x", "qx", "lx", "dx", "Lx", "Tx", "ex"], [900, 1500, 1400, 1400, 1400, 1560, 1200],
    D.extracto.map((r) => [r[0] === 100 ? "100+" : r[0], r[1].toFixed(6), n0(r[2]), n0(r[3]), n0(r[4]), n0(r[5]), f(r[6], 2)])),
  P("", { spacing: { after: 80 } }),
  ...figura("lx_sobrevivientes.png", 470, 261, "Figura 1. Sobrevivientes lx de una generación ficticia de 100,000 nacidos."),
  ...figura("qx_log.png", 470, 261, "Figura 2. Probabilidad de morir qx por edad (escala logarítmica)."),
  ...figura("dx_defunciones.png", 470, 261, "Figura 3. Defunciones dx de la tabla, por edad."),

  H1("4. Análisis"),
  P([run("Primer año de vida. ", { bold: true }), run(`De 100,000 nacidos mueren ${n0(D.q0_mil * 100)} antes de cumplir un año (${f(D.q0_mil)} por mil), el riesgo más alto de la vida hasta los ${D.edad_supera_q0 - 1} años. Aun así, la esperanza de vida casi no cambia entre el nacimiento (${f(D.e0, 2)}) y el primer año (${f(D.e1, 2)}), señal de que la mortalidad infantil ya es relativamente baja. El riesgo cae rápido y llega a su mínimo a los ${D.edad_q_min} años (qx = ${D.q_min.toFixed(5)}).`)]),
  P([run("Vida adulta. ", { bold: true }), run(`Desde la adolescencia qx crece de forma casi exponencial: ${D.q50.toFixed(4)} a los 50 años y ${D.q80.toFixed(3)} a los 80. Al menos la mitad de la generación sobrevive hasta los ${D.mediana} años y las defunciones adultas se concentran alrededor de los ${D.moda} años (Figura 3).`)]),
  P([run("Diferencias por sexo. ", { bold: true }), run(`Con las mismas fórmulas aplicadas por separado a hombres y mujeres, la e0 es ${f(D.e0_h)} y ${f(D.e0_m)} años: las mujeres viven ${f(D.e0_m - D.e0_h)} años más. Entre los 15 y los 39 años la probabilidad de morir de los hombres es en promedio ${f(D.razon_media)} veces la de las mujeres, con un máximo de ${f(D.razon_max)} veces a los ${D.razon_edad} años (Figura 4): en esas edades pesan las muertes violentas y por accidentes.`)]),
  ...figura("sobremortalidad_masculina.png", 470, 261, "Figura 4. Cociente entre la probabilidad de morir de hombres y de mujeres."),
  P([run("Comparación con CONAPO. ", { bold: true }), run(`Como referencia externa (no se usa en la tabla), CONAPO estima para 2019 una e0 de ${f(D.ref_a)} años (hombres ${f(D.ref_h)}; mujeres ${f(D.ref_m)}). La tabla con cifras de INEGI queda ${f(D.e0_a - D.ref_a)} años por encima (Figura 5). Es la dirección esperada si las defunciones registradas están algo subregistradas y si el Censo exagera la población de edades avanzadas, aunque con estas bases no puedo cuantificar cuánto aporta cada causa.`)]),
  ...figura("e0_comparacion_conapo.png", 470, 261, "Figura 5. e0 con datos de INEGI frente a la e0 oficial de CONAPO para 2019."),

  H1("5. Limitaciones"),
  bullet("Es una tabla de periodo (generación ficticia), no el seguimiento de una generación real: con las bases de INEGI las defunciones por edad solo existen desde 1990."),
  bullet(`El Censo sobrestima la población de edades avanzadas: a partir de ~90 años qx se aplana y ${pct(D.l100)} de los nacidos llegaría a los 100 años, lo cual no es creíble. Con qx = 1 en el grupo abierto, e100 = 0.5 años.`),
  bullet("Las defunciones y los nacimientos son registros administrativos: pueden tener subregistro, y los nacimientos de 2019 se siguen registrando después de 2021, por lo que q(0) puede estar algo sobrestimada."),
  bullet("Defunciones de 2019 y población de marzo de 2020 no son del mismo momento; la diferencia es de unos meses de crecimiento de la población."),
  P("Archivos: 2_Tabla_de_mortalidad_INEGI_2019.xlsx (tabla de la plantilla e insumos) y 1_Preparacion_datos_INEGI_2019.xlsx (datos de INEGI y su preparación), los gráficos en PNG y las bases originales de INEGI en bases_de_datos_INEGI/.", { spacing: { before: 120 } }),
];

const doc = new Document({ creator: "Tarea U2_P1", title: "Tabla de mortalidad de México, 2019 (INEGI)",
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1300, right: 1440, bottom: 1300, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Página ", PageNumber.CURRENT], font: FONT, size: 16, color: "808080" })] })] }) },
    children: hijos }] });
Packer.toBuffer(doc).then((b) => fs.writeFileSync(outPath, b));
