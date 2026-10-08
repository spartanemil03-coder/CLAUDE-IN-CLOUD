// Documento Word "Explicación paso a paso" de la tabla de mortalidad INEGI 2019 (lenguaje sencillo, con ejemplos numéricos).
// Uso: node scripts/06_explicacion.js datos.json salida.docx
const fs = require("fs");
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType, AlignmentType, HeadingLevel,
  BorderStyle, LevelFormat, Footer, PageNumber } = require("docx");
const [, , jsonPath, outPath] = process.argv;
const D = JSON.parse(fs.readFileSync(jsonPath, "utf8"));
const n0 = (x) => Math.round(x).toLocaleString("en-US");
const f = (x, n = 2) => Number(x).toFixed(n);
const FONT = "Arial", W = 9360;
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const P = (c, o = {}) => new Paragraph({ spacing: { after: 110, line: 276 }, ...o, children: Array.isArray(c) ? c : [run(c)] });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 110 }, keepNext: true, children: [new TextRun({ text: t, font: FONT, bold: true, size: 25, color: "1F3864" })] });
const bullet = (c) => new Paragraph({ numbering: { reference: "bul", level: 0 }, spacing: { after: 55, line: 264 }, children: Array.isArray(c) ? c : [run(c)] });
const b = (t) => run(t, { bold: true });
const mono = (t) => new Paragraph({ indent: { left: 540 }, spacing: { after: 45 }, children: [new TextRun({ text: t, font: "Consolas", size: 19 })] });
const borde = { style: BorderStyle.SINGLE, size: 4, color: "BFBFBF" }, bordes = { top: borde, bottom: borde, left: borde, right: borde };
const celda = (t, w, o = {}) => new TableCell({ width: { size: w, type: WidthType.DXA }, borders: bordes,
  shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: "auto" } : undefined, margins: { top: 45, bottom: 45, left: 80, right: 80 },
  children: [new Paragraph({ keepNext: true, alignment: o.left ? AlignmentType.LEFT : AlignmentType.CENTER, children: [new TextRun({ text: String(t), font: FONT, size: 18, bold: !!o.bold })] })] });
const tabla = (cols, anchos, filas) => new Table({ width: { size: W, type: WidthType.DXA }, columnWidths: anchos,
  rows: [new TableRow({ tableHeader: true, children: cols.map((c, i) => celda(c, anchos[i], { fill: "DCE3F0", bold: true })) }),
    ...filas.map((fila) => new TableRow({ children: fila.map((c, i) => celda(c, anchos[i], { left: i === 0 })) }))] });
const sp = () => P("", { spacing: { after: 50 } });

const hijos = [
  new Paragraph({ spacing: { after: 50 }, children: [new TextRun({ text: "Tabla de mortalidad de México, 2019: explicación paso a paso", font: FONT, bold: true, size: 34, color: "1F3864" })] }),
  new Paragraph({ spacing: { after: 160 }, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: "1F3864", space: 4 } },
    children: [run("U2_P1 · Tabla de mortalidad INEGI. Qué se hizo, en qué orden y por qué.", { color: "595959" })] }),

  H1("Resumen"),
  P(`Se construyó la tabla de mortalidad de México para 2019 (ambos sexos) con cifras de INEGI, siguiendo la plantilla de la profesora (columnas x, qx, lx, dx, Lx, Tx y ex, de 0 a 100 años). Se usaron tres bases de INEGI: defunciones (EDR), nacimientos (ENR) y población por edad (Censo 2020). Resultado: la esperanza de vida al nacer es de ${f(D.e0)} años y a los 65 años les quedan ${f(D.e65)} años más. Todo el Excel está hecho con fórmulas, para que se pueda revisar celda por celda.`),

  H1("1. Qué pide la tarea y qué es la plantilla"),
  bullet("La tarea pide: con cifras de INEGI sobre defunciones y nacimientos, generar la tabla de mortalidad, y hacer gráficos de la tabla como complemento del análisis."),
  bullet("La plantilla de la profesora (Cuadro 4.1, generación francesa de 1899) tiene siete columnas: x (edad), qx (probabilidad de morir entre x y x+1), lx (sobrevivientes), dx (defunciones), Lx (años vividos entre x y x+1), Tx (años vividos acumulados) y ex (esperanza de vida). Además trae una nota técnica: q(0) = defunciones / nacimientos."),
  bullet("Se parte de 100,000 recién nacidos (el radix) y se les aplica la mortalidad de cada edad."),

  H1("2. Idea general"),
  P("La tabla es de periodo: toma la mortalidad que se observó en un solo año (2019) y la aplica, edad por edad, a una generación imaginaria de 100,000 nacidos. Se eligió 2019 porque es el último año completo antes de la pandemia de COVID-19, que elevó mucho las defunciones en 2020 y 2021 y habría distorsionado la tabla."),
  P("Para calcular qx de cada edad hacen falta dos datos: cuántas personas murieron a esa edad y cuántas había expuestas al riesgo de morir. Por eso, además de defunciones y nacimientos, se necesita la población por edad."),

  H1("3. Paso 1: las bases de datos de INEGI"),
  tabla(["Base", "Archivos usados", "Qué se tomó de ella"], [2300, 3560, 3500], [
    ["Defunciones: EDR (Estadísticas de Defunciones Registradas), datos abiertos", "conjunto_de_datos_defunciones_registradas_2019, 2020 y 2021 (zip con CSV)", "Edad al morir, sexo y año de ocurrencia de cada defunción"],
    ["Nacimientos: ENR (Estadística de Nacimientos Registrados), datos abiertos", "conjunto_de_datos_natalidad_2019, 2020 y 2021 (zip con CSV)", "Sexo y año de nacimiento de cada nacimiento"],
    ["Población: Censo de Población y Vivienda 2020, tabulados del cuestionario básico", "cpv2020_b_eum_01_poblacion.xlsx (hoja 03)", "Población de México por edad (0 a 100 y más) y sexo"],
  ]),
  sp(),
  P([b("¿Por qué tres años de archivos de defunciones y de nacimientos? "), run(`Cada archivo de INEGI contiene los registros hechos en ese año, aunque el hecho haya ocurrido antes. Una defunción o un nacimiento de 2019 puede registrarse en 2019, 2020 o 2021. Por eso se abrieron los tres archivos y se contaron solo los hechos ocurridos en 2019. Ejemplo con nacimientos: ${n0(D.nac_arch[0])} se registraron en 2019, ${n0(D.nac_arch[1])} en 2020 y ${n0(D.nac_arch[2])} en 2021, en total ${n0(D.nacimientos)}.`)]),

  H1("4. Paso 2: leer los archivos y contar"),
  P("Cada archivo trae un renglón por defunción o por nacimiento (cientos de miles de renglones). Un programa en Python (scripts/01_descargar_inegi.py) los abre y hace lo siguiente:"),
  bullet([b("Defunciones: "), run("se queda con las del año de ocurrencia 2019. La edad viene codificada por INEGI: 1001 a 3999 significa menor de 1 año (horas, días o meses), 4001 a 4120 son años cumplidos (4001 = 1 año, 4065 = 65 años) y 4998 es edad no especificada. Se convierte a 0, 1, 2, …, 100 (100 = 100 y más) y se cuenta por edad y sexo.")]),
  bullet([b("Nacimientos: "), run("se queda con los nacidos en 2019 y se cuenta por sexo.")]),
  bullet([b("Población: "), run("se copia la población por edad y sexo del tabulado del Censo 2020.")]),
  P(`Totales de control: ${n0(D.def_total)} defunciones ocurridas en 2019 (${n0(D.def_arch[0])} + ${n0(D.def_arch[1])} + ${n0(D.def_arch[2])}), ${n0(D.nacimientos)} nacimientos y ${n0(D.pob_total)} personas en el Censo 2020. Estos conteos están en las hojas Defunciones, Nacimientos y Población del Excel (hay una celda de control que compara el total por archivo con el total por edad y debe dar 0).`),

  H1("5. Paso 3: dos ajustes a los datos"),
  P([b("Ajuste A: defunciones con edad no especificada. "), run(`${n0(D.ne)} defunciones (0.5 %) no tienen edad. Se reparten entre todas las edades en proporción a sus defunciones, multiplicando cada edad por un factor = total de defunciones / defunciones con edad = ${n0(D.def_total)} / ${n0(D.def_total - D.ne)} = ${f(D.factor, 4)}. Ejemplo: a la edad 0 había ${n0(D.d0_raw)} defunciones y quedan ${f(D.d0_aj, 1)}.`)]),
  P([b("Ajuste B: suavizado del Censo (opcional). "), run(`En el Censo mucha gente declara una edad “redonda”: hay ${n0(D.p30)} personas de 30 años pero solo ${n0(D.p31)} de 31. Eso hace que qx suba y baje sin sentido (qx a los 30 años saldría ${D.q30_crudo.toFixed(5)} y a los 31 saldría ${D.q31_crudo.toFixed(5)}). Para corregirlo, de los 10 a los 97 años se reemplaza cada dato por el promedio de 5 edades (la edad y dos hacia cada lado), tanto en defunciones como en población. Con el suavizado: ${D.q30_suav.toFixed(5)} y ${D.q31_suav.toFixed(5)}. En el Excel se apaga poniendo 0 en la celda amarilla “Suavizado del Censo” (L18); sin suavizar, e0 sería ${f(D.e0_crudo)} en vez de ${f(D.e0)}.`)]),

  H1("6. Paso 4: calcular qx"),
  P("La probabilidad de morir se calcula así:"),
  mono("q(0)     = defunciones de menores de 1 año / nacimientos"),
  mono("q(x)     = defunciones de la edad x / (población de la edad x + defunciones/2)   para x = 1, …, 99"),
  mono("q(100)   = 1    (a los 100 y más se cierra la tabla: todos mueren)"),
  sp(),
  P(`El “+ defunciones/2” del denominador es la forma usual de aproximar la población expuesta durante el año, porque quienes mueren dejan de estar expuestos a mitad de año en promedio.`),
  tabla(["Edad", "Defunciones usadas", "Población usada", "qx"], [1500, 2500, 2500, 2860], [
    ["0 (nacimientos)", `${f(D.d0_aj, 1)} (menores de 1 año)`, `${n0(D.nacimientos)} nacimientos`, `${f(D.d0_aj, 1)} / ${n0(D.nacimientos)} = ${f(D.q0, 6)}`],
    ["1", f(D.d1_aj, 1), n0(D.p1), `${f(D.d1_aj, 1)} / (${n0(D.p1)} + ${f(D.d1_aj / 2, 1)}) = ${f(D.q1, 6)}`],
    ["65", f(D.d65_us, 1), f(D.p65_us, 1), `${f(D.d65_us, 1)} / (${f(D.p65_us, 1)} + ${f(D.d65_us / 2, 1)}) = ${f(D.q65, 6)}`],
  ]),
  sp(),
  P("Por qué q(0) usa nacimientos: lo indica la nota técnica de la plantilla. Los que mueren antes de cumplir un año son del grupo de nacidos, y la población de “0 años” del Censo no representa bien a ese grupo."),

  H1("7. Paso 5: completar la tabla (lx, dx, Lx, Tx, ex)"),
  P("Con la columna qx lista, el resto sale con las mismas fórmulas de la plantilla. En el Excel (hoja Tabla de mortalidad, fila 5 = edad 0):"),
  tabla(["Columna", "Qué significa", "Fórmula en Excel", "Ejemplo (edad 0 → 1)"], [1300, 2300, 2700, 3060], [
    ["lx (D)", "Sobrevivientes a la edad x", "D5 = 100000; D6 = D5 − E5", `${n0(100000)} → ${n0(D.l1)}`],
    ["dx (E)", "Defunciones entre x y x+1", "E5 = ROUND(D5*C5, 0)", `${n0(100000)} × ${f(D.q0, 6)} = ${n0(D.dx0)}`],
    ["Lx (F)", "Años vividos entre x y x+1", "F5 = ROUND(D6 + (E5/2), 0)", `${n0(D.l1)} + ${n0(D.dx0)}/2 = ${n0(D.Lx0)}`],
    ["Tx (G)", "Años vividos de x en adelante", "G5 = SUMA(F5:F$105)", `${n0(D.Tx0)} (suma de todos los Lx)`],
    ["ex (H)", "Esperanza de vida a la edad x", "H5 = G5 / D5", `${n0(D.Tx0)} / ${n0(100000)} = ${f(D.e0)}`],
  ]),
  sp(),
  bullet("Los dx y Lx se redondean a enteros, igual que en la plantilla."),
  bullet("A los 100 años qx = 1, así que todos los que llegan mueren en esa fila; el último Lx es dx/2."),
  bullet("Control: la suma de todos los dx debe ser igual al radix (100,000). El Excel lo comprueba en la celda L25 y da 0."),
  bullet("En la plantilla original los Tx estaban escritos a mano y varios eran sumas de lx en lugar de Lx; aquí todos son fórmulas."),

  H1("8. Resultados"),
  tabla(["Indicador", "Valor"], [6360, 3000], [
    ["Esperanza de vida al nacer, e0 (años)", f(D.e0)], ["Esperanza de vida a los 65 años, e65 (años)", f(D.e65)],
    ["Mortalidad infantil, q(0) (por mil nacimientos)", f(D.q0 * 1000, 1)], ["Llegan con vida a los 65 años (de 100,000)", n0(D.l65)],
    ["Llegan con vida a los 80 años (de 100,000)", n0(D.l80)], ["Edad hasta la que sobrevive al menos la mitad", String(D.mediana) + " años"],
  ]),

  H1("9. Paso 6: gráficos"),
  P("Los gráficos salen de la propia tabla (están en el documento de análisis y en archivos PNG aparte):"),
  bullet([b("lx: "), run("cuántos de los 100,000 siguen vivos a cada edad. Muestra que casi todos sobreviven hasta los 50 años y que luego baja rápido.")]),
  bullet([b("qx (escala logarítmica): "), run("el riesgo de morir por edad: alto el primer año, mínimo alrededor de los 8 años y creciente después.")]),
  bullet([b("dx: "), run("en qué edades mueren las personas de la tabla; se concentran alrededor de los 83 años.")]),
  bullet([b("ex: "), run("cuántos años más se espera vivir a cada edad.")]),
  bullet([b("Sobremortalidad masculina y comparación con CONAPO: "), run("complementos del análisis.")]),

  H1("10. Paso 7: cómo se verificó que está bien"),
  bullet("Se recalculó toda la tabla por separado en Python y se comparó con el Excel: coinciden hasta el decimal 14."),
  bullet("Los totales del Excel (defunciones, nacimientos, población) se comparan con los totales de los archivos originales: diferencia 0."),
  bullet(`Se probó apagar el suavizado y recalcular: el Excel coincide también sin suavizar.`),
  bullet(`Referencia externa (no se usa en la tabla): la esperanza de vida oficial de CONAPO para 2019 es ${f(D.ref_a)} años; esta tabla da ${f(D.e0)}.`),

  H1("11. Cómo está organizado el Excel"),
  tabla(["Hoja", "Contenido"], [2600, 6760], [
    ["Tabla de mortalidad", "La plantilla de la profesora completa (x de 0 a 100) con fórmulas, nota técnica y la celda amarilla del suavizado."],
    ["Datos usados", "Para cada edad: defunciones, defunciones ajustadas, defunciones usadas, población, población usada y qx."],
    ["Defunciones", "Conteo de defunciones ocurridas en 2019 por edad y sexo, y cuántas aportó cada archivo."],
    ["Población", "Población del Censo 2020 por edad y sexo."],
    ["Nacimientos", "Nacimientos de 2019 por sexo y por archivo."],
    ["Fuentes", "Bases de INEGI usadas, qué campos y enlaces."],
  ]),

  H1("12. Preguntas que puede hacer la maestra"),
  P([b("¿Por qué se usó población del Censo si la tarea habla de defunciones y nacimientos? "), run("Con solo esos dos datos se puede calcular q(0), pero no el riesgo de morir de las demás edades: hace falta saber cuántas personas había de cada edad. El Censo 2020 es también de INEGI.")]),
  P([b("¿Las defunciones son de 2019 y el Censo es de 2020, no es un desfase? "), run("Sí. El Censo se levantó en marzo de 2020 y las defunciones corresponden a todo 2019, una diferencia de unos meses de crecimiento de la población. No se corrigió; es una aproximación normal en estos cálculos.")]),
  P([b("¿Es una tabla de generación? "), run("No: es de periodo (generación ficticia). Una generación real de 100 años no se puede seguir con estas bases, porque INEGI tiene defunciones por edad solo desde 1990.")]),
  P([b("¿Por qué q(100) = 1? "), run("Es el grupo abierto de 100 años y más; la tabla se cierra suponiendo que todos mueren. Por eso e100 queda en 0.5 años.")]),
  P([b("¿Por qué la esperanza de vida es mayor que la de CONAPO (76.0 contra 74.8)? "), run("Aquí se usan defunciones registradas y población censal sin correcciones. Es probable que haya algo de subregistro de defunciones y que el Censo exagere la población de edades avanzadas; con estas bases no se puede medir cuánto aporta cada causa.")]),
  P([b("¿Por qué a los 100 años llegan 3,536 de 100,000 (3.5 %)? "), run("Es un síntoma de la sobrestimación de la población de edades avanzadas en el Censo (a partir de ~90 años qx se aplana). Se señala como limitación.")]),

  H1("13. Archivos de la entrega"),
  bullet([b("T_Mortalidad_INEGI_Mexico_2019.xlsx: "), run("la tabla (hoja 1) y los datos usados (demás hojas).")]),
  bullet([b("Analisis_Tabla_Mortalidad_INEGI_2019.docx: "), run("método, resultados, gráficos y análisis.")]),
  bullet([b("Explicacion_paso_a_paso_INEGI_2019.docx: "), run("este documento.")]),
  bullet([b("Gráficos en PNG "), run("y, como respaldo, las bases originales de INEGI y el código en el repositorio de GitHub.")]),
];

const doc = new Document({ creator: "Tarea U2_P1", title: "Explicación paso a paso: tabla de mortalidad de México 2019 (INEGI)",
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: "bul", levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{ properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1300, right: 1440, bottom: 1300, left: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: ["Página ", PageNumber.CURRENT], font: FONT, size: 16, color: "808080" })] })] }) },
    children: hijos }] });
Packer.toBuffer(doc).then((x) => fs.writeFileSync(outPath, x));
