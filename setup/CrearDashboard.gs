/**
 * CrearDashboard.gs
 * ------------------------------------------------------------------
 * Genera el libro de finanzas personales de pareja (FM & Lucía).
 *
 * Ejecuta crearDashboard() UNA SOLA VEZ desde el editor de Apps Script.
 * Crea todas las hojas con cabeceras, formato, fórmulas y gráficos.
 *
 * Volver a ejecutarla BORRA y regenera las hojas (se pierden los datos).
 * Usa recrearDashboardSeguro() para que pida confirmación.
 *
 * El bot de Telegram depende del CONTRATO de nombres, columnas y celdas
 * definido aquí. No cambies claves de HOJAS, columnas de Movimientos ni
 * las filas B1..B20 de _Calc sin coordinar con el bot.
 * ------------------------------------------------------------------
 */

/* ============ CONTRATO: nombres de hoja (el bot referencia estas claves) ============ */
const HOJAS = {
  PANEL: 'Panel',
  MOVIMIENTOS: 'Movimientos',
  OBJETIVOS: 'Objetivos',
  AJUSTES: 'Ajustes',
  CALC: '_Calc',
  LOG: '_LogBot',
};

/* ============ Listas de dominio ============ */
const TIPOS = ['Ingreso', 'Aportación', 'Compartido fijo', 'Compartido variable', 'Individual'];
const PERSONAS = ['FM', 'Lucía', 'Bote'];

/* ============ Paleta sobria de producto ============ */
const COLOR = {
  fondo:       '#F7F8FA', // neutro claro de lienzo
  panel:       '#FFFFFF', // tarjetas
  acento:      '#2563EB', // azul único de acento
  acentoSuave: '#E8EFFE',
  tinta:       '#0F172A', // texto fuerte (slate-900)
  texto:       '#334155', // texto normal (slate-700)
  tenue:       '#94A3B8', // etiquetas / leyendas (slate-400)
  borde:       '#E2E8F0', // bordes sutiles (slate-200)
  verde:       '#16A34A', // positivo
  verdeSuave:  '#DCFCE7',
  rojo:        '#DC2626', // negativo
  rojoSuave:   '#FEE2E2',
  naranja:     '#D97706', // intermedio
  naranjaSuave:'#FEF3C7',
  cab:         '#0F172A', // cabeceras de tabla
  cabTxt:      '#FFFFFF',
};

const FUENTE = 'Inter';

/* ====================== ORQUESTADOR ====================== */

function crearDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Guarda el ID del libro para que el bot (Web App) pueda acceder.
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());

  // Pestaña temporal: evita el error "no se pueden borrar todas las hojas".
  const tmpExistente = ss.getSheetByName('__tmp_setup__');
  if (tmpExistente) ss.deleteSheet(tmpExistente);
  const tmp = ss.insertSheet('__tmp_setup__');

  // Borra las hojas del contrato si ya existían.
  Object.values(HOJAS).forEach(nombre => {
    const h = ss.getSheetByName(nombre);
    if (h) ss.deleteSheet(h);
  });

  // Orden lógico: primero las que otras necesitan (Ajustes, Movimientos),
  // luego cálculos y, por último, el Panel.
  crearAjustes(ss);
  crearMovimientos(ss);
  crearObjetivos(ss);
  crearCalc(ss);
  crearLogBot(ss);
  crearPanel(ss);

  // Ocultar hojas internas.
  ss.getSheetByName(HOJAS.CALC).hideSheet();
  ss.getSheetByName(HOJAS.LOG).hideSheet();

  // Panel como primera hoja y activa.
  const panel = ss.getSheetByName(HOJAS.PANEL);
  ss.setActiveSheet(panel);
  ss.moveActiveSheet(1);

  // Borra la temporal y cualquier "Hoja 1" vacía sobrante.
  ss.deleteSheet(tmp);
  const sobrante = ss.getSheets().find(s =>
    (s.getName() === 'Hoja 1' || s.getName() === 'Sheet1') && s.getLastRow() === 0
  );
  if (sobrante && ss.getSheets().length > 1) ss.deleteSheet(sobrante);

  SpreadsheetApp.flush();
  try {
    SpreadsheetApp.getUi().alert(
      'Libro creado correctamente.\n\n' +
      'Siguiente paso: abre la hoja "Ajustes" y rellena el token del bot ' +
      'y los chat IDs de Telegram autorizados. El "Mes activo" ya está fijado al mes en curso.'
    );
  } catch (e) {
    // Ejecutado desde el editor sin UI: no es un error real.
    console.log('Libro creado correctamente. Abre el Sheet para revisarlo.');
  }
}

function recrearDashboardSeguro() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert(
    'Recrear libro',
    'Esto BORRARÁ las hojas existentes con los mismos nombres y se perderán los datos.\n\n¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (r === ui.Button.YES) crearDashboard();
}

/* ====================== HOJA: AJUSTES ====================== */

function crearAjustes(ss) {
  const sh = ss.insertSheet(HOJAS.AJUSTES);

  // --- Bloque de parámetros (A:C). El bot lee la clave literal en A. ---
  const params = [
    ['Parámetro', 'Valor', 'Descripción'],
    ['Mes activo', mesActualISO(), 'Formato AAAA-MM. El Panel y los cálculos usan este mes como referencia.'],
    ['Objetivo ahorro mensual conjunto (€)', 200, 'Meta de ahorro del mes (ingresos − gastos). Editable desde el bot.'],
    ['Telegram chat IDs autorizados', '', 'Separa con coma. Solo estos chats pueden usar el bot.'],
    ['Telegram bot token', '', 'Pega aquí el token que te dé @BotFather.'],
  ];
  sh.getRange(1, 1, params.length, 3).setValues(params);
  sh.getRange('A1:C1')
    .setFontWeight('bold').setFontColor(COLOR.cabTxt).setBackground(COLOR.cab)
    .setVerticalAlignment('middle');
  sh.getRange('A2:A5').setFontWeight('bold').setFontColor(COLOR.tinta);
  sh.getRange('B2:B5').setBackground(COLOR.acentoSuave).setFontColor(COLOR.tinta);
  sh.getRange('B3').setNumberFormat('#,##0 €');
  sh.getRange('C2:C5').setFontColor(COLOR.texto).setWrap(true);
  sh.getRange('A1:C5').setBorder(true, true, true, true, true, true, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  // --- Bloque de categorías (D:F). El bot lee Ajustes!D2:F para desplegables. ---
  const cats = [
    ['Compartido fijo', 'Compartido variable', 'Individual'],
    ['Alquiler', 'Compra', 'Ocio'],
    ['Hipoteca', 'Restaurantes', 'Ropa'],
    ['Comunidad', 'Ocio', 'Transporte'],
    ['Luz', 'Transporte', 'Salud'],
    ['Agua', 'Viajes', 'Suscripciones'],
    ['Gas', 'Regalos', 'Caprichos'],
    ['Internet', 'Hogar', 'Otros'],
    ['Móvil', 'Otros', ''],
    ['Seguros', '', ''],
    ['Suscripciones', '', ''],
    ['Otros', '', ''],
  ];
  sh.getRange(1, 4, cats.length, 3).setValues(cats);
  sh.getRange('D1:F1')
    .setFontWeight('bold').setFontColor(COLOR.cabTxt).setBackground(COLOR.acento)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(2, 4, cats.length - 1, 3).setFontColor(COLOR.texto);
  sh.getRange(1, 4, cats.length, 3).setBorder(true, true, true, true, true, true, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  // Separación visual entre bloque parámetros y bloque categorías.
  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 170);
  sh.setColumnWidth(3, 380);
  sh.setColumnWidth(7, 24); // hueco visual (col G como margen tras categorías... realmente C-D pegadas)
  sh.setColumnWidths(4, 3, 180);
  sh.setRowHeight(1, 26);
  sh.setFrozenRows(1);
  sh.setHiddenGridlines(true);
  sh.getRange('A1:F60').setFontFamily(FUENTE);

  // Título-guía sobre el bloque de categorías para separarlo visualmente.
  sh.getRange('D14').setValue('Edita estas categorías a tu gusto. El bot y los desplegables las usan.')
    .setFontSize(9).setFontStyle('italic').setFontColor(COLOR.tenue);
}

/* ====================== HOJA: MOVIMIENTOS (libro mayor único) ====================== */

function crearMovimientos(ss) {
  const sh = ss.insertSheet(HOJAS.MOVIMIENTOS);

  // Cabecera EXACTA. El bot hace appendRow con este orden de columnas.
  const cab = ['Fecha', 'Tipo', 'Persona', 'Categoría', 'Concepto', 'Importe (€)', 'Nota'];
  sh.getRange(1, 1, 1, 7).setValues([cab]);
  sh.getRange('A1:G1')
    .setFontWeight('bold').setFontColor(COLOR.cabTxt).setBackground(COLOR.cab)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(1, 30);

  // Anchos generosos.
  sh.setColumnWidth(1, 120); // Fecha
  sh.setColumnWidth(2, 160); // Tipo
  sh.setColumnWidth(3, 100); // Persona
  sh.setColumnWidth(4, 170); // Categoría
  sh.setColumnWidth(5, 300); // Concepto
  sh.setColumnWidth(6, 130); // Importe
  sh.setColumnWidth(7, 320); // Nota

  // Formatos de columna.
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('F:F').setNumberFormat('#,##0.00 €');
  sh.getRange('A2:G').setFontColor(COLOR.texto).setVerticalAlignment('middle');

  // Validaciones desplegables (valores EXACTOS del contrato).
  validarLista(sh, 'B2:B', TIPOS);
  validarLista(sh, 'C2:C', PERSONAS);
  // Categoría: validación contra TODAS las categorías de Ajustes (D2:F), allowInvalid=true.
  validarRango(sh, 'D2:D', `${HOJAS.AJUSTES}!D2:F`, true);

  // Filas de ejemplo realistas (una de cada tipo) para verse poblado al abrir.
  const hoy = new Date();
  const y = hoy.getFullYear();
  const m = hoy.getMonth(); // 0-based
  const f = (dia) => new Date(y, m, dia);
  const ejemplos = [
    [f(1),  'Ingreso',             'FM',    'Otros',    'Nómina FM',           1850, 'Salario mensual'],
    [f(1),  'Ingreso',             'Lucía', 'Otros',    'Nómina Lucía',        1620, 'Salario mensual'],
    [f(2),  'Aportación',          'FM',    'Otros',    'Aportación al bote',   400, 'Parte gastos comunes'],
    [f(2),  'Aportación',          'Lucía', 'Otros',    'Aportación al bote',   400, 'Parte gastos comunes'],
    [f(3),  'Compartido fijo',     'Bote',  'Alquiler', 'Alquiler piso',        750, ''],
    [f(8),  'Compartido variable', 'Bote',  'Compra',   'Supermercado',         186, 'Semana 1'],
    [f(12), 'Individual',          'FM',    'Ocio',     'Cine',                  24, ''],
    [f(15), 'Individual',          'Lucía', 'Ropa',     'Zapatillas',            59, ''],
  ];
  sh.getRange(2, 1, ejemplos.length, 7).setValues(ejemplos);
  sh.getRange(2, 1, ejemplos.length, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, 6, ejemplos.length, 1).setNumberFormat('#,##0.00 €');

  // Filas cebra nativas.
  const banda = sh.getRange(1, 1, Math.max(ejemplos.length + 1, 50), 7)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banda.setHeaderRowColor(COLOR.cab);
  banda.setFirstRowColor(COLOR.panel);
  banda.setSecondRowColor('#F1F5F9');

  sh.setFrozenRows(1);
  sh.getRange('A:G').setFontFamily(FUENTE);
}

/* ====================== HOJA: OBJETIVOS (largo plazo) ====================== */

function crearObjetivos(ss) {
  const sh = ss.insertSheet(HOJAS.OBJETIVOS);

  const cab = ['Concepto', 'Meta (€)', 'Aportado (€)', '% Progreso', 'Fecha objetivo', 'Estado', 'Barra'];
  sh.getRange(1, 1, 1, 7).setValues([cab]);
  sh.getRange('A1:G1')
    .setFontWeight('bold').setFontColor(COLOR.cabTxt).setBackground(COLOR.cab)
    .setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.setRowHeight(1, 30);

  // Ejemplos con aportado parcial.
  const hoy = new Date();
  const enSeisMeses = new Date(hoy.getFullYear(), hoy.getMonth() + 6, hoy.getDate());
  const enUnAnio = new Date(hoy.getFullYear() + 1, hoy.getMonth(), hoy.getDate());
  const enDosAnios = new Date(hoy.getFullYear() + 2, hoy.getMonth(), hoy.getDate());

  const ejemplos = [
    ['Vacaciones',          2500, 1200, enSeisMeses],
    ['Coche',               8000, 3000, enDosAnios],
    ['Fondo de emergencia', 6000, 6000, enUnAnio],
  ];
  ejemplos.forEach((e, i) => {
    const r = i + 2;
    sh.getRange(r, 1).setValue(e[0]);
    sh.getRange(r, 2).setValue(e[1]);
    sh.getRange(r, 3).setValue(e[2]);
    sh.getRange(r, 5).setValue(e[3]);
  });

  // Fórmulas de % progreso, estado y barra (filas 2..50 listas para nuevas metas).
  for (let r = 2; r <= 50; r++) {
    sh.getRange(r, 4).setFormula(`=IFERROR(C${r}/B${r};0)`);
    sh.getRange(r, 6).setFormula(`=IF(B${r}="";"";IF(C${r}>=B${r};"✅ Cumplido";"En curso"))`);
    // Barra visual de 10 bloques.
    sh.getRange(r, 7).setFormula(
      `=IF(B${r}="";"";REPT("█";ROUND(MIN(C${r}/B${r};1)*10;0))&REPT("░";10-ROUND(MIN(C${r}/B${r};1)*10;0)))`
    );
  }

  // Anchos y formatos.
  sh.setColumnWidth(1, 220);
  sh.setColumnWidth(2, 120);
  sh.setColumnWidth(3, 120);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 140);
  sh.setColumnWidth(6, 130);
  sh.setColumnWidth(7, 160);
  sh.getRange('B2:C').setNumberFormat('#,##0.00 €');
  sh.getRange('D2:D').setNumberFormat('0.0%');
  sh.getRange('E2:E').setNumberFormat('yyyy-mm-dd');
  sh.getRange('A2:G').setFontColor(COLOR.texto).setVerticalAlignment('middle');
  sh.getRange('G2:G').setFontColor(COLOR.acento).setHorizontalAlignment('left');
  sh.getRange('F2:F').setHorizontalAlignment('center');
  sh.getRange('D2:D').setHorizontalAlignment('center');

  // Validación de fecha en la columna de fecha objetivo.
  const reglaFecha = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  sh.getRange('E2:E').setDataValidation(reglaFecha);

  // Filas cebra.
  const banda = sh.getRange(1, 1, 50, 7)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banda.setHeaderRowColor(COLOR.cab);
  banda.setFirstRowColor(COLOR.panel);
  banda.setSecondRowColor('#F1F5F9');

  // Formato condicional del estado.
  const reglaCumplido = SpreadsheetApp.newConditionalFormatRule()
    .whenTextContains('Cumplido').setBackground(COLOR.verdeSuave).setFontColor(COLOR.verde).setBold(true)
    .setRanges([sh.getRange('F2:F50')]).build();
  const reglaEnCurso = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('En curso').setBackground(COLOR.acentoSuave).setFontColor(COLOR.acento).setBold(true)
    .setRanges([sh.getRange('F2:F50')]).build();
  sh.setConditionalFormatRules([reglaCumplido, reglaEnCurso]);

  sh.setFrozenRows(1);
  sh.setHiddenGridlines(true);
  sh.getRange('A:G').setFontFamily(FUENTE);
}

/* ====================== HOJA: _Calc (oculta) ====================== */
/**
 * Métricas del MES ACTIVO en celdas FIJAS (B1..B20). El bot lee estas filas.
 * Tablas auxiliares para los visuales del Panel desde la columna E.
 */
function crearCalc(ss) {
  const sh = ss.insertSheet(HOJAS.CALC);

  const MOV = HOJAS.MOVIMIENTOS;
  const AJ = HOJAS.AJUSTES;

  // Mes activo y rango de fechas del mes (según contrato).
  const mesActivo = `VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE)`;
  const inicioMes = `DATE(VALUE(LEFT(${mesActivo};4));VALUE(MID(${mesActivo};6;2));1)`;
  const finMes = `EOMONTH(${inicioMes};0)+1`;

  // SUMIFS sobre Movimientos: B=Tipo, C=Persona, F=Importe, A=Fecha.
  // Filtra por tipo (+ persona opcional) y rango de fechas del mes activo.
  const sumar = (tipo, persona) => {
    let f = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};${MOV}!B:B;"${tipo}"`;
    if (persona) f += `;${MOV}!C:C;"${persona}"`;
    f += `)`;
    return `IFERROR(${f};0)`;
  };

  // Objetivo de ahorro desde Ajustes.
  const objetivo = `IFERROR(VLOOKUP("Objetivo ahorro mensual conjunto (€)";${AJ}!A:B;2;FALSE);0)`;

  // Etiqueta (A) | fórmula (B), filas 1..20 EXACTAS.
  const metricas = [
    ['Ingresos FM',                       `=${sumar('Ingreso', 'FM')}`],
    ['Ingresos Lucía',                    `=${sumar('Ingreso', 'Lucía')}`],
    ['Ingresos total',                    '=B1+B2'],
    ['Aportación al bote FM',             `=${sumar('Aportación', 'FM')}`],
    ['Aportación al bote Lucía',          `=${sumar('Aportación', 'Lucía')}`],
    ['Aportación al bote total',          '=B4+B5'],
    ['Gastos compartidos fijos',          `=${sumar('Compartido fijo', null)}`],
    ['Gastos compartidos variables',      `=${sumar('Compartido variable', null)}`],
    ['Gastos compartidos total',          '=B7+B8'],
    ['Gastos individuales FM',            `=${sumar('Individual', 'FM')}`],
    ['Gastos individuales Lucía',         `=${sumar('Individual', 'Lucía')}`],
    ['Gastos totales del mes',            '=B9+B10+B11'],
    ['Saldo FM',                          '=B1-B4-B10'],
    ['Saldo Lucía',                       '=B2-B5-B11'],
    ['Bote sobrante',                     '=B6-B9'],
    ['Objetivo ahorro',                   `=${objetivo}`],
    ['Ahorro real del mes',               '=B3-B12'],
    ['% objetivo cumplido',               '=IFERROR(B17/B16;0)'],
    ['Ahorro acumulado del año',          ''], // se rellena abajo (referencia a serie anual)
    ['Balance del mes',                   '=B17'],
  ];
  sh.getRange(1, 1, metricas.length, 2).setValues(metricas);

  // Formatos numéricos de la columna B.
  sh.getRange('B1:B17').setNumberFormat('#,##0.00 €');
  sh.getRange('B18').setNumberFormat('0.0%');
  sh.getRange('B19:B20').setNumberFormat('#,##0.00 €');
  sh.setColumnWidth(1, 240);
  sh.setColumnWidth(2, 130);
  sh.getRange('A1:B20').setFontFamily(FUENTE).setFontColor(COLOR.texto);

  /* ===== TABLA AUX 1: gasto por categoría del mes activo (E:F) ===== */
  // Suma de cualquier gasto (los 3 tipos de gasto) por categoría de Ajustes (col D:F).
  // Concatenamos las 3 columnas de categorías de Ajustes en una sola lista vertical.
  sh.getRange('E1').setValue('Categoría');
  sh.getRange('F1').setValue('Importe');
  // Lista única de categorías (apiladas) en E2 hacia abajo.
  const catsApiladas = `{${AJ}!D2:D;${AJ}!E2:E;${AJ}!F2:F}`;
  sh.getRange('E2').setFormula(
    `=IFERROR(SORT(UNIQUE(FILTER(${catsApiladas};${catsApiladas}<>"")));"")`
  );
  // Importe por categoría: suma de gastos (fijo+variable+individual) del mes con esa categoría.
  // Se calcula en F2:F40 con ARRAYFORMULA referenciando E (categoría) y Movimientos.
  sh.getRange('F2').setFormula(
    `=ARRAYFORMULA(IF(E2:E40="";"";` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};` +
    `${MOV}!D:D;E2:E40;${MOV}!B:B;"Compartido fijo");0)+` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};` +
    `${MOV}!D:D;E2:E40;${MOV}!B:B;"Compartido variable");0)+` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};` +
    `${MOV}!D:D;E2:E40;${MOV}!B:B;"Individual");0)))`
  );
  sh.getRange('F2:F40').setNumberFormat('#,##0.00 €');

  // Tabla compacta para el gráfico de categorías: solo categorías con gasto > 0 (H:I).
  sh.getRange('H1').setValue('Categoría (gasto)');
  sh.getRange('I1').setValue('Importe');
  sh.getRange('H2').setFormula(
    `=IFERROR(QUERY({E2:F40};"select Col1, Col2 where Col2 > 0 order by Col2 desc label Col2 ''";0);"")`
  );
  sh.getRange('I2:I40').setNumberFormat('#,##0.00 €');

  /* ===== TABLA AUX 2: ahorro acumulado por mes del año (K:M) ===== */
  // K = mes (texto AAAA-MM), L = ahorro real del mes, M = acumulado.
  sh.getRange('K1').setValue('Mes');
  sh.getRange('L1').setValue('Ahorro mes');
  sh.getRange('M1').setValue('Acumulado');
  const anio = `VALUE(LEFT(${mesActivo};4))`;
  for (let m = 1; m <= 12; m++) {
    const r = m + 1;
    const ini = `DATE(${anio};${m};1)`;
    const fin = `DATE(${anio};${m + 1};1)`;
    // Etiqueta de mes AAAA-MM.
    sh.getRange(r, 11).setFormula(`=TEXT(${ini};"yyyy-mm")`);
    // Ahorro real del mes = ingresos − gastos (fijos+variables+individuales) del mes.
    const ingMes = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Ingreso")`;
    const gastMes =
      `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Compartido fijo")+` +
      `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Compartido variable")+` +
      `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Individual")`;
    sh.getRange(r, 12).setFormula(`=IFERROR((${ingMes})-(${gastMes});0)`);
    // Acumulado hasta ese mes.
    sh.getRange(r, 13).setFormula(m === 1 ? '=L2' : `=M${r - 1}+L${r}`);
  }
  sh.getRange('L2:M13').setNumberFormat('#,##0.00 €');

  // B19 (Ahorro acumulado del año hasta el mes activo) = acumulado de la fila del mes activo.
  const idxMes = `VALUE(MID(${mesActivo};6;2))`;
  sh.getRange('B19').setFormula(`=IFERROR(INDEX(M2:M13;${idxMes});0)`);

  sh.setColumnWidth(5, 150);
  sh.setColumnWidth(6, 110);
  sh.setColumnWidth(8, 150);
  sh.setColumnWidth(9, 110);
  sh.setColumnWidth(11, 90);
  sh.setColumnWidth(12, 110);
  sh.setColumnWidth(13, 110);
}

/* ====================== HOJA: _LogBot (oculta) ====================== */

function crearLogBot(ss) {
  const sh = ss.insertSheet(HOJAS.LOG);
  sh.getRange(1, 1, 1, 4).setValues([['Fecha', 'Persona', 'Acción', 'Detalle']]);
  sh.getRange('A1:D1')
    .setFontWeight('bold').setFontColor(COLOR.cabTxt).setBackground(COLOR.cab)
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 100);
  sh.setColumnWidth(3, 180);
  sh.setColumnWidth(4, 420);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  sh.setFrozenRows(1);
  sh.getRange('A:D').setFontFamily(FUENTE);
}

/* ====================== HOJA: PANEL (dashboard) ====================== */
/**
 * Dashboard premium minimalista. TODO referencia _Calc!Bx (no recalcula).
 * Layout (cols B..K, 10 columnas de contenido; A y L márgenes):
 *   Fila 2-3   Cabecera (título, subtítulo mes, actualizado)
 *   Fila 6-9   4 KPIs en tarjetas
 *   Fila 12-18 Tabla "Por persona"
 *   Fila 21-25 Tabla "Bote común"
 *   Fila 28-33 Tabla "Objetivos"
 *   Fila 36+   2 gráficos (categorías + ahorro acumulado)
 */
function crearPanel(ss) {
  const sh = ss.insertSheet(HOJAS.PANEL);
  const C = `'${HOJAS.CALC}'`;
  const AJ = HOJAS.AJUSTES;

  sh.setHiddenGridlines(true);

  // Lienzo neutro y márgenes.
  sh.getRange('A1:L80').setBackground(COLOR.fondo).setFontFamily(FUENTE);
  sh.setColumnWidth(1, 24);  // margen izq
  sh.setColumnWidths(2, 10, 92); // B..K contenido
  sh.setColumnWidth(12, 24); // margen der

  /* ===== Cabecera ===== */
  sh.setRowHeight(2, 40);
  sh.getRange('B2:H2').merge()
    .setValue('Finanzas FM & Lucía')
    .setFontSize(24).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setBackground(COLOR.fondo);
  sh.getRange('I2:K2').merge()
    .setFormula(`="Actualizado " & TEXT(NOW();"yyyy-mm-dd HH:mm")`)
    .setFontSize(9).setFontColor(COLOR.tenue)
    .setHorizontalAlignment('right').setVerticalAlignment('bottom').setBackground(COLOR.fondo);
  sh.setRowHeight(3, 22);
  sh.getRange('B3:K3').merge()
    .setFormula(`="Resumen del mes " & VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE)`)
    .setFontSize(11).setFontColor(COLOR.tenue)
    .setVerticalAlignment('middle').setBackground(COLOR.fondo);

  /* ===== Fila de 4 KPIs (filas 6-9), tarjetas de 2 cols + 1 col de gap ===== */
  // Columnas: B:C | E:F | H:I... usamos bloques de 2 cols con gap de 1.
  // Layout 4 cards: B:C(6) D gap, E:F, G gap, H:I, J gap, ... no caben con gap.
  // Distribución: B:C, D:E, F:G, H:I (4 cards de 2 cols, K margen extra).
  pintarKPI(sh, 6, 2, 3, 'INGRESOS DEL MES', `=${C}!B3`, false,
    `="FM " & TEXT(${C}!B1;"#,##0 €") & "  ·  Lucía " & TEXT(${C}!B2;"#,##0 €")`, COLOR.acento);
  pintarKPI(sh, 6, 4, 5, 'GASTOS DEL MES', `=${C}!B12`, false,
    `="Compart. " & TEXT(${C}!B9;"#,##0 €") & "  ·  Indiv. " & TEXT(${C}!B10+${C}!B11;"#,##0 €")`, COLOR.acento);
  pintarKPI(sh, 6, 6, 7, 'AHORRO DEL MES', `=${C}!B17`, false,
    `=IF(${C}!B17>=0;"Superávit del mes";"Déficit del mes")`, COLOR.acento, 'signo');
  pintarKPI(sh, 6, 8, 9, '% OBJETIVO', `=${C}!B18`, true,
    `="Meta " & TEXT(${C}!B16;"#,##0 €") & "  ·  Real " & TEXT(${C}!B17;"#,##0 €")`, COLOR.acento, 'objetivo');

  /* ===== Sección "Por persona" (tabla limpia) ===== */
  etiquetaSeccion(sh, 11, 'POR PERSONA');
  // Cabecera de tabla en fila 12.
  const phead = ['', 'Ingresos', 'Aportación bote', 'Gastos', 'Saldo'];
  // Mapear a columnas B(label ancho B:C), D(ing), E(aport), F+G(gastos? ) -> usar B:C,D,E,F,G,H..
  // Tabla: B:C etiqueta | D:E Ingresos | F Aportación | G Gastos | H:I... Para limpieza, usamos:
  //   B:C (persona) | D:E (Ingresos) | F:G (Aportación) | H:I (Gastos) | J:K (Saldo)
  tablaPersonaCabecera(sh, 12);
  tablaPersonaFila(sh, 13, 'FM', `=${C}!B1`, `=${C}!B4`, `=${C}!B10`, `=${C}!B13`);
  tablaPersonaFila(sh, 14, 'Lucía', `=${C}!B2`, `=${C}!B5`, `=${C}!B11`, `=${C}!B14`);

  /* ===== Sección "Bote común" ===== */
  etiquetaSeccion(sh, 16, 'BOTE COMÚN');
  tablaBote(sh, 17, C);

  /* ===== Sección "Objetivos" ===== */
  etiquetaSeccion(sh, 20, 'OBJETIVOS DE AHORRO');
  tablaObjetivos(sh, 21);

  /* ===== Gráficos ===== */
  etiquetaSeccion(sh, 28, 'ANÁLISIS');
  insertarGraficos(sh, ss);

  sh.setFrozenRows(3);
  sh.setColumnWidth(12, 24);
}

/* ---------- KPIs ---------- */
/**
 * Tarjeta KPI ocupando columnas [colIni..colFin] x filas [fila..fila+3].
 * fila   = barra de acento (fina)
 * fila+1 = etiqueta pequeña en mayúsculas
 * fila+2 = número grande
 * fila+3 = subtítulo de contexto
 * modo: 'signo' -> verde/rojo según valor; 'objetivo' -> umbrales %.
 */
function pintarKPI(sh, fila, colIni, colFin, etiqueta, formulaValor, esPct, formulaDetalle, acento, modo) {
  const nCols = colFin - colIni + 1;
  const card = sh.getRange(fila, colIni, 4, nCols);
  card.setBackground(COLOR.panel)
    .setBorder(true, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  // Barra de acento superior (fina).
  sh.getRange(fila, colIni, 1, nCols).setBackground(acento);
  sh.setRowHeight(fila, 4);

  // Etiqueta.
  sh.getRange(fila + 1, colIni, 1, nCols).merge()
    .setValue(etiqueta)
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 1, 22);

  // Número grande.
  const valor = sh.getRange(fila + 2, colIni, 1, nCols).merge()
    .setFormula(formulaValor)
    .setFontSize(22).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBackground(COLOR.panel);
  valor.setNumberFormat(esPct ? '0.0%' : '#,##0.00 €');
  sh.setRowHeight(fila + 2, 40);

  // Subtítulo.
  sh.getRange(fila + 3, colIni, 1, nCols).merge()
    .setFormula(formulaDetalle)
    .setFontSize(9).setFontColor(COLOR.tenue)
    .setVerticalAlignment('top').setHorizontalAlignment('left')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 3, 24);

  // Sangría visual: pequeño margen interno con columna no, usamos alineación.
  const valCell = sh.getRange(fila + 2, colIni, 1, nCols);
  if (modo === 'signo') {
    aplicarPositivoNegativo(sh, valCell);
  } else if (modo === 'objetivo') {
    aplicarUmbralesObjetivo(sh, valCell);
  }
}

/* ---------- Etiqueta de sección ---------- */
function etiquetaSeccion(sh, fila, texto) {
  sh.setRowHeight(fila, 24);
  sh.getRange(fila, 2, 1, 9).merge()
    .setValue(texto)
    .setFontSize(10).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setVerticalAlignment('bottom').setBackground(COLOR.fondo);
}

/* ---------- Tabla "Por persona" ---------- */
function tablaPersonaCabecera(sh, fila) {
  sh.setRowHeight(fila, 28);
  const cols = [
    [2, 3, ''],            // B:C persona
    [4, 5, 'Ingresos'],    // D:E
    [6, 7, 'Aportación'],  // F:G
    [8, 9, 'Gastos'],      // H:I
    [10, 11, 'Saldo'],     // J:K
  ];
  cols.forEach(([c1, c2, txt], i) => {
    const r = sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setValue(txt)
      .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.cabTxt)
      .setBackground(COLOR.acento).setVerticalAlignment('middle');
    r.setHorizontalAlignment(i === 0 ? 'left' : 'right');
  });
}

function tablaPersonaFila(sh, fila, nombre, fIng, fApo, fGas, fSaldo) {
  sh.setRowHeight(fila, 30);
  const blanco = fila % 2 === 1 ? COLOR.panel : '#F8FAFC';
  sh.getRange(fila, 2, 1, 10).setBackground(blanco)
    .setBorder(false, false, true, false, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(fila, 2, 1, 2).merge().setValue(nombre)
    .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');

  const celdas = [
    [4, 5, fIng], [6, 7, fApo], [8, 9, fGas], [10, 11, fSaldo],
  ];
  celdas.forEach(([c1, c2, f], i) => {
    const cell = sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setFormula(f).setNumberFormat('#,##0.00 €')
      .setFontSize(11).setFontColor(COLOR.texto)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
    if (i === 3) { // Saldo: verde/rojo
      cell.setFontWeight('bold');
      aplicarPositivoNegativo(sh, sh.getRange(fila, c1, 1, c2 - c1 + 1));
    }
  });
}

/* ---------- Tabla "Bote común" ---------- */
function tablaBote(sh, fila, C) {
  // Cabecera
  sh.setRowHeight(fila, 28);
  const cols = [
    [2, 4, 'Aportado'],
    [5, 7, 'Gastos compartidos'],
    [8, 11, 'Sobrante'],
  ];
  cols.forEach(([c1, c2, txt]) => {
    sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setValue(txt)
      .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.cabTxt)
      .setBackground(COLOR.acento).setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
  });
  // Valores
  const fr = fila + 1;
  sh.setRowHeight(fr, 34);
  sh.getRange(fr, 2, 1, 10).setBackground(COLOR.panel)
    .setBorder(false, false, true, false, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  const vals = [
    [2, 4, `=${C}!B6`, false],
    [5, 7, `=${C}!B9`, false],
    [8, 11, `=${C}!B15`, true], // sobrante: verde/rojo
  ];
  vals.forEach(([c1, c2, f, signo]) => {
    const cell = sh.getRange(fr, c1, 1, c2 - c1 + 1).merge()
      .setFormula(f).setNumberFormat('#,##0.00 €')
      .setFontSize(15).setFontWeight('bold').setFontColor(COLOR.tinta)
      .setVerticalAlignment('middle').setHorizontalAlignment('center');
    if (signo) aplicarPositivoNegativo(sh, sh.getRange(fr, c1, 1, c2 - c1 + 1));
  });
}

/* ---------- Tabla "Objetivos" (refleja la hoja Objetivos) ---------- */
function tablaObjetivos(sh, fila) {
  const OBJ = `'${HOJAS.OBJETIVOS}'`;
  // Cabecera
  sh.setRowHeight(fila, 26);
  const cols = [
    [2, 4, 'Concepto', 'left'],
    [5, 6, 'Meta', 'right'],
    [7, 8, 'Aportado', 'right'],
    [9, 11, 'Progreso', 'left'],
  ];
  cols.forEach(([c1, c2, txt, al]) => {
    sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setValue(txt)
      .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.cabTxt)
      .setBackground(COLOR.acento).setVerticalAlignment('middle').setHorizontalAlignment(al);
  });
  // 3 filas espejo de la hoja Objetivos (filas 2,3,4).
  for (let i = 0; i < 3; i++) {
    const r = fila + 1 + i;
    const oRow = 2 + i;
    sh.setRowHeight(r, 26);
    const fondo = i % 2 === 0 ? COLOR.panel : '#F8FAFC';
    sh.getRange(r, 2, 1, 10).setBackground(fondo)
      .setBorder(false, false, true, false, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

    sh.getRange(r, 2, 1, 3).merge().setFormula(`=${OBJ}!A${oRow}`)
      .setFontSize(10).setFontColor(COLOR.tinta).setVerticalAlignment('middle');
    sh.getRange(r, 5, 1, 2).merge().setFormula(`=${OBJ}!B${oRow}`)
      .setNumberFormat('#,##0 €').setFontSize(10).setFontColor(COLOR.texto)
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    sh.getRange(r, 7, 1, 2).merge().setFormula(`=${OBJ}!C${oRow}`)
      .setNumberFormat('#,##0 €').setFontSize(10).setFontColor(COLOR.texto)
      .setHorizontalAlignment('right').setVerticalAlignment('middle');
    sh.getRange(r, 9, 1, 3).merge()
      .setFormula(`=${OBJ}!G${oRow} & "  " & TEXT(${OBJ}!D${oRow};"0%")`)
      .setFontSize(10).setFontColor(COLOR.acento).setVerticalAlignment('middle');
  }
}

/* ---------- Gráficos (máximo 2) ---------- */
function insertarGraficos(sh, ss) {
  const calc = ss.getSheetByName(HOJAS.CALC);

  // 1) Donut: gasto por categoría del mes (lee _Calc!H:I, tabla auxiliar).
  const donut = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(calc.getRange('H1:I20'))
    .setNumHeaders(1)
    .setOption('title', 'Gasto por categoría')
    .setOption('titleTextStyle', { color: COLOR.texto, fontSize: 12, bold: true })
    .setOption('pieHole', 0.6)
    .setOption('legend', { position: 'right', textStyle: { color: COLOR.texto, fontSize: 10 } })
    .setOption('pieSliceText', 'none')
    .setOption('backgroundColor', COLOR.panel)
    .setOption('colors', ['#2563EB', '#60A5FA', '#93C5FD', '#0EA5E9', '#38BDF8', '#1E40AF', '#3B82F6', '#7DD3FC', '#475569', '#94A3B8'])
    .setOption('chartArea', { left: 12, top: 40, width: '92%', height: '80%' })
    .setOption('width', 430).setOption('height', 280)
    .setPosition(30, 2, 0, 0) // fila 30, col B
    .build();
  sh.insertChart(donut);

  // 2) Línea: ahorro acumulado del año (lee _Calc!K:M, columna M acumulado).
  const linea = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(calc.getRange('K1:K13'))
    .addRange(calc.getRange('M1:M13'))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setOption('title', 'Ahorro acumulado del año')
    .setOption('titleTextStyle', { color: COLOR.texto, fontSize: 12, bold: true })
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COLOR.acento])
    .setOption('curveType', 'function')
    .setOption('lineWidth', 3)
    .setOption('pointSize', 4)
    .setOption('backgroundColor', COLOR.panel)
    .setOption('hAxis', { textStyle: { color: COLOR.tenue, fontSize: 9 } })
    .setOption('vAxis', { textStyle: { color: COLOR.tenue, fontSize: 9 }, format: '#,##0 €', gridlines: { color: COLOR.borde } })
    .setOption('chartArea', { left: 60, top: 40, width: '85%', height: '75%' })
    .setOption('width', 430).setOption('height', 280)
    .setPosition(30, 7, 0, 0) // fila 30, col G
    .build();
  sh.insertChart(linea);
}

/* ====================== FORMATO CONDICIONAL ====================== */

/** Verde si >=0, rojo si <0. Conserva reglas previas. */
function aplicarPositivoNegativo(sh, rango) {
  const reglas = sh.getConditionalFormatRules();
  reglas.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0).setFontColor(COLOR.verde)
      .setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0).setFontColor(COLOR.rojo)
      .setRanges([rango]).build()
  );
  sh.setConditionalFormatRules(reglas);
}

/** % objetivo: verde >=100%, naranja 50-99%, rojo <50%. */
function aplicarUmbralesObjetivo(sh, rango) {
  const reglas = sh.getConditionalFormatRules();
  reglas.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(1).setFontColor(COLOR.verde)
      .setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.5, 0.9999).setFontColor(COLOR.naranja)
      .setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.5).setFontColor(COLOR.rojo)
      .setRanges([rango]).build()
  );
  sh.setConditionalFormatRules(reglas);
}

/* ====================== HELPERS GENERALES ====================== */

/** Validación desplegable contra una lista literal de valores. */
function validarLista(sh, rangoA1, valores) {
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(valores, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(rangoA1).setDataValidation(regla);
}

/** Validación desplegable contra un rango de otra hoja (ej. categorías). */
function validarRango(sh, rangoA1, rangoFuenteA1, allowInvalid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const [nombreHoja, rango] = rangoFuenteA1.split('!');
  const fuente = ss.getSheetByName(nombreHoja).getRange(rango);
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInRange(fuente, true)
    .setAllowInvalid(allowInvalid !== false)
    .build();
  sh.getRange(rangoA1).setDataValidation(regla);
}

/** Mes actual en formato AAAA-MM. */
function mesActualISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}
