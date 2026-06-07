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
 * Contrato con el bot (no cambiar sin coordinar):
 *  - Nombres de HOJAS.
 *  - Orden de columnas de Movimientos (A:G).
 *  - Etiquetas literales en Ajustes!A (VLOOKUP por clave).
 *  - Celdas _Calc!B1..B20 (significado y posición).
 * ------------------------------------------------------------------
 */

/* ============ CONTRATO: nombres de hoja (el bot referencia estas claves) ============ */
const HOJAS = {
  COMPARTIDA: 'Compartida',
  FM: 'FM',
  LUCIA: 'Lucía',
  ANUAL: 'Resumen anual',
  TICKETS: 'Tickets',
  MOVIMIENTOS: 'Movimientos',
  OBJETIVOS: 'Objetivos',
  AJUSTES: 'Ajustes',
  CALC: '_Calc',
  LOG: '_LogBot',
};

/* ============ Listas de dominio ============ */
const TIPOS = ['Ingreso', 'Aportación', 'Compartido fijo', 'Compartido variable', 'Individual', 'Ahorro'];
const PERSONAS = ['FM', 'Lucía', 'Bote'];

// Reservas de filas en las listas dinámicas del Panel. Si el mes tiene más
// movimientos que la reserva, salen los primeros N y aparece "(+X más)".
const RESERVA = {
  COMP_FIJO: 8,
  COMP_VAR: 15,
  COMP_AHO: 4,
  IND: 10,
  IND_AHO: 4,
};

/* ============ Paleta pastel "Simple Budget" (estilo de las referencias) ============
 * Crema de fondo, verde salvia para ingresos, rosa empolvado para gastos,
 * arena para totales. Texto gris cálido. */
const COLOR = {
  fondo:       '#FBF8F3', // crema de lienzo
  panel:       '#FFFFFF', // tarjetas
  // Verde salvia (ingresos / aportaciones)
  verde:       '#8FAE8B',
  verdeClaro:  '#E7EFE5',
  // Rosa empolvado (gastos)
  rosa:        '#E0A9A6',
  rosaClaro:   '#F7E8E7',
  // Arena (totales / neutro)
  arena:       '#C9B68F',
  arenaClaro:  '#F2EBDD',
  // Texto
  tinta:       '#4A4A4A', // títulos
  texto:       '#6B6B6B', // normal
  tenue:       '#9A9A9A', // leyendas
  borde:       '#E5DFD6', // bordes suaves
  cabTxt:      '#FFFFFF',
  cebra:       '#FAF7F1', // filas alternas
  // Señal de signo
  positivo:    '#5E8C61',
  negativo:    '#C77',
  rojo:        '#C0706B',
  // Alias para las hojas de datos (Movimientos/Objetivos/Ajustes)
  cab:         '#7A746B', // cabecera oscura cálida
  acento:      '#8FAE8B', // = verde salvia
  acentoSuave: '#E7EFE5',
  verdeSuave:  '#E7EFE5',
  // Paleta del donut (categorías)
  donut: ['#8FAE8B', '#E0A9A6', '#C9B68F', '#9DB4C0', '#D3A9C9', '#A6C0B5', '#E8C7A0', '#B7AED0', '#CBB9A8', '#A9C7C2'],
};

const FUENTE = 'Montserrat';      // cuerpo de tablas
const FUENTE_TIT = 'Playfair Display'; // títulos elegantes

/* ====================== ORQUESTADOR ====================== */

/** Menú propio en la hoja (aparece al abrir el Sheet). */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('💰 Finanzas')
    .addItem('🧹 Limpiar datos (empezar de cero)', 'limpiarDatos')
    .addSeparator()
    .addItem('🤖 Probar Gemini API key', 'testGemini')
    .addItem('🗓 Activar resumen mensual + anomalías', 'configurarTriggersIA')
    .addSeparator()
    .addItem('🔄 Recrear libro completo', 'recrearDashboardSeguro')
    .addToUI();
}

function crearDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Guarda el ID del libro para que el bot (Web App) pueda acceder.
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());

  // Pestaña temporal: evita el error "no se pueden borrar todas las hojas".
  const tmpExistente = ss.getSheetByName('__tmp_setup__');
  if (tmpExistente) ss.deleteSheet(tmpExistente);
  const tmp = ss.insertSheet('__tmp_setup__');

  // Borra TODAS las hojas excepto la temporal: garantiza un libro limpio
  // y elimina cualquier hoja vieja de esquemas anteriores.
  ss.getSheets().forEach(h => {
    if (h.getName() !== '__tmp_setup__') ss.deleteSheet(h);
  });

  // Orden lógico: primero las que otras necesitan (Ajustes, Movimientos),
  // luego cálculos y, por último, las hojas-dashboard.
  crearAjustes(ss);
  crearMovimientos(ss);
  crearObjetivos(ss);
  crearCalc(ss);
  crearLogBot(ss);
  crearHojaCompartida(ss);
  crearHojaPersona(ss, HOJAS.FM, 'Cuenta de FM', 'FM');
  crearHojaPersona(ss, HOJAS.LUCIA, 'Cuenta de Lucía', 'Lucía');
  crearHojaAnual(ss);
  crearHojaTickets(ss);

  // Ocultar hojas internas.
  ss.getSheetByName(HOJAS.CALC).hideSheet();
  ss.getSheetByName(HOJAS.LOG).hideSheet();

  // Orden de pestañas: Compartida · FM · Lucía · Resumen anual · Tickets · Movimientos · Objetivos · Ajustes.
  ['Ajustes', 'Objetivos', 'Movimientos', HOJAS.TICKETS, HOJAS.ANUAL, HOJAS.LUCIA, HOJAS.FM, HOJAS.COMPARTIDA].forEach(n => {
    const s = ss.getSheetByName(n);
    if (s) { ss.setActiveSheet(s); ss.moveActiveSheet(1); }
  });
  ss.setActiveSheet(ss.getSheetByName(HOJAS.COMPARTIDA));

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

/**
 * Borra solo los DATOS (movimientos, metas y log), conservando todo el diseño,
 * las fórmulas y las validaciones. Úsala para empezar limpio sin regenerar el libro.
 */
function limpiarDatos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let confirmar = true;
  try {
    const ui = SpreadsheetApp.getUi();
    confirmar = ui.alert(
      'Limpiar datos',
      'Esto BORRARÁ todos los movimientos, metas y el log, pero conserva el diseño.\n\n¿Continuar?',
      ui.ButtonSet.YES_NO
    ) === ui.Button.YES;
  } catch (e) { /* sin UI (ejecutado desde editor): continúa */ }
  if (!confirmar) return;

  const mov = ss.getSheetByName(HOJAS.MOVIMIENTOS);
  if (mov && mov.getLastRow() > 1) {
    mov.getRange(2, 1, mov.getLastRow() - 1, 7).clearContent();
  }
  const obj = ss.getSheetByName(HOJAS.OBJETIVOS);
  if (obj && obj.getLastRow() > 1) {
    const n = obj.getLastRow() - 1;
    obj.getRange(2, 1, n, 1).clearContent(); // Concepto
    obj.getRange(2, 2, n, 2).clearContent(); // Meta, Aportado
    obj.getRange(2, 5, n, 1).clearContent(); // Fecha objetivo
  }
  const log = ss.getSheetByName(HOJAS.LOG);
  if (log && log.getLastRow() > 1) {
    log.getRange(2, 1, log.getLastRow() - 1, 4).clearContent();
  }
  SpreadsheetApp.flush();
  try {
    SpreadsheetApp.getUi().alert('Datos borrados. Las hojas están en limpio para empezar.');
  } catch (e) {
    console.log('Datos borrados.');
  }
}

/* ====================== HOJA: AJUSTES ====================== */

function crearAjustes(ss) {
  const sh = ss.insertSheet(HOJAS.AJUSTES);

  // --- Bloque de parámetros (A:C). El bot lee la clave literal en A vía VLOOKUP. ---
  const params = [
    ['Parámetro', 'Valor', 'Descripción'],
    ['Mes activo', mesActualISO(), 'Formato AAAA-MM. El Panel y los cálculos usan este mes como referencia.'],
    ['Objetivo ahorro mensual conjunto (€)', 200, 'Meta de ahorro del mes en euros. Editable desde el bot.'],
    ['Telegram chat IDs autorizados', '', 'Separa con coma. Solo estos chats pueden usar el bot.'],
    ['Telegram bot token', '', 'Pega aquí el token que te dé @BotFather.'],
    ['Modo aportación al bote', '50/50', 'Opciones: 50/50, Proporcional, Custom. Define cómo se reparten los gastos compartidos.'],
    ['Tasa de ahorro objetivo (%)', 0.2, 'Referencia para el KPI de tasa de ahorro (20% es el estándar saludable).'],
    ['Reserva de emergencia objetivo (meses)', 3, 'Meses de gastos esenciales que tu ahorro debería cubrir (3-6 estándar).'],
    ['Umbral settle-up (€)', 20, 'Solo en modo proporcional: a partir de este desvío se sugiere transferencia.'],
    ['Gemini API key', '', 'Tu API key de Google AI Studio (https://aistudio.google.com/apikey). Necesaria para OCR de tickets, categorización automática y resumen mensual con insights.'],
  ];
  sh.getRange(1, 1, params.length, 3).setValues(params);

  // Forzar B2 (Mes activo) a formato texto puro para que Sheets NUNCA lo interprete como fecha.
  // Si lo guardara como Date, LEFT/MID en _Calc devolverían el número de serie y todos los SUMIFS fallarían.
  sh.getRange('B2').setNumberFormat('@').setValue(mesActualISO());

  // Validación desplegable: últimos 24 meses (12 anteriores + actual + 11 futuros) en formato AAAA-MM.
  // Permite cambiar de mes en un clic. Texto puro, sin interpretación como fecha.
  validarLista(sh, 'B2', listaMesesISO(12, 11));

  // Cabecera
  sh.getRange('A1:C1')
    .setFontWeight('bold').setFontColor(COLOR.cabTxt).setBackground(COLOR.cab)
    .setVerticalAlignment('middle');

  // Estilo del bloque
  const ultFila = params.length;
  sh.getRange(2, 1, ultFila - 1, 1).setFontWeight('bold').setFontColor(COLOR.tinta);
  sh.getRange(2, 2, ultFila - 1, 1).setBackground(COLOR.acentoSuave).setFontColor(COLOR.tinta);
  sh.getRange(2, 3, ultFila - 1, 1).setFontColor(COLOR.texto).setWrap(true);
  sh.getRange(1, 1, ultFila, 3).setBorder(true, true, true, true, true, true, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  // Formatos específicos por fila
  sh.getRange('B3').setNumberFormat('#,##0 €');
  sh.getRange('B7').setNumberFormat('0%');             // % real, valor decimal 0..1
  sh.getRange('B8').setNumberFormat('0" meses"');
  sh.getRange('B9').setNumberFormat('#,##0 €');

  // Validación del modo de aportación
  validarLista(sh, 'B6', ['50/50', 'Proporcional', 'Custom']);

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

  // Anchos y diseño
  sh.setColumnWidth(1, 290);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 420);
  sh.setColumnWidths(4, 3, 180);
  sh.setRowHeight(1, 28);
  sh.setFrozenRows(1);
  sh.setHiddenGridlines(true);
  sh.getRange('A1:F60').setFontFamily(FUENTE);

  // Nota AYUDA fuera del rango D:F (D2:F lo apila _Calc como lista de categorías).
  sh.getRange('A14').setValue('💡 Edita las categorías de las columnas D/E/F a tu gusto. El bot y los desplegables las usan.')
    .setFontSize(9).setFontStyle('italic').setFontColor(COLOR.tenue)
    .setHorizontalAlignment('left');
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

  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 160);
  sh.setColumnWidth(3, 100);
  sh.setColumnWidth(4, 170);
  sh.setColumnWidth(5, 300);
  sh.setColumnWidth(6, 130);
  sh.setColumnWidth(7, 320);

  // Formato fecha+hora: las variables suelen registrarse al momento (super, restaurante…)
  // y la hora aporta contexto. En fijos/ingresos la hora será 00:00, irrelevante.
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd HH:mm');
  sh.getRange('F:F').setNumberFormat('#,##0.00 €');
  sh.getRange('A2:G').setFontColor(COLOR.texto).setVerticalAlignment('middle');

  validarLista(sh, 'B2:B', TIPOS);
  validarLista(sh, 'C2:C', PERSONAS);
  validarRango(sh, 'D2:D', `${HOJAS.AJUSTES}!D2:F`, true);

  // Libro mayor vacío: listo para registrar datos reales (desde el bot o a mano).
  const banda = sh.getRange(1, 1, 200, 7)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banda.setHeaderRowColor(COLOR.cab);
  banda.setFirstRowColor(COLOR.panel);
  banda.setSecondRowColor(COLOR.cebra);

  sh.setFrozenRows(1);
  sh.getRange('A:G').setFontFamily(FUENTE);
}

/* ====================== HOJA: TICKETS (productos desglosados de tickets de super) ====================== */

function crearHojaTickets(ss) {
  const sh = ss.insertSheet(HOJAS.TICKETS);

  const cab = ['Fecha', 'Tienda', 'Producto', 'Cantidad', 'Precio (€)', 'Precio/ud (€)', 'Categoría', 'Ticket #'];
  sh.getRange(1, 1, 1, 8).setValues([cab]);
  sh.getRange('A1:H1')
    .setFontWeight('bold').setFontColor(COLOR.cabTxt).setBackground(COLOR.cab)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.setRowHeight(1, 30);

  sh.setColumnWidth(1, 120);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 280);
  sh.setColumnWidth(4, 80);
  sh.setColumnWidth(5, 100);
  sh.setColumnWidth(6, 110);
  sh.setColumnWidth(7, 140);
  sh.setColumnWidth(8, 90);

  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd HH:mm');
  sh.getRange('E:F').setNumberFormat('#,##0.00 €');
  sh.getRange('A2:H').setFontColor(COLOR.texto).setVerticalAlignment('middle');

  const banda = sh.getRange(1, 1, 200, 8)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banda.setHeaderRowColor(COLOR.cab);
  banda.setFirstRowColor(COLOR.panel);
  banda.setSecondRowColor(COLOR.cebra);

  sh.setFrozenRows(1);
  sh.getRange('A:H').setFontFamily(FUENTE);
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

  // Sin metas de ejemplo: rellena Concepto · Meta · Aportado · Fecha en cada fila.
  // Las columnas % Progreso / Estado / Barra se calculan solas.
  for (let r = 2; r <= 50; r++) {
    sh.getRange(r, 4).setFormula(`=IFERROR(C${r}/B${r};0)`);
    sh.getRange(r, 6).setFormula(`=IF(B${r}="";"";IF(C${r}>=B${r};"✅ Cumplido";"En curso"))`);
    sh.getRange(r, 7).setFormula(
      `=IF(B${r}="";"";REPT("█";ROUND(MIN(C${r}/B${r};1)*10;0))&REPT("░";10-ROUND(MIN(C${r}/B${r};1)*10;0)))`
    );
  }

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

  const reglaFecha = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  sh.getRange('E2:E').setDataValidation(reglaFecha);

  const banda = sh.getRange(1, 1, 50, 7)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banda.setHeaderRowColor(COLOR.cab);
  banda.setFirstRowColor(COLOR.panel);
  banda.setSecondRowColor(COLOR.cebra);

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

/* ====================== HOJA: _Calc (oculta) ======================
 * Mapa de columnas (SIN solapamientos):
 *   A:B   Métricas. B1..B20 = CONTRATO con el bot (no mover posición ni significado).
 *   D:E   Categoría | importe del mes (suma de los 3 tipos de gasto).
 *   G:H   Categorías ORDENADAS desc por importe (para "Dónde se gasta").
 *   J:P   Serie 12 meses: mes, ahorro, acumulado, ingreso, gasto, fijos, variables.
 */
function crearCalc(ss) {
  const sh = ss.insertSheet(HOJAS.CALC);
  const MOV = HOJAS.MOVIMIENTOS;
  const AJ = HOJAS.AJUSTES;

  // Mes activo robusto a interpretación como fecha.
  const mesActivoRaw = `VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE)`;
  const mesActivo = `IF(ISNUMBER(${mesActivoRaw});TEXT(${mesActivoRaw};"yyyy-mm");${mesActivoRaw})`;
  const inicioMes = `DATE(VALUE(LEFT(${mesActivo};4));VALUE(MID(${mesActivo};6;2));1)`;
  const finMes = `EOMONTH(${inicioMes};0)+1`;
  const anio = `VALUE(LEFT(${mesActivo};4))`;

  const sumar = (tipo, persona) => {
    let f = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};${MOV}!B:B;"${tipo}"`;
    if (persona) f += `;${MOV}!C:C;"${persona}"`;
    f += `)`;
    return `IFERROR(${f};0)`;
  };
  const objetivo = `IFERROR(VLOOKUP("Objetivo ahorro mensual conjunto (€)";${AJ}!A:B;2;FALSE);0)`;

  /* ===== B1..B20: CONTRATO con el bot (enviarResumen lee B1:B20) ===== */
  const metricas = [
    ['Ingresos FM',                  `=${sumar('Ingreso', 'FM')}`],
    ['Ingresos Lucía',               `=${sumar('Ingreso', 'Lucía')}`],
    ['Ingresos total',               '=B1+B2'],
    ['Aportación al bote FM',        `=${sumar('Aportación', 'FM')}`],
    ['Aportación al bote Lucía',     `=${sumar('Aportación', 'Lucía')}`],
    ['Aportación al bote total',     '=B4+B5'],
    ['Gastos compartidos fijos',     `=${sumar('Compartido fijo', null)}`],
    ['Gastos compartidos variables', `=${sumar('Compartido variable', null)}`],
    ['Gastos compartidos total',     '=B7+B8'],
    ['Gastos individuales FM',       `=${sumar('Individual', 'FM')}`],
    ['Gastos individuales Lucía',    `=${sumar('Individual', 'Lucía')}`],
    ['Gastos totales del mes',       '=B9+B10+B11'],
    ['Saldo FM (en bolsillo)',       '=B1-B4-B10-B42'],
    ['Saldo Lucía (en bolsillo)',    '=B2-B5-B11-B43'],
    ['Queda en el bote',             '=B6-B9-B41'],
    ['Objetivo ahorro',              `=${objetivo}`],
    ['Ahorro real del mes',          '=B41+B42+B43'],
    ['% objetivo cumplido',          '=IFERROR(B17/B16;0)'],
    ['Ahorro acumulado del año',     ''],
    ['Balance del mes',              '=B3-B12'],
  ];
  sh.getRange(1, 1, metricas.length, 2).setValues(metricas);

  /* ===== B21..B47: auxiliares para el Panel (el bot NO las lee) ===== */
  // B21..B27: auxiliares contiguas.
  const extra = [
    ['% aporte real bote FM',    '=IFERROR(B4/B6;0)'],
    ['% aporte real bote Lucía', '=IFERROR(B5/B6;0)'],
    ['% ingreso FM al bote',     '=IFERROR(B4/B1;0)'],
    ['% ingreso Lucía al bote',  '=IFERROR(B5/B2;0)'],
    ['Tasa de ahorro real',      '=IFERROR(B17/B3;0)'],
    ['Tasa de ahorro objetivo',  `=IFERROR(VLOOKUP("Tasa de ahorro objetivo (%)";${AJ}!A:B;2;FALSE);2/10)`],
    ['Modo aportación bote',     `=IFERROR(VLOOKUP("Modo aportación al bote";${AJ}!A:B;2;FALSE);"50/50")`],
  ];
  sh.getRange(21, 1, extra.length, 2).setValues(extra);

  // B41..B47: por dirección explícita (evita errores de desplazamiento de fila).
  sh.getRange('A41').setValue('Ahorro compartido del mes');
  sh.getRange('B41').setFormula(`=${sumar('Ahorro', 'Bote')}`);
  sh.getRange('A42').setValue('Ahorro FM del mes');
  sh.getRange('B42').setFormula(`=${sumar('Ahorro', 'FM')}`);
  sh.getRange('A43').setValue('Ahorro Lucía del mes');
  sh.getRange('B43').setFormula(`=${sumar('Ahorro', 'Lucía')}`);
  sh.getRange('A44').setValue('Ahorro total del mes');
  sh.getRange('B44').setFormula('=B41+B42+B43');
  sh.getRange('A46').setValue('Inicio del mes activo');
  sh.getRange('B46').setFormula(`=${inicioMes}`);
  sh.getRange('A47').setValue('Fin del mes activo');
  sh.getRange('B47').setFormula(`=${finMes}`);

  // Formatos columna B
  sh.getRange('B1:B17').setNumberFormat('#,##0.00 €');
  sh.getRange('B18').setNumberFormat('0.0%');
  sh.getRange('B19:B20').setNumberFormat('#,##0.00 €');
  sh.getRange('B21:B26').setNumberFormat('0.0%');
  sh.getRange('B41:B44').setNumberFormat('#,##0.00 €');
  sh.getRange('B46:B47').setNumberFormat('yyyy-mm-dd');
  sh.setColumnWidth(1, 260);
  sh.setColumnWidth(2, 130);
  sh.getRange('A1:B47').setFontFamily(FUENTE).setFontColor(COLOR.texto);

  /* ===== D:E categoría | importe del mes ===== */
  sh.getRange('D1:E1').setValues([['Categoría', 'Importe mes']]);
  const catsApiladas = `{${AJ}!D2:D;${AJ}!E2:E;${AJ}!F2:F}`;
  sh.getRange('D2').setFormula(`=IFERROR(SORT(UNIQUE(FILTER(${catsApiladas};${catsApiladas}<>"")));"")`);
  sh.getRange('E2').setFormula(
    `=ARRAYFORMULA(IF(D2:D40="";"";` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&B46;${MOV}!A:A;"<"&B47;${MOV}!D:D;D2:D40;${MOV}!B:B;"Compartido fijo");0)+` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&B46;${MOV}!A:A;"<"&B47;${MOV}!D:D;D2:D40;${MOV}!B:B;"Compartido variable");0)+` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&B46;${MOV}!A:A;"<"&B47;${MOV}!D:D;D2:D40;${MOV}!B:B;"Individual");0)))`
  );
  sh.getRange('E2:E40').setNumberFormat('#,##0.00 €');

  /* ===== G:H categorías ordenadas desc (para "Dónde se gasta") ===== */
  sh.getRange('G1:H1').setValues([['Categoría', 'Importe']]);
  sh.getRange('G2').setFormula(
    `=IFERROR(QUERY(D2:E40;"select Col1, Col2 where Col2 > 0 order by Col2 desc";0);"")`
  );
  sh.getRange('H2:H40').setNumberFormat('#,##0.00 €');

  /* ===== J:P serie 12 meses del año activo ===== */
  sh.getRange('J1:P1').setValues([['Mes', 'Ahorro', 'Acumulado', 'Ingreso', 'Gasto', 'Fijos', 'Variables']]);
  for (let m = 1; m <= 12; m++) {
    const r = m + 1;
    const ini = `DATE(${anio};${m};1)`;
    const fin = `DATE(${anio};${m + 1};1)`;
    sh.getRange(r, 10).setFormula(`=TEXT(${ini};"yyyy-mm")`);
    const ahoMes = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Ahorro")`;
    const ingMes = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Ingreso")`;
    const fijMes = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Compartido fijo")`;
    const varMes = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Compartido variable")`;
    const indMes = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Individual")`;
    const gastTot = `(${fijMes})+(${varMes})+(${indMes})`;
    sh.getRange(r, 11).setFormula(`=IFERROR(${ahoMes};0)`);
    sh.getRange(r, 12).setFormula(m === 1 ? '=K2' : `=L${r - 1}+K${r}`);
    sh.getRange(r, 13).setFormula(`=IFERROR(${ingMes};0)`);
    sh.getRange(r, 14).setFormula(`=IFERROR(${gastTot};0)`);
    sh.getRange(r, 15).setFormula(`=IFERROR(${fijMes};0)`);
    sh.getRange(r, 16).setFormula(`=IFERROR(${varMes};0)`);
  }
  sh.getRange('K2:P13').setNumberFormat('#,##0.00 €');

  // B19 = ahorro explícito acumulado del año hasta el mes activo.
  const idxMes = `VALUE(MID(${mesActivo};6;2))`;
  sh.getRange('B19').setFormula(`=IFERROR(INDEX(L2:L13;${idxMes});0)`);

  /* ===== Desglose por categoría POR ÁMBITO (para los donuts) =====
   * Cada ámbito: par contiguo [categoría, importe] → QUERY filtrado >0 y ordenado.
   *   Compartida: R:S (raw) → U:V (donut)   [Compartido fijo + variable]
   *   FM:         W:X (raw) → Z:AA (donut)   [Individual + FM]
   *   Lucía:      AC:AD (raw) → AF:AG (donut) [Individual + Lucía]
   */
  const breakdown = (colCatRaw, colImpRaw, tipos, persona, colCatOut, colImpOut) => {
    sh.getRange(`${colCatRaw}2`).setFormula(`=ARRAYFORMULA(D2:D40)`);
    const suma = tipos.map(t => {
      let s = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&B46;${MOV}!A:A;"<"&B47;${MOV}!D:D;${colCatRaw}2:${colCatRaw}40;${MOV}!B:B;"${t}"`;
      if (persona) s += `;${MOV}!C:C;"${persona}"`;
      return `IFERROR(${s});0)`;
    }).join('+');
    sh.getRange(`${colImpRaw}2`).setFormula(`=ARRAYFORMULA(IF(${colCatRaw}2:${colCatRaw}40="";"";${suma}))`);
    sh.getRange(`${colImpRaw}2:${colImpRaw}40`).setNumberFormat('#,##0.00 €');
    sh.getRange(`${colCatOut}1`).setValue('Categoría');
    sh.getRange(`${colImpOut}1`).setValue('Importe');
    sh.getRange(`${colCatOut}2`).setFormula(
      `=IFERROR(QUERY(${colCatRaw}2:${colImpRaw}40;"select Col1, Col2 where Col2 > 0 order by Col2 desc";0);"")`
    );
    sh.getRange(`${colImpOut}2:${colImpOut}11`).setNumberFormat('#,##0.00 €');
  };
  breakdown('R', 'S', ['Compartido fijo', 'Compartido variable'], null, 'U', 'V');
  breakdown('W', 'X', ['Individual'], 'FM', 'Z', 'AA');
  breakdown('AC', 'AD', ['Individual'], 'Lucía', 'AF', 'AG');

  // Anchos auxiliares.
  sh.setColumnWidths(4, 13, 90);
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

/* ====================== HOJAS-DASHBOARD (estilo "Simple Budget") ======================
 * Tres hojas con el mismo estilo: Compartida, FM, Lucía.
 * Fondo crema, cabeceras pastel (verde salvia ingresos · rosa gastos · arena totales),
 * tarjetas KPI, tablas de movimientos desglosados con TOTAL, resumen y donut.
 * Las listas se leen en vivo de Movimientos con FILTER (reserva + overflow).
 */

function crearHojaCompartida(ss) {
  const sh = ss.insertSheet(HOJAS.COMPARTIDA);
  const C = `'${HOJAS.CALC}'`;
  lienzo(sh);
  let r = tituloHoja(sh, 'Cuenta Compartida');
  r += 1;

  // KPI cards
  tarjetaKPI(sh, r, 2, 4, 'APORTADO', COLOR.verde, `=${C}!B6`, false);
  tarjetaKPI(sh, r, 5, 7, 'GASTOS', COLOR.rosa, `=${C}!B7+${C}!B8`, false);
  tarjetaKPI(sh, r, 8, 11, 'RESULTADO', COLOR.arena, `=${C}!B6-${C}!B7-${C}!B8`, true);
  r += 3;

  // Aportaciones
  r = cabeceraTabla(sh, r, 'APORTACIONES AL BOTE', COLOR.verde);
  r = filaSimple(sh, r, 'FM', `=${C}!B4`, false, 0);
  r = filaSimple(sh, r, 'Lucía', `=${C}!B5`, false, 1);
  r = filaTotal(sh, r, 'Total aportado', `=${C}!B6`, COLOR.verdeClaro, false);
  r += 1;

  // Gastos fijos
  r = cabeceraTabla(sh, r, 'GASTOS FIJOS', COLOR.rosa);
  r = tablaMov(sh, r, 'Compartido fijo', null, `=${C}!B7`, RESERVA.COMP_FIJO);
  r += 1;

  // Gastos variables
  r = cabeceraTabla(sh, r, 'GASTOS VARIABLES', COLOR.rosa);
  r = tablaMov(sh, r, 'Compartido variable', null, `=${C}!B8`, RESERVA.COMP_VAR);
  r += 1;

  // Resumen
  r = cabeceraTabla(sh, r, 'RESUMEN', COLOR.arena);
  r = filaSimple(sh, r, 'Aportado al bote', `=${C}!B6`, false, 0);
  r = filaSimple(sh, r, 'Gastos fijos', `=-1*${C}!B7`, false, 1);
  r = filaSimple(sh, r, 'Gastos variables', `=-1*${C}!B8`, false, 0);
  r = filaTotal(sh, r, 'RESULTADO (ingresos − gastos)', `=${C}!B6-${C}!B7-${C}!B8`, COLOR.arenaClaro, true);
  r += 1;

  // Donut
  r = cabeceraTabla(sh, r, 'GASTOS POR CATEGORÍA', COLOR.rosa);
  donut(sh, ss, r, 'U1:V11');

  sh.setFrozenRows(3);
}

function crearHojaPersona(ss, nombreHoja, titulo, persona) {
  const sh = ss.insertSheet(nombreHoja);
  const C = `'${HOJAS.CALC}'`;
  lienzo(sh);
  let r = tituloHoja(sh, titulo);
  r += 1;

  const ingreso  = persona === 'FM' ? `${C}!B1`  : `${C}!B2`;
  const aport    = persona === 'FM' ? `${C}!B4`  : `${C}!B5`;
  const gastoInd = persona === 'FM' ? `${C}!B10` : `${C}!B11`;
  const ahorro   = persona === 'FM' ? `${C}!B42` : `${C}!B43`;
  const bolsillo = persona === 'FM' ? `${C}!B13` : `${C}!B14`;
  const donutRange = persona === 'FM' ? 'Z1:AA11' : 'AF1:AG11';

  // KPI cards
  tarjetaKPI(sh, r, 2, 4, 'INGRESO', COLOR.verde, `=${ingreso}`, false);
  tarjetaKPI(sh, r, 5, 7, 'SALIDAS', COLOR.rosa, `=${aport}+${gastoInd}+${ahorro}`, false);
  tarjetaKPI(sh, r, 8, 11, 'EN BOLSILLO', COLOR.arena, `=${bolsillo}`, true);
  r += 3;

  // Ingreso
  r = cabeceraTabla(sh, r, 'INGRESO', COLOR.verde);
  r = filaTotal(sh, r, 'Ingreso del mes', `=${ingreso}`, COLOR.verdeClaro, false);
  r += 1;

  // Gastos individuales
  r = cabeceraTabla(sh, r, 'GASTOS INDIVIDUALES', COLOR.rosa);
  r = tablaMov(sh, r, 'Individual', persona, `=${gastoInd}`, RESERVA.IND);
  r += 1;

  // Ahorro
  r = cabeceraTabla(sh, r, 'AHORRO', COLOR.verde);
  r = tablaMov(sh, r, 'Ahorro', persona, `=${ahorro}`, RESERVA.IND_AHO);
  r += 1;

  // Resumen
  r = cabeceraTabla(sh, r, 'RESUMEN', COLOR.arena);
  r = filaSimple(sh, r, 'Ingreso', `=${ingreso}`, false, 0);
  r = filaSimple(sh, r, 'Aportación al bote', `=-1*${aport}`, false, 1);
  r = filaSimple(sh, r, 'Gastos individuales', `=-1*${gastoInd}`, false, 0);
  r = filaSimple(sh, r, 'Ahorro', `=-1*${ahorro}`, false, 1);
  r = filaTotal(sh, r, 'EN BOLSILLO', `=${bolsillo}`, COLOR.arenaClaro, true);
  r += 1;

  // Donut
  r = cabeceraTabla(sh, r, 'GASTOS POR CATEGORÍA', COLOR.rosa);
  donut(sh, ss, r, donutRange);

  sh.setFrozenRows(3);
}

/* ---------- HOJA: RESUMEN ANUAL ---------- */
function crearHojaAnual(ss) {
  const sh = ss.insertSheet(HOJAS.ANUAL);
  const C = `'${HOJAS.CALC}'`;
  lienzo(sh);
  let r = tituloHoja(sh, 'Resumen Anual');
  r += 1;

  // KPI totales del año
  tarjetaKPI(sh, r, 2, 4, 'INGRESOS AÑO', COLOR.verde, `=SUM(${C}!M2:M13)`, false);
  tarjetaKPI(sh, r, 5, 7, 'GASTOS AÑO', COLOR.rosa, `=SUM(${C}!N2:N13)`, false);
  tarjetaKPI(sh, r, 8, 11, 'AHORRO AÑO', COLOR.arena, `=SUM(${C}!K2:K13)`, true);
  r += 3;

  // Tabla 12 meses
  r = cabeceraTabla(sh, r, 'EVOLUCIÓN MENSUAL', COLOR.arena);
  // Encabezado de columnas: Mes B:C · Ingresos D:E · Gastos F:G · Ahorro H:I · Resultado J:K
  sh.setRowHeight(r, 20);
  sh.getRange(r, 2, 1, 10).setBackground(COLOR.cebra)
    .setBorder(false, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  [[2, 2, 'Mes', 'left'], [4, 5, 'Ingresos', 'right'], [6, 7, 'Gastos', 'right'],
   [8, 9, 'Ahorro', 'right'], [10, 11, 'Resultado', 'right']].forEach(([c1, c2, txt, al]) => {
    sh.getRange(r, c1, 1, c2 - c1 + 1).merge().setValue(txt)
      .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
      .setVerticalAlignment('middle').setHorizontalAlignment(al);
  });
  r++;

  for (let m = 0; m < 12; m++) {
    const cr = 2 + m; // fila en _Calc (J2..J13)
    sh.setRowHeight(r, 22);
    const fondo = m % 2 === 0 ? COLOR.panel : COLOR.cebra;
    sh.getRange(r, 2, 1, 10).setBackground(fondo)
      .setBorder(false, true, false, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 2, 1, 2).merge().setFormula(`=${C}!J${cr}`)
      .setFontSize(10).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.getRange(r, 4, 1, 2).merge().setFormula(`=${C}!M${cr}`).setNumberFormat('#,##0 €')
      .setFontSize(10).setFontColor(COLOR.texto).setVerticalAlignment('middle').setHorizontalAlignment('right');
    sh.getRange(r, 6, 1, 2).merge().setFormula(`=${C}!N${cr}`).setNumberFormat('#,##0 €')
      .setFontSize(10).setFontColor(COLOR.texto).setVerticalAlignment('middle').setHorizontalAlignment('right');
    sh.getRange(r, 8, 1, 2).merge().setFormula(`=${C}!K${cr}`).setNumberFormat('#,##0 €')
      .setFontSize(10).setFontColor(COLOR.texto).setVerticalAlignment('middle').setHorizontalAlignment('right');
    const res = sh.getRange(r, 10, 1, 2).merge().setFormula(`=${C}!M${cr}-${C}!N${cr}`).setNumberFormat('#,##0 €')
      .setFontSize(10).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('right');
    aplicarPositivoNegativo(sh, sh.getRange(r, 10, 1, 2));
    r++;
  }

  // Total del año
  sh.setRowHeight(r, 26);
  sh.getRange(r, 2, 1, 10).setBackground(COLOR.arenaClaro)
    .setBorder(true, true, true, true, false, false, COLOR.arena, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(r, 2, 1, 2).merge().setValue('TOTAL').setFontSize(10).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle');
  sh.getRange(r, 4, 1, 2).merge().setFormula(`=SUM(${C}!M2:M13)`).setNumberFormat('#,##0 €').setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('right');
  sh.getRange(r, 6, 1, 2).merge().setFormula(`=SUM(${C}!N2:N13)`).setNumberFormat('#,##0 €').setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('right');
  sh.getRange(r, 8, 1, 2).merge().setFormula(`=SUM(${C}!K2:K13)`).setNumberFormat('#,##0 €').setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('right');
  sh.getRange(r, 10, 1, 2).merge().setFormula(`=SUM(${C}!M2:M13)-SUM(${C}!N2:N13)`).setNumberFormat('#,##0 €').setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('right');
  aplicarPositivoNegativo(sh, sh.getRange(r, 10, 1, 2));
  r += 2;

  // Gráfico anual
  r = cabeceraTabla(sh, r, 'GRÁFICO ANUAL', COLOR.verde);
  const calc = ss.getSheetByName(HOJAS.CALC);
  const ch = sh.newChart().setChartType(Charts.ChartType.COMBO)
    .addRange(calc.getRange('J1:J13'))
    .addRange(calc.getRange('M1:M13'))
    .addRange(calc.getRange('N1:N13'))
    .addRange(calc.getRange('K1:K13'))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setOption('title', '')
    .setOption('legend', { position: 'top', alignment: 'center', textStyle: { color: COLOR.texto, fontSize: 10, fontName: FUENTE } })
    .setOption('series', {
      0: { type: 'line', color: '#8FAE8B', lineWidth: 2, pointSize: 3 },
      1: { type: 'line', color: '#E0A9A6', lineWidth: 2, pointSize: 3 },
      2: { type: 'bars', color: '#C9B68F' },
    })
    .setOption('hAxis', { textStyle: { color: COLOR.tenue, fontSize: 9 } })
    .setOption('vAxis', { textStyle: { color: COLOR.tenue, fontSize: 9 }, format: '#,##0 €', gridlines: { color: COLOR.borde } })
    .setOption('backgroundColor', COLOR.fondo)
    .setOption('chartArea', { left: 60, top: 40, width: '88%', height: '72%' })
    .setOption('width', 900).setOption('height', 300)
    .setPosition(r, 2, 0, 0).build();
  sh.insertChart(ch);

  sh.setFrozenRows(3);
}

/* ---------- HELPERS DE ESTILO ---------- */

function lienzo(sh) {
  sh.setHiddenGridlines(true);
  sh.getRange('A1:L400').setBackground(COLOR.fondo).setFontFamily(FUENTE);
  sh.setColumnWidth(1, 28);
  sh.setColumnWidths(2, 10, 84);
  sh.setColumnWidth(12, 28);
}

function tituloHoja(sh, titulo) {
  const AJ = HOJAS.AJUSTES;
  sh.setRowHeight(1, 16);
  sh.setRowHeight(2, 40);
  sh.getRange('B2:G2').merge().setValue(titulo)
    .setFontFamily(FUENTE_TIT).setFontSize(28).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setBackground(COLOR.fondo);
  sh.getRange('H2:K2').merge()
    .setFormula(`="Mes  "&IF(ISNUMBER(VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE));TEXT(VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE);"yyyy-mm");VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE))`)
    .setFontFamily(FUENTE_TIT).setFontSize(13).setFontColor(COLOR.tenue)
    .setHorizontalAlignment('right').setVerticalAlignment('middle').setBackground(COLOR.fondo);
  sh.setRowHeight(3, 10);
  return 4;
}

function tarjetaKPI(sh, r, c1, c2, titulo, color, valorFormula, signo) {
  const n = c2 - c1 + 1;
  sh.setRowHeight(r, 22);
  sh.getRange(r, c1, 1, n).merge().setValue(titulo)
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.cabTxt)
    .setBackground(color).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(r + 1, 46);
  sh.getRange(r + 1, c1, 1, n).merge().setFormula(valorFormula).setNumberFormat('#,##0.00 €')
    .setFontSize(20).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setBackground(COLOR.panel).setHorizontalAlignment('center').setVerticalAlignment('middle')
    .setBorder(false, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  if (signo) aplicarPositivoNegativo(sh, sh.getRange(r + 1, c1, 1, n));
}

function cabeceraTabla(sh, r, titulo, color) {
  sh.setRowHeight(r, 26);
  sh.getRange(r, 2, 1, 10).merge().setValue(titulo)
    .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.cabTxt)
    .setBackground(color).setHorizontalAlignment('center').setVerticalAlignment('middle');
  return r + 1;
}

/** Tabla de movimientos desglosados: Concepto · Categoría · Fecha · Importe + TOTAL. */
function tablaMov(sh, r, tipo, persona, totalRef, reserva) {
  const MOV = HOJAS.MOVIMIENTOS;
  const C = `'${HOJAS.CALC}'`;
  let cond = `${MOV}!B2:B="${tipo}"`;
  if (persona) cond += `;${MOV}!C2:C="${persona}"`;
  const flt = `FILTER(HSTACK(${MOV}!E2:E;${MOV}!D2:D;${MOV}!A2:A;${MOV}!F2:F);${cond};${MOV}!A2:A>=${C}!B46;${MOV}!A2:A<${C}!B47)`;
  let cntA = `${MOV}!B:B;"${tipo}";${MOV}!A:A;">="&${C}!B46;${MOV}!A:A;"<"&${C}!B47`;
  if (persona) cntA += `;${MOV}!C:C;"${persona}"`;
  const cnt = `COUNTIFS(${cntA})`;

  // Encabezado de columnas
  sh.setRowHeight(r, 20);
  sh.getRange(r, 2, 1, 10).setBackground(COLOR.cebra)
    .setBorder(false, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(r, 2, 1, 4).merge().setValue('Concepto').setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue).setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.getRange(r, 6, 1, 2).merge().setValue('Categoría').setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue).setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.getRange(r, 8, 1, 2).merge().setValue('Fecha').setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue).setVerticalAlignment('middle').setHorizontalAlignment('center');
  sh.getRange(r, 10, 1, 2).merge().setValue('Importe').setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue).setVerticalAlignment('middle').setHorizontalAlignment('right');
  r++;

  for (let i = 0; i < reserva; i++) {
    const k = i + 1;
    sh.setRowHeight(r, 22);
    const fondo = i % 2 === 0 ? COLOR.panel : COLOR.cebra;
    sh.getRange(r, 2, 1, 10).setBackground(fondo)
      .setBorder(false, true, false, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 2, 1, 4).merge().setFormula(`=IFERROR(INDEX(${flt};${k};1);"")`)
      .setFontSize(10).setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.getRange(r, 6, 1, 2).merge().setFormula(`=IFERROR(INDEX(${flt};${k};2);"")`)
      .setFontSize(9).setFontColor(COLOR.texto).setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.getRange(r, 8, 1, 2).merge().setFormula(`=IFERROR(INDEX(${flt};${k};3);"")`)
      .setNumberFormat('dd-mmm HH:mm').setFontSize(9).setFontColor(COLOR.tenue).setVerticalAlignment('middle').setHorizontalAlignment('center');
    sh.getRange(r, 10, 1, 2).merge().setFormula(`=IFERROR(INDEX(${flt};${k};4);"")`)
      .setNumberFormat('#,##0.00 €').setFontSize(10).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('right');
    r++;
  }

  // Overflow
  sh.setRowHeight(r, 16);
  sh.getRange(r, 2, 1, 10).merge()
    .setFormula(`=IF(${cnt}>${reserva};"+"&(${cnt}-${reserva})&" más · ver Movimientos";"")`)
    .setFontSize(8).setFontStyle('italic').setFontColor(COLOR.tenue)
    .setHorizontalAlignment('right').setBackground(COLOR.fondo);
  r++;

  // Total
  sh.setRowHeight(r, 24);
  sh.getRange(r, 2, 1, 10).setBackground(COLOR.arenaClaro)
    .setBorder(true, true, true, true, false, false, COLOR.arena, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(r, 2, 1, 7).merge().setValue('TOTAL')
    .setFontSize(10).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle');
  sh.getRange(r, 10, 1, 2).merge().setFormula(totalRef).setNumberFormat('#,##0.00 €')
    .setFontSize(12).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('right');
  return r + 1;
}

function filaSimple(sh, r, label, valFormula, fuerte, idx) {
  sh.setRowHeight(r, 24);
  const fondo = idx % 2 === 0 ? COLOR.panel : COLOR.cebra;
  sh.getRange(r, 2, 1, 10).setBackground(fondo)
    .setBorder(false, true, false, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(r, 2, 1, 7).merge().setValue(label)
    .setFontSize(11).setFontWeight(fuerte ? 'bold' : 'normal').setFontColor(COLOR.texto)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.getRange(r, 10, 1, 2).merge().setFormula(valFormula).setNumberFormat('#,##0.00 €')
    .setFontSize(11).setFontWeight(fuerte ? 'bold' : 'normal').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setHorizontalAlignment('right');
  return r + 1;
}

function filaTotal(sh, r, label, valFormula, bg, signo) {
  sh.setRowHeight(r, 30);
  sh.getRange(r, 2, 1, 10).setBackground(bg || COLOR.arenaClaro)
    .setBorder(true, true, true, true, false, false, COLOR.arena, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(r, 2, 1, 7).merge().setValue(label)
    .setFontSize(12).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle');
  sh.getRange(r, 10, 1, 2).merge().setFormula(valFormula).setNumberFormat('#,##0.00 €')
    .setFontSize(14).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setHorizontalAlignment('right');
  if (signo) aplicarPositivoNegativo(sh, sh.getRange(r, 10, 1, 2));
  return r + 1;
}

function donut(sh, ss, anchorRow, rangeA1) {
  const calc = ss.getSheetByName(HOJAS.CALC);
  const ch = sh.newChart().setChartType(Charts.ChartType.PIE)
    .addRange(calc.getRange(rangeA1)).setNumHeaders(1)
    .setOption('pieHole', 0.6)
    .setOption('title', '')
    .setOption('legend', { position: 'right', textStyle: { color: COLOR.texto, fontSize: 10, fontName: FUENTE } })
    .setOption('colors', COLOR.donut)
    .setOption('pieSliceText', 'none')
    .setOption('backgroundColor', COLOR.fondo)
    .setOption('chartArea', { left: 10, top: 10, width: '94%', height: '88%' })
    .setOption('width', 660).setOption('height', 250)
    .setPosition(anchorRow, 2, 0, 0).build();
  sh.insertChart(ch);
}

/* ====================== FORMATO CONDICIONAL ====================== */

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

/* ====================== HELPERS GENERALES ====================== */

function validarLista(sh, rangoA1, valores) {
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInList(valores, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(rangoA1).setDataValidation(regla);
}

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

function mesActualISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

/** Lista de meses AAAA-MM centrada en el mes actual (atras meses hacia atrás, futuro hacia adelante). */
function listaMesesISO(atras, futuro) {
  const hoy = new Date();
  const meses = [];
  for (let i = -atras; i <= futuro; i++) {
    const d = new Date(hoy.getFullYear(), hoy.getMonth() + i, 1);
    meses.push(`${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`);
  }
  return meses;
}
