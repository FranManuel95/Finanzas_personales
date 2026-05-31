/**
 * CrearDashboard.gs
 * ------------------------------------------------------------------
 * Ejecuta crearDashboard() UNA SOLA VEZ desde el editor de Apps Script.
 * Genera todas las pestañas del libro con cabeceras, formato y fórmulas.
 *
 * Si vuelves a ejecutarla borra y regenera las pestañas (se pierden datos).
 * Usa recrearDashboardSeguro() si quieres conservar lo existente.
 * ------------------------------------------------------------------
 */

const HOJAS = {
  DASHBOARD: 'Dashboard',
  CONFIG: 'Config',
  CATEGORIAS: 'Categorias',
  INGRESOS: 'Ingresos',
  APORTACIONES: 'Aportaciones_Bote',
  GC_FIJOS: 'Gastos_Compartidos_Fijos',
  GC_VARIABLES: 'Gastos_Compartidos_Variables',
  G_FM: 'Gastos_FM',
  G_LUCIA: 'Gastos_Lucia',
  AHORRO: 'Ahorro',
  OBJETIVOS_LP: 'Objetivos_Largo_Plazo',
  ANALITICA_ANUAL: 'Analitica_Anual',
  LOG_BOT: 'Log_Bot',
};

/* ============ PALETA ============ */
const COLOR = {
  fondo:        '#F4F6FA',
  panel:        '#FFFFFF',
  acento:       '#2C5282', // azul corporativo
  acentoSuave:  '#E8EEF7',
  textoFuerte:  '#1A2A45',
  textoSuave:   '#5A6B85',
  verde:        '#1E8E5A',
  verdeSuave:   '#E1F4EA',
  rojo:         '#C5372B',
  rojoSuave:    '#FCE5E2',
  naranja:      '#D97706',
  naranjaSuave: '#FEF1DD',
  borde:        '#D6DEEB',
  cabHeader:    '#1A2A45',
  cabHeaderTxt: '#FFFFFF',
};

function crearDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Guarda el ID del Sheet para que el Web App pueda accederlo.
  PropertiesService.getScriptProperties().setProperty('SHEET_ID', ss.getId());

  // Crea una pestaña temporal para evitar el error "no se pueden borrar
  // todas las hojas" cuando borramos las pestañas existentes.
  const tmpExistente = ss.getSheetByName('__tmp_setup__');
  if (tmpExistente) ss.deleteSheet(tmpExistente);
  const tmp = ss.insertSheet('__tmp_setup__');

  Object.values(HOJAS).forEach(nombre => {
    const h = ss.getSheetByName(nombre);
    if (h) ss.deleteSheet(h);
  });

  crearConfig(ss);
  crearCategorias(ss);
  crearIngresos(ss);
  crearAportaciones(ss);
  crearGastosCompartidosFijos(ss);
  crearGastosCompartidosVariables(ss);
  crearGastosIndividuales(ss, HOJAS.G_FM, 'FM');
  crearGastosIndividuales(ss, HOJAS.G_LUCIA, 'Lucía');
  crearAhorro(ss);
  crearObjetivosLargoPlazo(ss);
  crearAnaliticaAnual(ss);
  crearLogBot(ss);
  crearDashboardResumen(ss);

  // Quita la pestaña temporal ahora que ya hay otras.
  ss.deleteSheet(tmp);

  // Borra la hoja inicial "Hoja 1" / "Sheet1" si sigue ahí vacía
  const sobrante = ss.getSheets().find(s =>
    (s.getName() === 'Hoja 1' || s.getName() === 'Sheet1') &&
    s.getLastRow() === 0
  );
  if (sobrante && ss.getSheets().length > 1) ss.deleteSheet(sobrante);

  SpreadsheetApp.flush();
  try {
    SpreadsheetApp.getUi().alert(
      'Dashboard creado correctamente.\n\n' +
      'Siguiente paso: revisa la pestaña "Config" y rellena tus valores ' +
      '(mes activo, objetivo de ahorro, chat IDs de Telegram).'
    );
  } catch (e) {
    // Si se ejecuta desde el editor sin Sheet activo, getUi() falla.
    // No es un error real: el dashboard ya está creado.
    console.log('Dashboard creado correctamente. Abre el Sheet para revisarlo.');
  }
}

function recrearDashboardSeguro() {
  // Wrapper que pide confirmación.
  const ui = SpreadsheetApp.getUi();
  const r = ui.alert(
    'Recrear dashboard',
    'Esto BORRARÁ las pestañas existentes con los mismos nombres y se perderán los datos.\n\n¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  if (r === ui.Button.YES) crearDashboard();
}

/* ====================== PESTAÑAS ====================== */

function crearConfig(ss) {
  const sh = ss.insertSheet(HOJAS.CONFIG);
  const filas = [
    ['Parámetro', 'Valor', 'Descripción'],
    ['Mes activo', mesActualISO(), 'Formato AAAA-MM. El Dashboard usa este mes como referencia.'],
    ['Objetivo ahorro mensual conjunto (€)', 200, 'Meta de ahorro del mes (ingresos - gastos). Modificable desde el bot con /objetivo.'],
    ['Telegram chat IDs autorizados', '', 'Separa con coma. Sólo estos chats pueden usar el bot.'],
    ['Telegram bot token', '', 'Pega aquí el token que te dé @BotFather.'],
  ];
  sh.getRange(1, 1, filas.length, 3).setValues(filas);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E8EAED');
  sh.setColumnWidths(1, 1, 260);
  sh.setColumnWidths(2, 1, 180);
  sh.setColumnWidths(3, 1, 520);
  sh.setFrozenRows(1);
}

function crearCategorias(ss) {
  const sh = ss.insertSheet(HOJAS.CATEGORIAS);
  const filas = [
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
  sh.getRange(1, 1, filas.length, 3).setValues(filas);
  sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E8EAED');
  sh.setColumnWidths(1, 3, 200);
  sh.setFrozenRows(1);
}

function crearIngresos(ss) {
  const sh = ss.insertSheet(HOJAS.INGRESOS);
  sh.getRange(1, 1, 1, 5).setValues([['Fecha', 'Persona', 'Concepto', 'Importe (€)', 'Recurrente']]);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#E6F4EA');
  sh.setColumnWidths(1, 1, 110);
  sh.setColumnWidths(2, 1, 100);
  sh.setColumnWidths(3, 1, 260);
  sh.setColumnWidths(4, 1, 110);
  sh.setColumnWidths(5, 1, 110);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('D:D').setNumberFormat('#,##0.00 €');
  validarLista(sh, 'B2:B', ['FM', 'Lucía']);
  validarLista(sh, 'E2:E', ['Sí', 'No']);
}

function crearAportaciones(ss) {
  const sh = ss.insertSheet(HOJAS.APORTACIONES);
  sh.getRange(1, 1, 1, 4).setValues([['Fecha', 'Persona', 'Concepto', 'Importe (€)']]);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#FFF2CC');
  sh.setColumnWidths(1, 1, 110);
  sh.setColumnWidths(2, 1, 100);
  sh.setColumnWidths(3, 1, 280);
  sh.setColumnWidths(4, 1, 110);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('D:D').setNumberFormat('#,##0.00 €');
  validarLista(sh, 'B2:B', ['FM', 'Lucía']);
}

function crearGastosCompartidosFijos(ss) {
  const sh = ss.insertSheet(HOJAS.GC_FIJOS);
  sh.getRange(1, 1, 1, 5).setValues([['Fecha', 'Concepto', 'Categoría', 'Importe (€)', 'Día de cargo']]);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#FCE8E6');
  sh.setColumnWidths(1, 1, 110);
  sh.setColumnWidths(2, 1, 260);
  sh.setColumnWidths(3, 1, 160);
  sh.setColumnWidths(4, 1, 110);
  sh.setColumnWidths(5, 1, 110);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('D:D').setNumberFormat('#,##0.00 €');
  validarRango(sh, 'C2:C', `${HOJAS.CATEGORIAS}!A2:A`);
}

function crearGastosCompartidosVariables(ss) {
  const sh = ss.insertSheet(HOJAS.GC_VARIABLES);
  sh.getRange(1, 1, 1, 5).setValues([['Fecha', 'Concepto', 'Categoría', 'Importe (€)', 'Pagado por']]);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#FEF7E0');
  sh.setColumnWidths(1, 1, 110);
  sh.setColumnWidths(2, 1, 260);
  sh.setColumnWidths(3, 1, 160);
  sh.setColumnWidths(4, 1, 110);
  sh.setColumnWidths(5, 1, 110);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('D:D').setNumberFormat('#,##0.00 €');
  validarRango(sh, 'C2:C', `${HOJAS.CATEGORIAS}!B2:B`);
  validarLista(sh, 'E2:E', ['Bote común', 'FM', 'Lucía']);
}

function crearGastosIndividuales(ss, nombre, persona) {
  const sh = ss.insertSheet(nombre);
  sh.getRange(1, 1, 1, 4).setValues([['Fecha', 'Concepto', 'Categoría', 'Importe (€)']]);
  const color = persona === 'FM' ? '#D2E3FC' : '#F8D7DA';
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground(color);
  sh.setColumnWidths(1, 1, 110);
  sh.setColumnWidths(2, 1, 280);
  sh.setColumnWidths(3, 1, 160);
  sh.setColumnWidths(4, 1, 110);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('D:D').setNumberFormat('#,##0.00 €');
  validarRango(sh, 'C2:C', `${HOJAS.CATEGORIAS}!C2:C`);
}

function crearAhorro(ss) {
  const sh = ss.insertSheet(HOJAS.AHORRO);
  sh.getRange(1, 1, 1, 4).setValues([['Mes', 'Objetivo (€)', 'Real (€)', 'Acumulado año (€)']]);
  sh.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#E8EAED');
  sh.setColumnWidths(1, 1, 100);
  sh.setColumnWidths(2, 4, 140);
  sh.setFrozenRows(1);

  // Pre-rellena los 12 meses del año en curso.
  const anio = new Date().getFullYear();
  const filas = [];
  for (let m = 1; m <= 12; m++) {
    const mesISO = `${anio}-${String(m).padStart(2, '0')}`;
    filas.push([
      mesISO,
      `=VLOOKUP("Objetivo ahorro mensual conjunto (€)";${HOJAS.CONFIG}!A:B;2;FALSE)`,
      `=SUMIFS(${HOJAS.INGRESOS}!D:D;${HOJAS.INGRESOS}!A:A;">="&DATE(${anio};${m};1);${HOJAS.INGRESOS}!A:A;"<"&DATE(${anio};${m + 1};1))` +
        `-SUMIFS(${HOJAS.GC_FIJOS}!D:D;${HOJAS.GC_FIJOS}!A:A;">="&DATE(${anio};${m};1);${HOJAS.GC_FIJOS}!A:A;"<"&DATE(${anio};${m + 1};1))` +
        `-SUMIFS(${HOJAS.GC_VARIABLES}!D:D;${HOJAS.GC_VARIABLES}!A:A;">="&DATE(${anio};${m};1);${HOJAS.GC_VARIABLES}!A:A;"<"&DATE(${anio};${m + 1};1))` +
        `-SUMIFS(${HOJAS.G_FM}!D:D;${HOJAS.G_FM}!A:A;">="&DATE(${anio};${m};1);${HOJAS.G_FM}!A:A;"<"&DATE(${anio};${m + 1};1))` +
        `-SUMIFS(${HOJAS.G_LUCIA}!D:D;${HOJAS.G_LUCIA}!A:A;">="&DATE(${anio};${m};1);${HOJAS.G_LUCIA}!A:A;"<"&DATE(${anio};${m + 1};1))`,
      m === 1 ? '=C2' : `=D${m + 0}+C${m + 1}`,
    ]);
  }
  sh.getRange(2, 1, 12, 4).setValues(filas);
  sh.getRange('B2:D13').setNumberFormat('#,##0.00 €');
}

function crearObjetivosLargoPlazo(ss) {
  const sh = ss.insertSheet(HOJAS.OBJETIVOS_LP);
  const cabeceras = [['Fecha_inicio', 'Concepto', 'Meta_€', 'Aportado_€', '%_Progreso', 'Fecha_objetivo', 'Estado']];
  sh.getRange(1, 1, 1, 7).setValues(cabeceras);
  sh.getRange(1, 1, 1, 7)
    .setFontWeight('bold')
    .setBackground(COLOR.cabHeader)
    .setFontColor(COLOR.cabHeaderTxt)
    .setHorizontalAlignment('center');

  // Filas de ejemplo (3)
  const hoy = new Date();
  const enUnAnio = new Date(hoy.getFullYear() + 1, hoy.getMonth(), hoy.getDate());
  const enDosAnios = new Date(hoy.getFullYear() + 2, hoy.getMonth(), hoy.getDate());
  const enSeisMeses = new Date(hoy.getFullYear(), hoy.getMonth() + 6, hoy.getDate());

  const ejemplos = [
    [hoy, 'Vacaciones',        2500,  0, null, enSeisMeses, null],
    [hoy, 'Coche',             8000,  0, null, enDosAnios,  null],
    [hoy, 'Fondo emergencia',  6000,  0, null, enUnAnio,    null],
  ];
  sh.getRange(2, 1, ejemplos.length, 7).setValues(ejemplos);

  // Fórmulas para % progreso y estado (filas 2..50 listas para nuevas metas).
  for (let r = 2; r <= 50; r++) {
    sh.getRange(r, 5).setFormula(`=IFERROR(D${r}/C${r};0)`);
    sh.getRange(r, 7).setFormula(`=IF(C${r}="";"";IF(D${r}>=C${r};"Cumplido";"En curso"))`);
  }

  sh.setColumnWidths(1, 1, 110);
  sh.setColumnWidths(2, 1, 220);
  sh.setColumnWidths(3, 1, 110);
  sh.setColumnWidths(4, 1, 110);
  sh.setColumnWidths(5, 1, 110);
  sh.setColumnWidths(6, 1, 130);
  sh.setColumnWidths(7, 1, 110);
  sh.setFrozenRows(1);

  sh.getRange('A2:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('F2:F').setNumberFormat('yyyy-mm-dd');
  sh.getRange('C2:D').setNumberFormat('#,##0.00 €');
  sh.getRange('E2:E').setNumberFormat('0.00%');

  // Validación de fecha en A y F
  const reglaFecha = SpreadsheetApp.newDataValidation()
    .requireDate()
    .setAllowInvalid(true)
    .build();
  sh.getRange('A2:A').setDataValidation(reglaFecha);
  sh.getRange('F2:F').setDataValidation(reglaFecha);

  // Formato condicional: barra de progreso visual con gradiente.
  const reglaGrad = SpreadsheetApp.newConditionalFormatRule()
    .setGradientMinpointWithValue('#FCE5E2', SpreadsheetApp.InterpolationType.NUMBER, '0')
    .setGradientMidpointWithValue('#FEF1DD', SpreadsheetApp.InterpolationType.NUMBER, '0.5')
    .setGradientMaxpointWithValue('#E1F4EA', SpreadsheetApp.InterpolationType.NUMBER, '1')
    .setRanges([sh.getRange('E2:E50')])
    .build();

  const reglaCumplido = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Cumplido')
    .setBackground(COLOR.verdeSuave)
    .setFontColor(COLOR.verde)
    .setBold(true)
    .setRanges([sh.getRange('G2:G50')])
    .build();
  const reglaEnCurso = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('En curso')
    .setBackground(COLOR.acentoSuave)
    .setFontColor(COLOR.acento)
    .setBold(true)
    .setRanges([sh.getRange('G2:G50')])
    .build();

  sh.setConditionalFormatRules([reglaGrad, reglaCumplido, reglaEnCurso]);
}

function crearAnaliticaAnual(ss) {
  const sh = ss.insertSheet(HOJAS.ANALITICA_ANUAL);
  const anio = new Date().getFullYear();

  const cab = [['Mes', 'Ingresos', 'Gastos compartidos', 'Gastos FM', 'Gastos Lucía', 'Ahorro', 'Variación vs mes anterior']];
  sh.getRange(1, 1, 1, 7).setValues(cab);
  sh.getRange(1, 1, 1, 7)
    .setFontWeight('bold')
    .setBackground(COLOR.cabHeader)
    .setFontColor(COLOR.cabHeaderTxt)
    .setHorizontalAlignment('center');

  const nombresMes = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const filas = [];
  for (let m = 1; m <= 12; m++) {
    const fila = m + 1;
    const ini = `DATE(${anio};${m};1)`;
    const fin = `DATE(${anio};${m + 1};1)`;
    filas.push([
      nombresMes[m - 1],
      `=IFERROR(SUMIFS(${HOJAS.INGRESOS}!D:D;${HOJAS.INGRESOS}!A:A;">="&${ini};${HOJAS.INGRESOS}!A:A;"<"&${fin});0)`,
      `=IFERROR(SUMIFS(${HOJAS.GC_FIJOS}!D:D;${HOJAS.GC_FIJOS}!A:A;">="&${ini};${HOJAS.GC_FIJOS}!A:A;"<"&${fin});0)+IFERROR(SUMIFS(${HOJAS.GC_VARIABLES}!D:D;${HOJAS.GC_VARIABLES}!A:A;">="&${ini};${HOJAS.GC_VARIABLES}!A:A;"<"&${fin});0)`,
      `=IFERROR(SUMIFS(${HOJAS.G_FM}!D:D;${HOJAS.G_FM}!A:A;">="&${ini};${HOJAS.G_FM}!A:A;"<"&${fin});0)`,
      `=IFERROR(SUMIFS(${HOJAS.G_LUCIA}!D:D;${HOJAS.G_LUCIA}!A:A;">="&${ini};${HOJAS.G_LUCIA}!A:A;"<"&${fin});0)`,
      `=B${fila}-C${fila}-D${fila}-E${fila}`,
      m === 1 ? '' : `=IFERROR((F${fila}-F${fila - 1})/ABS(F${fila - 1});0)`,
    ]);
  }
  sh.getRange(2, 1, 12, 7).setValues(filas);

  // Fila total
  sh.getRange('A14').setValue('TOTAL').setFontWeight('bold');
  for (let c = 2; c <= 6; c++) {
    sh.getRange(14, c).setFormula(`=SUM(${columnaLetra(c)}2:${columnaLetra(c)}13)`);
  }
  sh.getRange('A14:G14').setBackground(COLOR.acentoSuave).setFontWeight('bold');

  sh.getRange('B2:F14').setNumberFormat('#,##0.00 €');
  sh.getRange('G2:G14').setNumberFormat('0.00%;[Red]-0.00%');

  sh.setColumnWidths(1, 1, 80);
  sh.setColumnWidths(2, 5, 140);
  sh.setColumnWidths(7, 1, 180);
  sh.setFrozenRows(1);

  // Formato condicional ahorro (F) verde/rojo
  const reglaAhorroPos = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor(COLOR.verde)
    .setBold(true)
    .setRanges([sh.getRange('F2:F13')])
    .build();
  const reglaAhorroNeg = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor(COLOR.rojo)
    .setBold(true)
    .setRanges([sh.getRange('F2:F13')])
    .build();
  const reglaVarPos = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setFontColor(COLOR.verde)
    .setRanges([sh.getRange('G2:G13')])
    .build();
  const reglaVarNeg = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor(COLOR.rojo)
    .setRanges([sh.getRange('G2:G13')])
    .build();
  sh.setConditionalFormatRules([reglaAhorroPos, reglaAhorroNeg, reglaVarPos, reglaVarNeg]);

  // Gráficos a la derecha
  // 1) Línea ahorro mensual
  const chartAhorro = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(sh.getRange('A1:A13'))
    .addRange(sh.getRange('F1:F13'))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setOption('title', `Ahorro mensual ${anio}`)
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COLOR.acento])
    .setOption('width', 480)
    .setOption('height', 280)
    .setPosition(2, 9, 0, 0)
    .build();
  sh.insertChart(chartAhorro);

  // 2) Barras apiladas Ingresos vs Gastos
  const chartIngGas = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(sh.getRange('A1:E13'))
    .setOption('title', `Ingresos vs Gastos ${anio}`)
    .setOption('isStacked', true)
    .setOption('colors', [COLOR.verde, '#A8B4C9', '#5A86C7', '#E07A78'])
    .setOption('legend', { position: 'top' })
    .setOption('width', 480)
    .setOption('height', 280)
    .setPosition(18, 9, 0, 0)
    .build();
  sh.insertChart(chartIngGas);
}

function crearLogBot(ss) {
  const sh = ss.insertSheet(HOJAS.LOG_BOT);
  sh.getRange(1, 1, 1, 4).setValues([['Fecha', 'Persona', 'Acción', 'Detalle']]);
  sh.getRange(1, 1, 1, 4)
    .setFontWeight('bold')
    .setBackground(COLOR.cabHeader)
    .setFontColor(COLOR.cabHeaderTxt)
    .setHorizontalAlignment('center');
  sh.setColumnWidths(1, 1, 150);
  sh.setColumnWidths(2, 1, 100);
  sh.setColumnWidths(3, 1, 180);
  sh.setColumnWidths(4, 1, 400);
  sh.setFrozenRows(1);
  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

/* ====================== DASHBOARD ====================== */
/**
 * Dashboard rediseñado.
 *
 * IMPORTANTE: el bot lee celdas concretas de la columna B en el Dashboard
 * (B5, B6, B7, B15..B18, B21, B22, B25, B26, B29, B30, B32). Mantenemos
 * esos valores en su sitio exacto (bloque "datos legacy" en A:B) y
 * ocultamos esas columnas; el dashboard visual se construye en D+ y
 * referencia las celdas legacy. Así no rompemos Bot.gs ni Config.gs.
 */
function crearDashboardResumen(ss) {
  const sh = ss.insertSheet(HOJAS.DASHBOARD);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);

  // Helpers de fórmulas
  const mesActivo = `VLOOKUP("Mes activo";${HOJAS.CONFIG}!A:B;2;FALSE)`;
  const inicioMes = `DATE(VALUE(LEFT(${mesActivo};4));VALUE(MID(${mesActivo};6;2));1)`;
  const finMes = `EOMONTH(${inicioMes};0)+1`;

  const sumRango = (hoja, rImporte, rFecha) =>
    `IFERROR(SUMIFS(${hoja}!${rImporte};${hoja}!${rFecha};">="&${inicioMes};${hoja}!${rFecha};"<"&${finMes});0)`;
  const sumPersona = (persona) =>
    `IFERROR(SUMIFS(${HOJAS.INGRESOS}!D:D;${HOJAS.INGRESOS}!A:A;">="&${inicioMes};${HOJAS.INGRESOS}!A:A;"<"&${finMes};${HOJAS.INGRESOS}!B:B;"${persona}");0)`;
  const sumAportacion = (persona) =>
    `IFERROR(SUMIFS(${HOJAS.APORTACIONES}!D:D;${HOJAS.APORTACIONES}!A:A;">="&${inicioMes};${HOJAS.APORTACIONES}!A:A;"<"&${finMes};${HOJAS.APORTACIONES}!B:B;"${persona}");0)`;

  /* ===== BLOQUE LEGACY (cols A:B, filas 3-32) — el bot lee esto ===== */
  const bloques = [
    ['', ''],
    ['INGRESOS DEL MES', ''],
    ['Ingresos FM', `=${sumPersona('FM')}`],
    ['Ingresos Lucía', `=${sumPersona('Lucía')}`],
    ['Total ingresos', '=B5+B6'],
    ['', ''],
    ['BOTE COMÚN', ''],
    ['Aportación FM', `=${sumAportacion('FM')}`],
    ['Aportación Lucía', `=${sumAportacion('Lucía')}`],
    ['Total bote común', '=B10+B11'],
    ['', ''],
    ['GASTOS COMPARTIDOS', ''],
    ['Fijos', `=${sumRango(HOJAS.GC_FIJOS, 'D:D', 'A:A')}`],
    ['Variables', `=${sumRango(HOJAS.GC_VARIABLES, 'D:D', 'A:A')}`],
    ['Total compartidos', '=B15+B16'],
    ['Diferencia bote – gastos', '=B12-B17'],
    ['', ''],
    ['GASTOS INDIVIDUALES', ''],
    ['Gastos FM', `=${sumRango(HOJAS.G_FM, 'D:D', 'A:A')}`],
    ['Gastos Lucía', `=${sumRango(HOJAS.G_LUCIA, 'D:D', 'A:A')}`],
    ['', ''],
    ['BALANCE PERSONAL DEL MES', ''],
    ['Saldo FM (ingresos – bote – gastos)', '=B5-B10-B21'],
    ['Saldo Lucía (ingresos – bote – gastos)', '=B6-B11-B22'],
    ['', ''],
    ['AHORRO', ''],
    ['Objetivo del mes', `=VLOOKUP("Objetivo ahorro mensual conjunto (€)";${HOJAS.CONFIG}!A:B;2;FALSE)`],
    ['Ahorro real del mes', '=B7-B15-B16-B21-B22'],
    ['Cumplimiento %', '=IFERROR(B30/B29;0)'],
    ['Acumulado año', `=SUMIFS(${HOJAS.AHORRO}!C:C;${HOJAS.AHORRO}!A:A;"<="&${mesActivo})`],
  ];
  sh.getRange(3, 1, bloques.length, 2).setValues(bloques);
  sh.getRange('B5:B7').setNumberFormat('#,##0.00 €');
  sh.getRange('B10:B12').setNumberFormat('#,##0.00 €');
  sh.getRange('B15:B18').setNumberFormat('#,##0.00 €');
  sh.getRange('B21:B22').setNumberFormat('#,##0.00 €');
  sh.getRange('B25:B26').setNumberFormat('#,##0.00 €');
  sh.getRange('B29:B30').setNumberFormat('#,##0.00 €');
  sh.getRange('B31').setNumberFormat('0.00%');
  sh.getRange('B32').setNumberFormat('#,##0.00 €');

  /* ===== LAYOUT VISUAL en cols D..N =====
   *
   * Anchos:
   *   A:B  -> ocultas (datos legacy para el bot)
   *   C    -> margen
   *   D..N -> dashboard visible (11 columnas)
   */
  sh.setColumnWidths(1, 1, 240);   // A (texto legacy, oculta)
  sh.setColumnWidths(2, 1, 130);   // B (valor legacy, oculta)
  sh.setColumnWidths(3, 1, 20);    // C margen
  sh.setColumnWidths(4, 11, 95);   // D..N
  // Filas
  sh.setRowHeight(1, 12);
  sh.setRowHeight(2, 44);
  sh.setRowHeight(3, 28);
  sh.setRowHeight(4, 8);

  // Fondo general del dashboard visual
  sh.getRange('C1:N80').setBackground(COLOR.fondo);

  // ===== Cabecera =====
  sh.getRange('D2:N2').merge()
    .setValue('Finanzas FM & Lucía  ·  Dashboard mensual')
    .setFontSize(20).setFontWeight('bold')
    .setFontColor(COLOR.textoFuerte)
    .setBackground(COLOR.fondo)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');

  sh.getRange('D3:I3').merge()
    .setFormula(`="Mes activo:  " & ${mesActivo}`)
    .setFontSize(11).setFontStyle('italic')
    .setFontColor(COLOR.textoSuave)
    .setBackground(COLOR.fondo)
    .setVerticalAlignment('middle');

  sh.getRange('J3:N3').merge()
    .setFormula(`="Actualizado: " & TEXT(NOW();"yyyy-mm-dd HH:mm")`)
    .setFontSize(10).setFontStyle('italic')
    .setFontColor(COLOR.textoSuave)
    .setBackground(COLOR.fondo)
    .setHorizontalAlignment('right')
    .setVerticalAlignment('middle');

  /* ===== Fila 1 de KPIs (resumen general) — filas 5-9 =====
   * 4 tarjetas: Ingresos | Gastos totales | Balance | % Objetivo
   * Cada tarjeta ocupa ~2 cols, h=5 filas, separadas por 1 col vacía no, mejor todo pegado.
   * Layout: D5:F9 / G5:H9 (gap col? mejor 4 tarjetas equiespaciadas).
   *
   * Distribución:  D:F | G:H | I:K | L:N  -> 4 tarjetas de ~3,2,3,3 cols
   * Reajusto a 4 iguales: D:F (3), G:H+I (3 ya no), mejor:
   *   D5:F9, G5:I9, J5:K9, L5:N9
   * Total cols: 3+3+2+3=11 OK
   */
  pintarKPI(sh, 'D5:F9', 'INGRESOS DEL MES',  '=B7',           '=\"FM \" & TEXT(B5;\"#,##0 €\") & \"   ·   Lucía \" & TEXT(B6;\"#,##0 €\")', COLOR.verde);
  pintarKPI(sh, 'G5:I9', 'GASTOS TOTALES',    '=B15+B16+B21+B22', '=\"Compartidos \" & TEXT(B15+B16;\"#,##0 €\") & \"  ·  Indiv. \" & TEXT(B21+B22;\"#,##0 €\")', COLOR.rojo);
  pintarKPI(sh, 'J5:K9', 'BALANCE MES',       '=B7-(B15+B16+B21+B22)', '=IF(B7-(B15+B16+B21+B22)>=0;\"▲ Superávit\";\"▼ Déficit\")', COLOR.acento);
  pintarKPI(sh, 'L5:N9', 'CUMPLIMIENTO OBJ.', '=B31',           '=\"Objetivo: \" & TEXT(B29;\"#,##0 €\") & \"  ·  Real: \" & TEXT(B30;\"#,##0 €\")', COLOR.acento, true);

  /* ===== Fila 2: Por persona (FM, Lucía) y Ahorro — filas 11-17 =====
   * Tres paneles grandes lado a lado:
   *   D11:G17  FM
   *   H11:J17  Lucía
   *   K11:N17  Ahorro
   */
  pintarPanelPersona(sh, 'D11:G17', 'FM',     '=B5', '=B21', '=B10', '=B25', '#3A6FB8');
  pintarPanelPersona(sh, 'H11:J17', 'Lucía',  '=B6', '=B22', '=B11', '=B26', '#C04A6C');
  pintarPanelAhorro(sh,   'K11:N17',
    '=B30',  // ahorro mes
    '=B32',  // acumulado año
    '=B31',  // % objetivo
    '=B29'); // objetivo

  /* ===== Fila 3: Bote común — filas 19-24 (panel ancho) ===== */
  pintarPanelBote(sh, 'D19:N24',
    '=B12',          // aportaciones totales
    '=B15+B16',      // gastos compartidos
    '=B12-(B15+B16)',// sobrante / déficit
    '=IFERROR((B15+B16)/B12;0)'); // % cobertura del bote por gastos

  /* ===== Fila 4: Tendencia vs mes anterior — filas 26-29 (banner) ===== */
  pintarBanner(sh, 'D26:N29', mesActivo);

  /* ===== ZONA AUX para gráficos (oculta, columnas P..T) =====
   * P:Q  Top 5 categorías de gasto del mes
   * S:T  Aportaciones al bote por persona (donut)
   * V:W  Gastos del mes por categoría (donut)
   *
   * Estas tablas auxiliares se calculan con QUERY / SUMIFS y las usan
   * los gráficos.
   */
  prepararZonaAuxiliarYGraficos(sh, ss, inicioMes, finMes);

  // Bordes y limpieza visual
  sh.getRange('C1:N80').setFontFamily('Inter');
  sh.setHiddenGridlines(true);

  // Oculta cols legacy A:B (el bot las usa, pero el usuario no debe verlas).
  sh.hideColumns(1, 2);

  sh.setFrozenRows(3);
}

/* ====================== HELPERS DEL DASHBOARD ====================== */

/**
 * Pinta una tarjeta KPI en `rangoA1`.
 *  - titulo: texto pequeño arriba
 *  - formulaValor: fórmula del valor central grande (ej. '=B7')
 *  - formulaDetalle: fórmula del subtítulo (ej. '="FM " & ...')
 *  - colorAcento: barra superior de color
 *  - esPorcentaje: si true, formatea el valor central como %.
 */
function pintarKPI(sh, rangoA1, titulo, formulaValor, formulaDetalle, colorAcento, esPorcentaje) {
  const rng = sh.getRange(rangoA1);
  const fila = rng.getRow();
  const col = rng.getColumn();
  const numFilas = rng.getNumRows();
  const numCols = rng.getNumColumns();

  // Fondo de la tarjeta (panel blanco con bordes)
  rng.setBackground(COLOR.panel)
    .setBorder(true, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  // Barra de color superior (fila 1 de la card)
  sh.getRange(fila, col, 1, numCols).setBackground(colorAcento);
  sh.setRowHeight(fila, 4);

  // Título (fila 2 de la card)
  sh.getRange(fila + 1, col, 1, numCols).merge()
    .setValue(titulo)
    .setFontSize(9).setFontWeight('bold')
    .setFontColor(COLOR.textoSuave)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 1, 22);

  // Valor central (fila 3 de la card)
  const valorCell = sh.getRange(fila + 2, col, 1, numCols).merge()
    .setFormula(formulaValor)
    .setFontSize(22).setFontWeight('bold')
    .setFontColor(COLOR.textoFuerte)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(COLOR.panel);
  valorCell.setNumberFormat(esPorcentaje ? '0.0%' : '#,##0.00 €');
  sh.setRowHeight(fila + 2, 44);

  // Detalle (fila 4 de la card)
  sh.getRange(fila + 3, col, 1, numCols).merge()
    .setFormula(formulaDetalle)
    .setFontSize(9)
    .setFontColor(COLOR.textoSuave)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 3, 22);

  // Margen inferior
  if (numFilas >= 5) sh.setRowHeight(fila + 4, 8);
}

function pintarPanelPersona(sh, rangoA1, nombre, fIng, fGastos, fBote, fSaldo, colorAcento) {
  const rng = sh.getRange(rangoA1);
  const fila = rng.getRow();
  const col = rng.getColumn();
  const nCols = rng.getNumColumns();

  rng.setBackground(COLOR.panel)
    .setBorder(true, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  // Barra superior
  sh.getRange(fila, col, 1, nCols).setBackground(colorAcento);
  sh.setRowHeight(fila, 4);

  // Título
  sh.getRange(fila + 1, col, 1, nCols).merge()
    .setValue(nombre.toUpperCase())
    .setFontSize(11).setFontWeight('bold')
    .setFontColor(COLOR.textoFuerte)
    .setHorizontalAlignment('center')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 1, 24);

  // Lista de métricas (4 filas: ingresos / gastos / bote / saldo)
  const metricas = [
    ['Ingresos',           fIng],
    ['Gastos individ.',    fGastos],
    ['Aportación bote',    fBote],
    ['Saldo personal',     fSaldo],
  ];
  for (let i = 0; i < metricas.length; i++) {
    const r = fila + 2 + i;
    sh.setRowHeight(r, 22);
    const halfCols = Math.ceil(nCols / 2);
    sh.getRange(r, col, 1, halfCols).merge()
      .setValue('  ' + metricas[i][0])
      .setFontSize(10).setFontColor(COLOR.textoSuave)
      .setBackground(COLOR.panel)
      .setVerticalAlignment('middle');
    sh.getRange(r, col + halfCols, 1, nCols - halfCols).merge()
      .setFormula(metricas[i][1])
      .setNumberFormat('#,##0.00 €')
      .setFontSize(11).setFontWeight('bold')
      .setFontColor(i === 3 ? COLOR.acento : COLOR.textoFuerte)
      .setHorizontalAlignment('right')
      .setBackground(COLOR.panel)
      .setVerticalAlignment('middle');
  }

  // Formato condicional sólo a la celda del saldo (última fila, mitad derecha)
  const saldoCell = sh.getRange(fila + 5, col + Math.ceil(nCols / 2), 1, nCols - Math.ceil(nCols / 2));
  aplicarReglaPositivoNegativo(sh, saldoCell);
}

function pintarPanelAhorro(sh, rangoA1, fMes, fAcum, fPct, fObj) {
  const rng = sh.getRange(rangoA1);
  const fila = rng.getRow();
  const col = rng.getColumn();
  const nCols = rng.getNumColumns();

  rng.setBackground(COLOR.panel)
    .setBorder(true, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(fila, col, 1, nCols).setBackground(COLOR.verde);
  sh.setRowHeight(fila, 4);

  sh.getRange(fila + 1, col, 1, nCols).merge()
    .setValue('AHORRO')
    .setFontSize(11).setFontWeight('bold')
    .setFontColor(COLOR.textoFuerte)
    .setHorizontalAlignment('center')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 1, 24);

  // Valor principal: ahorro del mes (grande)
  sh.getRange(fila + 2, col, 1, nCols).merge()
    .setFormula(fMes)
    .setNumberFormat('#,##0.00 €')
    .setFontSize(20).setFontWeight('bold')
    .setFontColor(COLOR.verde)
    .setHorizontalAlignment('center')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 2, 38);

  // Sub-detalles: acumulado año y % objetivo
  sh.getRange(fila + 3, col, 1, nCols).merge()
    .setFormula(`="Acumulado año:  " & TEXT(${fAcum.replace('=', '')};"#,##0.00 €")`)
    .setFontSize(10)
    .setFontColor(COLOR.textoSuave)
    .setHorizontalAlignment('center')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 3, 20);

  sh.getRange(fila + 4, col, 1, nCols).merge()
    .setFormula(`="Cumplimiento objetivo: " & TEXT(${fPct.replace('=', '')};"0.0%") & "  (meta " & TEXT(${fObj.replace('=', '')};"#,##0 €") & ")"`)
    .setFontSize(10)
    .setFontColor(COLOR.textoSuave)
    .setHorizontalAlignment('center')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 4, 20);

  // Aplica regla de color al valor central (verde si positivo, rojo si negativo)
  const valorRng = sh.getRange(fila + 2, col, 1, nCols);
  aplicarReglaPositivoNegativo(sh, valorRng);
}

function pintarPanelBote(sh, rangoA1, fAport, fGasto, fDif, fCobertura) {
  const rng = sh.getRange(rangoA1);
  const fila = rng.getRow();
  const col = rng.getColumn();
  const nCols = rng.getNumColumns();

  rng.setBackground(COLOR.panel)
    .setBorder(true, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  // Barra superior amarilla/naranja
  sh.getRange(fila, col, 1, nCols).setBackground('#E8B84C');
  sh.setRowHeight(fila, 4);

  // Título
  sh.getRange(fila + 1, col, 1, nCols).merge()
    .setValue('BOTE COMÚN')
    .setFontSize(11).setFontWeight('bold')
    .setFontColor(COLOR.textoFuerte)
    .setHorizontalAlignment('center')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 1, 24);

  // Cuatro columnas internas: aport / gastos / diferencia / % cobertura
  const subCols = 4;
  const ancho = Math.floor(nCols / subCols);
  const datos = [
    ['Aportado',          fAport,      false, COLOR.verde],
    ['Gastos compart.',   fGasto,      false, COLOR.rojo],
    ['Sobrante / déficit', fDif,        false, COLOR.acento],
    ['% Cobertura',       fCobertura,  true,  COLOR.acento],
  ];

  for (let i = 0; i < subCols; i++) {
    const c = col + i * ancho;
    const w = i === subCols - 1 ? nCols - i * ancho : ancho;

    // Etiqueta arriba
    sh.getRange(fila + 2, c, 1, w).merge()
      .setValue(datos[i][0])
      .setFontSize(9).setFontWeight('bold')
      .setFontColor(COLOR.textoSuave)
      .setHorizontalAlignment('center')
      .setBackground(COLOR.panel);
    sh.setRowHeight(fila + 2, 20);

    // Valor
    const valorCell = sh.getRange(fila + 3, c, 2, w).merge()
      .setFormula(datos[i][1])
      .setFontSize(16).setFontWeight('bold')
      .setFontColor(datos[i][3])
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground(COLOR.panel);
    valorCell.setNumberFormat(datos[i][2] ? '0.0%' : '#,##0.00 €');
    sh.setRowHeight(fila + 3, 22);
    sh.setRowHeight(fila + 4, 22);
  }

  // Regla cobertura > 90% → naranja
  const covCell = sh.getRange(fila + 3, col + 3 * ancho, 2, nCols - 3 * ancho);
  const reglaCovAlta = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0.9)
    .setFontColor(COLOR.naranja)
    .setRanges([covCell])
    .build();
  const reglaCovOk = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberBetween(0, 0.9)
    .setFontColor(COLOR.verde)
    .setRanges([covCell])
    .build();
  // Regla diferencia (col 3)
  const difCell = sh.getRange(fila + 3, col + 2 * ancho, 2, ancho);
  const reglaDifPos = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0)
    .setFontColor(COLOR.verde)
    .setRanges([difCell])
    .build();
  const reglaDifNeg = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor(COLOR.rojo)
    .setRanges([difCell])
    .build();

  const reglas = sh.getConditionalFormatRules();
  reglas.push(reglaCovAlta, reglaCovOk, reglaDifPos, reglaDifNeg);
  sh.setConditionalFormatRules(reglas);
}

function pintarBanner(sh, rangoA1, mesActivoExpr) {
  const rng = sh.getRange(rangoA1);
  const fila = rng.getRow();
  const col = rng.getColumn();
  const nCols = rng.getNumColumns();

  rng.setBackground(COLOR.acentoSuave)
    .setBorder(true, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(fila, col, 1, nCols).merge()
    .setValue('TENDENCIA  ·  Mes activo vs mes anterior')
    .setFontSize(10).setFontWeight('bold')
    .setFontColor(COLOR.textoSuave)
    .setHorizontalAlignment('center')
    .setBackground(COLOR.acentoSuave);
  sh.setRowHeight(fila, 22);

  // Tres bloques: ingresos var, gastos var, ahorro var
  // Buscamos en Analitica_Anual por el mes activo (LEFT del mes activo da el año, MID 6;2 da el número del mes).
  const idxMes = `VALUE(MID(${mesActivoExpr};6;2))`;
  // Cada métrica: actual - anterior
  const expIng = `IFERROR(INDEX(${HOJAS.ANALITICA_ANUAL}!B:B;${idxMes}+1)-INDEX(${HOJAS.ANALITICA_ANUAL}!B:B;${idxMes});0)`;
  const expGas = `IFERROR((INDEX(${HOJAS.ANALITICA_ANUAL}!C:C;${idxMes}+1)+INDEX(${HOJAS.ANALITICA_ANUAL}!D:D;${idxMes}+1)+INDEX(${HOJAS.ANALITICA_ANUAL}!E:E;${idxMes}+1))-(INDEX(${HOJAS.ANALITICA_ANUAL}!C:C;${idxMes})+INDEX(${HOJAS.ANALITICA_ANUAL}!D:D;${idxMes})+INDEX(${HOJAS.ANALITICA_ANUAL}!E:E;${idxMes}));0)`;
  const expAho = `IFERROR(INDEX(${HOJAS.ANALITICA_ANUAL}!F:F;${idxMes}+1)-INDEX(${HOJAS.ANALITICA_ANUAL}!F:F;${idxMes});0)`;

  const items = [
    ['Ingresos', expIng, true],   // arrow up = bueno
    ['Gastos',   expGas, false],  // arrow up = malo
    ['Ahorro',   expAho, true],
  ];
  const ancho = Math.floor(nCols / items.length);
  for (let i = 0; i < items.length; i++) {
    const c = col + i * ancho;
    const w = i === items.length - 1 ? nCols - i * ancho : ancho;

    // Etiqueta
    sh.getRange(fila + 1, c, 1, w).merge()
      .setValue(items[i][0])
      .setFontSize(9).setFontWeight('bold')
      .setFontColor(COLOR.textoSuave)
      .setHorizontalAlignment('center')
      .setBackground(COLOR.acentoSuave);

    // Valor con flecha
    const buenoSiPositivo = items[i][2];
    const flecha = buenoSiPositivo
      ? `IF(${items[i][1]}>=0;"▲ ";"▼ ")`
      : `IF(${items[i][1]}<=0;"▲ ";"▼ ")`;
    sh.getRange(fila + 2, c, 2, w).merge()
      .setFormula(`=${flecha} & TEXT(${items[i][1]};"+#,##0 €;-#,##0 €;0 €")`)
      .setFontSize(14).setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBackground(COLOR.acentoSuave);

    // Formato condicional: la celda muestra "▲ ..." o "▼ ...";
    // coloreamos por el símbolo (interpretando "bueno" según buenoSiPositivo).
    const cell = sh.getRange(fila + 2, c, 2, w);
    const buenoSym = buenoSiPositivo ? '▲' : '▼';
    const maloSym  = buenoSiPositivo ? '▼' : '▲';
    const reglaBueno = SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith(buenoSym)
      .setFontColor(COLOR.verde).setRanges([cell]).build();
    const reglaMalo = SpreadsheetApp.newConditionalFormatRule()
      .whenTextStartsWith(maloSym)
      .setFontColor(COLOR.rojo).setRanges([cell]).build();
    const reglas = sh.getConditionalFormatRules();
    reglas.push(reglaBueno, reglaMalo);
    sh.setConditionalFormatRules(reglas);
  }

  sh.setRowHeight(fila + 1, 20);
  sh.setRowHeight(fila + 2, 24);
  sh.setRowHeight(fila + 3, 16);
}

/**
 * Crea las tablas auxiliares en columnas P..W (ocultas) y los gráficos.
 */
function prepararZonaAuxiliarYGraficos(sh, ss, inicioMes, finMes) {
  /* ===== AUX 1: Top 5 categorías de gasto del mes (cols P:Q, filas 2-7) =====
   * Unimos gastos compartidos (fijos+variables) e individuales y agrupamos por categoría.
   * Usamos QUERY sobre un rango concatenado con LLAVE { ; }.
   */
  sh.getRange('P1').setValue('Top categorías');
  sh.getRange('P2').setFormula(
    `=IFERROR(QUERY({` +
      `FILTER(${HOJAS.GC_FIJOS}!C:D;${HOJAS.GC_FIJOS}!A:A>=${inicioMes};${HOJAS.GC_FIJOS}!A:A<${finMes});` +
      `FILTER(${HOJAS.GC_VARIABLES}!C:D;${HOJAS.GC_VARIABLES}!A:A>=${inicioMes};${HOJAS.GC_VARIABLES}!A:A<${finMes});` +
      `FILTER(${HOJAS.G_FM}!C:D;${HOJAS.G_FM}!A:A>=${inicioMes};${HOJAS.G_FM}!A:A<${finMes});` +
      `FILTER(${HOJAS.G_LUCIA}!C:D;${HOJAS.G_LUCIA}!A:A>=${inicioMes};${HOJAS.G_LUCIA}!A:A<${finMes})` +
      `};"select Col1, sum(Col2) where Col1 is not null group by Col1 order by sum(Col2) desc limit 5 label sum(Col2) ''";0);"")`
  );

  /* ===== AUX 2: Aportaciones al bote por persona (cols S:T, filas 2-3) ===== */
  sh.getRange('S1').setValue('Aportaciones bote por persona');
  sh.getRange('S2:T3').setValues([
    ['FM',     ''],
    ['Lucía',  ''],
  ]);
  sh.getRange('T2').setFormula(
    `=IFERROR(SUMIFS(${HOJAS.APORTACIONES}!D:D;${HOJAS.APORTACIONES}!A:A;">="&${inicioMes};${HOJAS.APORTACIONES}!A:A;"<"&${finMes};${HOJAS.APORTACIONES}!B:B;"FM");0)`);
  sh.getRange('T3').setFormula(
    `=IFERROR(SUMIFS(${HOJAS.APORTACIONES}!D:D;${HOJAS.APORTACIONES}!A:A;">="&${inicioMes};${HOJAS.APORTACIONES}!A:A;"<"&${finMes};${HOJAS.APORTACIONES}!B:B;"Lucía");0)`);

  /* ===== AUX 3: Gastos del mes por categoría (cols V:W, dinámico) =====
   * Igual que AUX 1 pero sin limitar a 5 (alimenta el donut).
   */
  sh.getRange('V1').setValue('Gastos mes por categoría');
  sh.getRange('V2').setFormula(
    `=IFERROR(QUERY({` +
      `FILTER(${HOJAS.GC_FIJOS}!C:D;${HOJAS.GC_FIJOS}!A:A>=${inicioMes};${HOJAS.GC_FIJOS}!A:A<${finMes});` +
      `FILTER(${HOJAS.GC_VARIABLES}!C:D;${HOJAS.GC_VARIABLES}!A:A>=${inicioMes};${HOJAS.GC_VARIABLES}!A:A<${finMes});` +
      `FILTER(${HOJAS.G_FM}!C:D;${HOJAS.G_FM}!A:A>=${inicioMes};${HOJAS.G_FM}!A:A<${finMes});` +
      `FILTER(${HOJAS.G_LUCIA}!C:D;${HOJAS.G_LUCIA}!A:A>=${inicioMes};${HOJAS.G_LUCIA}!A:A<${finMes})` +
      `};"select Col1, sum(Col2) where Col1 is not null group by Col1 order by sum(Col2) desc label sum(Col2) ''";0);"")`
  );

  // Formato monetario en columnas aux
  sh.getRange('Q2:Q20').setNumberFormat('#,##0.00 €');
  sh.getRange('T2:T3').setNumberFormat('#,##0.00 €');
  sh.getRange('W2:W30').setNumberFormat('#,##0.00 €');

  /* ===== GRÁFICOS ===== */

  // 1) Donut de gastos del mes por categoría (debajo de la fila de KPIs, derecha)
  const chartDonutGastos = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('V2:W30'))
    .setOption('title', 'Gastos del mes por categoría')
    .setOption('pieHole', 0.55)
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
    .setOption('width', 380)
    .setOption('height', 260)
    .setOption('backgroundColor', COLOR.panel)
    .setOption('chartArea', { left: 10, top: 40, width: '95%', height: '80%' })
    .setPosition(31, 4, 0, 0)  // ~ fila 31, col D
    .build();
  sh.insertChart(chartDonutGastos);

  // 2) Donut de aportaciones al bote por persona
  const chartDonutBote = sh.newChart()
    .setChartType(Charts.ChartType.PIE)
    .addRange(sh.getRange('S2:T3'))
    .setOption('title', 'Aportaciones al bote por persona')
    .setOption('pieHole', 0.55)
    .setOption('legend', { position: 'right', textStyle: { fontSize: 10 } })
    .setOption('colors', ['#3A6FB8', '#C04A6C'])
    .setOption('width', 380)
    .setOption('height', 260)
    .setOption('backgroundColor', COLOR.panel)
    .setPosition(31, 10, 0, 0)  // fila 31, col J
    .build();
  sh.insertChart(chartDonutBote);

  // 3) Línea de evolución de ahorro acumulado (lee de Ahorro!D)
  const chartLineaAhorro = sh.newChart()
    .setChartType(Charts.ChartType.LINE)
    .addRange(ss.getSheetByName(HOJAS.AHORRO).getRange('A1:A13'))
    .addRange(ss.getSheetByName(HOJAS.AHORRO).getRange('D1:D13'))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setOption('title', 'Ahorro acumulado anual')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COLOR.verde])
    .setOption('width', 380)
    .setOption('height', 260)
    .setOption('backgroundColor', COLOR.panel)
    .setOption('areaOpacity', 0.2)
    .setPosition(46, 4, 0, 0)
    .build();
  sh.insertChart(chartLineaAhorro);

  // 4) Barras apiladas ingresos vs gastos por mes (Analitica_Anual)
  const an = ss.getSheetByName(HOJAS.ANALITICA_ANUAL);
  const chartBarrasMes = sh.newChart()
    .setChartType(Charts.ChartType.COLUMN)
    .addRange(an.getRange('A1:E13'))
    .setOption('title', 'Ingresos vs Gastos por mes')
    .setOption('isStacked', false)
    .setOption('colors', [COLOR.verde, '#A8B4C9', '#5A86C7', '#E07A78'])
    .setOption('legend', { position: 'top', textStyle: { fontSize: 9 } })
    .setOption('width', 380)
    .setOption('height', 260)
    .setOption('backgroundColor', COLOR.panel)
    .setPosition(46, 10, 0, 0)
    .build();
  sh.insertChart(chartBarrasMes);

  // 5) Barras horizontales Top 5 categorías del mes
  const chartTop5 = sh.newChart()
    .setChartType(Charts.ChartType.BAR)
    .addRange(sh.getRange('P2:Q6'))
    .setOption('title', 'Top 5 categorías de gasto (mes)')
    .setOption('legend', { position: 'none' })
    .setOption('colors', [COLOR.acento])
    .setOption('width', 760)
    .setOption('height', 260)
    .setOption('backgroundColor', COLOR.panel)
    .setOption('chartArea', { left: 120, top: 40, width: '70%', height: '75%' })
    .setPosition(61, 4, 0, 0)
    .build();
  sh.insertChart(chartTop5);

  // Ocultar columnas aux (P en adelante, hasta el final de la hoja).
  const ultima = sh.getMaxColumns();
  if (ultima >= 16) sh.hideColumns(16, ultima - 15);
}

/**
 * Añade reglas de formato condicional verde/rojo a un rango (positivo=verde,
 * negativo=rojo). Conserva las reglas previas de la hoja.
 */
function aplicarReglaPositivoNegativo(sh, rango) {
  const reglas = sh.getConditionalFormatRules();
  const reglaPos = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(0)
    .setFontColor(COLOR.verde)
    .setRanges([rango])
    .build();
  const reglaNeg = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(0)
    .setFontColor(COLOR.rojo)
    .setRanges([rango])
    .build();
  reglas.push(reglaPos, reglaNeg);
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

function validarRango(sh, rangoA1, rangoFuenteA1) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const [nombreHoja, rango] = rangoFuenteA1.split('!');
  const fuente = ss.getSheetByName(nombreHoja).getRange(rango);
  const regla = SpreadsheetApp.newDataValidation()
    .requireValueInRange(fuente, true)
    .setAllowInvalid(true)
    .build();
  sh.getRange(rangoA1).setDataValidation(regla);
}

function mesActualISO() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

/** Convierte un índice de columna (1-based) a letra A..Z..AA. */
function columnaLetra(n) {
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
