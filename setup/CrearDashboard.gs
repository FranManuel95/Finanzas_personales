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
  GC_FIJOS: 'Gastos_Compartidos_Fijos',
  GC_VARIABLES: 'Gastos_Compartidos_Variables',
  G_FM: 'Gastos_FM',
  G_LUCIA: 'Gastos_Lucia',
  AHORRO: 'Ahorro',
};

function crearDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Borra pestañas previas con los mismos nombres
  Object.values(HOJAS).forEach(nombre => {
    const h = ss.getSheetByName(nombre);
    if (h) ss.deleteSheet(h);
  });

  crearConfig(ss);
  crearCategorias(ss);
  crearIngresos(ss);
  crearGastosCompartidosFijos(ss);
  crearGastosCompartidosVariables(ss);
  crearGastosIndividuales(ss, HOJAS.G_FM, 'FM');
  crearGastosIndividuales(ss, HOJAS.G_LUCIA, 'Lucía');
  crearAhorro(ss);
  crearDashboardResumen(ss);

  // Borra la hoja inicial "Hoja 1" / "Sheet1" si sigue ahí vacía
  const sobrante = ss.getSheets().find(s =>
    (s.getName() === 'Hoja 1' || s.getName() === 'Sheet1') &&
    s.getLastRow() === 0
  );
  if (sobrante && ss.getSheets().length > 1) ss.deleteSheet(sobrante);

  SpreadsheetApp.getUi().alert(
    'Dashboard creado correctamente.\n\n' +
    'Siguiente paso: revisa la pestaña "Config" y rellena tus valores ' +
    '(mes activo, bote común, objetivo de ahorro, chat IDs de Telegram).'
  );
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
    ['Bote común FM (€)', 600, 'Cuánto aporta FM al bote común cada mes.'],
    ['Bote común Lucía (€)', 600, 'Cuánto aporta Lucía al bote común cada mes.'],
    ['Objetivo ahorro mensual conjunto (€)', 400, 'Meta de ahorro del mes (ingresos - gastos).'],
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

function crearDashboardResumen(ss) {
  const sh = ss.insertSheet(HOJAS.DASHBOARD);
  ss.setActiveSheet(sh);
  ss.moveActiveSheet(1);

  sh.getRange('A1').setValue('💶 Dashboard finanzas — FM & Lucía').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setFormula(`="Mes activo: " & VLOOKUP("Mes activo";${HOJAS.CONFIG}!A:B;2;FALSE)`).setFontStyle('italic');

  // Helpers de fórmulas
  const mesActivo = `VLOOKUP("Mes activo";${HOJAS.CONFIG}!A:B;2;FALSE)`;
  const inicioMes = `DATE(VALUE(LEFT(${mesActivo};4));VALUE(MID(${mesActivo};6;2));1)`;
  const finMes = `EOMONTH(${inicioMes};0)+1`;

  const sumRango = (hoja, rImporte, rFecha) =>
    `IFERROR(SUMIFS(${hoja}!${rImporte};${hoja}!${rFecha};">="&${inicioMes};${hoja}!${rFecha};"<"&${finMes});0)`;
  const sumPersona = (persona) =>
    `IFERROR(SUMIFS(${HOJAS.INGRESOS}!D:D;${HOJAS.INGRESOS}!A:A;">="&${inicioMes};${HOJAS.INGRESOS}!A:A;"<"&${finMes};${HOJAS.INGRESOS}!B:B;"${persona}");0)`;

  // Filas (empezamos a escribir en la 3):
  //  3 vacía · 4 cab Ingresos · 5 FM · 6 Lucía · 7 Total
  //  8 vacía · 9 cab Bote · 10 FM · 11 Lucía · 12 Total
  // 13 vacía · 14 cab GComp · 15 Fijos · 16 Var · 17 Total · 18 Dif
  // 19 vacía · 20 cab Indiv · 21 FM · 22 Lucía
  // 23 vacía · 24 cab Balance · 25 Saldo FM · 26 Saldo Lucía
  // 27 vacía · 28 cab Ahorro · 29 Objetivo · 30 Real · 31 % · 32 Acum
  const bloques = [
    ['', ''],
    ['INGRESOS DEL MES', ''],
    ['Ingresos FM', `=${sumPersona('FM')}`],
    ['Ingresos Lucía', `=${sumPersona('Lucía')}`],
    ['Total ingresos', '=B5+B6'],
    ['', ''],
    ['BOTE COMÚN', ''],
    ['Aportación FM', `=VLOOKUP("Bote común FM (€)";${HOJAS.CONFIG}!A:B;2;FALSE)`],
    ['Aportación Lucía', `=VLOOKUP("Bote común Lucía (€)";${HOJAS.CONFIG}!A:B;2;FALSE)`],
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

  // Formato monetario
  sh.getRange('B5:B7').setNumberFormat('#,##0.00 €');
  sh.getRange('B10:B12').setNumberFormat('#,##0.00 €');
  sh.getRange('B15:B18').setNumberFormat('#,##0.00 €');
  sh.getRange('B21:B22').setNumberFormat('#,##0.00 €');
  sh.getRange('B25:B26').setNumberFormat('#,##0.00 €');
  sh.getRange('B29:B30').setNumberFormat('#,##0.00 €');
  sh.getRange('B31').setNumberFormat('0.00%');
  sh.getRange('B32').setNumberFormat('#,##0.00 €');

  // Cabeceras de sección en negrita
  ['A4', 'A9', 'A14', 'A20', 'A24', 'A28'].forEach(c => {
    sh.getRange(c).setFontWeight('bold').setBackground('#E8EAED');
  });

  sh.setColumnWidths(1, 1, 320);
  sh.setColumnWidths(2, 1, 160);
  sh.setFrozenRows(2);
}

/* ====================== HELPERS ====================== */

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
