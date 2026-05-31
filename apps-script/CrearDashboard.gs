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
  PANEL: 'Panel',
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
  cebra:       '#F8FAFC', // filas alternas claras
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

  // Borra TODAS las hojas excepto la temporal: garantiza un libro limpio
  // y elimina cualquier hoja vieja de esquemas anteriores.
  ss.getSheets().forEach(h => {
    if (h.getName() !== '__tmp_setup__') ss.deleteSheet(h);
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

  const hoy = new Date();
  const y = hoy.getFullYear();
  const m = hoy.getMonth();
  const fDia = (dia) => new Date(y, m, dia, 0, 0);
  const fHora = (dia, h, min) => new Date(y, m, dia, h, min);
  const ejemplos = [
    [fDia(1),          'Ingreso',             'FM',    'Otros',    'Nómina FM',           1850, 'Salario mensual'],
    [fDia(1),          'Ingreso',             'Lucía', 'Otros',    'Nómina Lucía',        1620, 'Salario mensual'],
    [fDia(2),          'Aportación',          'FM',    'Otros',    'Aportación al bote',   400, 'Parte gastos comunes'],
    [fDia(2),          'Aportación',          'Lucía', 'Otros',    'Aportación al bote',   400, 'Parte gastos comunes'],
    [fDia(3),          'Compartido fijo',     'Bote',  'Alquiler', 'Alquiler piso',        750, ''],
    [fHora(8, 14, 32), 'Compartido variable', 'Bote',  'Compra',   'Supermercado',         186, 'Semana 1'],
    [fHora(12, 21, 5), 'Individual',          'FM',    'Ocio',     'Cine',                  24, ''],
    [fHora(15, 18, 40),'Individual',          'Lucía', 'Ropa',     'Zapatillas',            59, ''],
    [fDia(28),         'Ahorro',              'Bote',  'Otros',    'Ahorro común mensual',  50, 'Fondo emergencia'],
    [fDia(28),         'Ahorro',              'FM',    'Otros',    'Ahorro personal',      100, ''],
    [fDia(28),         'Ahorro',              'Lucía', 'Otros',    'Ahorro personal',       80, ''],
  ];
  sh.getRange(2, 1, ejemplos.length, 7).setValues(ejemplos);
  sh.getRange(2, 1, ejemplos.length, 1).setNumberFormat('yyyy-mm-dd HH:mm');
  sh.getRange(2, 6, ejemplos.length, 1).setNumberFormat('#,##0.00 €');

  const banda = sh.getRange(1, 1, Math.max(ejemplos.length + 1, 50), 7)
    .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  banda.setHeaderRowColor(COLOR.cab);
  banda.setFirstRowColor(COLOR.panel);
  banda.setSecondRowColor(COLOR.cebra);

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
  const extra = [
    ['% aporte real bote FM',    '=IFERROR(B4/B6;0)'],
    ['% aporte real bote Lucía', '=IFERROR(B5/B6;0)'],
    ['% ingreso FM al bote',     '=IFERROR(B4/B1;0)'],
    ['% ingreso Lucía al bote',  '=IFERROR(B5/B2;0)'],
    ['Tasa de ahorro real',      '=IFERROR(B17/B3;0)'],
    ['Tasa de ahorro objetivo',  `=IFERROR(VLOOKUP("Tasa de ahorro objetivo (%)";${AJ}!A:B;2;FALSE);2/10)`],
    ['Modo aportación bote',     `=IFERROR(VLOOKUP("Modo aportación al bote";${AJ}!A:B;2;FALSE);"50/50")`],
    ['', ''], ['', ''], ['', ''], ['', ''], ['', ''], ['', ''],
    ['', ''], ['', ''], ['', ''], ['', ''], ['', ''], ['', ''],
    ['Ahorro compartido del mes', `=${sumar('Ahorro', 'Bote')}`],
    ['Ahorro FM del mes',         `=${sumar('Ahorro', 'FM')}`],
    ['Ahorro Lucía del mes',      `=${sumar('Ahorro', 'Lucía')}`],
    ['Ahorro total del mes',      '=B41+B42+B43'],
    ['(reservado)',               ''],
    ['Inicio del mes activo',     `=${inicioMes}`],
    ['Fin del mes activo',        `=${finMes}`],
  ];
  sh.getRange(21, 1, extra.length, 2).setValues(extra);

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

/* ====================== HOJA: PANEL (dashboard) ======================
 * Diseño limpio, sin tarjetas KPI superiores. De arriba a abajo:
 *   🤝 COMPARTIDO  → aportado + gastos fijos + variables + ahorro común + queda en bote
 *   👤 FM / 👤 LUCÍA (apilados) → ingreso + bote + gastos + ahorro + en bolsillo
 *   📊 DÓNDE SE GASTA → categorías ordenadas con barras SPARKLINE nativas
 *   📈 EVOLUCIÓN → gráfico nativo 12 meses
 *   🎯 OBJETIVOS → progreso con barras SPARKLINE
 *
 * Rejilla de contenido: columnas B..K. Filas de transacción = concepto (B:F) ·
 * fecha (G:H) · importe (I:K). Las listas se leen en vivo de Movimientos con
 * FILTER (reserva de filas + overflow). Las barras son SPARKLINE (no REPT).
 */
function crearPanel(ss) {
  const sh = ss.insertSheet(HOJAS.PANEL);
  const C = `'${HOJAS.CALC}'`;
  const AJ = HOJAS.AJUSTES;

  sh.setHiddenGridlines(true);
  sh.getRange('A1:L220').setBackground(COLOR.fondo).setFontFamily(FUENTE);
  sh.setColumnWidth(1, 24);
  sh.setColumnWidths(2, 10, 86);
  sh.setColumnWidth(12, 24);

  /* Cabecera */
  sh.setRowHeight(2, 44);
  sh.getRange('B2:H2').merge().setValue('Finanzas FM & Lucía')
    .setFontSize(24).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setBackground(COLOR.fondo);
  sh.getRange('I2:K2').merge()
    .setFormula(`="Actualizado "&TEXT(NOW();"yyyy-mm-dd HH:mm")`)
    .setFontSize(9).setFontColor(COLOR.tenue)
    .setHorizontalAlignment('right').setVerticalAlignment('bottom').setBackground(COLOR.fondo);
  sh.setRowHeight(3, 22);
  sh.getRange('B3:K3').merge()
    .setFormula(`="Mes "&IF(ISNUMBER(VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE));TEXT(VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE);"yyyy-mm");VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE))`)
    .setFontSize(11).setFontColor(COLOR.tenue)
    .setVerticalAlignment('middle').setBackground(COLOR.fondo);

  let r = 5;

  /* ===== 🤝 COMPARTIDO ===== */
  r = seccion(sh, r, '🤝  COMPARTIDO');
  r = miniLista(sh, r, 'APORTADO AL BOTE', 'Total aportado', [
    ['FM', `=${C}!B4`],
    ['Lucía', `=${C}!B5`],
  ], `=${C}!B6`);
  r = subLista(sh, r, 'GASTOS FIJOS', 'Compartido fijo', null, `=${C}!B7`, RESERVA.COMP_FIJO);
  r = subLista(sh, r, 'GASTOS VARIABLES', 'Compartido variable', null, `=${C}!B8`, RESERVA.COMP_VAR);
  r = subLista(sh, r, 'AHORRO COMÚN', 'Ahorro', 'Bote', `=${C}!B41`, RESERVA.COMP_AHO);
  r = totalGrande(sh, r, 'QUEDA EN EL BOTE', `=${C}!B15`, true);
  r += 1;

  /* ===== 👤 FM ===== */
  r = seccion(sh, r, '👤  FM');
  r = filaResumen(sh, r, 'INGRESO', `=${C}!B1`, false, true);
  r = filaResumen(sh, r, 'Aportación al bote', `=${C}!B4`, true, false);
  r = subLista(sh, r, 'GASTOS INDIVIDUALES', 'Individual', 'FM', `=${C}!B10`, RESERVA.IND);
  r = subLista(sh, r, 'AHORRO INDIVIDUAL', 'Ahorro', 'FM', `=${C}!B42`, RESERVA.IND_AHO);
  r = totalGrande(sh, r, 'EN BOLSILLO', `=${C}!B13`, true);
  r += 1;

  /* ===== 👤 LUCÍA ===== */
  r = seccion(sh, r, '👤  LUCÍA');
  r = filaResumen(sh, r, 'INGRESO', `=${C}!B2`, false, true);
  r = filaResumen(sh, r, 'Aportación al bote', `=${C}!B5`, true, false);
  r = subLista(sh, r, 'GASTOS INDIVIDUALES', 'Individual', 'Lucía', `=${C}!B11`, RESERVA.IND);
  r = subLista(sh, r, 'AHORRO INDIVIDUAL', 'Ahorro', 'Lucía', `=${C}!B43`, RESERVA.IND_AHO);
  r = totalGrande(sh, r, 'EN BOLSILLO', `=${C}!B14`, true);
  r += 1;

  /* ===== 📊 DÓNDE SE GASTA ===== */
  r = seccion(sh, r, '📊  DÓNDE SE GASTA');
  r = tablaCategorias(sh, r, C);
  r += 1;

  /* ===== 📈 EVOLUCIÓN 12 MESES ===== */
  r = seccion(sh, r, '📈  EVOLUCIÓN 12 MESES');
  insertarGraficoEvolucionEn(sh, ss, r);
  r += 16;

  /* ===== 🎯 OBJETIVOS ===== */
  r = seccion(sh, r, '🎯  OBJETIVOS A LARGO PLAZO');
  r = barrasObjetivos(sh, r);

  sh.setFrozenRows(3);
}

/* ====================== HELPERS DE PANEL ====================== */

/** Título de sección con subrayado de acento. Devuelve la siguiente fila libre. */
function seccion(sh, r, texto) {
  sh.setRowHeight(r, 14);
  r++;
  sh.setRowHeight(r, 30);
  sh.getRange(r, 2, 1, 10).merge().setValue(texto)
    .setFontSize(14).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setBackground(COLOR.fondo)
    .setBorder(false, false, true, false, false, false, COLOR.acento, SpreadsheetApp.BorderStyle.SOLID_THICK);
  return r + 1;
}

/** Mini-lista de filas etiqueta→valor + total resaltado (para "Aportado al bote"). */
function miniLista(sh, r, etiqueta, etiquetaTotal, filas, totalFormula) {
  sh.setRowHeight(r, 22);
  sh.getRange(r, 2, 1, 10).merge().setValue(etiqueta)
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setBackground(COLOR.fondo).setVerticalAlignment('bottom').setHorizontalAlignment('left');
  r++;
  filas.forEach(([lbl, f], i) => {
    sh.setRowHeight(r, 24);
    const fondo = i % 2 === 0 ? COLOR.panel : COLOR.cebra;
    sh.getRange(r, 2, 1, 10).setBackground(fondo)
      .setBorder(false, true, false, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 2, 1, 7).merge().setValue(lbl)
      .setFontSize(11).setFontColor(COLOR.texto).setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.getRange(r, 9, 1, 3).merge().setFormula(f).setNumberFormat('#,##0.00 €')
      .setFontSize(11).setFontColor(COLOR.texto).setVerticalAlignment('middle').setHorizontalAlignment('right');
    r++;
  });
  sh.setRowHeight(r, 26);
  sh.getRange(r, 2, 1, 10).setBackground(COLOR.acentoSuave)
    .setBorder(true, true, true, true, false, false, COLOR.acento, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(r, 2, 1, 7).merge().setValue(etiquetaTotal)
    .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta).setVerticalAlignment('middle');
  sh.getRange(r, 9, 1, 3).merge().setFormula(totalFormula).setNumberFormat('#,##0.00 €')
    .setFontSize(13).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setHorizontalAlignment('right');
  return r + 1;
}

/** Fila resumen de una sola línea (INGRESO grande, o "− Aportación" normal). */
function filaResumen(sh, r, label, valFormula, negativo, grande) {
  sh.setRowHeight(r, grande ? 32 : 24);
  sh.getRange(r, 2, 1, 10).setBackground(grande ? COLOR.acentoSuave : COLOR.panel)
    .setBorder(grande, true, false, true, false, false, grande ? COLOR.acento : COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(r, 2, 1, 7).merge().setValue(grande ? label : ('     − ' + label))
    .setFontSize(grande ? 12 : 11).setFontWeight(grande ? 'bold' : 'normal')
    .setFontColor(grande ? COLOR.tinta : COLOR.texto)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  const f = negativo ? `=-1*(${valFormula.slice(1)})` : valFormula;
  sh.getRange(r, 9, 1, 3).merge().setFormula(f).setNumberFormat('#,##0.00 €')
    .setFontSize(grande ? 15 : 11).setFontWeight(grande ? 'bold' : 'normal')
    .setFontColor(grande ? COLOR.tinta : COLOR.texto)
    .setVerticalAlignment('middle').setHorizontalAlignment('right');
  return r + 1;
}

/** Sub-sección con LISTA de transacciones en vivo (FILTER) + total. Devuelve fila libre. */
function subLista(sh, r, etiqueta, tipo, persona, totalFormula, reserva) {
  const MOV = HOJAS.MOVIMIENTOS;
  const C = `'${HOJAS.CALC}'`;
  let cond = `${MOV}!B2:B="${tipo}"`;
  if (persona) cond += `;${MOV}!C2:C="${persona}"`;
  const flt = `FILTER(HSTACK(${MOV}!E2:E;${MOV}!A2:A;${MOV}!F2:F);${cond};${MOV}!A2:A>=${C}!B46;${MOV}!A2:A<${C}!B47)`;
  let cntA = `${MOV}!B:B;"${tipo}";${MOV}!A:A;">="&${C}!B46;${MOV}!A:A;"<"&${C}!B47`;
  if (persona) cntA += `;${MOV}!C:C;"${persona}"`;
  const cnt = `COUNTIFS(${cntA})`;

  sh.setRowHeight(r, 22);
  sh.getRange(r, 2, 1, 10).merge().setValue(etiqueta)
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setBackground(COLOR.fondo).setVerticalAlignment('bottom').setHorizontalAlignment('left');
  r++;

  for (let i = 0; i < reserva; i++) {
    const k = i + 1;
    sh.setRowHeight(r, 22);
    const fondo = i % 2 === 0 ? COLOR.panel : COLOR.cebra;
    sh.getRange(r, 2, 1, 10).setBackground(fondo)
      .setBorder(false, true, false, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(r, 2, 1, 5).merge().setFormula(`=IFERROR(INDEX(${flt};${k};1);"")`)
      .setFontSize(10).setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.getRange(r, 7, 1, 2).merge().setFormula(`=IFERROR(INDEX(${flt};${k};2);"")`)
      .setNumberFormat('dd-mmm HH:mm').setFontSize(9).setFontColor(COLOR.tenue)
      .setVerticalAlignment('middle').setHorizontalAlignment('center');
    sh.getRange(r, 9, 1, 3).merge().setFormula(`=IFERROR(INDEX(${flt};${k};3);"")`)
      .setNumberFormat('#,##0.00 €').setFontSize(11).setFontColor(COLOR.texto)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
    r++;
  }

  sh.setRowHeight(r, 18);
  sh.getRange(r, 2, 1, 10).merge()
    .setFormula(`=IF(${cnt}>${reserva};"+"&(${cnt}-${reserva})&" más · ver hoja Movimientos";"")`)
    .setFontSize(9).setFontStyle('italic').setFontColor(COLOR.tenue)
    .setBackground(COLOR.fondo).setHorizontalAlignment('right').setVerticalAlignment('middle');
  r++;

  sh.setRowHeight(r, 26);
  sh.getRange(r, 2, 1, 10).setBackground(COLOR.cab);
  sh.getRange(r, 2, 1, 7).merge().setValue('Total ' + etiqueta.toLowerCase())
    .setFontSize(10).setFontWeight('bold').setFontColor(COLOR.cabTxt).setVerticalAlignment('middle');
  sh.getRange(r, 9, 1, 3).merge().setFormula(totalFormula).setNumberFormat('#,##0.00 €')
    .setFontSize(12).setFontWeight('bold').setFontColor(COLOR.cabTxt)
    .setVerticalAlignment('middle').setHorizontalAlignment('right');
  return r + 2;
}

/** Total grande en tarjeta clara con número coloreado por signo. */
function totalGrande(sh, r, label, valFormula, signo) {
  sh.setRowHeight(r, 10);
  r++;
  sh.setRowHeight(r, 48);
  sh.getRange(r, 2, 1, 10).setBackground(COLOR.panel)
    .setBorder(true, true, true, true, false, false, COLOR.tinta, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sh.getRange(r, 2, 1, 6).merge().setValue(label)
    .setFontSize(13).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');
  sh.getRange(r, 8, 1, 4).merge().setFormula(valFormula).setNumberFormat('#,##0.00 €')
    .setFontSize(22).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setHorizontalAlignment('right');
  if (signo) aplicarPositivoNegativo(sh, sh.getRange(r, 8, 1, 4));
  return r + 1;
}

/** Tabla "Dónde se gasta": categorías ordenadas con barra SPARKLINE nativa. */
function tablaCategorias(sh, r, C) {
  const n = 8;
  sh.setRowHeight(r, 22);
  sh.getRange(r, 2, 1, 10).setBackground(COLOR.fondo)
    .setBorder(false, false, true, false, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(r, 2, 1, 3).merge().setValue('Categoría')
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue).setVerticalAlignment('bottom');
  sh.getRange(r, 9, 1, 2).merge().setValue('Importe')
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setHorizontalAlignment('right').setVerticalAlignment('bottom');
  sh.getRange(r, 11).setValue('%')
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setHorizontalAlignment('right').setVerticalAlignment('bottom');
  r++;
  for (let i = 0; i < n; i++) {
    const cr = 2 + i;
    sh.setRowHeight(r, 24);
    const fondo = i % 2 === 0 ? COLOR.panel : COLOR.cebra;
    sh.getRange(r, 2, 1, 10).setBackground(fondo);
    sh.getRange(r, 2, 1, 3).merge().setFormula(`=IFERROR(${C}!G${cr};"")`)
      .setFontSize(11).setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.getRange(r, 5, 1, 4).merge()
      .setFormula(`=IF(${C}!H${cr}="";"";IFERROR(SPARKLINE(${C}!H${cr};{"charttype"\\"bar";"max"\\${C}!H$2;"color1"\\"#2563EB";"empty"\\"zero"});""))`)
      .setVerticalAlignment('middle');
    sh.getRange(r, 9, 1, 2).merge().setFormula(`=IFERROR(${C}!H${cr};"")`).setNumberFormat('#,##0 €')
      .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
    sh.getRange(r, 11).setFormula(`=IFERROR(${C}!H${cr}/${C}!B12;"")`).setNumberFormat('0%')
      .setFontSize(10).setFontColor(COLOR.tenue).setVerticalAlignment('middle').setHorizontalAlignment('right');
    r++;
  }
  return r;
}

/** Objetivos a largo plazo con barra de progreso SPARKLINE. */
function barrasObjetivos(sh, r) {
  const OBJ = `'${HOJAS.OBJETIVOS}'`;
  sh.setRowHeight(r, 22);
  sh.getRange(r, 2, 1, 10).setBackground(COLOR.fondo)
    .setBorder(false, false, true, false, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(r, 2, 1, 3).merge().setValue('Objetivo')
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue).setVerticalAlignment('bottom');
  sh.getRange(r, 9, 1, 3).merge().setValue('Aportado / Meta')
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setHorizontalAlignment('right').setVerticalAlignment('bottom');
  r++;
  for (let i = 0; i < 5; i++) {
    const o = 2 + i;
    sh.setRowHeight(r, 26);
    const fondo = i % 2 === 0 ? COLOR.panel : COLOR.cebra;
    sh.getRange(r, 2, 1, 10).setBackground(fondo);
    sh.getRange(r, 2, 1, 3).merge().setFormula(`=IFERROR(${OBJ}!A${o};"")`)
      .setFontSize(11).setFontColor(COLOR.tinta).setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.getRange(r, 5, 1, 3).merge()
      .setFormula(`=IF(${OBJ}!A${o}="";"";IFERROR(SPARKLINE(${OBJ}!D${o};{"charttype"\\"bar";"max"\\1;"color1"\\"#2563EB"});""))`)
      .setVerticalAlignment('middle');
    sh.getRange(r, 8).setFormula(`=IF(${OBJ}!A${o}="";"";TEXT(${OBJ}!D${o};"0%"))`)
      .setFontSize(10).setFontColor(COLOR.acento).setVerticalAlignment('middle').setHorizontalAlignment('left');
    sh.getRange(r, 9, 1, 3).merge()
      .setFormula(`=IF(${OBJ}!A${o}="";"";TEXT(${OBJ}!C${o};"#,##0 €")&" / "&TEXT(${OBJ}!B${o};"#,##0 €"))`)
      .setFontSize(10).setFontColor(COLOR.texto).setVerticalAlignment('middle').setHorizontalAlignment('right');
    r++;
  }
  return r;
}

/** Gráfico combinado 12 meses (líneas ingresos/gastos + barras ahorro). */
function insertarGraficoEvolucionEn(sh, ss, filaPos) {
  const calc = ss.getSheetByName(HOJAS.CALC);
  const grafico = sh.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(calc.getRange('J1:J13'))
    .addRange(calc.getRange('M1:M13'))
    .addRange(calc.getRange('N1:N13'))
    .addRange(calc.getRange('K1:K13'))
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setOption('title', '')
    .setOption('legend', { position: 'top', alignment: 'center', textStyle: { color: COLOR.texto, fontSize: 10 } })
    .setOption('series', {
      0: { type: 'line', color: '#16A34A', lineWidth: 2, pointSize: 3 },
      1: { type: 'line', color: '#DC2626', lineWidth: 2, pointSize: 3 },
      2: { type: 'bars', color: '#2563EB' },
    })
    .setOption('hAxis', { textStyle: { color: COLOR.tenue, fontSize: 9 } })
    .setOption('vAxis', { textStyle: { color: COLOR.tenue, fontSize: 9 }, format: '#,##0 €', gridlines: { color: COLOR.borde } })
    .setOption('backgroundColor', COLOR.panel)
    .setOption('chartArea', { left: 60, top: 40, width: '88%', height: '72%' })
    .setOption('width', 900)
    .setOption('height', 300)
    .setPosition(filaPos, 2, 0, 0)
    .build();
  sh.insertChart(grafico);
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
