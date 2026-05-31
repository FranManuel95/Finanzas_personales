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

  sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  sh.getRange('F:F').setNumberFormat('#,##0.00 €');
  sh.getRange('A2:G').setFontColor(COLOR.texto).setVerticalAlignment('middle');

  validarLista(sh, 'B2:B', TIPOS);
  validarLista(sh, 'C2:C', PERSONAS);
  validarRango(sh, 'D2:D', `${HOJAS.AJUSTES}!D2:F`, true);

  const hoy = new Date();
  const y = hoy.getFullYear();
  const m = hoy.getMonth();
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
 * Estructura:
 *  - B1..B20:  CONTRATO con el bot. NO CAMBIAR.
 *  - B21..B40: Métricas adicionales para el Panel rediseñado.
 *  - E:G:      Tabla "categoría / importe mes / importe media 3M".
 *  - H:J:      Top 10 categorías ordenadas (cat / importe / delta vs media 3M).
 *  - K:Q:      Series mensuales 12 meses (mes, ahorro, acumulado, ingreso, gasto, fijos, variables).
 *  - R:U:      Top 3 categorías de gastos individuales por persona.
 */
function crearCalc(ss) {
  const sh = ss.insertSheet(HOJAS.CALC);

  const MOV = HOJAS.MOVIMIENTOS;
  const AJ = HOJAS.AJUSTES;

  // Mes activo y rango de fechas del mes (según contrato).
  // Robusto a que Sheets haya interpretado el "2026-05" como fecha: si es número (Date),
  // lo convertimos a texto "yyyy-mm" antes de hacer LEFT/MID.
  const mesActivoRaw = `VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE)`;
  const mesActivo = `IF(ISNUMBER(${mesActivoRaw});TEXT(${mesActivoRaw};"yyyy-mm");${mesActivoRaw})`;
  const inicioMes = `DATE(VALUE(LEFT(${mesActivo};4));VALUE(MID(${mesActivo};6;2));1)`;
  const finMes = `EOMONTH(${inicioMes};0)+1`;

  // SUMIFS sobre Movimientos del mes activo.
  const sumar = (tipo, persona) => {
    let f = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};${MOV}!B:B;"${tipo}"`;
    if (persona) f += `;${MOV}!C:C;"${persona}"`;
    f += `)`;
    return `IFERROR(${f};0)`;
  };

  const objetivo = `IFERROR(VLOOKUP("Objetivo ahorro mensual conjunto (€)";${AJ}!A:B;2;FALSE);0)`;

  /* ===== B1..B20: CONTRATO con el bot. NO CAMBIAR significado ni posición ===== */
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
    ['Ahorro acumulado del año',          ''], // rellenado abajo via INDEX a serie 12M
    ['Balance del mes',                   '=B17'],
  ];
  sh.getRange(1, 1, metricas.length, 2).setValues(metricas);

  /* ===== B21..B40: Métricas adicionales para el Panel rediseñado ===== */
  const extra = [
    ['% aporte real bote FM',             '=IFERROR(B4/B6;0)'],
    ['% aporte real bote Lucía',          '=IFERROR(B5/B6;0)'],
    ['% ingreso FM al bote',              '=IFERROR(B4/B1;0)'],
    ['% ingreso Lucía al bote',           '=IFERROR(B5/B2;0)'],
    ['Tasa de ahorro del mes',            '=IFERROR(B17/B3;0)'],
    ['Tasa de ahorro objetivo',           `=IFERROR(VLOOKUP("Tasa de ahorro objetivo (%)";${AJ}!A:B;2;FALSE);0.2)`],
    ['Ingresos media 3M',                 ''], // rellenado abajo
    ['Gastos media 3M',                   ''],
    ['Ahorro media 3M',                   ''],
    ['Delta ingresos vs media 3M',        '=IFERROR(IF(B27=0;0;B3/B27-1);0)'],
    ['Delta gastos vs media 3M',          '=IFERROR(IF(B28=0;0;B12/B28-1);0)'],
    ['% gastos fijos sobre total',        '=IFERROR(B7/B12;0)'],
    ['% gastos variables sobre total',    '=IFERROR(B8/B12;0)'],
    ['% gastos individuales sobre total', '=IFERROR((B10+B11)/B12;0)'],
    ['Gastos individuales total',         '=B10+B11'],
    ['Gastos esenciales media 3M',        ''], // rellenado abajo (fijos media)
    ['Reserva emergencia (meses cub.)',   '=IFERROR(IF(B36=0;0;B19/B36);0)'],
    ['Reserva emergencia objetivo',       `=IFERROR(VLOOKUP("Reserva de emergencia objetivo (meses)";${AJ}!A:B;2;FALSE);3)`],
    ['Modo aportación bote',              `=IFERROR(VLOOKUP("Modo aportación al bote";${AJ}!A:B;2;FALSE);"50/50")`],
    ['Settle-up sugerido (€)',            '=IF(B39<>"Proporcional";0;ROUND(ABS((B23-B24)*B3/2);2))'],
  ];
  sh.getRange(21, 1, extra.length, 2).setValues(extra);

  // Formatos de la columna B
  sh.getRange('B1:B17').setNumberFormat('#,##0.00 €');
  sh.getRange('B18').setNumberFormat('0.0%');
  sh.getRange('B19:B20').setNumberFormat('#,##0.00 €');
  sh.getRange('B21:B25').setNumberFormat('0.0%');
  sh.getRange('B26').setNumberFormat('0.0%');
  sh.getRange('B27:B29').setNumberFormat('#,##0.00 €');
  sh.getRange('B30:B31').setNumberFormat('+0.0%;-0.0%;0.0%');
  sh.getRange('B32:B34').setNumberFormat('0.0%');
  sh.getRange('B35').setNumberFormat('#,##0.00 €');
  sh.getRange('B36').setNumberFormat('#,##0.00 €');
  sh.getRange('B37').setNumberFormat('0.0" meses"');
  sh.getRange('B38').setNumberFormat('0" meses"');
  sh.getRange('B40').setNumberFormat('#,##0.00 €');

  sh.setColumnWidth(1, 280);
  sh.setColumnWidth(2, 140);
  sh.getRange('A1:B40').setFontFamily(FUENTE).setFontColor(COLOR.texto);

  /* ===== TABLA AUX: gasto por categoría (E:G) ===== */
  sh.getRange('E1:G1').setValues([['Categoría', 'Importe mes', 'Importe media 3M']]);
  const catsApiladas = `{${AJ}!D2:D;${AJ}!E2:E;${AJ}!F2:F}`;
  sh.getRange('E2').setFormula(
    `=IFERROR(SORT(UNIQUE(FILTER(${catsApiladas};${catsApiladas}<>"")));"")`
  );

  // Importe del mes activo por categoría (suma de los 3 tipos de gasto).
  sh.getRange('F2').setFormula(
    `=ARRAYFORMULA(IF(E2:E40="";"";` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};` +
    `${MOV}!D:D;E2:E40;${MOV}!B:B;"Compartido fijo");0)+` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};` +
    `${MOV}!D:D;E2:E40;${MOV}!B:B;"Compartido variable");0)+` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};` +
    `${MOV}!D:D;E2:E40;${MOV}!B:B;"Individual");0)))`
  );

  // Importe media 3 meses anteriores (para el delta del Top categorías).
  const ini3M = `DATE(VALUE(LEFT(${mesActivo};4));VALUE(MID(${mesActivo};6;2))-3;1)`;
  sh.getRange('G2').setFormula(
    `=ARRAYFORMULA(IF(E2:E40="";"";` +
    `IFERROR((` +
    `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini3M};${MOV}!A:A;"<"&${inicioMes};${MOV}!D:D;E2:E40;${MOV}!B:B;"Compartido fijo")+` +
    `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini3M};${MOV}!A:A;"<"&${inicioMes};${MOV}!D:D;E2:E40;${MOV}!B:B;"Compartido variable")+` +
    `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini3M};${MOV}!A:A;"<"&${inicioMes};${MOV}!D:D;E2:E40;${MOV}!B:B;"Individual")` +
    `)/3;0)))`
  );
  sh.getRange('F2:G40').setNumberFormat('#,##0.00 €');

  /* ===== Delta vs media 3M por categoría (H, precalculado) ===== */
  // Se calcula aparte para que la QUERY del top NO necesite IFERROR/aritmética en su SELECT.
  sh.getRange('H1').setValue('Delta vs 3M');
  sh.getRange('H2').setFormula(
    `=ARRAYFORMULA(IF(E2:E40="";"";IFERROR(F2:F40/G2:G40-1;0)))`
  );
  sh.getRange('H2:H40').setNumberFormat('0.0%');

  /* ===== TOP 10 categorías ordenadas (I:K) ===== */
  sh.getRange('I1:K1').setValues([['Categoría top', 'Importe top', 'Delta top']]);
  // QUERY simple sobre {E, F, G, H}: selecciona cat, importe y delta donde importe > 0,
  // ordena desc por importe y limita a 10. Sin IFERROR en SELECT (no es válido en QUERY).
  sh.getRange('I2').setFormula(
    `=IFERROR(QUERY({E2:H40};"select Col1, Col2, Col4 where Col2 > 0 order by Col2 desc limit 10";0);"")`
  );
  sh.getRange('J2:J11').setNumberFormat('#,##0.00 €');
  sh.getRange('K2:K11').setNumberFormat('+0.0%;-0.0%;"="');

  /* ===== Series mensuales 12 meses del año activo (K:Q) ===== */
  sh.getRange('K1:Q1').setValues([['Mes', 'Ahorro', 'Acumulado', 'Ingreso', 'Gasto', 'Fijos', 'Variables']]);
  const anio = `VALUE(LEFT(${mesActivo};4))`;
  for (let m = 1; m <= 12; m++) {
    const r = m + 1;
    const ini = `DATE(${anio};${m};1)`;
    const fin = `DATE(${anio};${m + 1};1)`;
    sh.getRange(r, 11).setFormula(`=TEXT(${ini};"yyyy-mm")`);

    const ingMes = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Ingreso")`;
    const fijMes = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Compartido fijo")`;
    const varMes = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Compartido variable")`;
    const indMes = `SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${ini};${MOV}!A:A;"<"&${fin};${MOV}!B:B;"Individual")`;
    const gastTot = `(${fijMes})+(${varMes})+(${indMes})`;

    sh.getRange(r, 12).setFormula(`=IFERROR((${ingMes})-(${gastTot});0)`);            // L: Ahorro
    sh.getRange(r, 13).setFormula(m === 1 ? '=L2' : `=M${r - 1}+L${r}`);              // M: Acumulado
    sh.getRange(r, 14).setFormula(`=IFERROR(${ingMes};0)`);                            // N: Ingreso
    sh.getRange(r, 15).setFormula(`=IFERROR(${gastTot};0)`);                           // O: Gasto total
    sh.getRange(r, 16).setFormula(`=IFERROR(${fijMes};0)`);                            // P: Fijos
    sh.getRange(r, 17).setFormula(`=IFERROR(${varMes};0)`);                            // Q: Variables
  }
  sh.getRange('L2:Q13').setNumberFormat('#,##0.00 €');

  // B19 = acumulado anual hasta el mes activo.
  const idxMes = `VALUE(MID(${mesActivo};6;2))`;
  sh.getRange('B19').setFormula(`=IFERROR(INDEX(M2:M13;${idxMes});0)`);

  // Medias móviles 3M usando OFFSET sobre la tabla 12 meses.
  // Si idxMes <= 1, no hay meses anteriores → 0.
  const mediaPrev = (col, idxLetra) => {
    // Promedio de los hasta-3 meses anteriores al activo (sin incluirlo).
    return `=IFERROR(IF(${idxLetra}<=1;0;AVERAGE(OFFSET(${col}2;MAX(0;${idxLetra}-4);0;MIN(3;${idxLetra}-1);1)));0)`;
  };
  sh.getRange('B27').setFormula(mediaPrev('N', idxMes));  // Ingresos 3M
  sh.getRange('B28').setFormula(mediaPrev('O', idxMes));  // Gastos 3M
  sh.getRange('B29').setFormula(mediaPrev('L', idxMes));  // Ahorro 3M
  sh.getRange('B36').setFormula(mediaPrev('P', idxMes));  // Fijos 3M (esenciales)

  /* ===== Top 3 gastos individuales por persona (R:U) ===== */
  sh.getRange('R1:U1').setValues([['Categoría FM', 'Importe FM', 'Categoría Lucía', 'Importe Lucía']]);

  // Importes por categoría restringidos a Tipo=Individual y Persona=FM/Lucía.
  // Lo hacemos vía tabla auxiliar en V/W (FM) y X/Y (Lucía), luego QUERY top 3.
  sh.getRange('V1:W1').setValues([['Cat', 'FM']]);
  sh.getRange('X1:Y1').setValues([['Cat', 'Lucía']]);

  // V = lista de categorías del bloque Individual de Ajustes (col F).
  sh.getRange('V2').setFormula(`=IFERROR(SORT(UNIQUE(FILTER(${AJ}!F2:F;${AJ}!F2:F<>"")));"")`);
  sh.getRange('X2').setFormula(`=IFERROR(SORT(UNIQUE(FILTER(${AJ}!F2:F;${AJ}!F2:F<>"")));"")`);

  // W = importe individual FM por categoría del mes activo.
  sh.getRange('W2').setFormula(
    `=ARRAYFORMULA(IF(V2:V20="";"";` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};` +
    `${MOV}!D:D;V2:V20;${MOV}!B:B;"Individual";${MOV}!C:C;"FM");0)))`
  );
  sh.getRange('Y2').setFormula(
    `=ARRAYFORMULA(IF(X2:X20="";"";` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};` +
    `${MOV}!D:D;X2:X20;${MOV}!B:B;"Individual";${MOV}!C:C;"Lucía");0)))`
  );

  // R:S = top 3 FM, T:U = top 3 Lucía.
  sh.getRange('R2').setFormula(
    `=IFERROR(QUERY({V2:W20};"select Col1, Col2 where Col2 > 0 order by Col2 desc limit 3 label Col1 '', Col2 ''";0);"")`
  );
  sh.getRange('T2').setFormula(
    `=IFERROR(QUERY({X2:Y20};"select Col1, Col2 where Col2 > 0 order by Col2 desc limit 3 label Col1 '', Col2 ''";0);"")`
  );
  sh.getRange('S2:S4').setNumberFormat('#,##0.00 €');
  sh.getRange('U2:U4').setNumberFormat('#,##0.00 €');
  sh.getRange('W2:W20').setNumberFormat('#,##0.00 €');
  sh.getRange('Y2:Y20').setNumberFormat('#,##0.00 €');

  // Anchos de las columnas auxiliares (todas ocultas dentro de _Calc, pero por limpieza).
  sh.setColumnWidth(5, 160);
  sh.setColumnWidth(6, 110);
  sh.setColumnWidth(7, 110);
  sh.setColumnWidth(8, 160);
  sh.setColumnWidth(9, 110);
  sh.setColumnWidth(10, 100);
  sh.setColumnWidth(11, 90);
  sh.setColumnWidth(12, 110);
  sh.setColumnWidth(13, 110);
  sh.setColumnWidth(14, 110);
  sh.setColumnWidth(15, 110);
  sh.setColumnWidth(16, 110);
  sh.setColumnWidth(17, 110);
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

/* ====================== HOJA: PANEL (dashboard rediseñado) ======================
 * Layout (cols B..K, 10 cols de contenido; A y L márgenes laterales):
 *  Filas 2-3    Cabecera + sub-cabecera
 *  Filas 5-9    Bloque A — 5 KPI cards
 *  Filas 11-16  Bloque B — Reparto por persona (tabla densa)
 *  Filas 18-22  Bloque C — Bote común
 *  Filas 24-31  Bloque D1 — Top categorías del mes
 *  Filas 33-37  Bloque D2 — Fijos/Variables/Individuales (barra apilada)
 *  Filas 39-45  Bloque D3 — Gastos individuales por persona (lado a lado)
 *  Filas 47-51  Bloque F — Objetivos top 3
 *  Filas 53+    Bloque G — Gráfico evolución 12 meses
 */
function crearPanel(ss) {
  const sh = ss.insertSheet(HOJAS.PANEL);
  const C = `'${HOJAS.CALC}'`;
  const AJ = HOJAS.AJUSTES;

  sh.setHiddenGridlines(true);
  sh.getRange('A1:L90').setBackground(COLOR.fondo).setFontFamily(FUENTE);
  sh.setColumnWidth(1, 20);            // margen izq
  sh.setColumnWidths(2, 10, 88);       // B..K contenido
  sh.setColumnWidth(12, 20);           // margen der

  /* ===== Cabecera ===== */
  sh.setRowHeight(2, 42);
  sh.getRange('B2:H2').merge()
    .setValue('Finanzas FM & Lucía')
    .setFontSize(24).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setBackground(COLOR.fondo);
  sh.getRange('I2:K2').merge()
    .setFormula(`="Actualizado " & TEXT(NOW();"yyyy-mm-dd HH:mm")`)
    .setFontSize(9).setFontColor(COLOR.tenue)
    .setHorizontalAlignment('right').setVerticalAlignment('bottom').setBackground(COLOR.fondo);
  sh.setRowHeight(3, 22);
  // Wrap con TEXT por si Sheets interpretara B2 como fecha (defensa en profundidad).
  sh.getRange('B3:K3').merge()
    .setFormula(
      `="Resumen del mes " & IF(ISNUMBER(VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE));` +
      `TEXT(VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE);"yyyy-mm");VLOOKUP("Mes activo";${AJ}!A:B;2;FALSE))`)
    .setFontSize(11).setFontColor(COLOR.tenue)
    .setVerticalAlignment('middle').setBackground(COLOR.fondo);

  /* ===== BLOQUE A — 5 KPI cards (filas 5-9, cols B..K en bloques de 2) ===== */
  // B:C | D:E | F:G | H:I | J:K  → 5 tarjetas de 2 cols cada una.
  // Todos los subtextos van envueltos en IFERROR para que un cálculo todavía sin datos
  // no rompa la tarjeta con #ERROR.
  pintarKPI(sh, 5, 2, 3,  'INGRESOS DEL MES', `=${C}!B3`,  'euro',
    `=IFERROR("vs media 3M  " & TEXT(${C}!B30;"+0.0%;-0.0%;0.0%");"")`, 'delta');
  pintarKPI(sh, 5, 4, 5,  'GASTOS DEL MES',   `=${C}!B12`, 'euro',
    `=IFERROR("vs media 3M  " & TEXT(${C}!B31;"+0.0%;-0.0%;0.0%");"")`, 'deltaGasto');
  pintarKPI(sh, 5, 6, 7,  'TASA DE AHORRO',   `=${C}!B25`, 'pct',
    `=IFERROR("objetivo  " & TEXT(${C}!B26;"0%");"")`, 'tasaAhorro');
  pintarKPI(sh, 5, 8, 9,  'BOTE SOBRANTE',    `=${C}!B15`, 'euro',
    `=IFERROR(IF(${C}!B6=0;"sin aportaciones";TEXT(${C}!B15/${C}!B6;"0.0%") & " del bote");"")`, 'signo');
  pintarKPI(sh, 5, 10, 11,'RESERVA EMERG.',   `=${C}!B37`, 'meses',
    `=IFERROR("objetivo  " & TEXT(${C}!B38;"0") & " meses";"")`, 'reserva');

  /* ===== BLOQUE B — Reparto por persona (filas 11-16) ===== */
  etiquetaSeccion(sh, 11, 'REPARTO DEL MES POR PERSONA');
  bloquePersonas(sh, 12, C);

  /* ===== BLOQUE C — Bote común (filas 18-22) ===== */
  etiquetaSeccion(sh, 18, 'BOTE COMÚN');
  bloqueBote(sh, 19, C);

  /* ===== BLOQUE D1 — Top categorías (filas 24-31) ===== */
  etiquetaSeccion(sh, 24, 'TOP CATEGORÍAS DEL MES — DÓNDE SE GASTA');
  bloqueTopCategorias(sh, 25, C);

  /* ===== BLOQUE D2 — Fijos vs Variables vs Individuales (filas 33-37) ===== */
  etiquetaSeccion(sh, 33, 'COMPOSICIÓN DEL GASTO — DÓNDE PUEDES RECORTAR');
  bloqueComposicion(sh, 34, C);

  /* ===== BLOQUE D3 — Gastos individuales por persona (filas 39-45) ===== */
  etiquetaSeccion(sh, 39, 'GASTOS INDIVIDUALES — DÓNDE SE VA LO PERSONAL');
  bloqueIndividuales(sh, 40, C);

  /* ===== BLOQUE F — Objetivos top 3 (filas 47-51) ===== */
  etiquetaSeccion(sh, 47, 'OBJETIVOS A LARGO PLAZO');
  bloqueObjetivos(sh, 48);

  /* ===== BLOQUE G — Evolución 12 meses (gráfico) ===== */
  etiquetaSeccion(sh, 53, 'EVOLUCIÓN 12 MESES');
  insertarGraficoEvolucion(sh, ss);

  sh.setFrozenRows(3);
}

/* ====================== HELPERS DE PANEL ====================== */

/**
 * Tarjeta KPI ocupando columnas [colIni..colFin] x filas [fila..fila+3].
 * fila   = barra de acento (fina)
 * fila+1 = etiqueta pequeña
 * fila+2 = valor grande
 * fila+3 = subtexto con delta/objetivo
 * formato: 'euro' | 'pct' | 'meses'
 * modo: 'delta' (verde sube ing, rojo baja) | 'deltaGasto' (rojo sube, verde baja) |
 *       'tasaAhorro' | 'signo' | 'reserva'
 */
function pintarKPI(sh, fila, colIni, colFin, etiqueta, formulaValor, formato, formulaDetalle, modo) {
  const nCols = colFin - colIni + 1;
  const card = sh.getRange(fila, colIni, 4, nCols);
  card.setBackground(COLOR.panel)
    .setBorder(true, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  // Barra de acento superior
  sh.getRange(fila, colIni, 1, nCols).setBackground(COLOR.acento);
  sh.setRowHeight(fila, 4);

  // Etiqueta
  sh.getRange(fila + 1, colIni, 1, nCols).merge()
    .setValue(etiqueta)
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 1, 22);

  // Valor
  const valor = sh.getRange(fila + 2, colIni, 1, nCols).merge()
    .setFormula(formulaValor)
    .setFontSize(20).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBackground(COLOR.panel);
  if (formato === 'pct') valor.setNumberFormat('0.0%');
  else if (formato === 'meses') valor.setNumberFormat('0.0" meses"');
  else valor.setNumberFormat('#,##0 €');
  sh.setRowHeight(fila + 2, 36);

  // Subtexto
  sh.getRange(fila + 3, colIni, 1, nCols).merge()
    .setFormula(formulaDetalle)
    .setFontSize(9).setFontColor(COLOR.tenue)
    .setVerticalAlignment('top').setHorizontalAlignment('left')
    .setBackground(COLOR.panel);
  sh.setRowHeight(fila + 3, 24);

  // Color del valor según modo
  const valCell = sh.getRange(fila + 2, colIni, 1, nCols);
  if (modo === 'signo') {
    aplicarPositivoNegativo(sh, valCell);
  } else if (modo === 'tasaAhorro') {
    aplicarTasaAhorro(sh, valCell);
  } else if (modo === 'reserva') {
    aplicarReserva(sh, valCell);
  }
}

/* ---------- BLOQUE B: Reparto por persona ---------- */
function bloquePersonas(sh, fila, C) {
  // Cabecera
  sh.setRowHeight(fila, 28);
  const cab = [
    [2, 3, ''],
    [4, 5, 'Ingreso neto'],
    [6, 7, 'Aportación bote'],
    [8, 9, 'Gastos individuales'],
    [10, 11, 'Saldo personal'],
  ];
  cab.forEach(([c1, c2, txt], i) => {
    sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setValue(txt)
      .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.cabTxt)
      .setBackground(COLOR.acento).setVerticalAlignment('middle')
      .setHorizontalAlignment(i === 0 ? 'left' : 'right');
  });

  // Dos filas: FM y Lucía. Saldo = ingreso − aportación al bote − gastos individuales.
  filaPersona(sh, fila + 1, 'FM',    `=${C}!B1`, `=${C}!B4`, `=${C}!B10`, `=${C}!B13`, COLOR.panel);
  filaPersona(sh, fila + 2, 'Lucía', `=${C}!B2`, `=${C}!B5`, `=${C}!B11`, `=${C}!B14`, COLOR.cebra);

  // Banda resumen: una frase con la verdad agregada del mes (no intenta cuadrar columnas).
  const fr = fila + 3;
  sh.setRowHeight(fr, 30);
  sh.getRange(fr, 2, 1, 10).merge()
    .setFormula(
      `="Conjunto: ingresos " & TEXT(${C}!B3;"#,##0 €") & ` +
      `"  ·  bote pagó " & TEXT(${C}!B9;"#,##0 €") & " en compartidos" & ` +
      `"  ·  AHORRO DEL MES " & TEXT(${C}!B17;"#,##0 €")`)
    .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setBackground(COLOR.acentoSuave).setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBorder(true, true, true, true, false, false, COLOR.acento, SpreadsheetApp.BorderStyle.SOLID);
}

function filaPersona(sh, fila, nombre, fIng, fApo, fGas, fSaldo, fondo, esTotal) {
  sh.setRowHeight(fila, 30);
  sh.getRange(fila, 2, 1, 10).setBackground(fondo)
    .setBorder(false, false, true, false, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  sh.getRange(fila, 2, 1, 2).merge().setValue(nombre)
    .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta)
    .setVerticalAlignment('middle').setHorizontalAlignment('left');

  const celdas = [[4, 5, fIng], [6, 7, fApo], [8, 9, fGas], [10, 11, fSaldo]];
  celdas.forEach(([c1, c2, f], i) => {
    const cell = sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setFormula(f).setNumberFormat('#,##0.00 €')
      .setFontSize(11).setFontColor(COLOR.texto)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
    if (i === 3) {
      cell.setFontWeight('bold');
      aplicarPositivoNegativo(sh, sh.getRange(fila, c1, 1, c2 - c1 + 1));
    }
    if (esTotal) cell.setFontWeight('bold');
  });
  if (esTotal) {
    sh.getRange(fila, 2, 1, 10)
      .setBorder(true, null, true, null, null, null, COLOR.tinta, SpreadsheetApp.BorderStyle.SOLID);
  }
}

/* ---------- BLOQUE C: Bote común ---------- */
function bloqueBote(sh, fila, C) {
  // Cabecera 4 celdas
  sh.setRowHeight(fila, 26);
  const cab = [
    [2, 3,  'Aportado al bote'],
    [4, 6,  'Gastos fijos'],
    [7, 8,  'Gastos variables'],
    [9, 11, 'Sobrante'],
  ];
  cab.forEach(([c1, c2, txt]) => {
    sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setValue(txt)
      .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.cabTxt)
      .setBackground(COLOR.acento).setVerticalAlignment('middle')
      .setHorizontalAlignment('center');
  });

  // Valores
  const fr = fila + 1;
  sh.setRowHeight(fr, 38);
  sh.getRange(fr, 2, 1, 10).setBackground(COLOR.panel)
    .setBorder(false, false, true, false, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  const vals = [
    [2, 3,  `=${C}!B6`,  false],
    [4, 6,  `=${C}!B7`,  false],
    [7, 8,  `=${C}!B8`,  false],
    [9, 11, `=${C}!B15`, true],
  ];
  vals.forEach(([c1, c2, f, signo]) => {
    const cell = sh.getRange(fr, c1, 1, c2 - c1 + 1).merge()
      .setFormula(f).setNumberFormat('#,##0.00 €')
      .setFontSize(16).setFontWeight('bold').setFontColor(COLOR.tinta)
      .setVerticalAlignment('middle').setHorizontalAlignment('center');
    if (signo) aplicarPositivoNegativo(sh, sh.getRange(fr, c1, 1, c2 - c1 + 1));
  });

  // Sub-texto debajo: explica el bote según modo
  sh.setRowHeight(fr + 1, 22);
  sh.getRange(fr + 1, 2, 1, 10).merge()
    .setFormula(
      `=IF(${C}!B39="50/50";` +
      `"Modo 50/50 — cada uno aporta lo acordado. Sobrante " & TEXT(IFERROR(${C}!B15/${C}!B6;0);"0%") & " del aportado.";` +
      `IF(${C}!B39="Proporcional";` +
      `"Modo proporcional — FM aporta " & TEXT(${C}!B23;"0%") & " de su ingreso, Lucía " & TEXT(${C}!B24;"0%") & ".";` +
      `"Modo custom — ratio definido manualmente."))`)
    .setFontSize(10).setFontColor(COLOR.tenue)
    .setVerticalAlignment('middle').setHorizontalAlignment('center')
    .setBackground(COLOR.fondo);
}

/* ---------- BLOQUE D1: Top 5 categorías ---------- */
function bloqueTopCategorias(sh, fila, C) {
  // Cabecera
  sh.setRowHeight(fila, 24);
  const cab = [
    [2, 5,  'Categoría'],
    [6, 8,  'Barra'],
    [9, 10, 'Importe'],
    [11, 11,'vs 3M'],
  ];
  cab.forEach(([c1, c2, txt], i) => {
    sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setValue(txt)
      .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
      .setBackground(COLOR.fondo).setVerticalAlignment('middle')
      .setHorizontalAlignment(i === 0 ? 'left' : (i === 1 ? 'left' : 'right'))
      .setBorder(null, null, true, null, null, null, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  });

  // Filas 1..5 (top 5). El top vive en _Calc!I:K (cat / importe / delta).
  for (let i = 0; i < 5; i++) {
    const r = fila + 1 + i;
    const calcRow = 2 + i; // _Calc!I2..I6
    sh.setRowHeight(r, 24);
    const fondo = i % 2 === 0 ? COLOR.panel : COLOR.cebra;
    sh.getRange(r, 2, 1, 10).setBackground(fondo);

    // Categoría
    sh.getRange(r, 2, 1, 4).merge()
      .setFormula(`=IFERROR(${C}!I${calcRow};"")`)
      .setFontSize(11).setFontColor(COLOR.tinta)
      .setVerticalAlignment('middle').setHorizontalAlignment('left');

    // Barra horizontal proporcional al top 1 (REPT con bloque Unicode)
    sh.getRange(r, 6, 1, 3).merge()
      .setFormula(
        `=IFERROR(IF(${C}!J${calcRow}="";"";REPT("█";MAX(1;ROUND(${C}!J${calcRow}/${C}!J$2*22;0))));"")`
      )
      .setFontSize(10).setFontColor(COLOR.acento)
      .setVerticalAlignment('middle').setHorizontalAlignment('left');

    // Importe
    sh.getRange(r, 9, 1, 2).merge()
      .setFormula(`=IFERROR(${C}!J${calcRow};"")`)
      .setNumberFormat('#,##0 €')
      .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');

    // Delta vs 3M
    const delta = sh.getRange(r, 11)
      .setFormula(`=IFERROR(${C}!K${calcRow};"")`)
      .setNumberFormat('+0%;-0%;"="')
      .setFontSize(10).setFontColor(COLOR.texto)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
    aplicarDeltaGasto(sh, delta);
  }

  // Borde inferior bajo la tabla
  sh.getRange(fila + 5, 2, 1, 10)
    .setBorder(null, null, true, null, null, null, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
}

/* ---------- BLOQUE D2: Composición fijos/variables/individuales ---------- */
function bloqueComposicion(sh, fila, C) {
  // Tres celdas-stack horizontales mostrando % y €, una al lado de otra.
  sh.setRowHeight(fila, 22);
  sh.setRowHeight(fila + 1, 34);
  sh.setRowHeight(fila + 2, 20);
  sh.setRowHeight(fila + 3, 24);

  // Cabecera de cada bloque (3 bloques: B:E, F:H, I:K)
  const bloques = [
    [2, 5, 'GASTOS FIJOS',         `=${C}!B7`,  `=${C}!B32`],
    [6, 8, 'GASTOS VARIABLES',     `=${C}!B8`,  `=${C}!B33`],
    [9, 11,'GASTOS INDIVIDUALES',  `=${C}!B35`, `=${C}!B34`],
  ];
  bloques.forEach(([c1, c2, titulo, fEuro, fPct]) => {
    // Etiqueta
    sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setValue(titulo)
      .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
      .setBackground(COLOR.fondo).setVerticalAlignment('middle').setHorizontalAlignment('center');
    // Importe €
    sh.getRange(fila + 1, c1, 1, c2 - c1 + 1).merge()
      .setFormula(fEuro).setNumberFormat('#,##0 €')
      .setFontSize(18).setFontWeight('bold').setFontColor(COLOR.tinta)
      .setBackground(COLOR.panel).setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(true, true, false, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
    // %
    sh.getRange(fila + 2, c1, 1, c2 - c1 + 1).merge()
      .setFormula(`="(" & TEXT(${fPct.replace('=', '')};"0%") & " del gasto del mes)"`)
      .setFontSize(9).setFontColor(COLOR.tenue)
      .setBackground(COLOR.panel).setVerticalAlignment('middle').setHorizontalAlignment('center')
      .setBorder(false, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  });

  // Barra apilada visual: 3 celdas adyacentes, una por tipo, cada una con su color.
  // Cada celda repinta proporcionalmente al % de su tipo (B32/B33/B34).
  // azul = fijos · naranja = variables · rojo = individuales
  sh.getRange(fila + 3, 2, 1, 4).merge()
    .setFormula(`=IFERROR(REPT("█";ROUND(${C}!B32*40;0));"")`)
    .setFontSize(11).setFontColor(COLOR.acento)
    .setBackground(COLOR.fondo).setHorizontalAlignment('right').setVerticalAlignment('middle');
  sh.getRange(fila + 3, 6, 1, 3).merge()
    .setFormula(`=IFERROR(REPT("█";ROUND(${C}!B33*40;0));"")`)
    .setFontSize(11).setFontColor(COLOR.naranja)
    .setBackground(COLOR.fondo).setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.getRange(fila + 3, 9, 1, 3).merge()
    .setFormula(`=IFERROR(REPT("█";ROUND(${C}!B34*40;0));"")`)
    .setFontSize(11).setFontColor(COLOR.rojo)
    .setBackground(COLOR.fondo).setHorizontalAlignment('left').setVerticalAlignment('middle');
}

/* ---------- BLOQUE D3: Gastos individuales por persona ---------- */
function bloqueIndividuales(sh, fila, C) {
  // Dos sub-tablas lado a lado: B:F = FM, G:K = Lucía
  const tablas = [
    [2, 6,  'FM',    'R', 'S'],
    [7, 11, 'Lucía', 'T', 'U'],
  ];
  tablas.forEach(([c1, c2, titulo, colCat, colImp]) => {
    // Encabezado
    sh.setRowHeight(fila, 26);
    sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setValue('Top 3 categorías — ' + titulo)
      .setFontSize(10).setFontWeight('bold').setFontColor(COLOR.cabTxt)
      .setBackground(COLOR.acento).setVerticalAlignment('middle').setHorizontalAlignment('left');

    // 3 filas top
    for (let i = 0; i < 3; i++) {
      const r = fila + 1 + i;
      const cr = 2 + i; // _Calc fila 2..4
      sh.setRowHeight(r, 24);
      const fondo = i % 2 === 0 ? COLOR.panel : COLOR.cebra;
      sh.getRange(r, c1, 1, c2 - c1 + 1).setBackground(fondo)
        .setBorder(false, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

      // Categoría (3 cols)
      sh.getRange(r, c1, 1, 3).merge()
        .setFormula(`=IFERROR(${C}!${colCat}${cr};"—")`)
        .setFontSize(11).setFontColor(COLOR.tinta)
        .setVerticalAlignment('middle').setHorizontalAlignment('left');
      // Importe (2 cols)
      sh.getRange(r, c1 + 3, 1, 2).merge()
        .setFormula(`=IFERROR(${C}!${colImp}${cr};"")`)
        .setNumberFormat('#,##0 €')
        .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.texto)
        .setVerticalAlignment('middle').setHorizontalAlignment('right');
    }
  });
}

/* ---------- BLOQUE F: Objetivos top 3 ---------- */
function bloqueObjetivos(sh, fila) {
  const OBJ = `'${HOJAS.OBJETIVOS}'`;
  // Cabecera
  sh.setRowHeight(fila, 24);
  const cab = [
    [2, 4,  'Concepto'],
    [5, 8,  'Progreso'],
    [9, 10, 'Aportado'],
    [11,11, 'Meta'],
  ];
  cab.forEach(([c1, c2, txt], i) => {
    sh.getRange(fila, c1, 1, c2 - c1 + 1).merge()
      .setValue(txt)
      .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
      .setBackground(COLOR.fondo).setVerticalAlignment('middle')
      .setHorizontalAlignment(i === 0 ? 'left' : (i === 1 ? 'left' : 'right'))
      .setBorder(null, null, true, null, null, null, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  });

  // 3 filas (Objetivos!A2..A4)
  for (let i = 0; i < 3; i++) {
    const r = fila + 1 + i;
    const oRow = 2 + i;
    sh.setRowHeight(r, 26);
    const fondo = i % 2 === 0 ? COLOR.panel : COLOR.cebra;
    sh.getRange(r, 2, 1, 10).setBackground(fondo);

    sh.getRange(r, 2, 1, 3).merge()
      .setFormula(`=${OBJ}!A${oRow}`)
      .setFontSize(11).setFontColor(COLOR.tinta).setVerticalAlignment('middle');

    sh.getRange(r, 5, 1, 4).merge()
      .setFormula(`=IFERROR(${OBJ}!G${oRow} & "   " & TEXT(${OBJ}!D${oRow};"0%");"")`)
      .setFontSize(11).setFontColor(COLOR.acento).setVerticalAlignment('middle')
      .setHorizontalAlignment('left');

    sh.getRange(r, 9, 1, 2).merge()
      .setFormula(`=IFERROR(${OBJ}!C${oRow};"")`)
      .setNumberFormat('#,##0 €').setFontSize(11).setFontColor(COLOR.texto)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');

    sh.getRange(r, 11)
      .setFormula(`=IFERROR(${OBJ}!B${oRow};"")`)
      .setNumberFormat('#,##0 €').setFontSize(11).setFontColor(COLOR.tenue)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
  }
  sh.getRange(fila + 3, 2, 1, 10)
    .setBorder(null, null, true, null, null, null, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
}

/* ---------- BLOQUE G: Gráfico evolución 12 meses ---------- */
function insertarGraficoEvolucion(sh, ss) {
  const calc = ss.getSheetByName(HOJAS.CALC);

  // Línea de gasto mensual (eje izq €) + tasa de ahorro derivada (sin segundo eje para mantenerlo simple).
  // Series: Gasto (O) y Ahorro (L) sobre los 12 meses.
  const grafico = sh.newChart()
    .setChartType(Charts.ChartType.COMBO)
    .addRange(calc.getRange('K1:K13'))   // X: meses
    .addRange(calc.getRange('N1:N13'))   // Y1: Ingresos
    .addRange(calc.getRange('O1:O13'))   // Y1: Gastos
    .addRange(calc.getRange('L1:L13'))   // Y1: Ahorro
    .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_COLUMNS)
    .setNumHeaders(1)
    .setOption('title', '')
    .setOption('legend', { position: 'top', alignment: 'center', textStyle: { color: COLOR.texto, fontSize: 10 } })
    .setOption('series', {
      0: { type: 'line', color: '#16A34A', lineWidth: 2, pointSize: 4 },   // Ingresos verde
      1: { type: 'line', color: '#DC2626', lineWidth: 2, pointSize: 4 },   // Gastos rojo
      2: { type: 'bars', color: '#2563EB' },                                // Ahorro barra azul
    })
    .setOption('hAxis', { textStyle: { color: COLOR.tenue, fontSize: 9 } })
    .setOption('vAxis', { textStyle: { color: COLOR.tenue, fontSize: 9 }, format: '#,##0 €', gridlines: { color: COLOR.borde } })
    .setOption('backgroundColor', COLOR.panel)
    .setOption('chartArea', { left: 60, top: 40, width: '90%', height: '75%' })
    .setOption('width', 920)
    .setOption('height', 300)
    .setPosition(54, 2, 0, 0)
    .build();
  sh.insertChart(grafico);
}

/* ---------- Etiqueta de sección ---------- */
function etiquetaSeccion(sh, fila, texto) {
  sh.setRowHeight(fila, 22);
  sh.getRange(fila, 2, 1, 10).merge()
    .setValue(texto)
    .setFontSize(10).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setVerticalAlignment('bottom').setBackground(COLOR.fondo);
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

/** Delta de gasto: subir es malo (rojo), bajar es bueno (verde). */
function aplicarDeltaGasto(sh, rango) {
  const reglas = sh.getConditionalFormatRules();
  reglas.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0.1).setFontColor(COLOR.rojo).setBold(true)
      .setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.01, 0.1).setFontColor(COLOR.naranja)
      .setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThanOrEqualTo(-0.05).setFontColor(COLOR.verde)
      .setRanges([rango]).build()
  );
  sh.setConditionalFormatRules(reglas);
}

/** Tasa de ahorro: >=20% verde, 10-20% naranja, <10% rojo. */
function aplicarTasaAhorro(sh, rango) {
  const reglas = sh.getConditionalFormatRules();
  reglas.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(0.2).setFontColor(COLOR.verde)
      .setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(0.1, 0.1999).setFontColor(COLOR.naranja)
      .setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0.1).setFontColor(COLOR.rojo)
      .setRanges([rango]).build()
  );
  sh.setConditionalFormatRules(reglas);
}

/** Reserva de emergencia: >=3 meses verde, 1-3 meses naranja, <1 mes rojo. */
function aplicarReserva(sh, rango) {
  const reglas = sh.getConditionalFormatRules();
  reglas.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThanOrEqualTo(3).setFontColor(COLOR.verde)
      .setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberBetween(1, 2.999).setFontColor(COLOR.naranja)
      .setRanges([rango]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(1).setFontColor(COLOR.rojo)
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
