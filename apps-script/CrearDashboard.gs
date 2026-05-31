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

  /* ===== Importe COMPARTIDO por categoría y Top 3 compartidas (Z:AC) =====
   * Para el nuevo Panel: las top categorías del bote (fijo+variable, sin individuales).
   */
  sh.getRange('Z1:AA1').setValues([['Cat', 'Importe compartido']]);
  // Z: copia de las categorías de E2:E40 para tener rango contiguo con AA en QUERY.
  sh.getRange('Z2').setFormula(`=ARRAYFORMULA(E2:E40)`);
  // AA: suma SOLO de Compartido fijo + Compartido variable (sin Individual).
  sh.getRange('AA2').setFormula(
    `=ARRAYFORMULA(IF(E2:E40="";"";` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};` +
    `${MOV}!D:D;E2:E40;${MOV}!B:B;"Compartido fijo");0)+` +
    `IFERROR(SUMIFS(${MOV}!F:F;${MOV}!A:A;">="&${inicioMes};${MOV}!A:A;"<"&${finMes};` +
    `${MOV}!D:D;E2:E40;${MOV}!B:B;"Compartido variable");0)))`
  );
  sh.getRange('AA2:AA40').setNumberFormat('#,##0.00 €');

  // AB:AC = Top 3 categorías compartidas (cat, importe), ordenadas desc.
  sh.getRange('AB1:AC1').setValues([['Top cat compartida', 'Importe']]);
  sh.getRange('AB2').setFormula(
    `=IFERROR(QUERY(Z2:AA40;"select Col1, Col2 where Col2 > 0 order by Col2 desc limit 3";0);"")`
  );
  sh.getRange('AC2:AC4').setNumberFormat('#,##0.00 €');

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

  /* ===== BLOQUE B — 🤝 COMPARTIDO (bote común), ancho completo (filas 11-25) ===== */
  etiquetaSeccion(sh, 11, '🤝  COMPARTIDO  (bote común)');
  bloqueCompartido(sh, 12, C);

  /* ===== BLOQUE C — 👤 FM (B..F) y 👤 LUCÍA (G..K) lado a lado (filas 27-41) ===== */
  etiquetaSeccionRango(sh, 27, 2, 6,  '👤  FM');
  etiquetaSeccionRango(sh, 27, 7, 11, '👤  LUCÍA');
  bloquePersona(sh, 28, 2, 6,  C, 'FM',    `=${C}!B1`, `=${C}!B4`, `=${C}!B10`, `=${C}!B13`, 'R', 'S');
  bloquePersona(sh, 28, 7, 11, C, 'Lucía', `=${C}!B2`, `=${C}!B5`, `=${C}!B11`, `=${C}!B14`, 'T', 'U');

  /* ===== BLOQUE F — Objetivos top 3 (filas 43-47) ===== */
  etiquetaSeccion(sh, 43, '🎯  OBJETIVOS A LARGO PLAZO');
  bloqueObjetivos(sh, 44);

  /* ===== BLOQUE G — Evolución 12 meses (gráfico) ===== */
  etiquetaSeccion(sh, 49, '📈  EVOLUCIÓN 12 MESES');
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

/* ---------- Etiqueta de sección de ANCHO PARCIAL (para títulos lado a lado) ---------- */
function etiquetaSeccionRango(sh, fila, colIni, colFin, texto) {
  sh.setRowHeight(fila, 28);
  sh.getRange(fila, colIni, 1, colFin - colIni + 1).merge()
    .setValue(texto)
    .setFontSize(12).setFontWeight('bold').setFontColor(COLOR.cabTxt)
    .setBackground(COLOR.cab).setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
}

/* ---------- BLOQUE COMPARTIDO (ancho completo B:K) ----------
 * Layout vertical (14 filas a partir de `fila`):
 *   fila   0  : sub-cabeceras "APORTACIONES AL BOTE" | "GASTOS DEL BOTE"
 *   fila   1-3: 3 filas dato (FM/Lucía/Total) | (Fijos/Variables/Total)
 *   fila   5  : SOBRANTE DEL BOTE — número grande ocupando todo
 *   fila   7  : etiqueta "Top categorías del bote"
 *   fila   8-10: 3 filas top categorías compartidas (cat / barra / importe)
 */
function bloqueCompartido(sh, fila, C) {
  // === fila 0 — Sub-cabeceras lado a lado (B:E = aportes, G:K = gastos; F separador) ===
  sh.setRowHeight(fila, 24);
  sh.getRange(fila, 2, 1, 4).merge()
    .setValue('APORTACIONES AL BOTE')
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.cabTxt)
    .setBackground(COLOR.acento).setVerticalAlignment('middle')
    .setHorizontalAlignment('center');
  sh.getRange(fila, 7, 1, 5).merge()
    .setValue('GASTOS DEL BOTE')
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.cabTxt)
    .setBackground(COLOR.acento).setVerticalAlignment('middle')
    .setHorizontalAlignment('center');

  // === filas 1-3 — Datos lado a lado ===
  // Izquierda: FM | Lucía | TOTAL APORTADO     Derecha: Fijos | Variables | TOTAL DEL BOTE
  const filasDatos = [
    ['FM',    `=${C}!B4`, false, 'Fijos',     `=${C}!B7`, false],
    ['Lucía', `=${C}!B5`, false, 'Variables', `=${C}!B8`, false],
    ['TOTAL APORTADO', `=${C}!B6`, true,  'TOTAL DEL BOTE', `=${C}!B9`, true],
  ];
  filasDatos.forEach(([lblL, fL, totL, lblR, fR, totR], i) => {
    const r = fila + 1 + i;
    sh.setRowHeight(r, 26);
    const fondo = totL ? COLOR.acentoSuave : (i % 2 === 0 ? COLOR.panel : COLOR.cebra);
    sh.getRange(r, 2, 1, 10).setBackground(fondo);

    // Izquierda: etiqueta B:C | valor D:E
    sh.getRange(r, 2, 1, 2).merge().setValue(lblL)
      .setFontSize(11).setFontWeight(totL ? 'bold' : 'normal')
      .setFontColor(COLOR.tinta).setVerticalAlignment('middle')
      .setHorizontalAlignment('left');
    sh.getRange(r, 4, 1, 2).merge().setFormula(fL).setNumberFormat('#,##0.00 €')
      .setFontSize(12).setFontWeight(totL ? 'bold' : 'normal')
      .setFontColor(COLOR.tinta).setVerticalAlignment('middle')
      .setHorizontalAlignment('right');

    // Derecha: etiqueta G:H | valor I:K
    sh.getRange(r, 7, 1, 2).merge().setValue(lblR)
      .setFontSize(11).setFontWeight(totR ? 'bold' : 'normal')
      .setFontColor(COLOR.tinta).setVerticalAlignment('middle')
      .setHorizontalAlignment('left');
    sh.getRange(r, 9, 1, 3).merge().setFormula(fR).setNumberFormat('#,##0.00 €')
      .setFontSize(12).setFontWeight(totR ? 'bold' : 'normal')
      .setFontColor(COLOR.tinta).setVerticalAlignment('middle')
      .setHorizontalAlignment('right');

    // Borde inferior bajo la fila TOTAL
    if (totL) {
      sh.getRange(r, 2, 1, 10)
        .setBorder(true, null, true, null, null, null, COLOR.acento, SpreadsheetApp.BorderStyle.SOLID);
    }
  });

  // === fila 5 — SOBRANTE DEL BOTE: número grande ocupando todo el bloque ===
  const frSob = fila + 4;
  sh.setRowHeight(frSob, 16);
  sh.getRange(frSob, 2, 1, 10).setBackground(COLOR.fondo);

  const frSobVal = fila + 5;
  sh.setRowHeight(frSobVal, 44);
  sh.getRange(frSobVal, 2, 1, 4).merge()
    .setValue('SOBRANTE DEL BOTE')
    .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setBackground(COLOR.panel).setVerticalAlignment('middle').setHorizontalAlignment('right')
    .setBorder(true, true, true, false, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  const sobCell = sh.getRange(frSobVal, 6, 1, 3).merge()
    .setFormula(`=${C}!B15`).setNumberFormat('#,##0.00 €')
    .setFontSize(22).setFontWeight('bold')
    .setBackground(COLOR.panel).setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBorder(true, false, true, false, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);
  aplicarPositivoNegativo(sh, sh.getRange(frSobVal, 6, 1, 3));
  // Sub-texto a la derecha: "X% del aportado" o aviso
  sh.getRange(frSobVal, 9, 1, 3).merge()
    .setFormula(
      `=IFERROR(IF(${C}!B6=0;"sin aportaciones aún";` +
      `IF(${C}!B15<0;"⚠ déficit · gastasteis " & TEXT(ABS(${C}!B15/${C}!B6);"0%") & " más";` +
      `TEXT(${C}!B15/${C}!B6;"0%") & " del aportado");"")`)
    .setFontSize(10).setFontColor(COLOR.tenue)
    .setBackground(COLOR.panel).setVerticalAlignment('middle').setHorizontalAlignment('left')
    .setBorder(true, false, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

  // === fila 7 — etiqueta "Top categorías compartidas" ===
  const frTopTitle = fila + 6;
  sh.setRowHeight(frTopTitle, 24);
  sh.getRange(frTopTitle, 2, 1, 10).merge()
    .setValue('TOP CATEGORÍAS DEL BOTE')
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setBackground(COLOR.fondo).setVerticalAlignment('bottom').setHorizontalAlignment('left');

  // === filas 8-10 — Top 3 compartidas (cat | barra | importe) ===
  for (let i = 0; i < 3; i++) {
    const r = fila + 7 + i;
    const cr = 2 + i; // _Calc!AB2..AB4
    sh.setRowHeight(r, 24);
    const fondo = i % 2 === 0 ? COLOR.panel : COLOR.cebra;
    sh.getRange(r, 2, 1, 10).setBackground(fondo)
      .setBorder(false, false, true, false, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

    // Categoría (B:D)
    sh.getRange(r, 2, 1, 3).merge()
      .setFormula(`=IFERROR(${C}!AB${cr};"—")`)
      .setFontSize(11).setFontColor(COLOR.tinta)
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    // Barra (E:I), proporcional al top 1 (_Calc!AC$2)
    sh.getRange(r, 5, 1, 5).merge()
      .setFormula(
        `=IFERROR(IF(${C}!AC${cr}="";"";REPT("█";MAX(1;ROUND(${C}!AC${cr}/${C}!AC$2*30;0))));"")`
      )
      .setFontSize(10).setFontColor(COLOR.acento)
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    // Importe (J:K)
    sh.getRange(r, 10, 1, 2).merge()
      .setFormula(`=IFERROR(${C}!AC${cr};"")`)
      .setNumberFormat('#,##0 €')
      .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.tinta)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
  }
}

/* ---------- BLOQUE PERSONA (medio ancho: B:F o G:K) ----------
 * Layout vertical (10 filas a partir de `fila`):
 *   fila   0-3 : 4 líneas resumen (Ingresos / − Bote / − Gastos / EN BOLSILLO)
 *   fila   5   : etiqueta "Sus gastos individuales (top 3)"
 *   fila   6-8 : top 3 categorías individuales (cat | importe)
 *
 * colCat/colImp: letras de columna en _Calc donde viven las top 3 categorías individuales
 * de esta persona (R/S para FM, T/U para Lucía).
 */
function bloquePersona(sh, fila, colIni, colFin, C, nombre, fIng, fApo, fGas, fSaldo, colCat, colImp) {
  const ancho = colFin - colIni + 1;
  const colLbl = colIni;
  const colVal = colIni + Math.floor(ancho / 2);
  const nLbl = colVal - colIni;
  const nVal = colFin - colVal + 1;

  // 4 líneas resumen
  const lineas = [
    ['Ingresos',         fIng,   false],
    ['− Aportación bote',fApo,   false],
    ['− Sus gastos',     fGas,   false],
    ['EN BOLSILLO',      fSaldo, true],
  ];
  lineas.forEach(([lbl, f, total], i) => {
    const r = fila + i;
    sh.setRowHeight(r, 28);
    const fondo = total ? COLOR.acentoSuave : (i % 2 === 0 ? COLOR.panel : COLOR.cebra);
    sh.getRange(r, colIni, 1, ancho).setBackground(fondo)
      .setBorder(false, true, false, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

    // Etiqueta
    sh.getRange(r, colLbl, 1, nLbl).merge().setValue(lbl)
      .setFontSize(11).setFontWeight(total ? 'bold' : 'normal')
      .setFontColor(total ? COLOR.tinta : COLOR.texto)
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    // Valor
    const valCell = sh.getRange(r, colVal, 1, nVal).merge().setFormula(f)
      .setNumberFormat('#,##0.00 €')
      .setFontSize(total ? 14 : 12).setFontWeight(total ? 'bold' : 'normal')
      .setFontColor(total ? COLOR.tinta : COLOR.texto)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
    if (total) {
      sh.getRange(r, colIni, 1, ancho)
        .setBorder(true, true, true, true, false, false, COLOR.acento, SpreadsheetApp.BorderStyle.SOLID);
      aplicarPositivoNegativo(sh, sh.getRange(r, colVal, 1, nVal));
    }
  });

  // Espacio
  sh.setRowHeight(fila + 4, 12);

  // Etiqueta "Sus gastos individuales (top 3)"
  const frTit = fila + 5;
  sh.setRowHeight(frTit, 22);
  sh.getRange(frTit, colIni, 1, ancho).merge()
    .setValue('SUS GASTOS INDIVIDUALES (TOP 3)')
    .setFontSize(9).setFontWeight('bold').setFontColor(COLOR.tenue)
    .setBackground(COLOR.fondo).setVerticalAlignment('bottom').setHorizontalAlignment('left');

  // 3 filas top categorías individuales
  for (let i = 0; i < 3; i++) {
    const r = fila + 6 + i;
    const cr = 2 + i; // _Calc filas 2..4
    sh.setRowHeight(r, 24);
    const fondo = i % 2 === 0 ? COLOR.panel : COLOR.cebra;
    sh.getRange(r, colIni, 1, ancho).setBackground(fondo)
      .setBorder(false, true, true, true, false, false, COLOR.borde, SpreadsheetApp.BorderStyle.SOLID);

    // Categoría
    sh.getRange(r, colLbl, 1, nLbl).merge()
      .setFormula(`=IFERROR(${C}!${colCat}${cr};"—")`)
      .setFontSize(11).setFontColor(COLOR.tinta)
      .setVerticalAlignment('middle').setHorizontalAlignment('left');
    // Importe
    sh.getRange(r, colVal, 1, nVal).merge()
      .setFormula(`=IFERROR(${C}!${colImp}${cr};"")`)
      .setNumberFormat('#,##0 €')
      .setFontSize(11).setFontWeight('bold').setFontColor(COLOR.texto)
      .setVerticalAlignment('middle').setHorizontalAlignment('right');
  }
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
    .setPosition(50, 2, 0, 0)
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
