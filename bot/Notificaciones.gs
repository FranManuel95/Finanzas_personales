/**
 * Notificaciones.gs
 * ------------------------------------------------------------------
 * Triggers programados y notificaciones automáticas a los chats
 * autorizados. Usa enviar(), _config(), formatoEur(), leerMovimientos()
 * de Bot.gs / Config.gs y el esquema consolidado (Movimientos, _Calc, Ajustes).
 *
 * Acciones a ejecutar desde el editor:
 *  - configurarTriggers(): instala los triggers periódicos.
 *  - eliminarTriggers(): los borra.
 * ------------------------------------------------------------------
 */

/* ============== TRIGGERS ============== */

function configurarTriggers() {
  eliminarTriggers();

  // Día 1 a las 9:00 → resumen del mes anterior
  ScriptApp.newTrigger('notificacionResumenMensual')
    .timeBased()
    .onMonthDay(1)
    .atHour(9)
    .create();

  // Día 25 a las 20:00 → recordatorio pre-cierre
  ScriptApp.newTrigger('notificacionRevisionPreCierre')
    .timeBased()
    .onMonthDay(25)
    .atHour(20)
    .create();

  console.log('Triggers instalados: notificacionResumenMensual (día 1, 9h) y notificacionRevisionPreCierre (día 25, 20h).');
}

function eliminarTriggers() {
  const objetivo = ['notificacionResumenMensual', 'notificacionRevisionPreCierre'];
  ScriptApp.getProjectTriggers().forEach(t => {
    if (objetivo.indexOf(t.getHandlerFunction()) !== -1) ScriptApp.deleteTrigger(t);
  });
}

/* ============== NOTIFICACIONES ============== */

/**
 * Resumen del mes anterior (agregado sobre Movimientos) y avance del
 * Mes activo en Ajustes al mes en curso.
 * Se ejecuta el día 1: el "mes anterior" es el que acaba de cerrar.
 */
function notificacionResumenMensual() {
  const ss = _ss();
  const hoy = new Date();
  const anioAnt = hoy.getMonth() === 0 ? hoy.getFullYear() - 1 : hoy.getFullYear();
  const mesAnt = hoy.getMonth() === 0 ? 12 : hoy.getMonth();
  const mesAntISO = `${anioAnt}-${String(mesAnt).padStart(2, '0')}`;

  const inicio = new Date(anioAnt, mesAnt - 1, 1);
  const fin = new Date(anioAnt, mesAnt, 1);

  let ingresos = 0, gastos = 0;
  const porTipo = { 'Compartido fijo': 0, 'Compartido variable': 0, 'Individual': 0 };
  leerMovimientos().forEach(m => {
    if (!(m.fecha instanceof Date) || m.fecha < inicio || m.fecha >= fin) return;
    if (m.tipo === 'Ingreso') ingresos += m.importe;
    else if (porTipo.hasOwnProperty(m.tipo)) { porTipo[m.tipo] += m.importe; gastos += m.importe; }
  });
  const ahorro = ingresos - gastos;

  const objetivo = Number(_config('Objetivo ahorro mensual conjunto (€)')) || 0;
  const cumplido = objetivo > 0 ? (ahorro / objetivo) : 0;

  const lineas = [
    `📅 <b>Cierre de ${mesAntISO}</b>`,
    '',
    `Ingresos: ${formatoEur(ingresos)}`,
    `Gastos: ${formatoEur(gastos)}`,
    `  Comp. fijos: ${formatoEur(porTipo['Compartido fijo'])}`,
    `  Comp. variables: ${formatoEur(porTipo['Compartido variable'])}`,
    `  Individuales: ${formatoEur(porTipo['Individual'])}`,
    '',
    `Ahorro: <b>${formatoEur(ahorro)}</b>`,
    `Objetivo: ${formatoEur(objetivo)} · cumplimiento ${(cumplido * 100).toFixed(0)}%`,
    '',
    '¡Empieza el nuevo mes! Usa /nuevo para registrar movimientos.',
  ];

  // Avanzar el "Mes activo" en Ajustes al mes en curso.
  try {
    const sh = ss.getSheetByName(HOJAS.AJUSTES);
    const datos = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    const idx = datos.findIndex(([k]) => k === 'Mes activo');
    if (idx !== -1) {
      const nuevo = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}`;
      sh.getRange(idx + 2, 2).setValue(nuevo);
    }
  } catch (err) {
    console.error('No pude actualizar Mes activo', err);
  }

  enviarATodos(lineas.join('\n'));
}

/**
 * Recordatorio de revisión antes de fin de mes (día 25).
 * Estado actual leído de _Calc (mes activo).
 */
function notificacionRevisionPreCierre() {
  const ss = _ss();
  const hoy = new Date();
  const ultimoDia = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).getDate();
  const dias = ultimoDia - hoy.getDate();

  const calc = ss.getSheetByName(HOJAS.CALC);
  let estado = '';
  if (calc) {
    const v = calc.getRange('B1:B20').getValues().map(r => Number(r[0]) || 0);
    estado = [
      `Ingresos: ${formatoEur(v[2])}`,       // B3 ingresos total
      `Gastos comp.: ${formatoEur(v[8])}`,    // B9 GC total
      `Gastos FM: ${formatoEur(v[9])}`,       // B10
      `Gastos Lucía: ${formatoEur(v[10])}`,   // B11
      `Ahorro real: ${formatoEur(v[16])}`,    // B17
      `Objetivo: ${formatoEur(v[15])}`,       // B16
    ].join('\n');
  }

  const lineas = [
    `⏰ <b>Faltan ${dias} días para fin de mes</b>`,
    '',
    'Buen momento para repasar gastos pendientes de registrar.',
    '',
    '<b>Estado actual</b>',
    estado || '(sin datos de _Calc)',
    '',
    '/resumen para ver el detalle.',
  ];
  enviarATodos(lineas.join('\n'));
}

/**
 * Aviso opcional de gasto inusual. Se llama desde guardarMovimiento() /
 * ejecutarComandoRapido(). Compara contra la media de la categoría en los
 * últimos 6 meses, calculada sobre Movimientos.
 * Recibe { tipoTxt, categoria, importe, concepto }.
 */
function notificacionGastoAtipico(g) {
  if (!g || !g.categoria || !g.importe) return;
  const TIPOS = ['Compartido fijo', 'Compartido variable', 'Individual'];
  if (TIPOS.indexOf(g.tipoTxt) === -1) return;

  const hoy = new Date();
  const hace6m = new Date(hoy.getFullYear(), hoy.getMonth() - 6, 1);
  const importes = [];
  leerMovimientos().forEach(m => {
    if (TIPOS.indexOf(m.tipo) === -1) return;
    if (!(m.fecha instanceof Date) || m.fecha < hace6m) return;
    if (m.categoria !== g.categoria) return;
    if (m.importe > 0) importes.push(m.importe);
  });
  if (importes.length < 3) return; // sin muestra suficiente (incluye el recién añadido)
  const media = importes.reduce((a, b) => a + b, 0) / importes.length;
  // Umbral: 2x la media y > 30€ para evitar ruido en cifras pequeñas.
  if (g.importe < 30 || g.importe < media * 2) return;

  const texto = [
    `⚠️ <b>Gasto inusual</b>`,
    `${g.tipoTxt} · ${g.categoria}`,
    `Importe: <b>${formatoEur(g.importe)}</b>${g.concepto ? ' — ' + g.concepto : ''}`,
    `Media en esta categoría (últimos 6 meses): ${formatoEur(media)}`,
  ].join('\n');
  enviarATodos(texto);
}

/* ============== HELPERS ============== */

function enviarATodos(texto, teclado) {
  const lista = getChatsAutorizados();
  if (lista.length === 0) {
    console.warn('No hay chats autorizados; no se envía notificación.');
    return;
  }
  lista.forEach(chatId => {
    try {
      enviar(chatId, texto, teclado);
    } catch (err) {
      console.error('enviarATodos error chat=' + chatId, err);
    }
  });
}
