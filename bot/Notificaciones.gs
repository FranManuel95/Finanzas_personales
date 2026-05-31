/**
 * Notificaciones.gs
 * ------------------------------------------------------------------
 * Triggers programados y notificaciones automáticas a los chats
 * autorizados. Usa enviar() y _config() de Bot.gs / Config.gs.
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
 * Resumen del mes anterior + avance del Mes activo al mes en curso.
 */
function notificacionResumenMensual() {
  const ss = _ss();
  const hoy = new Date();
  // El mes anterior (porque se ejecuta el día 1 del nuevo mes).
  const anioAnt = hoy.getMonth() === 0 ? hoy.getFullYear() - 1 : hoy.getFullYear();
  const mesAnt = hoy.getMonth() === 0 ? 12 : hoy.getMonth();
  const mesAntISO = `${anioAnt}-${String(mesAnt).padStart(2, '0')}`;

  const inicio = new Date(anioAnt, mesAnt - 1, 1);
  const fin = new Date(anioAnt, mesAnt, 1);

  const sumar = (nombre, colImp) => {
    const sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() < 2) return 0;
    const filas = sh.getRange(2, 1, sh.getLastRow() - 1, colImp).getValues();
    let total = 0;
    filas.forEach(f => {
      const fecha = f[0];
      if (fecha instanceof Date && fecha >= inicio && fecha < fin) total += Number(f[colImp - 1]) || 0;
    });
    return total;
  };

  const ingresos = sumar(HOJAS.INGRESOS, 4);
  const gcf = sumar(HOJAS.GC_FIJOS, 4);
  const gcv = sumar(HOJAS.GC_VARIABLES, 4);
  const gfm = sumar(HOJAS.G_FM, 4);
  const glu = sumar(HOJAS.G_LUCIA, 4);
  const gastos = gcf + gcv + gfm + glu;
  const ahorro = ingresos - gastos;

  const objetivo = Number(_config('Objetivo ahorro mensual conjunto (€)')) || 0;
  const cumplido = objetivo > 0 ? (ahorro / objetivo) : 0;

  const lineas = [
    `📅 <b>Cierre de ${mesAntISO}</b>`,
    '',
    `Ingresos: ${formatoEur(ingresos)}`,
    `Gastos: ${formatoEur(gastos)}`,
    `  Comp. fijos: ${formatoEur(gcf)}`,
    `  Comp. variables: ${formatoEur(gcv)}`,
    `  FM: ${formatoEur(gfm)}`,
    `  Lucía: ${formatoEur(glu)}`,
    '',
    `Ahorro: <b>${formatoEur(ahorro)}</b>`,
    `Objetivo: ${formatoEur(objetivo)} · cumplimiento ${(cumplido * 100).toFixed(0)}%`,
    '',
    '¡Empieza el nuevo mes! Usa /nuevo para registrar movimientos.',
  ];
  const texto = lineas.join('\n');

  // Avanzar el Mes activo al mes en curso.
  try {
    const sh = ss.getSheetByName(HOJAS.CONFIG);
    const datos = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    const idx = datos.findIndex(([k]) => k === 'Mes activo');
    if (idx !== -1) {
      const nuevo = `${hoy.getFullYear()}-${String(hoy.getMonth() + 1).padStart(2, '0')}`;
      sh.getRange(idx + 2, 2).setValue(nuevo);
    }
  } catch (err) {
    console.error('No pude actualizar Mes activo', err);
  }

  enviarATodos(texto);
}

/**
 * Recordatorio de revisión antes de fin de mes (día 25).
 */
function notificacionRevisionPreCierre() {
  const ss = _ss();
  const hoy = new Date();
  const ultimoDia = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).getDate();
  const dias = ultimoDia - hoy.getDate();

  const dash = ss.getSheetByName(HOJAS.DASHBOARD);
  let estado = '';
  if (dash) {
    const get = (c) => dash.getRange(c).getValue();
    estado = [
      `Ingresos: ${formatoEur(get('B7'))}`,
      `Gastos comp.: ${formatoEur(get('B17'))}`,
      `Gastos FM: ${formatoEur(get('B21'))}`,
      `Gastos Lucía: ${formatoEur(get('B22'))}`,
      `Ahorro real: ${formatoEur(get('B30'))}`,
      `Objetivo: ${formatoEur(get('B29'))}`,
    ].join('\n');
  }

  const lineas = [
    `⏰ <b>Faltan ${dias} días para fin de mes</b>`,
    '',
    'Buen momento para repasar gastos pendientes de registrar.',
    '',
    '<b>Estado actual</b>',
    estado || '(sin datos del dashboard)',
    '',
    '/resumen para ver el detalle.',
  ];
  enviarATodos(lineas.join('\n'));
}

/**
 * Aviso opcional de gasto inusual. Se llama desde guardarMovimiento().
 * Compara contra la media de la categoría en los últimos 6 meses.
 */
function notificacionGastoAtipico(g) {
  if (!g || !g.categoria || !g.importe) return;
  const hojaPorTipo = {
    gc_fijo: HOJAS.GC_FIJOS,
    gc_variable: HOJAS.GC_VARIABLES,
    gasto_fm: HOJAS.G_FM,
    gasto_lucia: HOJAS.G_LUCIA,
  };
  const hoja = hojaPorTipo[g.tipo];
  if (!hoja) return;
  const sh = _ss().getSheetByName(hoja);
  if (!sh || sh.getLastRow() < 2) return;

  const hoy = new Date();
  const hace6m = new Date(hoy.getFullYear(), hoy.getMonth() - 6, 1);
  const filas = sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues();
  const importes = [];
  filas.forEach(f => {
    if (!(f[0] instanceof Date) || f[0] < hace6m) return;
    if (f[2] !== g.categoria) return;
    const v = Number(f[3]);
    if (v > 0) importes.push(v);
  });
  if (importes.length < 3) return; // sin muestra suficiente
  const media = importes.reduce((a, b) => a + b, 0) / importes.length;
  // Umbral: 2x la media y > 30€ para evitar ruido en cifras pequeñas.
  if (g.importe < 30 || g.importe < media * 2) return;

  const texto = [
    `⚠️ <b>Gasto inusual</b>`,
    `${etiquetaTipo(g.tipo)} · ${g.categoria}`,
    `Importe: <b>${formatoEur(g.importe)}</b>${g.concepto ? ' — ' + g.concepto : ''}`,
    `Tu media en esta categoría (últimos 6 meses): ${formatoEur(media)}`,
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
