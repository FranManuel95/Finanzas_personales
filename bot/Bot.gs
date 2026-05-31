/**
 * Bot.gs
 * ------------------------------------------------------------------
 * Webhook de Telegram que conduce un flujo guiado por botones inline
 * para registrar ingresos y gastos en el Sheet. SIN LLM: todo el
 * "entendimiento" se hace por estado finito guardado en
 * PropertiesService (clave = chat_id).
 *
 * Endpoints:
 *  - doPost(e): lo llama Telegram (webhook).
 *  - doGet(e): respuesta amistosa al abrir la URL en el navegador.
 *
 * Acciones utilitarias (ejecutar desde el editor):
 *  - registrarWebhook(): registra esta Web App como webhook de Telegram.
 *  - eliminarWebhook(): la deshace.
 *  - obtenerMiChatId(): te dice qué chat IDs te han escrito (logs).
 * ------------------------------------------------------------------
 */

const API = 'https://api.telegram.org/bot';

/* ============== ENTRADA HTTP ============== */

function doGet() {
  return ContentService.createTextOutput('Bot de finanzas activo.');
}

function doPost(e) {
  try {
    const update = JSON.parse(e.postData.contents);
    if (update.message) manejarMensaje(update.message);
    else if (update.callback_query) manejarCallback(update.callback_query);
  } catch (err) {
    console.error('doPost error', err, e && e.postData && e.postData.contents);
  }
  return ContentService.createTextOutput('ok');
}

/* ============== MANEJO DE EVENTOS ============== */

function manejarMensaje(msg) {
  const chatId = msg.chat.id;
  if (!chatAutorizado(chatId)) {
    enviar(chatId, `Este chat no está autorizado. Pide a FM que añada tu chat ID (${chatId}) en la pestaña Config.`);
    return;
  }

  const texto = (msg.text || '').trim();

  if (texto === '/start' || texto === '/nuevo' || texto === '/menu') {
    limpiarEstado(chatId);
    mostrarMenuPrincipal(chatId);
    return;
  }
  if (texto === '/cancelar') {
    limpiarEstado(chatId);
    enviar(chatId, 'Operación cancelada. Usa /nuevo para empezar otra.');
    return;
  }
  if (texto === '/resumen') {
    enviarResumen(chatId);
    return;
  }
  if (texto === '/id') {
    enviar(chatId, `Tu chat ID es: <code>${chatId}</code>`);
    return;
  }

  const estado = leerEstado(chatId);
  if (!estado) {
    enviar(chatId, 'No te entendí. Pulsa /nuevo para registrar un movimiento o /resumen para ver el mes.');
    return;
  }

  // Estamos esperando texto en un flujo activo.
  switch (estado.paso) {
    case 'esperar_importe': return procesarImporte(chatId, texto, estado);
    case 'esperar_concepto': return procesarConcepto(chatId, texto, estado);
    default:
      enviar(chatId, 'Estoy esperando un botón. Si te has perdido, pulsa /cancelar.');
  }
}

function manejarCallback(cb) {
  const chatId = cb.message.chat.id;
  if (!chatAutorizado(chatId)) return;
  const data = cb.data;
  responderCallback(cb.id); // quita el "cargando" del botón

  const estado = leerEstado(chatId) || {};

  if (data.startsWith('tipo:')) {
    const tipo = data.slice(5);
    estado.tipo = tipo;
    guardarEstado(chatId, estado);
    if (tipo === 'ingreso' || tipo === 'aportacion') {
      estado.paso = 'elegir_persona';
      guardarEstado(chatId, estado);
      return mostrarPersonas(chatId, tipo);
    }
    estado.paso = 'elegir_categoria';
    guardarEstado(chatId, estado);
    if (tipo === 'gasto_fm') { estado.persona = 'FM'; guardarEstado(chatId, estado); return mostrarCategorias(chatId, 'Individual'); }
    if (tipo === 'gasto_lucia') { estado.persona = 'Lucía'; guardarEstado(chatId, estado); return mostrarCategorias(chatId, 'Individual'); }
    if (tipo === 'gc_fijo') return mostrarCategorias(chatId, 'Compartido fijo');
    if (tipo === 'gc_variable') return mostrarCategorias(chatId, 'Compartido variable');
  }

  if (data.startsWith('persona:')) {
    estado.persona = data.slice(8);
    estado.paso = 'esperar_importe';
    guardarEstado(chatId, estado);
    const verbo = estado.tipo === 'aportacion' ? 'aportación al bote' : 'ingreso';
    return enviar(chatId, `Persona: <b>${estado.persona}</b>\n\nEscribe el importe de la ${verbo} en € (ej: 600):`);
  }

  if (data.startsWith('cat:')) {
    estado.categoria = data.slice(4);
    estado.paso = 'esperar_importe';
    guardarEstado(chatId, estado);
    return enviar(chatId, `Categoría: <b>${estado.categoria}</b>\n\nEscribe el importe en € (ej: 42.50):`);
  }

  if (data.startsWith('pagador:')) {
    estado.pagador = data.slice(8);
    estado.paso = 'esperar_concepto';
    guardarEstado(chatId, estado);
    return enviar(chatId, `Pagado por: <b>${estado.pagador}</b>\n\nEscribe el concepto (ej: Cena con amigos):`);
  }

  if (data === 'confirmar:si') {
    return guardarMovimiento(chatId, estado);
  }
  if (data === 'confirmar:no') {
    limpiarEstado(chatId);
    return enviar(chatId, 'Descartado. /nuevo para empezar otro.');
  }

  if (data === 'menu') {
    limpiarEstado(chatId);
    return mostrarMenuPrincipal(chatId);
  }
}

/* ============== FLUJO ============== */

function mostrarMenuPrincipal(chatId) {
  const teclado = [
    [btn('💰 Ingreso', 'tipo:ingreso')],
    [btn('🏦 Aportación al bote', 'tipo:aportacion')],
    [btn('🏠 Gasto compartido fijo', 'tipo:gc_fijo'), btn('🛒 Gasto compartido variable', 'tipo:gc_variable')],
    [btn('👤 Gasto FM', 'tipo:gasto_fm'), btn('👤 Gasto Lucía', 'tipo:gasto_lucia')],
  ];
  enviar(chatId, '¿Qué quieres registrar?', teclado);
}

function mostrarPersonas(chatId) {
  enviar(chatId, '¿De quién es el ingreso?', [[btn('FM', 'persona:FM'), btn('Lucía', 'persona:Lucía')]]);
}

function mostrarCategorias(chatId, columna) {
  const sh = _ss().getSheetByName(HOJAS.CATEGORIAS);
  const datos = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
  const idxCol = { 'Compartido fijo': 0, 'Compartido variable': 1, 'Individual': 2 }[columna];
  const cats = datos.map(f => f[idxCol]).filter(Boolean);
  const teclado = [];
  for (let i = 0; i < cats.length; i += 2) {
    const fila = [btn(cats[i], 'cat:' + cats[i])];
    if (cats[i + 1]) fila.push(btn(cats[i + 1], 'cat:' + cats[i + 1]));
    teclado.push(fila);
  }
  teclado.push([btn('« Volver', 'menu')]);
  enviar(chatId, 'Elige categoría:', teclado);
}

function procesarImporte(chatId, texto, estado) {
  const limpio = texto.replace(',', '.').replace(/[^\d.]/g, '');
  const valor = Number(limpio);
  if (!valor || valor <= 0) {
    return enviar(chatId, 'No reconozco ese importe. Escribe un número, ej: 42.50');
  }
  estado.importe = valor;

  // Para compartido variable preguntamos quién pagó antes del concepto.
  if (estado.tipo === 'gc_variable' && !estado.pagador) {
    guardarEstado(chatId, estado);
    return enviar(chatId, `Importe: <b>${formatoEur(valor)}</b>\n\n¿Quién lo pagó?`, [
      [btn('Bote común', 'pagador:Bote común')],
      [btn('FM', 'pagador:FM'), btn('Lucía', 'pagador:Lucía')],
    ]);
  }

  estado.paso = 'esperar_concepto';
  guardarEstado(chatId, estado);
  enviar(chatId, `Importe: <b>${formatoEur(valor)}</b>\n\nEscribe el concepto:`);
}

function procesarConcepto(chatId, texto, estado) {
  estado.concepto = texto.slice(0, 120);
  guardarEstado(chatId, estado);
  pedirConfirmacion(chatId, estado);
}

function pedirConfirmacion(chatId, estado) {
  const lineas = ['<b>Confirmar:</b>'];
  lineas.push('Tipo: ' + etiquetaTipo(estado.tipo));
  if (estado.persona) lineas.push('Persona: ' + estado.persona);
  if (estado.categoria) lineas.push('Categoría: ' + estado.categoria);
  if (estado.pagador) lineas.push('Pagado por: ' + estado.pagador);
  lineas.push('Importe: ' + formatoEur(estado.importe));
  lineas.push('Concepto: ' + estado.concepto);

  enviar(chatId, lineas.join('\n'), [[
    btn('✅ Guardar', 'confirmar:si'),
    btn('❌ Descartar', 'confirmar:no'),
  ]]);
}

function guardarMovimiento(chatId, estado) {
  const hoy = new Date();
  const ss = _ss();

  switch (estado.tipo) {
    case 'ingreso':
      ss.getSheetByName(HOJAS.INGRESOS).appendRow([hoy, estado.persona, estado.concepto, estado.importe, 'No']);
      break;
    case 'aportacion':
      ss.getSheetByName(HOJAS.APORTACIONES).appendRow([hoy, estado.persona, estado.concepto, estado.importe]);
      break;
    case 'gc_fijo':
      ss.getSheetByName(HOJAS.GC_FIJOS).appendRow([hoy, estado.concepto, estado.categoria, estado.importe, hoy.getDate()]);
      break;
    case 'gc_variable':
      ss.getSheetByName(HOJAS.GC_VARIABLES).appendRow([hoy, estado.concepto, estado.categoria, estado.importe, estado.pagador]);
      break;
    case 'gasto_fm':
      ss.getSheetByName(HOJAS.G_FM).appendRow([hoy, estado.concepto, estado.categoria, estado.importe]);
      break;
    case 'gasto_lucia':
      ss.getSheetByName(HOJAS.G_LUCIA).appendRow([hoy, estado.concepto, estado.categoria, estado.importe]);
      break;
  }
  limpiarEstado(chatId);
  enviar(chatId, `Guardado ✅\n\n<b>${formatoEur(estado.importe)}</b> — ${estado.concepto}\n\n/nuevo para añadir otro · /resumen`);
}

/* ============== RESUMEN ============== */

function enviarResumen(chatId) {
  const ss = _ss();
  const dash = ss.getSheetByName(HOJAS.DASHBOARD);
  const get = (cell) => dash.getRange(cell).getValue();

  const lineas = [
    `📊 <b>Resumen ${mesActivoConfig()}</b>`,
    '',
    `<b>Ingresos</b>`,
    `  FM: ${formatoEur(get('B5'))}`,
    `  Lucía: ${formatoEur(get('B6'))}`,
    `  Total: ${formatoEur(get('B7'))}`,
    '',
    `<b>Gastos compartidos</b>`,
    `  Fijos: ${formatoEur(get('B15'))}`,
    `  Variables: ${formatoEur(get('B16'))}`,
    `  Total: ${formatoEur(get('B17'))}`,
    `  Bote – gastos: ${formatoEur(get('B18'))}`,
    '',
    `<b>Gastos individuales</b>`,
    `  FM: ${formatoEur(get('B21'))}`,
    `  Lucía: ${formatoEur(get('B22'))}`,
    '',
    `<b>Saldo del mes</b>`,
    `  FM: ${formatoEur(get('B25'))}`,
    `  Lucía: ${formatoEur(get('B26'))}`,
    '',
    `<b>Ahorro</b>`,
    `  Objetivo: ${formatoEur(get('B29'))}`,
    `  Real: ${formatoEur(get('B30'))}`,
    `  Acumulado año: ${formatoEur(get('B32'))}`,
  ];
  enviar(chatId, lineas.join('\n'));
}

/* ============== ESTADO POR CHAT ============== */

function _propClave(chatId) { return 'estado_' + chatId; }

function leerEstado(chatId) {
  const v = PropertiesService.getScriptProperties().getProperty(_propClave(chatId));
  return v ? JSON.parse(v) : null;
}

function guardarEstado(chatId, estado) {
  PropertiesService.getScriptProperties().setProperty(_propClave(chatId), JSON.stringify(estado));
}

function limpiarEstado(chatId) {
  PropertiesService.getScriptProperties().deleteProperty(_propClave(chatId));
}

/* ============== TELEGRAM HTTP ============== */

function enviar(chatId, texto, teclado) {
  const payload = {
    chat_id: chatId,
    text: texto,
    parse_mode: 'HTML',
  };
  if (teclado) payload.reply_markup = JSON.stringify({ inline_keyboard: teclado });

  UrlFetchApp.fetch(API + getBotToken() + '/sendMessage', {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true,
  });
}

function responderCallback(callbackQueryId) {
  UrlFetchApp.fetch(API + getBotToken() + '/answerCallbackQuery', {
    method: 'post',
    payload: { callback_query_id: callbackQueryId },
    muteHttpExceptions: true,
  });
}

function btn(texto, data) { return { text: texto, callback_data: data }; }

function etiquetaTipo(t) {
  return ({
    ingreso: '💰 Ingreso',
    aportacion: '🏦 Aportación al bote',
    gc_fijo: '🏠 Compartido fijo',
    gc_variable: '🛒 Compartido variable',
    gasto_fm: '👤 Gasto FM',
    gasto_lucia: '👤 Gasto Lucía',
  })[t] || t;
}

function formatoEur(v) {
  return Number(v || 0).toLocaleString('es-ES', { style: 'currency', currency: 'EUR' });
}

/* ============== ADMINISTRACIÓN ============== */

/**
 * Despliega primero la Web App, copia la URL y pégala dentro de
 * URL_WEBAPP antes de ejecutar esta función.
 */
function registrarWebhook() {
  const URL_WEBAPP = ''; // <-- pega aquí la URL "/exec" de tu despliegue
  if (!URL_WEBAPP) throw new Error('Edita registrarWebhook() y pega la URL de la Web App.');
  const r = UrlFetchApp.fetch(API + getBotToken() + '/setWebhook?url=' + encodeURIComponent(URL_WEBAPP));
  console.log(r.getContentText());
}

function eliminarWebhook() {
  const r = UrlFetchApp.fetch(API + getBotToken() + '/deleteWebhook');
  console.log(r.getContentText());
}
