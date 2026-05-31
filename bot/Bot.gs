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

function doGet(e) {
  if (e && e.parameter && e.parameter.debug === '1') {
    try {
      const r = UrlFetchApp.fetch(API + getBotToken() + '/getWebhookInfo');
      return ContentService.createTextOutput(r.getContentText());
    } catch (err) {
      return ContentService.createTextOutput('Error: ' + err.message);
    }
  }
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
  return ContentService
    .createTextOutput(JSON.stringify({ok: true}))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ============== MANEJO DE EVENTOS ============== */

function manejarMensaje(msg) {
  const chatId = msg.chat.id;
  if (!chatAutorizado(chatId)) {
    enviar(chatId, `Este chat no está autorizado. Pide a FM que añada tu chat ID (${chatId}) en la pestaña Config.`);
    return;
  }

  // Foto: si llega una foto, la asociamos al último movimiento si así se pidió.
  if (msg.photo && msg.photo.length > 0) {
    return manejarFoto(chatId, msg);
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
  if (texto === '/resumen_anual') {
    enviarResumenAnual(chatId);
    return;
  }
  if (texto === '/id') {
    enviar(chatId, `Tu chat ID es: <code>${chatId}</code>`);
    return;
  }
  if (texto === '/objetivo') {
    iniciarCambioObjetivo(chatId);
    return;
  }
  if (texto === '/ultimos') {
    return enviarUltimos(chatId);
  }
  if (texto === '/top') {
    return enviarTopMes(chatId);
  }
  if (texto === '/borrar_ultimo') {
    return pedirConfirmacionBorrarUltimo(chatId);
  }
  if (texto === '/foto') {
    return iniciarFoto(chatId);
  }
  if (texto === '/objetivos') {
    return mostrarMenuObjetivos(chatId);
  }
  if (texto === '/ayuda' || texto === '/help') {
    return enviarAyuda(chatId);
  }

  // Comandos rápidos (regex).
  const cmdRapido = parsearComandoRapido(texto);
  if (cmdRapido) {
    return ejecutarComandoRapido(chatId, cmdRapido);
  }

  const estado = leerEstado(chatId);
  if (!estado) {
    enviar(chatId, 'No te entendí. Pulsa /nuevo para registrar un movimiento, /resumen para ver el mes, o /ayuda para ver todos los comandos.');
    return;
  }

  // Estamos esperando texto en un flujo activo.
  switch (estado.paso) {
    case 'esperar_importe': return procesarImporte(chatId, texto, estado);
    case 'esperar_concepto': return procesarConcepto(chatId, texto, estado);
    case 'esperar_objetivo': return procesarNuevoObjetivo(chatId, texto);
    case 'obj_esperar_concepto': return procesarObjLPConcepto(chatId, texto, estado);
    case 'obj_esperar_meta': return procesarObjLPMeta(chatId, texto, estado);
    case 'obj_esperar_fecha': return procesarObjLPFecha(chatId, texto, estado);
    case 'obj_esperar_aportacion': return procesarObjLPAportacion(chatId, texto, estado);
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

  if (data === 'cfg:objetivo') {
    return iniciarCambioObjetivo(chatId);
  }

  if (data === 'ver:ultimos') return enviarUltimos(chatId);
  if (data === 'ver:top') return enviarTopMes(chatId);
  if (data === 'borrar:ultimo') return pedirConfirmacionBorrarUltimo(chatId);
  if (data === 'borrar:si') return ejecutarBorrarUltimo(chatId);
  if (data === 'borrar:no') return enviar(chatId, 'Borrado cancelado.');

  // Objetivos largo plazo
  if (data === 'obj:menu') return mostrarMenuObjetivos(chatId);
  if (data === 'obj:ver') return verObjetivosLP(chatId);
  if (data === 'obj:add') return iniciarAddObjetivoLP(chatId);
  if (data === 'obj:aportar') return listarObjetivosLP(chatId, 'aportar');
  if (data === 'obj:cumplir') return listarObjetivosLP(chatId, 'cumplir');
  if (data.startsWith('obj:apor:')) return iniciarAportarObjetivoLP(chatId, Number(data.slice(9)));
  if (data.startsWith('obj:cump:')) return marcarObjetivoCumplido(chatId, Number(data.slice(9)));
}

function iniciarCambioObjetivo(chatId) {
  const actual = _config('Objetivo ahorro mensual conjunto (€)');
  guardarEstado(chatId, { tipo: 'cfg_objetivo', paso: 'esperar_objetivo' });
  enviar(chatId, `Objetivo actual: <b>${formatoEur(actual)}</b>\n\nEscribe el nuevo objetivo de ahorro mensual en € (ej: 250):`);
}

function procesarNuevoObjetivo(chatId, texto) {
  const limpio = texto.replace(',', '.').replace(/[^\d.]/g, '');
  const valor = Number(limpio);
  if (!valor || valor < 0) {
    return enviar(chatId, 'No reconozco ese importe. Escribe un número, ej: 250');
  }
  const sh = _ss().getSheetByName(HOJAS.CONFIG);
  const datos = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  const idx = datos.findIndex(([k]) => k === 'Objetivo ahorro mensual conjunto (€)');
  if (idx === -1) {
    limpiarEstado(chatId);
    return enviar(chatId, '⚠️ No encuentro la fila del objetivo en Config. Revísalo.');
  }
  sh.getRange(idx + 2, 2).setValue(valor);
  limpiarEstado(chatId);
  enviar(chatId, `Objetivo actualizado ✅\n\nNuevo objetivo: <b>${formatoEur(valor)}</b> / mes\n\n/nuevo para seguir`);
}

/* ============== FLUJO ============== */

function mostrarMenuPrincipal(chatId) {
  const teclado = [
    [btn('💰 Ingreso', 'tipo:ingreso'), btn('🏦 Aportación al bote', 'tipo:aportacion')],
    [btn('🏠 Gasto comp. fijo', 'tipo:gc_fijo'), btn('🛒 Gasto comp. variable', 'tipo:gc_variable')],
    [btn('👤 Gasto FM', 'tipo:gasto_fm'), btn('👤 Gasto Lucía', 'tipo:gasto_lucia')],
    [btn('🎯 Objetivos', 'obj:menu'), btn('👀 Últimos movimientos', 'ver:ultimos')],
    [btn('🏆 Top gastos del mes', 'ver:top'), btn('↩ Borrar último', 'borrar:ultimo')],
    [btn('⚙️ Cambiar objetivo de ahorro', 'cfg:objetivo')],
  ];
  enviar(chatId, '¿Qué quieres hacer?', teclado);
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
  let hojaDestino = '';
  let filaEscrita = 0;

  switch (estado.tipo) {
    case 'ingreso': {
      const sh = ss.getSheetByName(HOJAS.INGRESOS);
      sh.appendRow([hoy, estado.persona, estado.concepto, estado.importe, 'No']);
      hojaDestino = HOJAS.INGRESOS; filaEscrita = sh.getLastRow();
      break;
    }
    case 'aportacion': {
      const sh = ss.getSheetByName(HOJAS.APORTACIONES);
      sh.appendRow([hoy, estado.persona, estado.concepto, estado.importe]);
      hojaDestino = HOJAS.APORTACIONES; filaEscrita = sh.getLastRow();
      break;
    }
    case 'gc_fijo': {
      const sh = ss.getSheetByName(HOJAS.GC_FIJOS);
      sh.appendRow([hoy, estado.concepto, estado.categoria, estado.importe, hoy.getDate()]);
      hojaDestino = HOJAS.GC_FIJOS; filaEscrita = sh.getLastRow();
      break;
    }
    case 'gc_variable': {
      const sh = ss.getSheetByName(HOJAS.GC_VARIABLES);
      sh.appendRow([hoy, estado.concepto, estado.categoria, estado.importe, estado.pagador]);
      hojaDestino = HOJAS.GC_VARIABLES; filaEscrita = sh.getLastRow();
      break;
    }
    case 'gasto_fm': {
      const sh = ss.getSheetByName(HOJAS.G_FM);
      sh.appendRow([hoy, estado.concepto, estado.categoria, estado.importe]);
      hojaDestino = HOJAS.G_FM; filaEscrita = sh.getLastRow();
      break;
    }
    case 'gasto_lucia': {
      const sh = ss.getSheetByName(HOJAS.G_LUCIA);
      sh.appendRow([hoy, estado.concepto, estado.categoria, estado.importe]);
      hojaDestino = HOJAS.G_LUCIA; filaEscrita = sh.getLastRow();
      break;
    }
  }

  const persona = estado.persona || estado.pagador || personaPorChat(chatId);
  const detalle = `${formatoEur(estado.importe)} - ${estado.concepto || ''} -> ${hojaDestino}:${filaEscrita}`;
  logBot(persona, etiquetaTipo(estado.tipo), detalle);

  limpiarEstado(chatId);
  enviar(chatId, `Guardado ✅\n\n<b>${formatoEur(estado.importe)}</b> — ${estado.concepto}\n\n/nuevo para añadir otro · /resumen`);

  // Notificación de gasto atípico (no bloqueante).
  try {
    if (typeof notificacionGastoAtipico === 'function' &&
        ['gc_fijo','gc_variable','gasto_fm','gasto_lucia'].indexOf(estado.tipo) !== -1) {
      notificacionGastoAtipico({
        chatId, tipo: estado.tipo, categoria: estado.categoria,
        importe: estado.importe, concepto: estado.concepto,
      });
    }
  } catch (err) { console.error('notificacionGastoAtipico', err); }
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

function enviarResumenAnual(chatId) {
  const ss = _ss();
  const anio = new Date().getFullYear();
  const inicio = new Date(anio, 0, 1);
  const fin = new Date(anio + 1, 0, 1);

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
  const gastosTotal = gcf + gcv + gfm + glu;
  const ahorro = ingresos - gastosTotal;

  // Top 3 categorías del año (todos los gastos).
  const acumCat = {};
  const acumular = (nombre, colCat, colImp) => {
    const sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() < 2) return;
    const filas = sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(colCat, colImp)).getValues();
    filas.forEach(f => {
      const fecha = f[0];
      if (!(fecha instanceof Date) || fecha < inicio || fecha >= fin) return;
      const cat = f[colCat - 1] || 'Sin categoría';
      acumCat[cat] = (acumCat[cat] || 0) + (Number(f[colImp - 1]) || 0);
    });
  };
  acumular(HOJAS.GC_FIJOS, 3, 4);
  acumular(HOJAS.GC_VARIABLES, 3, 4);
  acumular(HOJAS.G_FM, 3, 4);
  acumular(HOJAS.G_LUCIA, 3, 4);

  const top = Object.keys(acumCat)
    .map(k => [k, acumCat[k]])
    .sort((a, b) => b[1] - a[1])
    .slice(0, 3);

  const lineas = [
    `📅 <b>Resumen anual ${anio}</b>`,
    '',
    `Ingresos: <b>${formatoEur(ingresos)}</b>`,
    `Gastos: <b>${formatoEur(gastosTotal)}</b>`,
    `  Comp. fijos: ${formatoEur(gcf)}`,
    `  Comp. variables: ${formatoEur(gcv)}`,
    `  FM: ${formatoEur(gfm)}`,
    `  Lucía: ${formatoEur(glu)}`,
    '',
    `Ahorro: <b>${formatoEur(ahorro)}</b>`,
    '',
    '<b>Top 3 categorías del año</b>',
  ];
  if (top.length === 0) lineas.push('  (sin datos)');
  else top.forEach((t, i) => lineas.push(`  ${i + 1}. ${t[0]}: ${formatoEur(t[1])}`));

  enviar(chatId, lineas.join('\n'));
}

/* ============== ÚLTIMOS MOVIMIENTOS ============== */

function enviarUltimos(chatId) {
  const movs = recopilarUltimosMovimientos(5);
  if (movs.length === 0) return enviar(chatId, 'No hay movimientos registrados todavía.');
  const lineas = ['👀 <b>Últimos movimientos</b>', ''];
  movs.forEach(m => {
    lineas.push(`<b>${etiquetaTipo(m.tipo) || m.hoja}</b> · ${fechaCorta(m.fecha)}`);
    lineas.push(`  ${m.concepto || '(sin concepto)'} — <b>${formatoEur(m.importe)}</b>`);
  });
  enviar(chatId, lineas.join('\n'));
}

function recopilarUltimosMovimientos(n) {
  const ss = _ss();
  const acc = [];
  const fuentes = [
    {hoja: HOJAS.INGRESOS, tipo: 'ingreso', fechaCol: 1, conceptoCol: 3, importeCol: 4},
    {hoja: HOJAS.APORTACIONES, tipo: 'aportacion', fechaCol: 1, conceptoCol: 3, importeCol: 4},
    {hoja: HOJAS.GC_FIJOS, tipo: 'gc_fijo', fechaCol: 1, conceptoCol: 2, importeCol: 4},
    {hoja: HOJAS.GC_VARIABLES, tipo: 'gc_variable', fechaCol: 1, conceptoCol: 2, importeCol: 4},
    {hoja: HOJAS.G_FM, tipo: 'gasto_fm', fechaCol: 1, conceptoCol: 2, importeCol: 4},
    {hoja: HOJAS.G_LUCIA, tipo: 'gasto_lucia', fechaCol: 1, conceptoCol: 2, importeCol: 4},
  ];
  fuentes.forEach(src => {
    const sh = ss.getSheetByName(src.hoja);
    if (!sh || sh.getLastRow() < 2) return;
    const ult = Math.min(n, sh.getLastRow() - 1);
    const ancho = Math.max(src.fechaCol, src.conceptoCol, src.importeCol);
    const filas = sh.getRange(sh.getLastRow() - ult + 1, 1, ult, ancho).getValues();
    filas.forEach(f => {
      const fecha = f[src.fechaCol - 1];
      if (!(fecha instanceof Date)) return;
      acc.push({
        hoja: src.hoja,
        tipo: src.tipo,
        fecha,
        concepto: f[src.conceptoCol - 1],
        importe: Number(f[src.importeCol - 1]) || 0,
      });
    });
  });
  acc.sort((a, b) => b.fecha - a.fecha);
  return acc.slice(0, n);
}

/* ============== TOP GASTOS DEL MES ============== */

function enviarTopMes(chatId) {
  const ss = _ss();
  const mes = mesActivoConfig(); // 'AAAA-MM'
  const [anio, m] = mes.split('-').map(Number);
  const inicio = new Date(anio, m - 1, 1);
  const fin = new Date(anio, m, 1);

  const acc = {};
  const fuentes = [
    {hoja: HOJAS.GC_FIJOS, catCol: 3, impCol: 4},
    {hoja: HOJAS.GC_VARIABLES, catCol: 3, impCol: 4},
    {hoja: HOJAS.G_FM, catCol: 3, impCol: 4},
    {hoja: HOJAS.G_LUCIA, catCol: 3, impCol: 4},
  ];
  fuentes.forEach(src => {
    const sh = ss.getSheetByName(src.hoja);
    if (!sh || sh.getLastRow() < 2) return;
    const ancho = Math.max(src.catCol, src.impCol);
    const filas = sh.getRange(2, 1, sh.getLastRow() - 1, ancho).getValues();
    filas.forEach(f => {
      const fecha = f[0];
      if (!(fecha instanceof Date) || fecha < inicio || fecha >= fin) return;
      const cat = f[src.catCol - 1] || 'Sin categoría';
      acc[cat] = (acc[cat] || 0) + (Number(f[src.impCol - 1]) || 0);
    });
  });
  const top = Object.keys(acc).map(k => [k, acc[k]]).sort((a, b) => b[1] - a[1]).slice(0, 5);
  if (top.length === 0) return enviar(chatId, `🏆 Sin gastos registrados en ${mes}.`);
  const lineas = [`🏆 <b>Top gastos ${mes}</b>`, ''];
  top.forEach((t, i) => lineas.push(`${i + 1}. ${t[0]}: <b>${formatoEur(t[1])}</b>`));
  enviar(chatId, lineas.join('\n'));
}

/* ============== BORRAR ÚLTIMO ============== */

function pedirConfirmacionBorrarUltimo(chatId) {
  const ult = ultimoLogBot();
  if (!ult) return enviar(chatId, 'No hay nada que borrar en el registro del bot.');
  enviar(chatId,
    `¿Borrar el último movimiento?\n\n<b>${ult.accion}</b>\n${ult.detalle}\nFecha: ${fechaCorta(ult.fecha)}`,
    [[btn('✅ Sí, borrar', 'borrar:si'), btn('❌ No', 'borrar:no')]]);
}

function ejecutarBorrarUltimo(chatId) {
  const ss = _ss();
  const logSh = ss.getSheetByName(HOJAS.LOG_BOT);
  if (!logSh || logSh.getLastRow() < 2) return enviar(chatId, 'No hay registro de bot que borrar.');
  const fila = logSh.getLastRow();
  const valores = logSh.getRange(fila, 1, 1, 4).getValues()[0];
  const detalle = String(valores[3] || '');
  // Detalle tiene forma: "12,34 € - concepto -> Hoja:N"
  const m = detalle.match(/->\s*([^:]+):(\d+)\s*$/);
  if (!m) {
    logSh.deleteRow(fila);
    return enviar(chatId, 'Log borrado, pero no pude identificar la fila del movimiento. Revísalo a mano.');
  }
  const hoja = m[1].trim();
  const numFila = Number(m[2]);
  const sh = ss.getSheetByName(hoja);
  if (sh && numFila > 1 && numFila <= sh.getLastRow()) {
    sh.deleteRow(numFila);
  }
  logSh.deleteRow(fila);
  // Recalcular las referencias de filas posteriores del Log_Bot: cualquier
  // entrada que apuntara a una fila > numFila en la misma hoja queda con un
  // índice desplazado. Las reajustamos.
  reindexarLogBot(hoja, numFila);
  enviar(chatId, `Borrado ✅\n\n<b>${hoja}</b> fila ${numFila} eliminada.`);
}

function reindexarLogBot(hoja, filaBorrada) {
  const ss = _ss();
  const logSh = ss.getSheetByName(HOJAS.LOG_BOT);
  if (!logSh || logSh.getLastRow() < 2) return;
  const filas = logSh.getLastRow() - 1;
  const datos = logSh.getRange(2, 4, filas, 1).getValues();
  let cambios = 0;
  for (let i = 0; i < datos.length; i++) {
    const d = String(datos[i][0] || '');
    const mm = d.match(/->\s*([^:]+):(\d+)\s*$/);
    if (!mm) continue;
    if (mm[1].trim() !== hoja) continue;
    const n = Number(mm[2]);
    if (n > filaBorrada) {
      const nuevo = d.replace(/->\s*([^:]+):(\d+)\s*$/, `-> ${hoja}:${n - 1}`);
      datos[i][0] = nuevo;
      cambios++;
    }
  }
  if (cambios > 0) logSh.getRange(2, 4, filas, 1).setValues(datos);
}

function ultimoLogBot() {
  const ss = _ss();
  const sh = ss.getSheetByName(HOJAS.LOG_BOT);
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getRange(sh.getLastRow(), 1, 1, 4).getValues()[0];
  return { fecha: v[0], persona: v[1], accion: v[2], detalle: v[3] };
}

function logBot(persona, accion, detalle) {
  try {
    const sh = _ss().getSheetByName(HOJAS.LOG_BOT);
    if (!sh) return; // si la pestaña aún no existe, no rompemos el flujo
    sh.appendRow([new Date(), persona || '', accion || '', detalle || '']);
  } catch (err) {
    console.error('logBot error', err);
  }
}

/* ============== COMANDOS RÁPIDOS ============== */

/**
 * Devuelve {tipo, importe, concepto} o null.
 * Formatos:
 *   /g 12,50 cafe
 *   /gc 30 cena
 *   /i 600 nomina
 *   /a 200
 */
function parsearComandoRapido(texto) {
  const m = texto.match(/^\/(g|gc|i|a)\s+([\d.,]+)(?:\s+(.+))?$/i);
  if (!m) return null;
  const cmd = m[1].toLowerCase();
  const importe = Number(m[2].replace(',', '.'));
  if (!importe || importe <= 0) return { _error: 'Importe inválido. Ej: /g 12,50 cafe' };
  return { cmd, importe, concepto: (m[3] || '').slice(0, 120) };
}

function ejecutarComandoRapido(chatId, c) {
  if (c._error) return enviar(chatId, c._error);
  const persona = personaPorChat(chatId);
  const hoy = new Date();
  const ss = _ss();
  let tipoEstado = '';
  let hojaDestino = '';
  let filaEscrita = 0;
  let concepto = c.concepto;

  switch (c.cmd) {
    case 'g': {
      // Gasto individual del que envía
      const hoja = persona === 'Lucía' ? HOJAS.G_LUCIA : HOJAS.G_FM;
      const sh = ss.getSheetByName(hoja);
      sh.appendRow([hoy, concepto || '(sin concepto)', 'Otros', c.importe]);
      tipoEstado = persona === 'Lucía' ? 'gasto_lucia' : 'gasto_fm';
      hojaDestino = hoja; filaEscrita = sh.getLastRow();
      break;
    }
    case 'gc': {
      const sh = ss.getSheetByName(HOJAS.GC_VARIABLES);
      sh.appendRow([hoy, concepto || '(sin concepto)', 'Otros', c.importe, 'Bote común']);
      tipoEstado = 'gc_variable';
      hojaDestino = HOJAS.GC_VARIABLES; filaEscrita = sh.getLastRow();
      break;
    }
    case 'i': {
      const sh = ss.getSheetByName(HOJAS.INGRESOS);
      sh.appendRow([hoy, persona, concepto || '(sin concepto)', c.importe, 'No']);
      tipoEstado = 'ingreso';
      hojaDestino = HOJAS.INGRESOS; filaEscrita = sh.getLastRow();
      break;
    }
    case 'a': {
      const sh = ss.getSheetByName(HOJAS.APORTACIONES);
      sh.appendRow([hoy, persona, '', c.importe]);
      tipoEstado = 'aportacion';
      hojaDestino = HOJAS.APORTACIONES; filaEscrita = sh.getLastRow();
      concepto = '';
      break;
    }
  }

  const detalle = `${formatoEur(c.importe)} - ${concepto || ''} -> ${hojaDestino}:${filaEscrita}`;
  logBot(persona, etiquetaTipo(tipoEstado), detalle);

  enviar(chatId, `Guardado ✅ (${etiquetaTipo(tipoEstado)})\n<b>${formatoEur(c.importe)}</b>${concepto ? ' — ' + concepto : ''}`);
}

/**
 * Devuelve "FM" o "Lucía" según el chatId.
 * Prioridad:
 *  1) Filas en Config "Chat FM" y "Chat Lucía".
 *  2) Orden de "Telegram chat IDs autorizados" (1º=FM, 2º=Lucía).
 *  3) Fallback: FM.
 */
function personaPorChat(chatId) {
  const fm = _config('Chat FM');
  const lu = _config('Chat Lucía');
  if (fm && String(fm).trim() === String(chatId)) return 'FM';
  if (lu && String(lu).trim() === String(chatId)) return 'Lucía';
  const autor = getChatsAutorizados();
  const i = autor.indexOf(String(chatId));
  if (i === 0) return 'FM';
  if (i === 1) return 'Lucía';
  return 'FM';
}

/* ============== OBJETIVOS LARGO PLAZO ============== */

function mostrarMenuObjetivos(chatId) {
  enviar(chatId, '🎯 <b>Objetivos a largo plazo</b>', [
    [btn('👁 Ver', 'obj:ver'), btn('➕ Añadir', 'obj:add')],
    [btn('💸 Aportar a uno', 'obj:aportar'), btn('🏁 Marcar cumplido', 'obj:cumplir')],
    [btn('« Menú', 'menu')],
  ]);
}

function verObjetivosLP(chatId) {
  const objs = leerObjetivosLP();
  if (objs.length === 0) return enviar(chatId, 'Aún no hay objetivos. Usa Añadir.');
  const lineas = ['🎯 <b>Objetivos</b>', ''];
  objs.forEach(o => {
    const pct = o.meta > 0 ? Math.min(1, o.aportado / o.meta) : 0;
    lineas.push(`<b>${o.concepto}</b> — ${o.estado}`);
    lineas.push(`  ${barra(pct)} ${(pct * 100).toFixed(0)}%`);
    lineas.push(`  ${formatoEur(o.aportado)} / ${formatoEur(o.meta)}` + (o.fechaObjetivo ? ` · ${fechaCorta(o.fechaObjetivo)}` : ''));
    lineas.push('');
  });
  enviar(chatId, lineas.join('\n'));
}

function leerObjetivosLP() {
  const sh = _ss().getSheetByName(HOJAS.OBJETIVOS_LP);
  if (!sh || sh.getLastRow() < 2) return [];
  const filas = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  return filas.map((f, i) => ({
    fila: i + 2,
    fechaInicio: f[0],
    concepto: f[1],
    meta: Number(f[2]) || 0,
    aportado: Number(f[3]) || 0,
    porcentaje: f[4],
    fechaObjetivo: f[5],
    estado: f[6] || 'En curso',
  })).filter(o => o.concepto);
}

function iniciarAddObjetivoLP(chatId) {
  guardarEstado(chatId, { tipo: 'obj_lp_add', paso: 'obj_esperar_concepto' });
  enviar(chatId, '🎯 Nuevo objetivo.\n\nEscribe el <b>concepto</b> (ej: Viaje Japón):');
}

function procesarObjLPConcepto(chatId, texto, estado) {
  estado.concepto = texto.slice(0, 120);
  estado.paso = 'obj_esperar_meta';
  guardarEstado(chatId, estado);
  enviar(chatId, `Concepto: <b>${estado.concepto}</b>\n\nEscribe la <b>meta</b> en € (ej: 3000):`);
}

function procesarObjLPMeta(chatId, texto, estado) {
  const v = Number(texto.replace(',', '.').replace(/[^\d.]/g, ''));
  if (!v || v <= 0) return enviar(chatId, 'Importe inválido. Escribe un número, ej: 3000');
  estado.meta = v;
  estado.paso = 'obj_esperar_fecha';
  guardarEstado(chatId, estado);
  enviar(chatId, `Meta: <b>${formatoEur(v)}</b>\n\nFecha objetivo (AAAA-MM-DD) o "-" para sin fecha:`);
}

function procesarObjLPFecha(chatId, texto, estado) {
  let fecha = '';
  if (texto !== '-' && texto !== '') {
    const m = texto.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return enviar(chatId, 'Formato inválido. Usa AAAA-MM-DD o "-" para omitir.');
    fecha = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  const sh = _ss().getSheetByName(HOJAS.OBJETIVOS_LP);
  if (!sh) {
    limpiarEstado(chatId);
    return enviar(chatId, '⚠️ Falta la pestaña Objetivos_Largo_Plazo. Ejecuta crearDashboard().');
  }
  const hoy = new Date();
  sh.appendRow([hoy, estado.concepto, estado.meta, 0, 0, fecha || '', 'En curso']);
  logBot(personaPorChat(chatId), 'Objetivo LP creado',
    `${formatoEur(estado.meta)} - ${estado.concepto} -> ${HOJAS.OBJETIVOS_LP}:${sh.getLastRow()}`);
  limpiarEstado(chatId);
  enviar(chatId, `Objetivo creado ✅\n<b>${estado.concepto}</b> · meta ${formatoEur(estado.meta)}`);
}

function listarObjetivosLP(chatId, accion) {
  const objs = leerObjetivosLP().filter(o => o.estado !== 'Cumplido');
  if (objs.length === 0) return enviar(chatId, 'No hay objetivos en curso.');
  const teclado = objs.map(o => [btn(o.concepto, `obj:${accion === 'aportar' ? 'apor' : 'cump'}:${o.fila}`)]);
  teclado.push([btn('« Volver', 'obj:menu')]);
  enviar(chatId, accion === 'aportar' ? '¿A qué objetivo aportas?' : '¿Cuál marcas como cumplido?', teclado);
}

function iniciarAportarObjetivoLP(chatId, fila) {
  guardarEstado(chatId, { tipo: 'obj_lp_aportar', paso: 'obj_esperar_aportacion', fila });
  enviar(chatId, 'Escribe el importe a aportar en € (ej: 50):');
}

function procesarObjLPAportacion(chatId, texto, estado) {
  const v = Number(texto.replace(',', '.').replace(/[^\d.]/g, ''));
  if (!v || v <= 0) return enviar(chatId, 'Importe inválido. Escribe un número, ej: 50');
  const sh = _ss().getSheetByName(HOJAS.OBJETIVOS_LP);
  const fila = estado.fila;
  const valores = sh.getRange(fila, 1, 1, 7).getValues()[0];
  const concepto = valores[1];
  const meta = Number(valores[2]) || 0;
  const aportadoPrev = Number(valores[3]) || 0;
  const nuevo = aportadoPrev + v;
  sh.getRange(fila, 4).setValue(nuevo);
  if (meta > 0) sh.getRange(fila, 5).setValue(nuevo / meta);
  logBot(personaPorChat(chatId), 'Aportación a objetivo LP',
    `${formatoEur(v)} - ${concepto} -> ${HOJAS.OBJETIVOS_LP}:${fila}`);
  limpiarEstado(chatId);
  const pct = meta > 0 ? Math.min(1, nuevo / meta) : 0;
  enviar(chatId, `Aportado ✅\n<b>${concepto}</b>\n${barra(pct)} ${(pct * 100).toFixed(0)}%\n${formatoEur(nuevo)} / ${formatoEur(meta)}`);
}

function marcarObjetivoCumplido(chatId, fila) {
  const sh = _ss().getSheetByName(HOJAS.OBJETIVOS_LP);
  const concepto = sh.getRange(fila, 2).getValue();
  sh.getRange(fila, 7).setValue('Cumplido');
  logBot(personaPorChat(chatId), 'Objetivo LP cumplido', `- ${concepto} -> ${HOJAS.OBJETIVOS_LP}:${fila}`);
  enviar(chatId, `🏁 <b>${concepto}</b> marcado como cumplido. ¡Enhorabuena!`);
}

function barra(pct) {
  const total = 10;
  const llenos = Math.round(Math.max(0, Math.min(1, pct)) * total);
  return '█'.repeat(llenos) + '░'.repeat(total - llenos);
}

/* ============== FOTOS DE TICKETS ============== */

function iniciarFoto(chatId) {
  const ult = ultimoLogBot();
  if (!ult) return enviar(chatId, 'No tengo movimiento reciente al que asociar la foto.');
  guardarEstado(chatId, { tipo: 'foto_pendiente' });
  enviar(chatId, `Envía la foto del ticket; la asociaré al último movimiento:\n<b>${ult.accion}</b> · ${ult.detalle}`);
}

function manejarFoto(chatId, msg) {
  const estado = leerEstado(chatId);
  if (!estado || estado.tipo !== 'foto_pendiente') {
    return enviar(chatId, 'Recibida la foto, pero no hay /foto activo. Usa /foto antes de mandarla.');
  }
  const file = msg.photo[msg.photo.length - 1]; // mayor resolución
  let url = '';
  try {
    const r = UrlFetchApp.fetch(API + getBotToken() + '/getFile?file_id=' + encodeURIComponent(file.file_id));
    const j = JSON.parse(r.getContentText());
    if (j.ok) url = 'https://api.telegram.org/file/bot' + getBotToken() + '/' + j.result.file_path;
  } catch (err) {
    console.error('getFile', err);
  }
  // Anexar al último log
  const sh = _ss().getSheetByName(HOJAS.LOG_BOT);
  if (sh && sh.getLastRow() >= 2) {
    const fila = sh.getLastRow();
    const det = sh.getRange(fila, 4).getValue();
    sh.getRange(fila, 4).setValue(det + ' | foto: ' + (url || file.file_id));
  }
  limpiarEstado(chatId);
  enviar(chatId, 'Foto asociada al último movimiento ✅');
}

/* ============== AYUDA ============== */

function enviarAyuda(chatId) {
  const lineas = [
    '<b>Comandos</b>',
    '/nuevo · menú principal',
    '/resumen · resumen del mes activo',
    '/resumen_anual · resumen del año',
    '/ultimos · últimos 5 movimientos',
    '/top · top gastos del mes',
    '/borrar_ultimo · borra el último movimiento',
    '/objetivos · objetivos a largo plazo',
    '/objetivo · cambia objetivo de ahorro mensual',
    '/foto · asocia la siguiente foto al último gasto',
    '/cancelar · cancela el flujo actual',
    '/id · tu chat ID',
    '',
    '<b>Comandos rápidos</b>',
    '/g 12,50 cafe — gasto individual tuyo',
    '/gc 30 cena — gasto compartido variable',
    '/i 600 nomina — ingreso tuyo',
    '/a 200 — aportación al bote',
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
  if (teclado) payload.reply_markup = { inline_keyboard: teclado };

  UrlFetchApp.fetch(API + getBotToken() + '/sendMessage', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
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

function fechaCorta(f) {
  if (!(f instanceof Date)) return String(f || '');
  return Utilities.formatDate(f, Session.getScriptTimeZone() || 'Europe/Madrid', 'yyyy-MM-dd HH:mm');
}

/* ============== ADMINISTRACIÓN ============== */

/**
 * Despliega primero la Web App, copia la URL y pégala dentro de
 * URL_WEBAPP antes de ejecutar esta función.
 */
function registrarWebhook() {
  const URL_WEBAPP = ''; // <-- pega aquí la URL "/exec" de tu despliegue
  if (!URL_WEBAPP) throw new Error('Edita registrarWebhook() y pega la URL de la Web App.');
  const r = UrlFetchApp.fetch(API + getBotToken() + '/setWebhook?url=' + encodeURIComponent(URL_WEBAPP) + '&drop_pending_updates=true');
  console.log(r.getContentText());
}

function eliminarWebhook() {
  const r = UrlFetchApp.fetch(API + getBotToken() + '/deleteWebhook');
  console.log(r.getContentText());
}
