/**
 * Bot.gs
 * ------------------------------------------------------------------
 * Webhook de Telegram que conduce un flujo guiado por botones inline
 * para registrar movimientos en el libro mayor único (hoja Movimientos).
 * SIN LLM: todo el "entendimiento" se hace por estado finito guardado en
 * PropertiesService (clave = estado_<chatId>).
 *
 * Esquema consolidado (ver CrearDashboard.gs, contrato cerrado):
 *  - Movimientos: A Fecha | B Tipo | C Persona | D Categoría | E Concepto | F Importe | G Nota
 *    Tipo ∈ {Ingreso, Aportación, Compartido fijo, Compartido variable, Individual}
 *    Persona ∈ {FM, Lucía, Bote}
 *  - Ajustes: parámetros clave-valor en A:B; categorías en D:F (cabecera D1:F1).
 *  - _Calc: métricas del mes activo en B1..B20 (celdas fijas).
 *  - Objetivos: A Concepto | B Meta | C Aportado | D %Progreso | E Fecha | F Estado | G Barra
 *  - _LogBot: A Fecha | B Persona | C Acción | D Detalle.
 *
 * Endpoints:
 *  - doPost(e): lo llama Telegram (webhook).
 *  - doGet(e): respuesta amistosa al abrir la URL en el navegador.
 *
 * Acciones utilitarias (ejecutar desde el editor):
 *  - registrarWebhook(): registra esta Web App como webhook de Telegram.
 *  - eliminarWebhook(): la deshace.
 * ------------------------------------------------------------------
 */

const API = 'https://api.telegram.org/bot';

// Textos EXACTOS del contrato para la columna Tipo de Movimientos.
const TIPO_TXT = {
  ingreso: 'Ingreso',
  aportacion: 'Aportación',
  gc_fijo: 'Compartido fijo',
  gc_variable: 'Compartido variable',
  gasto_fm: 'Individual',
  gasto_lucia: 'Individual',
  aho_compartido: 'Ahorro',
  aho_fm: 'Ahorro',
  aho_lucia: 'Ahorro',
};

// Tipos de gasto que cuentan para top / atípicos.
const TIPOS_GASTO = ['Compartido fijo', 'Compartido variable', 'Individual'];

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
    enviar(chatId, `Este chat no está autorizado. Pide a FM que añada tu chat ID (${chatId}) en la pestaña Ajustes.`);
    return;
  }

  // FOTO: ticket de super, notificación bancaria, etc. La procesa Gemini.
  if (msg.photo && msg.photo.length) {
    return iaProcesarFoto(chatId, msg.photo, msg.caption || '');
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
    // No hay flujo activo: probar lenguaje natural con Gemini ("20 cena amigos", etc.)
    if (texto.length > 0 && /\d/.test(texto)) {
      return iaProcesarTextoLibre(chatId, texto);
    }
    enviar(chatId, 'No te entendí. Pulsa /nuevo para registrar un movimiento, /resumen para ver el mes, o /ayuda para ver todos los comandos.');
    return;
  }

  // Estamos esperando texto en un flujo activo.
  switch (estado.paso) {
    case 'esperar_importe': return procesarImporte(chatId, texto, estado);
    case 'esperar_concepto': return procesarConcepto(chatId, texto, estado);
    case 'esperar_objetivo': return procesarNuevoObjetivo(chatId, texto);
    case 'obj_esperar_concepto': return procesarObjConcepto(chatId, texto, estado);
    case 'obj_esperar_meta': return procesarObjMeta(chatId, texto, estado);
    case 'obj_esperar_fecha': return procesarObjFecha(chatId, texto, estado);
    case 'obj_esperar_aportacion': return procesarObjAportacion(chatId, texto, estado);
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
    if (tipo === 'gasto_fm') { estado.persona = 'FM'; guardarEstado(chatId, estado); return mostrarCategorias(chatId, 'Individual'); }
    if (tipo === 'gasto_lucia') { estado.persona = 'Lucía'; guardarEstado(chatId, estado); return mostrarCategorias(chatId, 'Individual'); }
    if (tipo === 'gc_fijo') { estado.persona = 'Bote'; guardarEstado(chatId, estado); return mostrarCategorias(chatId, 'Compartido fijo'); }
    if (tipo === 'gc_variable') { guardarEstado(chatId, estado); return mostrarCategorias(chatId, 'Compartido variable'); }
    // Ahorro: persona fija según botón, categoría 'Otros' por defecto, saltamos a pedir importe.
    if (tipo === 'aho_compartido' || tipo === 'aho_fm' || tipo === 'aho_lucia') {
      estado.persona = (tipo === 'aho_compartido') ? 'Bote' : (tipo === 'aho_fm' ? 'FM' : 'Lucía');
      estado.categoria = 'Otros';
      estado.paso = 'esperar_importe';
      guardarEstado(chatId, estado);
      const etiqueta = (tipo === 'aho_compartido') ? 'ahorro común' :
                       (tipo === 'aho_fm') ? 'ahorro personal de FM' : 'ahorro personal de Lucía';
      return enviar(chatId, `Tipo: <b>Ahorro</b> · ${estado.persona}\n\nEscribe el importe del ${etiqueta} en € (ej: 100):`);
    }
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
    estado.persona = data.slice(8); // pagador = Persona del movimiento
    estado.paso = 'esperar_concepto';
    guardarEstado(chatId, estado);
    return enviar(chatId, `Pagado por: <b>${estado.persona}</b>\n\nEscribe el concepto (ej: Cena con amigos):`);
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
  // Borrar movimiento por número de fila (desde /ultimos con botones)
  if (data.startsWith('borrarid:')) return pedirConfirmacionBorrarFila(chatId, Number(data.slice(9)));
  if (data.startsWith('borrarOK:')) return ejecutarBorrarFila(chatId, Number(data.slice(9)));

  // Objetivos a largo plazo
  if (data === 'obj:menu') return mostrarMenuObjetivos(chatId);
  if (data === 'obj:ver') return verObjetivos(chatId);
  if (data === 'obj:add') return iniciarAddObjetivo(chatId);
  if (data === 'obj:aportar') return listarObjetivos(chatId, 'aportar');
  if (data === 'obj:cumplir') return listarObjetivos(chatId, 'cumplir');
  if (data.startsWith('obj:apor:')) return iniciarAportarObjetivo(chatId, Number(data.slice(9)));
  if (data.startsWith('obj:cump:')) return marcarObjetivoCumplido(chatId, Number(data.slice(9)));

  // Confirmaciones IA (ticket de super, notificación bancaria, texto natural)
  if (data === 'ia:ticket:guardar') return iaGuardarTicket(chatId, estado);
  if (data === 'ia:gasto:guardar') return iaGuardarGastoSimple(chatId, estado);
  if (data === 'ia:cancelar')      { limpiarEstado(chatId); return enviar(chatId, 'Descartado.'); }
  if (data.startsWith('ia:cat:')) {
    // El usuario elige una categoría sugerida (o "otra" para escribirla)
    const idx = Number(data.slice(7));
    if (estado && estado.iaSugerencias && estado.iaSugerencias[idx]) {
      estado.categoria = estado.iaSugerencias[idx];
      guardarEstado(chatId, estado);
      return iaContinuarTrasCategoria(chatId, estado);
    }
  }
  if (data === 'ia:persona:FM' || data === 'ia:persona:Lucía') {
    estado.persona = data.slice(11);
    guardarEstado(chatId, estado);
    return iaContinuarTrasPersona(chatId, estado);
  }
}

/* ============== CAMBIO DE OBJETIVO DE AHORRO ============== */

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
  const sh = _ss().getSheetByName(HOJAS.AJUSTES);
  const datos = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
  const idx = datos.findIndex(([k]) => k === 'Objetivo ahorro mensual conjunto (€)');
  if (idx === -1) {
    limpiarEstado(chatId);
    return enviar(chatId, '⚠️ No encuentro la fila del objetivo en Ajustes. Revísalo.');
  }
  sh.getRange(idx + 2, 2).setValue(valor);
  logBot(personaPorChat(chatId), 'Objetivo de ahorro cambiado', formatoEur(valor) + '/mes');
  limpiarEstado(chatId);
  enviar(chatId, `Objetivo actualizado ✅\n\nNuevo objetivo: <b>${formatoEur(valor)}</b> / mes\n\n/nuevo para seguir`);
}

/* ============== FLUJO ============== */

function mostrarMenuPrincipal(chatId) {
  const teclado = [
    [btn('💰 Ingreso', 'tipo:ingreso'), btn('🏦 Aportación bote', 'tipo:aportacion')],
    [btn('🏠 Gasto comp. fijo', 'tipo:gc_fijo'), btn('🛒 Gasto comp. variable', 'tipo:gc_variable')],
    [btn('👤 Gasto FM', 'tipo:gasto_fm'), btn('👤 Gasto Lucía', 'tipo:gasto_lucia')],
    [btn('🐷 Compartido', 'tipo:aho_compartido'), btn('🐷 FM', 'tipo:aho_fm'), btn('🐷 Lucía', 'tipo:aho_lucia')],
    [btn('🎯 Objetivos', 'obj:menu'), btn('👀 Últimos', 'ver:ultimos')],
    [btn('🏆 Top', 'ver:top'), btn('↩ Borrar último', 'borrar:ultimo')],
    [btn('⚙️ Objetivo de ahorro', 'cfg:objetivo')],
  ];
  enviar(chatId, '¿Qué quieres hacer?\n\n📸 <i>Manda una foto de un ticket o notificación bancaria y lo registro automáticamente.</i>\n💬 <i>O escribe el gasto en lenguaje natural: "20 cena amigos", "53,40 mercadona compra"…</i>', teclado);
}

function mostrarPersonas(chatId, tipo) {
  const q = tipo === 'aportacion' ? '¿Quién aporta al bote?' : '¿De quién es el ingreso?';
  enviar(chatId, q, [[btn('FM', 'persona:FM'), btn('Lucía', 'persona:Lucía')]]);
}

/**
 * Lee categorías de Ajustes!D2:F según ámbito.
 * Cabecera D1:F1 = Compartido fijo | Compartido variable | Individual.
 */
function mostrarCategorias(chatId, ambito) {
  const sh = _ss().getSheetByName(HOJAS.AJUSTES);
  const idxCol = { 'Compartido fijo': 0, 'Compartido variable': 1, 'Individual': 2 }[ambito];
  let cats = [];
  if (sh && sh.getLastRow() >= 2) {
    const datos = sh.getRange(2, 4, sh.getLastRow() - 1, 3).getValues(); // D:F desde fila 2
    cats = datos.map(f => f[idxCol]).filter(c => c !== '' && c != null).map(String);
  }
  const teclado = [];
  for (let i = 0; i < cats.length; i += 2) {
    const fila = [btn(cats[i], 'cat:' + cats[i])];
    if (cats[i + 1]) fila.push(btn(cats[i + 1], 'cat:' + cats[i + 1]));
    teclado.push(fila);
  }
  // Comodín "Otros" siempre disponible.
  teclado.push([btn('Otros', 'cat:Otros')]);
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
  if (estado.tipo === 'gc_variable' && !estado.persona) {
    guardarEstado(chatId, estado);
    return enviar(chatId, `Importe: <b>${formatoEur(valor)}</b>\n\n¿Quién lo pagó?`, [
      [btn('Bote', 'pagador:Bote')],
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
  lineas.push('Importe: ' + formatoEur(estado.importe));
  lineas.push('Concepto: ' + estado.concepto);

  enviar(chatId, lineas.join('\n'), [[
    btn('✅ Guardar', 'confirmar:si'),
    btn('❌ Descartar', 'confirmar:no'),
  ]]);
}

/**
 * Escribe TODO en Movimientos con el orden de columnas del contrato:
 * [Fecha, Tipo, Persona, Categoría, Concepto, Importe, Nota]
 */
function guardarMovimiento(chatId, estado) {
  const hoy = new Date();
  const sh = _ss().getSheetByName(HOJAS.MOVIMIENTOS);
  const tipoTxt = TIPO_TXT[estado.tipo] || estado.tipo;
  const persona = personaMovimiento(estado, chatId);
  const categoria = categoriaMovimiento(estado);
  const concepto = estado.concepto || '';

  sh.appendRow([hoy, tipoTxt, persona, categoria, concepto, estado.importe, '']);

  logBot(persona, tipoTxt, `${formatoEur(estado.importe)} - ${concepto || '(sin concepto)'}`);

  limpiarEstado(chatId);
  enviar(chatId, `Guardado ✅\n\n<b>${tipoTxt}</b> · ${persona}\n<b>${formatoEur(estado.importe)}</b>${concepto ? ' — ' + concepto : ''}\n\n/nuevo para añadir otro · /resumen`);

  // Notificación de gasto atípico (no bloqueante).
  try {
    if (typeof notificacionGastoAtipico === 'function' && TIPOS_GASTO.indexOf(tipoTxt) !== -1) {
      notificacionGastoAtipico({
        tipoTxt, categoria, importe: estado.importe, concepto,
      });
    }
  } catch (err) { console.error('notificacionGastoAtipico', err); }
}

// Determina la Persona (FM/Lucía/Bote) del movimiento según el flujo.
function personaMovimiento(estado, chatId) {
  if (estado.persona) return estado.persona;
  if (estado.tipo === 'gc_fijo') return 'Bote';
  return personaPorChat(chatId);
}

// Categoría según contrato (vacía/"Aportación" para aportaciones, libre para ingresos).
function categoriaMovimiento(estado) {
  if (estado.tipo === 'aportacion') return estado.categoria || 'Aportación';
  if (estado.tipo === 'ingreso') return estado.categoria || 'Otros';
  return estado.categoria || 'Otros';
}

/* ============== RESUMEN (lee _Calc B1..B20) ============== */

function enviarResumen(chatId) {
  const calc = _ss().getSheetByName(HOJAS.CALC);
  if (!calc) return enviar(chatId, '⚠️ No encuentro la hoja _Calc. Ejecuta crearDashboard().');
  const v = calc.getRange('B1:B20').getValues().map(r => Number(r[0]) || 0);
  // Índices: B1=v[0] ... B20=v[19]
  const lineas = [
    `📊 <b>Resumen ${mesActivoConfig()}</b>`,
    '',
    `<b>Ingresos</b>`,
    `  FM: ${formatoEur(v[0])}`,
    `  Lucía: ${formatoEur(v[1])}`,
    `  Total: ${formatoEur(v[2])}`,
    '',
    `<b>Aportaciones al bote</b>`,
    `  FM: ${formatoEur(v[3])}`,
    `  Lucía: ${formatoEur(v[4])}`,
    `  Total: ${formatoEur(v[5])}`,
    '',
    `<b>Gastos compartidos</b>`,
    `  Fijos: ${formatoEur(v[6])}`,
    `  Variables: ${formatoEur(v[7])}`,
    `  Total: ${formatoEur(v[8])}`,
    '',
    `<b>Gastos individuales</b>`,
    `  FM: ${formatoEur(v[9])}`,
    `  Lucía: ${formatoEur(v[10])}`,
    `  Gastos totales: ${formatoEur(v[11])}`,
    '',
    `<b>Saldos del mes</b>`,
    `  FM: ${formatoEur(v[12])}`,
    `  Lucía: ${formatoEur(v[13])}`,
    `  Bote sobrante: ${formatoEur(v[14])}`,
    '',
    `<b>Ahorro</b>`,
    `  Objetivo: ${formatoEur(v[15])}`,
    `  Real: ${formatoEur(v[16])}`,
    `  % objetivo: ${Math.round(v[17] * 100)}%`,
    `  Acumulado año: ${formatoEur(v[18])}`,
    `  Balance del mes: ${formatoEur(v[19])}`,
  ];
  enviar(chatId, lineas.join('\n'));
}

/* ============== RESUMEN ANUAL (agrega sobre Movimientos) ============== */

function enviarResumenAnual(chatId) {
  const anio = new Date().getFullYear();
  const movs = leerMovimientos();
  let ingresos = 0, gastos = 0;
  const acumCat = {};
  movs.forEach(m => {
    if (!(m.fecha instanceof Date) || m.fecha.getFullYear() !== anio) return;
    if (m.tipo === 'Ingreso') ingresos += m.importe;
    else if (TIPOS_GASTO.indexOf(m.tipo) !== -1) {
      gastos += m.importe;
      const cat = m.categoria || 'Sin categoría';
      acumCat[cat] = (acumCat[cat] || 0) + m.importe;
    }
  });
  const ahorro = ingresos - gastos;
  const top = Object.keys(acumCat).map(k => [k, acumCat[k]]).sort((a, b) => b[1] - a[1]).slice(0, 3);

  const lineas = [
    `📅 <b>Resumen anual ${anio}</b>`,
    '',
    `Ingresos: <b>${formatoEur(ingresos)}</b>`,
    `Gastos: <b>${formatoEur(gastos)}</b>`,
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
  const sh = _ss().getSheetByName(HOJAS.MOVIMIENTOS);
  if (!sh || sh.getLastRow() < 2) return enviar(chatId, 'No hay movimientos registrados todavía.');
  const ultimaFila = sh.getLastRow();
  const n = Math.min(10, ultimaFila - 1);
  const filas = sh.getRange(ultimaFila - n + 1, 1, n, 7).getValues();
  const filaInicio = ultimaFila - n + 1;

  const lineas = ['👀 <b>Últimos movimientos</b>', '<i>pulsa 🗑 para borrar uno</i>', ''];
  const teclado = [];
  // Mostrar de más reciente a más antiguo
  for (let i = filas.length - 1; i >= 0; i--) {
    const v = filas[i];
    const filaReal = filaInicio + i;
    lineas.push(`<b>${v[1]}</b> · ${v[2]} · ${fechaCorta(v[0])}`);
    lineas.push(`  ${v[4] || '(sin concepto)'} — <b>${formatoEur(v[5])}</b>`);
    teclado.push([btn(`🗑 ${truncar(v[4] || v[3], 18)} · ${formatoEur(v[5])}`, `borrarid:${filaReal}`)]);
  }
  lineas.push('');
  lineas.push('<i>Para EDITAR un movimiento: abre la hoja Movimientos en el Sheet y edita la celda directamente.</i>');
  enviar(chatId, lineas.join('\n'), teclado);
}

function truncar(s, n) {
  s = String(s || '');
  return s.length > n ? s.slice(0, n - 1) + '…' : s;
}

/* ---------- Borrar movimiento por nº de fila ---------- */

function pedirConfirmacionBorrarFila(chatId, fila) {
  const sh = _ss().getSheetByName(HOJAS.MOVIMIENTOS);
  if (!sh || fila < 2 || fila > sh.getLastRow()) return enviar(chatId, '⚠️ Ese movimiento ya no existe.');
  const v = sh.getRange(fila, 1, 1, 7).getValues()[0];
  enviar(chatId,
    `¿Borrar este movimiento?\n\n<b>${v[1]}</b> · ${v[2]}\n${v[4] || '(sin concepto)'} — <b>${formatoEur(v[5])}</b>\nFecha: ${fechaCorta(v[0])}`,
    [[btn('✅ Sí, borrar', `borrarOK:${fila}`), btn('❌ No', 'borrar:no')]]);
}

function ejecutarBorrarFila(chatId, fila) {
  const sh = _ss().getSheetByName(HOJAS.MOVIMIENTOS);
  if (!sh || fila < 2 || fila > sh.getLastRow()) return enviar(chatId, '⚠️ Ese movimiento ya no existe.');
  const v = sh.getRange(fila, 1, 1, 7).getValues()[0];
  sh.deleteRow(fila);
  logBot(personaPorChat(chatId), 'Borrado', `${v[1]} · ${formatoEur(v[5])} - ${v[4] || '(sin concepto)'}`);
  enviar(chatId, `Borrado ✅\n\n<b>${v[1]}</b> · ${v[2]}\n${v[4] || '(sin concepto)'} — <b>${formatoEur(v[5])}</b>`);
}

/* ============== TOP GASTOS DEL MES (sobre Movimientos) ============== */

function enviarTopMes(chatId) {
  const mes = mesActivoConfig(); // 'AAAA-MM'
  const [anio, m] = mes.split('-').map(Number);
  const inicio = new Date(anio, m - 1, 1);
  const fin = new Date(anio, m, 1);

  const acc = {};
  leerMovimientos().forEach(mov => {
    if (TIPOS_GASTO.indexOf(mov.tipo) === -1) return;
    if (!(mov.fecha instanceof Date) || mov.fecha < inicio || mov.fecha >= fin) return;
    const cat = mov.categoria || 'Sin categoría';
    acc[cat] = (acc[cat] || 0) + mov.importe;
  });
  const top = Object.keys(acc).map(k => [k, acc[k]]).sort((a, b) => b[1] - a[1]).slice(0, 5);
  if (top.length === 0) return enviar(chatId, `🏆 Sin gastos registrados en ${mes}.`);
  const lineas = [`🏆 <b>Top gastos ${mes}</b>`, ''];
  top.forEach((t, i) => lineas.push(`${i + 1}. ${t[0]}: <b>${formatoEur(t[1])}</b>`));
  enviar(chatId, lineas.join('\n'));
}

/* ============== BORRAR ÚLTIMO ============== */

function pedirConfirmacionBorrarUltimo(chatId) {
  const sh = _ss().getSheetByName(HOJAS.MOVIMIENTOS);
  if (!sh || sh.getLastRow() < 2) return enviar(chatId, 'No hay movimientos que borrar.');
  const fila = sh.getLastRow();
  const v = sh.getRange(fila, 1, 1, 7).getValues()[0];
  enviar(chatId,
    `¿Borrar el último movimiento?\n\n<b>${v[1]}</b> · ${v[2]}\n${v[4] || '(sin concepto)'} — <b>${formatoEur(v[5])}</b>\nFecha: ${fechaCorta(v[0])}`,
    [[btn('✅ Sí, borrar', 'borrar:si'), btn('❌ No', 'borrar:no')]]);
}

function ejecutarBorrarUltimo(chatId) {
  const sh = _ss().getSheetByName(HOJAS.MOVIMIENTOS);
  if (!sh || sh.getLastRow() < 2) return enviar(chatId, 'No hay movimientos que borrar.');
  const fila = sh.getLastRow();
  const v = sh.getRange(fila, 1, 1, 7).getValues()[0];
  sh.deleteRow(fila);
  logBot(personaPorChat(chatId), 'Borrado', `${v[1]} · ${formatoEur(v[5])} - ${v[4] || '(sin concepto)'}`);
  enviar(chatId, `Borrado ✅\n\n<b>${v[1]}</b> · ${v[2]}\n${v[4] || '(sin concepto)'} — <b>${formatoEur(v[5])}</b>`);
}

/* ============== LECTURA DE MOVIMIENTOS ============== */

/**
 * Devuelve todas las filas de datos de Movimientos como objetos.
 * Orden = orden de la hoja (la última fila es la más reciente).
 */
function leerMovimientos() {
  const sh = _ss().getSheetByName(HOJAS.MOVIMIENTOS);
  if (!sh || sh.getLastRow() < 2) return [];
  const filas = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  return filas.map((f, i) => ({
    fila: i + 2,
    fecha: f[0],
    tipo: String(f[1] || ''),
    persona: String(f[2] || ''),
    categoria: String(f[3] || ''),
    concepto: String(f[4] || ''),
    importe: Number(f[5]) || 0,
    nota: f[6],
  })).filter(m => m.tipo);
}

/* ============== LOG BOT ============== */

function logBot(persona, accion, detalle) {
  try {
    const sh = _ss().getSheetByName(HOJAS.LOG);
    if (!sh) return; // si la pestaña aún no existe, no rompemos el flujo
    sh.appendRow([new Date(), persona || '', accion || '', detalle || '']);
  } catch (err) {
    console.error('logBot error', err);
  }
}

/* ============== COMANDOS RÁPIDOS ============== */

/**
 * Devuelve {cmd, importe, concepto} o null.
 * Formatos:
 *   /g 12,50 cafe   → gasto individual del que envía
 *   /gc 30 cena     → gasto compartido variable (paga el bote)
 *   /i 600 nomina   → ingreso del que envía
 *   /a 200          → aportación al bote del que envía
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
  const sh = _ss().getSheetByName(HOJAS.MOVIMIENTOS);
  const hoy = new Date();
  let tipoTxt = '', mvPersona = persona, categoria = 'Otros', concepto = c.concepto;

  switch (c.cmd) {
    case 'g': // Individual del que envía
      tipoTxt = 'Individual';
      break;
    case 'gc': // Compartido variable, paga el bote
      tipoTxt = 'Compartido variable';
      mvPersona = 'Bote';
      break;
    case 'i': // Ingreso del que envía
      tipoTxt = 'Ingreso';
      categoria = 'Otros';
      break;
    case 'a': // Aportación al bote del que envía
      tipoTxt = 'Aportación';
      categoria = 'Aportación';
      concepto = '';
      break;
  }

  sh.appendRow([hoy, tipoTxt, mvPersona, categoria, concepto || '', c.importe, '']);
  logBot(persona, tipoTxt, `${formatoEur(c.importe)} - ${concepto || '(sin concepto)'}`);
  enviar(chatId, `Guardado ✅ (${tipoTxt})\n<b>${formatoEur(c.importe)}</b>${concepto ? ' — ' + concepto : ''}`);

  try {
    if (typeof notificacionGastoAtipico === 'function' && TIPOS_GASTO.indexOf(tipoTxt) !== -1) {
      notificacionGastoAtipico({ tipoTxt, categoria, importe: c.importe, concepto });
    }
  } catch (err) { console.error('notificacionGastoAtipico', err); }
}

/**
 * Devuelve "FM" o "Lucía" según el chatId.
 * Prioridad:
 *  1) Filas en Ajustes "Chat FM" y "Chat Lucía".
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

function verObjetivos(chatId) {
  const objs = leerObjetivos();
  if (objs.length === 0) return enviar(chatId, 'Aún no hay objetivos. Usa Añadir.');
  const lineas = ['🎯 <b>Objetivos</b>', ''];
  objs.forEach(o => {
    const pct = o.meta > 0 ? Math.min(1, o.aportado / o.meta) : 0;
    lineas.push(`<b>${o.concepto}</b>${o.estado ? ' — ' + o.estado : ''}`);
    lineas.push(`  ${barra(pct)} ${(pct * 100).toFixed(0)}%`);
    lineas.push(`  ${formatoEur(o.aportado)} / ${formatoEur(o.meta)}` + (o.fechaObjetivo instanceof Date ? ` · ${fechaCorta(o.fechaObjetivo)}` : ''));
    lineas.push('');
  });
  enviar(chatId, lineas.join('\n'));
}

/**
 * Objetivos: A Concepto | B Meta | C Aportado | D %Progreso | E Fecha | F Estado | G Barra
 */
function leerObjetivos() {
  const sh = _ss().getSheetByName(HOJAS.OBJETIVOS);
  if (!sh || sh.getLastRow() < 2) return [];
  const filas = sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues();
  return filas.map((f, i) => ({
    fila: i + 2,
    concepto: String(f[0] || ''),
    meta: Number(f[1]) || 0,
    aportado: Number(f[2]) || 0,
    porcentaje: f[3],
    fechaObjetivo: f[4],
    estado: f[5] || 'En curso',
  })).filter(o => o.concepto);
}

function iniciarAddObjetivo(chatId) {
  guardarEstado(chatId, { tipo: 'obj_add', paso: 'obj_esperar_concepto' });
  enviar(chatId, '🎯 Nuevo objetivo.\n\nEscribe el <b>concepto</b> (ej: Viaje Japón):');
}

function procesarObjConcepto(chatId, texto, estado) {
  estado.concepto = texto.slice(0, 120);
  estado.paso = 'obj_esperar_meta';
  guardarEstado(chatId, estado);
  enviar(chatId, `Concepto: <b>${estado.concepto}</b>\n\nEscribe la <b>meta</b> en € (ej: 3000):`);
}

function procesarObjMeta(chatId, texto, estado) {
  const v = Number(texto.replace(',', '.').replace(/[^\d.]/g, ''));
  if (!v || v <= 0) return enviar(chatId, 'Importe inválido. Escribe un número, ej: 3000');
  estado.meta = v;
  estado.paso = 'obj_esperar_fecha';
  guardarEstado(chatId, estado);
  enviar(chatId, `Meta: <b>${formatoEur(v)}</b>\n\nFecha objetivo (AAAA-MM-DD) o "-" para sin fecha:`);
}

function procesarObjFecha(chatId, texto, estado) {
  let fecha = '';
  if (texto !== '-' && texto !== '') {
    const m = texto.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return enviar(chatId, 'Formato inválido. Usa AAAA-MM-DD o "-" para omitir.');
    fecha = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  }
  const sh = _ss().getSheetByName(HOJAS.OBJETIVOS);
  if (!sh) {
    limpiarEstado(chatId);
    return enviar(chatId, '⚠️ Falta la pestaña Objetivos. Ejecuta crearDashboard().');
  }
  // appendRow [concepto, meta, 0, '', fecha, '', ''] — fórmulas D/F/G las gestiona la hoja.
  sh.appendRow([estado.concepto, estado.meta, 0, '', fecha || '', '', '']);
  logBot(personaPorChat(chatId), 'Objetivo creado', `${formatoEur(estado.meta)} - ${estado.concepto}`);
  limpiarEstado(chatId);
  enviar(chatId, `Objetivo creado ✅\n<b>${estado.concepto}</b> · meta ${formatoEur(estado.meta)}`);
}

function listarObjetivos(chatId, accion) {
  const objs = leerObjetivos().filter(o => o.estado !== 'Cumplido');
  if (objs.length === 0) return enviar(chatId, 'No hay objetivos en curso.');
  const teclado = objs.map(o => [btn(o.concepto, `obj:${accion === 'aportar' ? 'apor' : 'cump'}:${o.fila}`)]);
  teclado.push([btn('« Volver', 'obj:menu')]);
  enviar(chatId, accion === 'aportar' ? '¿A qué objetivo aportas?' : '¿Cuál marcas como cumplido?', teclado);
}

function iniciarAportarObjetivo(chatId, fila) {
  guardarEstado(chatId, { tipo: 'obj_aportar', paso: 'obj_esperar_aportacion', fila });
  enviar(chatId, 'Escribe el importe a aportar en € (ej: 50):');
}

function procesarObjAportacion(chatId, texto, estado) {
  const v = Number(texto.replace(',', '.').replace(/[^\d.]/g, ''));
  if (!v || v <= 0) return enviar(chatId, 'Importe inválido. Escribe un número, ej: 50');
  const sh = _ss().getSheetByName(HOJAS.OBJETIVOS);
  const fila = estado.fila;
  const valores = sh.getRange(fila, 1, 1, 3).getValues()[0];
  const concepto = valores[0];
  const meta = Number(valores[1]) || 0;
  const aportadoPrev = Number(valores[2]) || 0;
  const nuevo = aportadoPrev + v;
  sh.getRange(fila, 3).setValue(nuevo); // columna C Aportado
  logBot(personaPorChat(chatId), 'Aportación a objetivo', `${formatoEur(v)} - ${concepto}`);
  limpiarEstado(chatId);
  const pct = meta > 0 ? Math.min(1, nuevo / meta) : 0;
  enviar(chatId, `Aportado ✅\n<b>${concepto}</b>\n${barra(pct)} ${(pct * 100).toFixed(0)}%\n${formatoEur(nuevo)} / ${formatoEur(meta)}`);
}

function marcarObjetivoCumplido(chatId, fila) {
  const sh = _ss().getSheetByName(HOJAS.OBJETIVOS);
  const valores = sh.getRange(fila, 1, 1, 2).getValues()[0];
  const concepto = valores[0];
  const meta = Number(valores[1]) || 0;
  // Aportado = Meta y estado = Cumplido (la columna F puede ser fórmula; lo escribimos igualmente).
  sh.getRange(fila, 3).setValue(meta);
  sh.getRange(fila, 6).setValue('Cumplido');
  logBot(personaPorChat(chatId), 'Objetivo cumplido', String(concepto));
  enviar(chatId, `🏁 <b>${concepto}</b> marcado como cumplido. ¡Enhorabuena!`);
}

function barra(pct) {
  const total = 10;
  const llenos = Math.round(Math.max(0, Math.min(1, pct)) * total);
  return '█'.repeat(llenos) + '░'.repeat(total - llenos);
}

/* ============== AYUDA ============== */

function enviarAyuda(chatId) {
  const lineas = [
    '<b>Comandos</b>',
    '/nuevo · menú principal',
    '/resumen · resumen del mes activo',
    '/resumen_anual · resumen del año',
    '/ultimos · últimos 10 movimientos (con botón 🗑 para borrar cualquiera)',
    '/top · top gastos del mes',
    '/borrar_ultimo · borra el último movimiento (atajo)',
    '<i>Para editar un movimiento: abre la hoja Movimientos y edita la celda.</i>',
    '/objetivos · objetivos a largo plazo',
    '/objetivo · cambia objetivo de ahorro mensual',
    '/cancelar · cancela el flujo actual',
    '/id · tu chat ID',
    '',
    '<b>Comandos rápidos</b>',
    '/g 12,50 cafe — gasto individual tuyo',
    '/gc 30 cena — gasto compartido variable (paga el bote)',
    '/i 600 nomina — ingreso tuyo',
    '/a 200 — aportación al bote',
    '',
    '<b>📸 Con foto</b>',
    'Mándame una foto de un <b>ticket</b> de super o de una <b>notificación bancaria</b> y la registro automáticamente.',
    '',
    '<b>💬 Lenguaje natural</b>',
    'Escríbelo directo: "20 cena con amigos", "53,40 mercadona compra semanal", "600 nómina". Lo entiendo y pido confirmación.',
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
    gasto_fm: '👤 Gasto FM (Individual)',
    gasto_lucia: '👤 Gasto Lucía (Individual)',
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
