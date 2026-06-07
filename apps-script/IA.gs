/**
 * IA.gs
 * ------------------------------------------------------------------
 * Lógica del bot que usa Gemini (definido en Gemini.gs):
 *   - iaProcesarFoto:        Foto → detecta tipo (ticket o notif bancaria) → extrae datos → confirma
 *   - iaProcesarTextoLibre:  Texto tipo "20 cena amigos" → estructura movimiento → confirma
 *   - iaGuardarTicket:       Guarda ticket confirmado en Movimientos + Tickets
 *   - iaGuardarGastoSimple:  Guarda movimiento simple confirmado
 *   - resumenMensualGemini:  Análisis narrativo del mes anterior (trigger día 1)
 *   - detectarAnomaliasGemini: Detecta cambios raros (trigger diario opcional)
 *
 * Todo pide confirmación al usuario ANTES de escribir en hojas.
 * ------------------------------------------------------------------
 */

/* ============================================================
 * SPRINT 1+2: Foto → ticket o notificación bancaria
 * ============================================================ */

/**
 * Llamado cuando el bot recibe un mensaje con foto.
 * msgPhotos: array de PhotoSize de Telegram (varias resoluciones).
 * caption: texto opcional que acompaña la foto.
 */
function iaProcesarFoto(chatId, msgPhotos, caption) {
  enviar(chatId, '📸 Procesando foto con IA…');
  let foto;
  try {
    // La última posición es la mayor resolución.
    const ph = msgPhotos[msgPhotos.length - 1];
    foto = descargarFotoTelegram(ph.file_id);
  } catch (e) {
    return enviar(chatId, '⚠️ No pude descargar la foto: ' + e.message);
  }

  // Esquema de respuesta esperada (responseSchema de Gemini).
  const esquema = {
    type: 'object',
    properties: {
      tipo_documento: { type: 'string', enum: ['ticket_super', 'notificacion_bancaria', 'desconocido'] },
      tienda:    { type: 'string' },
      fecha:     { type: 'string', description: 'AAAA-MM-DD HH:MM o AAAA-MM-DD si no hay hora' },
      total:     { type: 'number' },
      moneda:    { type: 'string' },
      productos: {
        type: 'array',
        items: {
          type: 'object',
          properties: {
            nombre:      { type: 'string' },
            cantidad:    { type: 'number' },
            precio:      { type: 'number' },
            precio_unidad: { type: 'number' },
            categoria:   { type: 'string', description: 'Una de las categorías estándar (Compra, Alquiler, Ocio, Restaurantes, Transporte, Hogar, Salud, Caprichos, Suscripciones, Otros)' },
          },
          required: ['nombre', 'precio'],
        },
      },
      categoria_movimiento: { type: 'string', description: 'Si es notificación bancaria, una categoría del movimiento' },
      concepto: { type: 'string', description: 'Concepto corto si es notificación bancaria' },
    },
    required: ['tipo_documento', 'total'],
  };

  const prompt = `Analiza esta imagen. Puede ser:
  (a) un ticket de supermercado/tienda: extrae cada producto con su precio, la tienda, la fecha y el total.
  (b) una notificación bancaria/captura de cargo: extrae comercio (campo "tienda"), fecha, importe total, concepto corto.
  (c) otra cosa: tipo_documento = "desconocido".

  Devuelve JSON con el esquema dado. Categoriza productos/movimientos usando estas categorías:
  Alquiler, Hipoteca, Comunidad, Luz, Agua, Gas, Internet, Móvil, Seguros, Suscripciones,
  Compra, Restaurantes, Ocio, Transporte, Viajes, Regalos, Hogar, Ropa, Salud, Caprichos, Otros.

  Si no encuentras la fecha, usa la fecha actual. Importes en euros, números sin separadores de miles, punto decimal.`;

  let datos;
  try {
    datos = geminiVisionJSON(foto.b64, foto.mime, prompt, esquema);
  } catch (e) {
    return enviar(chatId, '⚠️ Error con Gemini: ' + e.message + '\n\nRevisa que la "Gemini API key" está bien en Ajustes.');
  }

  if (datos.tipo_documento === 'desconocido') {
    return enviar(chatId, '🤔 No reconozco esta foto como ticket ni como notificación bancaria. ¿Puedes registrarlo a mano con /nuevo?');
  }

  if (datos.tipo_documento === 'ticket_super') {
    return iaPreviewTicket(chatId, datos);
  }
  if (datos.tipo_documento === 'notificacion_bancaria') {
    return iaPreviewNotifBanco(chatId, datos);
  }
}

/** Vista previa de ticket con botones de confirmar. */
function iaPreviewTicket(chatId, datos) {
  const productos = (datos.productos || []).slice(0, 30);
  const total = datos.total || productos.reduce((a, p) => a + (p.precio || 0), 0);
  const lineas = [
    '🧾 <b>Ticket detectado</b>',
    `Tienda: <b>${datos.tienda || '?'}</b>`,
    `Fecha: ${datos.fecha || 'hoy'}`,
    `Total: <b>${formatoEur(total)}</b>`,
    '',
    `<b>Productos (${productos.length}):</b>`,
  ];
  productos.slice(0, 15).forEach(p => {
    lineas.push(`  • ${p.nombre} — ${formatoEur(p.precio)}`);
  });
  if (productos.length > 15) lineas.push(`  …y ${productos.length - 15} más`);
  lineas.push('');
  lineas.push('¿Lo registro como gasto <b>compartido variable</b> (paga el bote) y guardo cada producto en la hoja Tickets?');

  // Guardo el dato pendiente en el estado.
  guardarEstado(chatId, {
    paso: 'esperar_confirm_ia_ticket',
    iaTicket: {
      tienda: datos.tienda || '',
      fecha: datos.fecha || '',
      total: total,
      productos: productos,
    },
  });

  enviar(chatId, lineas.join('\n'), [[
    btn('✅ Sí, guardar', 'ia:ticket:guardar'),
    btn('❌ Descartar', 'ia:cancelar'),
  ]]);
}

/** Vista previa de notificación bancaria con botones. */
function iaPreviewNotifBanco(chatId, datos) {
  const total = datos.total || 0;
  const tienda = datos.tienda || datos.concepto || 'Cargo';
  const lineas = [
    '🏦 <b>Cargo detectado</b>',
    `Comercio: <b>${tienda}</b>`,
    `Fecha: ${datos.fecha || 'hoy'}`,
    `Categoría sugerida: ${datos.categoria_movimiento || 'Otros'}`,
    `Importe: <b>${formatoEur(total)}</b>`,
    '',
    '¿De quién es este gasto?',
  ];
  guardarEstado(chatId, {
    paso: 'esperar_confirm_ia_gasto',
    iaGasto: {
      tipo: 'Individual',
      persona: null, // se elige con botón
      categoria: datos.categoria_movimiento || 'Otros',
      concepto: tienda,
      importe: total,
      fecha: datos.fecha || '',
    },
  });
  enviar(chatId, lineas.join('\n'), [[
    btn('👤 FM', 'ia:persona:FM'),
    btn('👤 Lucía', 'ia:persona:Lucía'),
  ], [
    btn('❌ Descartar', 'ia:cancelar'),
  ]]);
}

/* ============================================================
 * SPRINT 2: Texto libre tipo "20 cena amigos"
 * ============================================================ */

function iaProcesarTextoLibre(chatId, texto) {
  const esquema = {
    type: 'object',
    properties: {
      tipo:     { type: 'string', enum: ['Ingreso', 'Aportación', 'Compartido fijo', 'Compartido variable', 'Individual', 'Ahorro', 'desconocido'] },
      persona:  { type: 'string', enum: ['FM', 'Lucía', 'Bote', 'desconocido'] },
      categoria:{ type: 'string' },
      concepto: { type: 'string' },
      importe:  { type: 'number' },
    },
    required: ['tipo', 'importe'],
  };
  const prompt = `Interpreta este texto del usuario como un movimiento financiero de una pareja (FM y Lucía):
  """${texto}"""

  Devuelve JSON:
  - tipo: uno de Ingreso, Aportación, Compartido fijo, Compartido variable, Individual, Ahorro, desconocido.
  - persona: FM, Lucía, Bote, o desconocido (Bote = gastos compartidos que salen del bote común).
  - categoria: una de Alquiler, Hipoteca, Comunidad, Luz, Agua, Gas, Internet, Móvil, Seguros, Suscripciones, Compra, Restaurantes, Ocio, Transporte, Viajes, Regalos, Hogar, Ropa, Salud, Caprichos, Otros.
  - concepto: descripción breve (ej. "Cena con amigos").
  - importe: número en euros (punto decimal).

  Si el texto menciona "super" o "mercadona" o similar sin más contexto, asume Compartido variable · Bote · Compra.
  Si menciona "cena", "restaurante", "café": Individual del que escribe (no lo sabes → desconocido).
  Si no consigues interpretarlo, tipo = "desconocido".`;

  let d;
  try {
    d = geminiJSON(prompt, esquema);
  } catch (e) {
    return enviar(chatId, '⚠️ No entendí: ' + e.message + '\n\nUsa /nuevo para el menú clásico.');
  }
  if (d.tipo === 'desconocido' || !d.importe) {
    return enviar(chatId, '🤔 No conseguí interpretarlo. Usa /nuevo para el menú clásico.');
  }

  // Si persona es "desconocido" y el tipo lo necesita, preguntamos.
  if (d.persona === 'desconocido' && (d.tipo === 'Individual' || d.tipo === 'Ahorro' || d.tipo === 'Ingreso' || d.tipo === 'Aportación')) {
    guardarEstado(chatId, {
      paso: 'esperar_confirm_ia_gasto',
      iaGasto: {
        tipo: d.tipo,
        persona: null,
        categoria: d.categoria || 'Otros',
        concepto: d.concepto || '',
        importe: d.importe,
        fecha: '',
      },
    });
    return enviar(chatId,
      `💬 <b>${d.tipo}</b> · ${formatoEur(d.importe)}\nCategoría: ${d.categoria || 'Otros'}\nConcepto: ${d.concepto || '(sin concepto)'}\n\n¿De quién?`,
      [[btn('👤 FM', 'ia:persona:FM'), btn('👤 Lucía', 'ia:persona:Lucía')], [btn('❌ Descartar', 'ia:cancelar')]]
    );
  }

  // Preview directo con persona ya conocida.
  guardarEstado(chatId, {
    paso: 'esperar_confirm_ia_gasto',
    iaGasto: {
      tipo: d.tipo,
      persona: d.persona,
      categoria: d.categoria || 'Otros',
      concepto: d.concepto || '',
      importe: d.importe,
      fecha: '',
    },
  });
  enviar(chatId,
    `💬 <b>${d.tipo}</b> · ${d.persona}\nImporte: <b>${formatoEur(d.importe)}</b>\nCategoría: ${d.categoria || 'Otros'}\nConcepto: ${d.concepto || '(sin concepto)'}\n\n¿Guardar?`,
    [[btn('✅ Guardar', 'ia:gasto:guardar'), btn('❌ Descartar', 'ia:cancelar')]]
  );
}

function iaContinuarTrasPersona(chatId, estado) {
  if (!estado || !estado.iaGasto) return enviar(chatId, '⚠️ Estado perdido. Empieza de nuevo.');
  estado.iaGasto.persona = estado.persona;
  guardarEstado(chatId, estado);
  const g = estado.iaGasto;
  enviar(chatId,
    `Persona: <b>${g.persona}</b>\n<b>${g.tipo}</b> · ${formatoEur(g.importe)}\nCategoría: ${g.categoria}\nConcepto: ${g.concepto || '(sin concepto)'}\n\n¿Guardar?`,
    [[btn('✅ Guardar', 'ia:gasto:guardar'), btn('❌ Descartar', 'ia:cancelar')]]
  );
}

function iaContinuarTrasCategoria(chatId, estado) {
  if (!estado || !estado.iaGasto) return enviar(chatId, '⚠️ Estado perdido.');
  estado.iaGasto.categoria = estado.categoria;
  guardarEstado(chatId, estado);
  return iaContinuarTrasPersona(chatId, estado);
}

/* ============================================================
 * Guardar tras confirmación
 * ============================================================ */

function iaGuardarTicket(chatId, estado) {
  if (!estado || !estado.iaTicket) return enviar(chatId, '⚠️ No hay ticket pendiente.');
  const t = estado.iaTicket;
  const sheetMov = _ss().getSheetByName(HOJAS.MOVIMIENTOS);
  const sheetTic = _ss().getSheetByName(HOJAS.TICKETS);

  const fechaMov = parsearFechaIA(t.fecha) || new Date();
  // Movimiento agregado del total como Compartido variable del Bote.
  sheetMov.appendRow([
    fechaMov, 'Compartido variable', 'Bote', 'Compra',
    `${t.tienda || 'Ticket'} (${(t.productos || []).length} prod.)`,
    Number(t.total) || 0,
    'Ticket OCR',
  ]);

  // Genera un Ticket # único para identificar las líneas.
  const ticketId = 'T' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMddHHmmss');
  if (sheetTic) {
    const filas = (t.productos || []).map(p => [
      fechaMov,
      t.tienda || '',
      p.nombre || '',
      Number(p.cantidad) || 1,
      Number(p.precio) || 0,
      Number(p.precio_unidad) || 0,
      p.categoria || '',
      ticketId,
    ]);
    if (filas.length) sheetTic.getRange(sheetTic.getLastRow() + 1, 1, filas.length, 8).setValues(filas);
  }

  logBot(personaPorChat(chatId), 'IA ticket guardado', `${t.tienda} · ${formatoEur(t.total)} · ${(t.productos || []).length} prod.`);
  limpiarEstado(chatId);
  enviar(chatId, `✅ Ticket guardado.\n\n<b>${t.tienda}</b> · ${formatoEur(t.total)}\n${(t.productos || []).length} productos en la hoja Tickets · Ticket # ${ticketId}`);
}

function iaGuardarGastoSimple(chatId, estado) {
  if (!estado || !estado.iaGasto) return enviar(chatId, '⚠️ No hay gasto pendiente.');
  const g = estado.iaGasto;
  if (!g.persona) return enviar(chatId, '⚠️ Falta indicar la persona.');
  const sh = _ss().getSheetByName(HOJAS.MOVIMIENTOS);
  const fecha = parsearFechaIA(g.fecha) || new Date();
  sh.appendRow([fecha, g.tipo, g.persona, g.categoria || 'Otros', g.concepto || '', Number(g.importe) || 0, 'IA']);
  logBot(personaPorChat(chatId), 'IA gasto guardado', `${g.tipo} · ${g.persona} · ${formatoEur(g.importe)}`);
  limpiarEstado(chatId);
  enviar(chatId, `✅ Guardado.\n<b>${g.tipo}</b> · ${g.persona} · ${formatoEur(g.importe)}\n${g.concepto || '(sin concepto)'}`);
}

/* ============================================================
 * SPRINT 3: Resumen mensual con insights + detección de anomalías
 * ============================================================ */

/**
 * Genera y envía un análisis narrativo del MES ANTERIOR al chat.
 * Lo invoca un trigger el día 1 a las 10:00.
 */
function resumenMensualGemini() {
  const ss = _ss();
  const ahora = new Date();
  const inicioMesAct = new Date(ahora.getFullYear(), ahora.getMonth(), 1);
  const inicioMesAnt = new Date(ahora.getFullYear(), ahora.getMonth() - 1, 1);
  const movs = leerMovimientos().filter(m => {
    const f = new Date(m.fecha);
    return f >= inicioMesAnt && f < inicioMesAct;
  });
  if (!movs.length) {
    console.log('resumenMensualGemini: sin movimientos del mes anterior');
    return;
  }
  const mesTxt = Utilities.formatDate(inicioMesAnt, Session.getScriptTimeZone(), 'yyyy-MM');

  // Agrega para no mandar toda la lista (tokens).
  const ingresos = movs.filter(m => m.tipo === 'Ingreso').reduce((a, m) => a + m.importe, 0);
  const aportes = movs.filter(m => m.tipo === 'Aportación').reduce((a, m) => a + m.importe, 0);
  const gFijos = movs.filter(m => m.tipo === 'Compartido fijo').reduce((a, m) => a + m.importe, 0);
  const gVar = movs.filter(m => m.tipo === 'Compartido variable').reduce((a, m) => a + m.importe, 0);
  const gFM = movs.filter(m => m.tipo === 'Individual' && m.persona === 'FM').reduce((a, m) => a + m.importe, 0);
  const gLu = movs.filter(m => m.tipo === 'Individual' && m.persona === 'Lucía').reduce((a, m) => a + m.importe, 0);
  const aho = movs.filter(m => m.tipo === 'Ahorro').reduce((a, m) => a + m.importe, 0);
  // Top categorías
  const porCat = {};
  movs.filter(m => /Compartido|Individual/.test(m.tipo)).forEach(m => {
    porCat[m.categoria] = (porCat[m.categoria] || 0) + m.importe;
  });
  const topCat = Object.entries(porCat).sort((a, b) => b[1] - a[1]).slice(0, 5)
    .map(([c, v]) => `${c}: ${v.toFixed(2)}€`).join(', ');

  const prompt = `Eres el asistente financiero de una pareja (FM y Lucía). Escribe un resumen breve y útil del mes ${mesTxt} en español, con tono cercano pero no condescendiente.

  Datos del mes:
  - Ingresos totales: ${ingresos.toFixed(2)} €
  - Aportaciones al bote común: ${aportes.toFixed(2)} €
  - Gastos compartidos fijos: ${gFijos.toFixed(2)} €
  - Gastos compartidos variables: ${gVar.toFixed(2)} €
  - Gastos individuales FM: ${gFM.toFixed(2)} €
  - Gastos individuales Lucía: ${gLu.toFixed(2)} €
  - Ahorro explícito del mes: ${aho.toFixed(2)} €
  - Top categorías de gasto: ${topCat}

  Estructura (máx. 8 líneas):
  1) Frase resumen del mes (positiva/neutra/atenta según los números).
  2) 2-3 observaciones concretas con números (qué subió, qué bajó si tienes con qué comparar).
  3) 1 recomendación accionable.

  NO repitas todos los números, elige los relevantes. Usa HTML simple para Telegram (<b>, <i>).`;

  let texto;
  try { texto = geminiText(prompt); }
  catch (e) { console.error('Resumen mensual Gemini error', e); return; }

  // Envía a todos los chats autorizados.
  const ids = (_config('Telegram chat IDs autorizados') || '').split(',').map(s => s.trim()).filter(Boolean);
  ids.forEach(id => enviar(id, `📊 <b>Resumen ${mesTxt}</b>\n\n${texto}`));
}

/**
 * Detecta anomalías comparando el día anterior con la media de ese día de la semana
 * en los últimos 30 días. Trigger diario opcional a las 21:00.
 */
function detectarAnomaliasGemini() {
  const ss = _ss();
  const ahora = new Date();
  const ayer = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() - 1);
  const haceMes = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() - 30);
  const movs = leerMovimientos();
  const movsAyer = movs.filter(m => sameDay(new Date(m.fecha), ayer));
  if (!movsAyer.length) return;
  const movsMes = movs.filter(m => new Date(m.fecha) >= haceMes && new Date(m.fecha) < ayer);

  // Resumen muy compacto, sin mandar todo a Gemini.
  const lineaAyer = movsAyer.map(m => `${m.tipo}|${m.categoria}|${m.importe}|${m.concepto || ''}`).join('\n');
  const totalesMes = {};
  movsMes.forEach(m => {
    const k = m.categoria || 'Otros';
    totalesMes[k] = (totalesMes[k] || 0) + m.importe;
  });
  const mediaCat = Object.entries(totalesMes).map(([c, v]) => `${c}: media diaria ${(v / 30).toFixed(2)}€`).join(', ');

  const prompt = `Eres detector de anomalías financieras. Compara los movimientos de AYER con las medias del último mes. Si NO hay anomalías serias (cargos duplicados, subidas mayores del 80%, gastos inusuales), responde EXACTAMENTE: "OK".
  Si hay alguna anomalía, responde con un único mensaje breve (max 4 líneas) en español, comenzando con "⚠️", explicando qué y por qué llama la atención.

  Movimientos de ayer:
  ${lineaAyer}

  Medias del último mes:
  ${mediaCat}`;

  let r;
  try { r = geminiText(prompt); }
  catch (e) { return; }
  if (!r || r.trim() === 'OK') return;
  const ids = (_config('Telegram chat IDs autorizados') || '').split(',').map(s => s.trim()).filter(Boolean);
  ids.forEach(id => enviar(id, r));
}

/* ============================================================
 * Helpers
 * ============================================================ */

function descargarFotoTelegram(fileId) {
  const token = getBotToken();
  const info = JSON.parse(UrlFetchApp.fetch(`${API}${token}/getFile?file_id=${fileId}`).getContentText());
  if (!info.ok) throw new Error('getFile fallo');
  const path = info.result.file_path;
  const blob = UrlFetchApp.fetch(`https://api.telegram.org/file/bot${token}/${path}`).getBlob();
  // Telegram a veces devuelve 'application/octet-stream' que Gemini rechaza.
  // Inferimos el MIME real desde la extensión del path.
  let mime = blob.getContentType() || '';
  if (!mime || mime === 'application/octet-stream') {
    const ext = (path.split('.').pop() || 'jpg').toLowerCase();
    mime = ext === 'png'  ? 'image/png'
         : ext === 'webp' ? 'image/webp'
         : ext === 'heic' ? 'image/heic'
         : ext === 'gif'  ? 'image/gif'
         :                  'image/jpeg';
  }
  return { b64: Utilities.base64Encode(blob.getBytes()), mime: mime };
}

function parsearFechaIA(s) {
  if (!s) return null;
  try {
    // Acepta "2026-05-31", "2026-05-31 14:32", "2026-05-31T14:32"
    const d = new Date(String(s).replace(' ', 'T'));
    return isNaN(d) ? null : d;
  } catch (e) { return null; }
}

function sameDay(a, b) {
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate();
}

/**
 * Para activar los triggers del Sprint 3 ejecuta esta función UNA VEZ
 * desde el editor (o llámala desde el menú).
 */
function configurarTriggersIA() {
  // Borra triggers IA previos
  ScriptApp.getProjectTriggers().forEach(t => {
    const h = t.getHandlerFunction();
    if (h === 'resumenMensualGemini' || h === 'detectarAnomaliasGemini') ScriptApp.deleteTrigger(t);
  });
  // Día 1 de cada mes a las 10:00
  ScriptApp.newTrigger('resumenMensualGemini').timeBased().onMonthDay(1).atHour(10).create();
  // Diario a las 21:00 (anomalías)
  ScriptApp.newTrigger('detectarAnomaliasGemini').timeBased().everyDays(1).atHour(21).create();
  try { SpreadsheetApp.getUi().alert('Triggers IA activados: resumen mensual (día 1, 10:00) y anomalías (diario, 21:00).'); }
  catch (e) {}
}
