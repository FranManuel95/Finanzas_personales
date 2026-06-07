/**
 * Gemini.gs
 * ------------------------------------------------------------------
 * Cliente mínimo de Google Gemini API (tier gratuito).
 *
 * - geminiVision(b64, mime, prompt)  → analiza una imagen y devuelve texto
 * - geminiText(prompt)               → llamada solo de texto
 * - geminiJSON(prompt, esquema)      → fuerza respuesta JSON (con esquema opcional)
 *
 * La API key se lee de Ajustes!A:B (clave "Gemini API key") con cache.
 * Modelo por defecto: gemini-2.5-flash (rápido + tier gratuito generoso).
 * Docs: https://ai.google.dev/api/generate-content
 * ------------------------------------------------------------------
 */

const GEMINI_MODEL = 'gemini-2.5-flash';
const GEMINI_BASE = 'https://generativelanguage.googleapis.com/v1beta/models/';

/** Lee la API key desde Ajustes con cache de 5 min. */
function getGeminiKey() {
  const cache = CacheService.getScriptCache();
  let k = cache.get('GEMINI_KEY');
  if (k) return k;
  const sh = _ss().getSheetByName(HOJAS.AJUSTES);
  if (!sh) throw new Error('No encuentro hoja Ajustes');
  const finder = sh.getRange('A:B').createTextFinder('Gemini API key').matchEntireCell(true).findNext();
  if (!finder) throw new Error('Falta la fila "Gemini API key" en Ajustes');
  k = String(sh.getRange(finder.getRow(), 2).getValue() || '').trim();
  if (!k) throw new Error('Pega tu API key de Gemini en Ajustes!B (fila "Gemini API key").');
  cache.put('GEMINI_KEY', k, 300);
  return k;
}

/** Llamada genérica a Gemini. Devuelve el texto de la primera respuesta. */
function _geminiCall(parts, generationConfig) {
  const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${getGeminiKey()}`;
  const payload = {
    contents: [{ parts: parts }],
    generationConfig: generationConfig || { temperature: 0.1 },
  };
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
  });
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) {
    console.error('Gemini error', code, body.slice(0, 400));
    throw new Error(`Gemini ${code}: ${body.slice(0, 200)}`);
  }
  const json = JSON.parse(body);
  const cand = (json.candidates || [])[0];
  if (!cand) throw new Error('Gemini sin candidatos');
  const text = ((cand.content || {}).parts || []).map(p => p.text || '').join('').trim();
  if (!text) throw new Error('Gemini respuesta vacía');
  return text;
}

/** Analiza una imagen (base64) con un prompt. */
function geminiVision(b64, mime, prompt) {
  return _geminiCall([
    { inlineData: { mimeType: mime || 'image/jpeg', data: b64 } },
    { text: prompt },
  ]);
}

/** Llamada solo texto. */
function geminiText(prompt) {
  return _geminiCall([{ text: prompt }]);
}

/**
 * Pide a Gemini una respuesta JSON estructurada. Devuelve el objeto parseado.
 * Si pasas un esquema, lo usa para forzar el formato (responseSchema).
 * Si no, intenta parsear el texto de salida.
 */
function geminiJSON(prompt, esquema) {
  const config = { temperature: 0.1, responseMimeType: 'application/json' };
  if (esquema) config.responseSchema = esquema;
  const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${getGeminiKey()}`;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: config,
    }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(`Gemini ${res.getResponseCode()}: ${res.getContentText().slice(0, 200)}`);
  }
  const json = JSON.parse(res.getContentText());
  const txt = (((json.candidates || [])[0] || {}).content || {}).parts[0].text;
  return JSON.parse(txt);
}

/**
 * Versión visión + JSON: imagen + prompt, devuelve JSON.
 * Usa responseMimeType para forzar JSON en la respuesta.
 */
function geminiVisionJSON(b64, mime, prompt, esquema) {
  const config = { temperature: 0.1, responseMimeType: 'application/json' };
  if (esquema) config.responseSchema = esquema;
  const url = `${GEMINI_BASE}${GEMINI_MODEL}:generateContent?key=${getGeminiKey()}`;
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({
      contents: [{
        parts: [
          { inlineData: { mimeType: mime || 'image/jpeg', data: b64 } },
          { text: prompt },
        ],
      }],
      generationConfig: config,
    }),
    muteHttpExceptions: true,
  });
  if (res.getResponseCode() !== 200) {
    throw new Error(`Gemini ${res.getResponseCode()}: ${res.getContentText().slice(0, 200)}`);
  }
  const json = JSON.parse(res.getContentText());
  const txt = (((json.candidates || [])[0] || {}).content || {}).parts[0].text;
  return JSON.parse(txt);
}

/** Test rápido (ejecuta esta función desde el editor para verificar la key). */
function testGemini() {
  try {
    const r = geminiText('Responde solo "OK" si me lees.');
    Logger.log('Respuesta Gemini: ' + r);
    SpreadsheetApp.getUi().alert('Gemini conectado ✅\n\nRespuesta: ' + r);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Error: ' + e.message);
  }
}
