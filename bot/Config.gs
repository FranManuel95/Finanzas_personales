/**
 * Config.gs
 * ------------------------------------------------------------------
 * Helpers de configuración: lectura de token, chat IDs autorizados y
 * referencias a las pestañas del Sheet. Toda la config sensible vive
 * en la pestaña "Config" del Sheet, no en código.
 * ------------------------------------------------------------------
 */

function _ss() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function _config(clave) {
  const sh = _ss().getSheetByName(HOJAS.CONFIG);
  const datos = sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues();
  const fila = datos.find(([k]) => k === clave);
  return fila ? fila[1] : null;
}

function getBotToken() {
  const t = _config('Telegram bot token');
  if (!t) throw new Error('Falta el token de Telegram en la pestaña Config.');
  return String(t).trim();
}

function getChatsAutorizados() {
  const raw = _config('Telegram chat IDs autorizados');
  if (!raw) return [];
  return String(raw)
    .split(',')
    .map(s => s.trim())
    .filter(Boolean);
}

function chatAutorizado(chatId) {
  const lista = getChatsAutorizados();
  if (lista.length === 0) return true; // sin lista = abierto (útil al probar)
  return lista.includes(String(chatId));
}

function mesActivoConfig() {
  return String(_config('Mes activo')).trim();
}
