/**
 * Config.gs
 * ------------------------------------------------------------------
 * Helpers de configuración: lectura de token, chat IDs autorizados y
 * referencias a las pestañas del Sheet. Toda la config sensible vive
 * en la pestaña "Config" del Sheet, no en código.
 * ------------------------------------------------------------------
 */

const SHEET_ID_KEY = 'SHEET_ID';

/**
 * Guarda el ID del Sheet activo en ScriptProperties.
 * Ejecuta esta función UNA VEZ desde el editor antes de desplegar el
 * Web App. Sin esto, el Web App público no sabe a qué Sheet acceder.
 */
function inicializar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('Esta función debe ejecutarse desde el editor del Sheet, no como Web App.');
  }
  const id = ss.getId();
  PropertiesService.getScriptProperties().setProperty(SHEET_ID_KEY, id);
  console.log('SHEET_ID guardado: ' + id);
}

function _ss() {
  const activa = SpreadsheetApp.getActiveSpreadsheet();
  if (activa) return activa;
  const id = PropertiesService.getScriptProperties().getProperty(SHEET_ID_KEY);
  if (!id) {
    throw new Error('Falta SHEET_ID: ejecuta la función inicializar() desde el editor.');
  }
  return SpreadsheetApp.openById(id);
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
