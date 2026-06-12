/**
 * Planning & Analytics — Backend de Progreso (Apps Script Web App)
 *
 * Guarda y lee los % de progreso de objetivos en una pestaña "Progreso"
 * del mismo Excel. Compartido en vivo entre todos los que abran la web.
 *
 * USO (1 sola vez):
 *  1. https://script.google.com → "Nuevo proyecto"
 *  2. Borra el código de ejemplo y pega TODO este archivo
 *  3. Guarda (Ctrl+S), ponle nombre "Backend Progreso"
 *  4. "Implementar" → "Nueva implementación"
 *  5. Icono engranaje ⚙ → "Aplicación web"
 *  6. Configura:
 *       • Ejecutar como: "Yo"
 *       • Quién tiene acceso: "Cualquiera"
 *  7. "Implementar" → acepta permisos
 *  8. Copia "URL de la aplicación web" (termina en /exec)
 *  9. Pásamela y la conecto a objetivos.html.
 *
 * Después puedes hacer la hoja del Excel privada otra vez:
 * este script tiene su propio acceso porque corre con tu cuenta.
 */

const SPREADSHEET_ID = '1EtrDQwfxlRcdGGo_eU4zTHYE4Ih6LDxpzJ5sr-c_bHI';
const PROGRESS_SHEET_NAME = 'Progreso';

function getOrCreateSheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(PROGRESS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(PROGRESS_SHEET_NAME);
    sheet.appendRow(['Key', 'Value', 'Updated']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const progress = {};
    for (let i = 1; i < data.length; i++) {
      const key = String(data[i][0] || '').trim();
      const value = Number(data[i][1]);
      if (key && !isNaN(value)) progress[key] = value;
    }
    return jsonResponse({ progress: progress, count: Object.keys(progress).length });
  } catch (err) {
    return jsonResponse({ error: String(err) });
  }
}

function doPost(e) {
  try {
    let body;
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else {
      body = e.parameter;
    }
    const updates = body.updates || [];
    if (!Array.isArray(updates) || !updates.length) {
      return jsonResponse({ error: 'No hay updates' });
    }

    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const keyToRow = {};
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][0] || '').trim();
      if (k) keyToRow[k] = i + 1;
    }

    const now = new Date().toISOString();
    const newRows = [];
    let updated = 0;

    for (const u of updates) {
      const k = String(u.key || '').trim();
      const v = Number(u.value);
      if (!k || isNaN(v)) continue;
      if (keyToRow[k]) {
        sheet.getRange(keyToRow[k], 2, 1, 2).setValues([[v, now]]);
      } else {
        newRows.push([k, v, now]);
      }
      updated++;
    }
    if (newRows.length) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow + 1, 1, newRows.length, 3).setValues(newRows);
    }
    return jsonResponse({ saved: updated });
  } catch (err) {
    return jsonResponse({ error: String(err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
