/**
 * Planning & Analytics — Proxy de respuestas (Apps Script Web App)
 *
 * Mantén tu hoja PRIVADA. Este script corre con TU cuenta y sirve un JSON
 * con las respuestas. La web (plan-accion.html) llama a este endpoint,
 * no directamente al spreadsheet.
 *
 * USO (1 sola vez):
 *  1. https://script.google.com → "Nuevo proyecto"
 *  2. Borra el código de ejemplo y pega TODO este archivo
 *  3. Guarda (Ctrl+S) y ponle nombre (ej. "Proxy Respuestas")
 *  4. Botón "Implementar" (arriba derecha) → "Nueva implementación"
 *  5. Icono engranaje ⚙ → "Aplicación web"
 *  6. Configura:
 *       • Descripción: cualquier cosa
 *       • Ejecutar como: "Yo" (tu cuenta)
 *       • Quién tiene acceso: "Cualquiera"
 *  7. "Implementar" → acepta permisos
 *  8. Copia "URL de la aplicación web" (termina en /exec)
 *  9. Pásamela y la cableo a la web.
 *
 * Vuelve a desplegar (paso 4-8) cada vez que cambies este script.
 */

const SPREADSHEET_ID = '1EtrDQwfxlRcdGGo_eU4zTHYE4Ih6LDxpzJ5sr-c_bHI';
const DEFAULT_GID = 474424277; // pestaña de respuestas del form Objetivos+Canción

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const gid = params.gid ? parseInt(params.gid, 10) : DEFAULT_GID;

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets().find(s => s.getSheetId() === gid);
    if (!sheet) {
      return jsonResponse({ error: 'Pestaña no encontrada para gid ' + gid });
    }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return jsonResponse({ responses: [], count: 0 });
    }

    const headers = data[0].map(h => String(h).trim());
    const responses = data.slice(1)
      .filter(row => row.some(c => c !== '' && c !== null))
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          const val = row[i];
          obj[h] = val != null ? String(val).trim() : '';
        });
        return obj;
      });

    return jsonResponse({
      count: responses.length,
      responses: responses
    });
  } catch (err) {
    return jsonResponse({ error: String(err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
