/**
 * Crevy - Encuesta de clientes -> Google Sheets
 *
 * Cómo usarlo:
 *  1. Abrí una Google Sheet nueva (o usá una existente).
 *  2. Extensiones -> Apps Script.
 *  3. Pegá TODO este archivo en el editor (reemplazando lo que haya).
 *  4. Guardá (Ctrl+S) y ponele un nombre al proyecto.
 *  5. Implementar -> Nueva implementación -> Tipo: "Aplicación web".
 *       - Ejecutar como: "Yo".
 *       - Quién tiene acceso: "Cualquier persona" (es necesario para que la web pueda postear).
 *     Copiá la URL que te da (termina en /exec).
 *  6. En index.html, pegá esa URL dentro de:
 *        const WEBHOOK_URL = "<<acá>>";
 *  7. Listo. Cada respuesta cae como fila nueva en la pestaña "Respuestas".
 *
 * Si más adelante cambiás el código y querés que los cambios surtan efecto,
 * tenés que hacer "Implementar -> Administrar implementaciones -> Editar (lápiz)
 * -> Versión: Nueva versión -> Implementar". Si solo guardás, no se actualiza.
 */

const SHEET_NAME = "Respuestas";

// Orden de columnas. Si agregás campos en la encuesta, agregalos también acá.
const COLS = [
  "fecha",
  "nombre",
  "empresa",
  "email",
  "rubro",
  "meses",
  "horas_antes",
  "horas_despues",
  "horas_ahorradas_semana",
  "ventas_antes",
  "ventas_despues",
  "delta_ventas",
  "ventas_atribuidas",
  "pct_ventas",
  "ticket",
  "facturacion_atribuida",
  "ahorro_dinero",
  "tareas",
  "impacto",
  "nps",
  "permiso_testimonio",
  "testimonio",
  "dolor_antes",
  "cambio",
  "extra"
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(COLS);
      sheet.getRange(1, 1, 1, COLS.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }

    const row = COLS.map(c => {
      const v = data[c];
      if (Array.isArray(v)) return v.join("; ");
      if (v === undefined || v === null) return "";
      return v;
    });
    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Ping rápido para verificar que la URL responde
function doGet() {
  return ContentService
    .createTextOutput("Crevy webhook OK")
    .setMimeType(ContentService.MimeType.TEXT);
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}
