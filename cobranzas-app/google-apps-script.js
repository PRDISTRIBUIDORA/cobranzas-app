/**
 * ═══════════════════════════════════════════════════════
 *  GESTIÓN APP — Google Apps Script
 *  Pegar en: Extensiones > Apps Script > Reemplazar todo
 *  Luego: Implementar > Nueva implementación > Aplicación web
 *         - Ejecutar como: Yo
 *         - Quién puede acceder: Cualquier usuario
 * ═══════════════════════════════════════════════════════
 *
 *  Estructura esperada en Google Sheets:
 *  Hoja "Cobranzas": id | cliente | monto | fechaVencimiento | estado | notas
 *  Hoja "Visitas":   id | cliente | direccion | fecha | hora | estado | notas
 */

const SHEETS = {
  Cobranzas: ["id","cliente","monto","fechaVencimiento","estado","notas"],
  Visitas:   ["id","cliente","direccion","fecha","hora","estado","notas"],
};

/* ── Respuesta JSON con cabeceras CORS ── */
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ── GET: getData?sheet=Cobranzas ── */
function doGet(e) {
  try {
    const action = e.parameter.action;
    const sheetName = e.parameter.sheet;

    if (action === "getData" && sheetName) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return jsonResponse({ success: false, error: `Hoja "${sheetName}" no encontrada` });

      const cols = SHEETS[sheetName] || [];
      const lastRow = sheet.getLastRow();

      if (lastRow <= 1) return jsonResponse({ success: true, data: [] });

      const range = sheet.getRange(2, 1, lastRow - 1, cols.length);
      const values = range.getValues();

      const data = values
        .filter(row => row[0] !== "")
        .map(row => {
          const obj = {};
          cols.forEach((col, i) => { obj[col] = row[i]; });
          return obj;
        });

      return jsonResponse({ success: true, data });
    }

    return jsonResponse({ success: false, error: "Acción no válida" });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

/* ── POST: addRow | updateRow | deleteRow ── */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, sheet: sheetName, id, data } = body;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return jsonResponse({ success: false, error: `Hoja "${sheetName}" no encontrada` });

    const cols = SHEETS[sheetName] || [];

    /* ── addRow ── */
    if (action === "addRow") {
      const newId = Utilities.getUuid();
      const row = cols.map(col => col === "id" ? newId : (data[col] ?? ""));
      sheet.appendRow(row);
      return jsonResponse({ success: true, id: newId });
    }

    /* ── updateRow / deleteRow ── */
    if (action === "updateRow" || action === "deleteRow") {
      if (!id) return jsonResponse({ success: false, error: "ID requerido" });

      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return jsonResponse({ success: false, error: "Sin datos" });

      const idCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      let targetRow = -1;
      idCol.forEach((r, i) => { if (String(r[0]) === String(id)) targetRow = i + 2; });

      if (targetRow === -1) return jsonResponse({ success: false, error: "Registro no encontrado" });

      if (action === "deleteRow") {
        sheet.deleteRow(targetRow);
        return jsonResponse({ success: true });
      }

      if (action === "updateRow") {
        const updatedRow = cols.map(col => col === "id" ? id : (data[col] ?? ""));
        sheet.getRange(targetRow, 1, 1, cols.length).setValues([updatedRow]);
        return jsonResponse({ success: true });
      }
    }

    return jsonResponse({ success: false, error: `Acción "${action}" desconocida` });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

/* ── Función de inicialización (opcional) ── */
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.entries(SHEETS).forEach(([name, cols]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) { sheet = ss.insertSheet(name); }
    // Escribir encabezados solo si la hoja está vacía
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
      sheet.getRange(1, 1, 1, cols.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
  });
  SpreadsheetApp.getUi().alert("✅ Hojas inicializadas correctamente");
}
