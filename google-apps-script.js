/**
 * ═══════════════════════════════════════════════════════
 *  GESTIÓN APP v3 — Google Apps Script
 *  Pegar en: Extensiones > Apps Script > Reemplazar todo
 *  Ejecutar initSheets() una vez para crear las hojas
 * ═══════════════════════════════════════════════════════
 */

const SHEETS = {
  Clientes:     ["id","codigo","nombre","localidad"],
  Deudores:     ["id","codigo","nombre","localidad","saldo"],
  Comprobantes: ["id","codigo","cliente","fecha","comprobante","importe"],
  Cobranzas:    ["id","codigo","cliente","monto","fechaVencimiento","estado","cobrador","notas"],
  Visitas:      ["id","codigo","cliente","direccion","fecha","hora","estado","cobrador","notas"],
};

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  try {
    const action = e.parameter.action;
    const sheetName = e.parameter.sheet;
    if (action === "getData" && sheetName) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return jsonResponse({ success: false, error: "Hoja no encontrada: " + sheetName });
      const cols = SHEETS[sheetName] || [];
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return jsonResponse({ success: true, data: [] });
      const values = sheet.getRange(2, 1, lastRow - 1, cols.length).getValues();
      const data = values.filter(row => row[0] !== "").map(row => {
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

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const { action, sheet: sheetName, id, data, rows } = body;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // ── bulkUpsert: insert or update by codigo ──
    if (action === "bulkUpsert") {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return jsonResponse({ success: false, error: "Hoja no encontrada: " + sheetName });
      const cols = SHEETS[sheetName] || [];
      const codigoIdx = cols.indexOf("codigo");

      let added = 0, updated = 0;
      const lastRow = sheet.getLastRow();
      let existing = {};

      if (lastRow > 1) {
        const allData = sheet.getRange(2, 1, lastRow - 1, cols.length).getValues();
        allData.forEach((row, i) => {
          if (row[0] !== "" && codigoIdx >= 0) {
            existing[String(row[codigoIdx])] = i + 2;
          }
        });
      }

      rows.forEach(rowData => {
        const codigo = String(rowData.codigo || "");
        const rowValues = cols.map(col => col === "id" ? (rowData.id || Utilities.getUuid()) : (rowData[col] ?? ""));
        if (codigoIdx >= 0 && codigo && existing[codigo]) {
          sheet.getRange(existing[codigo], 1, 1, cols.length).setValues([rowValues]);
          updated++;
        } else {
          if (!rowValues[0]) rowValues[0] = Utilities.getUuid();
          sheet.appendRow(rowValues);
          added++;
        }
      });

      return jsonResponse({ success: true, added, updated });
    }

    // ── clearAndInsert: replace all rows ──
    if (action === "clearAndInsert") {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return jsonResponse({ success: false, error: "Hoja no encontrada: " + sheetName });
      const cols = SHEETS[sheetName] || [];
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      rows.forEach(rowData => {
        const rowValues = cols.map(col => col === "id" ? Utilities.getUuid() : (rowData[col] ?? ""));
        sheet.appendRow(rowValues);
      });
      return jsonResponse({ success: true, inserted: rows.length });
    }

    // ── standard CRUD ──
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return jsonResponse({ success: false, error: "Hoja no encontrada: " + sheetName });
    const cols = SHEETS[sheetName] || [];

    if (action === "addRow") {
      const newId = Utilities.getUuid();
      sheet.appendRow(cols.map(col => col === "id" ? newId : (data[col] ?? "")));
      return jsonResponse({ success: true, id: newId });
    }

    if (action === "updateRow" || action === "deleteRow") {
      if (!id) return jsonResponse({ success: false, error: "ID requerido" });
      const lastRow = sheet.getLastRow();
      if (lastRow <= 1) return jsonResponse({ success: false, error: "Sin datos" });
      const idCol = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      let targetRow = -1;
      idCol.forEach((r, i) => { if (String(r[0]) === String(id)) targetRow = i + 2; });
      if (targetRow === -1) return jsonResponse({ success: false, error: "Registro no encontrado" });
      if (action === "deleteRow") { sheet.deleteRow(targetRow); return jsonResponse({ success: true }); }
      sheet.getRange(targetRow, 1, 1, cols.length).setValues([cols.map(col => col === "id" ? id : (data[col] ?? ""))]);
      return jsonResponse({ success: true });
    }

    return jsonResponse({ success: false, error: "Acción desconocida" });
  } catch (err) {
    return jsonResponse({ success: false, error: err.message });
  }
}

function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.entries(SHEETS).forEach(([name, cols]) => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, cols.length).setValues([cols]);
      sheet.getRange(1, 1, 1, cols.length).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
  });
  SpreadsheetApp.getUi().alert("✅ Hojas creadas: Clientes, Deudores, Comprobantes, Cobranzas, Visitas");
}
