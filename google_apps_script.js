// ═══════════════════════════════════════════════════════════════
//  Google Apps Script — RSVP для Лианы (Қыз Ұзату 15.08.2026)
//  Вставьте этот код в Apps Script и задеплойте как Web App
// ═══════════════════════════════════════════════════════════════

const SHEET_NAME = "RSVP";

// ─── Принять ответ от гостя (POST из index.html) ───
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getOrCreateSheet();

    sheet.appendRow([
      sheet.getLastRow(),           // № строки
      data.name  || "",             // Имя
      data.att   || "",             // Статус (Приду / Не смогу / С супругой)
      data.wish  || "",             // Пожелание
      new Date().toLocaleString("ru-RU", { timeZone: "Asia/Almaty" })  // Время
    ]);

    return jsonResponse({ ok: true });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ─── Отдать все ответы (GET из admin.html) ───
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const rows  = sheet.getDataRange().getValues();

    // Пропускаем заголовок (первая строка)
    const data = rows.slice(1).map(r => ({
      num:  r[0],
      name: r[1],
      att:  r[2],
      wish: r[3],
      time: r[4]
    }));

    return jsonResponse(data);
  } catch(err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ─── Вспомогательные функции ───
function getOrCreateSheet() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Заголовки
    sheet.appendRow(["№", "Имя", "Статус", "Пожелание", "Время"]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold").setBackground("#8b1a1a").setFontColor("#ffffff");
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
