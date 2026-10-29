// Приём заявок с сайта в Google Таблицу.
// Развёртывание: см. README в конце файла index.html-инструкции (или сообщение в чате).
const SHEET_ID = "1TJpoUq8VDNQuF1ms0fSUfazlPHnkpyGuoVEabrkHwCU";

function doPost(e) {
  const p = (e && e.parameter) || {};
  if (p.website) return json({ ok: true }); // honeypot: боты заполняют скрытое поле

  const name = String(p.name || "").trim().slice(0, 100);
  const phone = String(p.phone || "").trim().slice(0, 30);
  if (!name || !phone) return json({ ok: false, error: "empty" });

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Дата", "Имя", "Телефон"]);
  }
  // Апостроф-префикс не нужен: setNumberFormat("@") сохраняет телефон как текст
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 3).setValues([[new Date(), name, phone]]);
  sheet.getRange(row, 3).setNumberFormat("@").setValue(phone);
  return json({ ok: true });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON,
  );
}
