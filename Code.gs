// ==========================================
// Google Apps Script — บันทึกรายรับรายจ่าย
// วางโค้ดนี้ใน Google Apps Script แล้ว Deploy เป็น Web App
// ==========================================

const SHEET_ID = "13O1b3pB70s7TDXWgfLfax77yoPKbkAEGkuHt7Yenep0"; // ← ใส่ ID ของ Google Sheet ตรงนี้
// จาก URL ของ Google Sheet
// 'https://docs.google.com/spreadsheets/d/XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX/edit
// ให้เอาเฉพาะส่วนนี้ XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX มาใส่

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SHEET_ID);
    writeData(ss, data);
    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const action = e.parameter.action;

    if (action === "getAll") {
      const expenses  = getSheetData(ss, "รายจ่าย");
      const incomes   = getSheetData(ss, "รายรับ");
      const transfers = getSheetData(ss, "โอนเงิน");
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, expenses, incomes, transfers }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === "write") {
      const data = JSON.parse(e.parameter.payload);
      writeData(ss, data);
      return ContentService
        .createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: "Unknown action" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function writeData(ss, data) {
  const type = data.type;
  if (type === "expense") {
    const sheet = ss.getSheetByName("รายจ่าย") || ss.insertSheet("รายจ่าย");
    ensureExpenseHeader(sheet);
    appendByHeaders(sheet, {
      Date: data.date,
      Category: data.category,
      Type: data.txType,
      Detail: data.detail,
      Account: data.account,
      Amount: Number(data.amount),
      Note: data.note || ""
    });
  } else if (type === "income") {
    const sheet = ss.getSheetByName("รายรับ") || ss.insertSheet("รายรับ");
    ensureIncomeHeader(sheet);
    appendByHeaders(sheet, {
      Date: data.date,
      Category: data.category,
      Type: data.txType,
      Detail: data.detail,
      Account: data.account,
      Amount: Number(data.amount)
    });
  } else if (type === "transfer") {
    const sheet = ss.getSheetByName("โอนเงิน") || ss.insertSheet("โอนเงิน");
    ensureTransferHeader(sheet);
    appendByHeaders(sheet, {
      Date: data.date,
      Category: data.category || "โอนระหว่างธนาคาร",
      "From account": data.fromAccount,
      Detail: "to",
      "To account": data.toAccount,
      Amount: Number(data.amount)
    });
  }
}

function getSheetData(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return [];

  const headers = rows[0];
  const tz = 'Asia/Bangkok';

  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      const v = row[i];
      obj[h] = (v instanceof Date) ? Utilities.formatDate(v, tz, "yyyy-MM-dd") : v;
    });
    return obj;
  });
}

// เขียน row โดยจับคู่กับ header ชื่อแทนตำแหน่ง — ป้องกัน column order ต่างกัน
function appendByHeaders(sheet, dataObj) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = headers.map(h => (dataObj[h] !== undefined ? dataObj[h] : ''));
  sheet.appendRow(row);
}

function ensureExpenseHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Category", "Type", "Detail", "Account", "Amount", "Note"]);
  }
}

function ensureIncomeHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Category", "Type", "Detail", "Account", "Amount"]);
  }
}

function ensureTransferHeader(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Date", "Category", "From account", "Detail", "To account", "Amount"]);
  }
}
