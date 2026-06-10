/**
 * Google Apps Script สำหรับ Stock Checklist
 *
 * วิธีใช้งาน:
 * 1. ไปที่ Google Sheets > Extensions > Apps Script
 * 2. ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดนี้
 * 3. กด Deploy > New deployment
 * 4. เลือก Type: Web app
 * 5. Execute as: Me
 * 6. Who has access: Anyone
 * 7. กด Deploy แล้วคัดลอก URL ไปใส่ในไฟล์ index.html (SHEET_URL)
 *
 * โครงสร้าง Google Sheet:
 * Row 1: Header (date, checker, userId, picture, ok, warn, warnItems)
 * Row 2+: ข้อมูล
 */

// ดึงข้อมูลทั้งหมด
function doGet(e) {
  const action = e.parameter.action;

  if (action === 'get') {
    return getHistory();
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// เพิ่ม/แก้ไข/ลบข้อมูล
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;

    if (action === 'update') {
      return updateHistory(data);
    } else if (action === 'delete') {
      return deleteHistory(data);
    } else {
      // Default: เพิ่มข้อมูลใหม่
      return addHistory(data);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ดึงประวัติทั้งหมด
function getHistory() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    return ContentService.createTextOutput(JSON.stringify([]))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  // ตรวจสอบจำนวนคอลัมน์เพื่อรองรับทั้งข้อมูลเก่าและใหม่
  const result = data.map((row, index) => {
    // รูปแบบใหม่ 7 คอลัมน์: date, checker, userId, picture, ok, warn, warnItems
    if (lastCol >= 7) {
      return {
        id: index + 2,
        date: formatDateForOutput(row[0]),
        checker: row[1] || '-',
        userId: row[2] || '-',
        picture: row[3] || '-',
        ok: parseInt(row[4]) || 0,
        warn: parseInt(row[5]) || 0,
        warnItems: row[6] || '-'
      };
    }
    // รูปแบบเก่า 5 คอลัมน์: date, checker, ok, warn, warnItems
    else {
      return {
        id: index + 2,
        date: formatDateForOutput(row[0]),
        checker: row[1] || '-',
        userId: '-',
        picture: '-',
        ok: parseInt(row[2]) || 0,
        warn: parseInt(row[3]) || 0,
        warnItems: row[4] || '-'
      };
    }
  });

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// แปลงวันที่เป็น string ก่อนส่ง
function formatDateForOutput(dateValue) {
  if (!dateValue) return '-';

  // ถ้าเป็น Date object
  if (dateValue instanceof Date) {
    const day = dateValue.getDate();
    const month = dateValue.getMonth();
    let year = dateValue.getFullYear();

    // แปลงเป็น พ.ศ. เฉพาะถ้าเป็นปี ค.ศ. (< 2500)
    if (year < 2500) {
      year = year + 543;
    }

    const months = ['ม.ค.', 'ก.พ.', 'มี.ค.', 'เม.ย.', 'พ.ค.', 'มิ.ย.', 'ก.ค.', 'ส.ค.', 'ก.ย.', 'ต.ค.', 'พ.ย.', 'ธ.ค.'];
    return day + ' ' + months[month] + ' ' + year;
  }

  // ถ้าเป็น string ตรวจสอบว่ามีปีผิดปกติไหม (เช่น 2569)
  const str = String(dateValue);

  // ถ้าเป็น string วันที่ไทยที่มีปี > 2500 แล้ว ให้ return เลย
  if (/\d{1,2}\s+[ก-ฮ]+\.?\s+\d{4}/.test(str)) {
    return str;
  }

  return str;
}

// เพิ่มประวัติใหม่
function addHistory(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // ถ้ายังไม่มี header ให้สร้าง
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['date', 'checker', 'userId', 'picture', 'ok', 'warn', 'warnItems']);
  }

  sheet.appendRow([
    data.date,
    data.checker,
    data.userId || '-',
    data.picture || '-',
    data.ok,
    data.warn,
    data.warnItems
  ]);

  return ContentService.createTextOutput(JSON.stringify({ success: true, action: 'add' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// แก้ไขประวัติ
function updateHistory(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = data.row + 1; // +1 because row 1 is header

  // ตรวจสอบว่า row อยู่ในช่วงที่ถูกต้อง
  if (row < 2 || row > sheet.getLastRow()) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid row number' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // อัปเดตข้อมูล
  sheet.getRange(row, 1).setValue(data.date);
  sheet.getRange(row, 2).setValue(data.checker);
  // คอลัมน์ 3-4 (userId, picture) ไม่ต้องแก้ไข เพราะเป็นข้อมูลตายตัว
  sheet.getRange(row, 5).setValue(data.ok);
  sheet.getRange(row, 6).setValue(data.warn);
  sheet.getRange(row, 7).setValue(data.warnItems);

  return ContentService.createTextOutput(JSON.stringify({ success: true, action: 'update', row: row }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ลบประวัติ
function deleteHistory(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = data.row + 1; // +1 because row 1 is header

  // ตรวจสอบว่า row อยู่ในช่วงที่ถูกต้อง
  if (row < 2 || row > sheet.getLastRow()) {
    return ContentService.createTextOutput(JSON.stringify({ error: 'Invalid row number' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // ลบแถว
  sheet.deleteRow(row);

  return ContentService.createTextOutput(JSON.stringify({ success: true, action: 'delete', row: row }))
    .setMimeType(ContentService.MimeType.JSON);
}
