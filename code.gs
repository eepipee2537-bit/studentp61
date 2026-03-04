const SHEET_ID = '11q_1hyXzz_JPjrnsLOKGMZC05xNct67qq9GV9ButzI8';

// ฟังก์ชันสำหรับเปิดหน้าเว็บ HTML
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('ระบบเช็คชื่อนักเรียน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ฟังก์ชันดึงรายชื่อนักเรียนจากแผ่นงาน 'name'
function getStudentNames() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('name');
  // สมมติว่ารายชื่อเริ่มที่ A2 (A1 เป็นหัวข้อ) ถ้าไม่มีหัวข้อให้เปลี่ยนเป็น A1:A
  const data = sheet.getRange('A2:A').getValues(); 
  return data.flat().filter(String); // ตัดแถวที่ว่างออก
}

// ฟังก์ชันบันทึกข้อมูลการเช็คชื่อลงแผ่นงาน 'check in'
function saveAttendance(recordData) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName('check in');
  
  // จัดรูปแบบวันที่ปัจจุบัน (เช่น 04/03/2026)
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  // อ่านข้อมูลหัวคอลัมน์แถวที่ 1 ทั้งหมดเพื่อหาวันที่
  let headers = [];
  if (sheet.getLastColumn() > 0) {
    headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  
  // ตรวจสอบว่ามีคอลัมน์ของวันนี้หรือยัง ถ้ายังให้สร้างใหม่
  let dateColIndex = headers.indexOf(today) + 1;
  if (dateColIndex === 0) {
    dateColIndex = (sheet.getLastColumn() || 1) + 1;
    sheet.getRange(1, dateColIndex).setValue(today);
  }
  
  // อ่านรายชื่อนักเรียนในคอลัมน์ A ของแผ่นงาน 'check in'
  let numRows = sheet.getLastRow();
  let nameList = [];
  if (numRows > 1) {
    nameList = sheet.getRange(2, 1, numRows - 1, 1).getValues().flat();
  }
  
  // วนลูปบันทึกข้อมูลทีละคน
  recordData.forEach(item => {
    let rowIndex = nameList.indexOf(item.name);
    if (rowIndex === -1) {
      // ถ้ารายชื่อนี้ยังไม่มีในคอลัมน์ A ให้เพิ่มชื่อลงไปก่อน
      sheet.getRange(numRows + 1, 1).setValue(item.name);
      sheet.getRange(numRows + 1, dateColIndex).setValue(item.status);
      nameList.push(item.name); // อัปเดตรายการชื่อในหน่วยความจำ
      numRows++;
    } else {
      // ถ้ามีชื่ออยู่แล้ว ให้บันทึกสถานะในแถวของคนนั้น และคอลัมน์ของวันนี้ (บวก 2 เพราะ index เริ่มที่ 0 และข้อมูลเริ่มแถว 2)
      sheet.getRange(rowIndex + 2, dateColIndex).setValue(item.status);
    }
  });
  
  return "บันทึกข้อมูลการเช็คชื่อเรียบร้อยแล้ว!";
}
