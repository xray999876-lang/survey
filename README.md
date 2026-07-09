/**
 * Backend สำหรับแบบสำรวจความพึงพอใจ - กลุ่มงานรังสีวิทยา โรงพยาบาลโพนทราย
 *
 * วิธีใช้:
 * 1. เปิด Google Sheet ที่จะใช้เก็บข้อมูล (สร้างใหม่ก็ได้)
 * 2. เมนู Extensions > Apps Script
 * 3. ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดนี้แทน
 * 4. กด Deploy > New deployment
 *    - Select type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 5. คัดลอก Web app URL ที่ได้ ไปใส่ในตัวแปร SCRIPT_URL ของไฟล์ survey.html
 */

function doPost(e) {
  var sheet = getResponseSheet_();
  var d = JSON.parse(e.postData.contents);

  sheet.appendRow([
    new Date(),
    d.serviceType || "",
    d.gender || "",
    d.ageRange || "",
    d.education || "",
    d.personType || "",
    d.process_1 || "", d.process_2 || "", d.process_3 || "", d.process_4 || "",
    d.staff_1 || "", d.staff_2 || "", d.staff_3 || "", d.staff_4 || "", d.staff_5 || "",
    d.facility_1 || "", d.facility_2 || "", d.facility_3 || "", d.facility_4 || "", d.facility_5 || "",
    d.comment1 || "", d.comment2 || "", d.comment3 || ""
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput("Survey backend is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}

function getResponseSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Responses");
  if (!sheet) {
    sheet = ss.insertSheet("Responses");
    sheet.appendRow([
      "เวลาที่ส่ง",
      "ผู้รับบริการ", "เพศ", "อายุ", "ระดับการศึกษา", "ประเภทบุคคล",
      "กระบวนการ 1", "กระบวนการ 2", "กระบวนการ 3", "กระบวนการ 4",
      "เจ้าหน้าที่ 1", "เจ้าหน้าที่ 2", "เจ้าหน้าที่ 3", "เจ้าหน้าที่ 4", "เจ้าหน้าที่ 5",
      "สิ่งอำนวยความสะดวก 1", "สิ่งอำนวยความสะดวก 2", "สิ่งอำนวยความสะดวก 3", "สิ่งอำนวยความสะดวก 4", "สิ่งอำนวยความสะดวก 5",
      "ข้อเสนอแนะ: หน่วยงาน", "ข้อเสนอแนะ: เจ้าหน้าที่", "ข้อคิดเห็นอื่นๆ"
    ]);
    sheet.getRange(1, 1, 1, 23).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}
