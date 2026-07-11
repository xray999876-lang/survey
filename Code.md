/**
 * Backend สำหรับแบบสำรวจความพึงพอใจ - กลุ่มงานรังสีวิทยา โรงพยาบาลโพนทราย
 *
 * วิธีใช้:
 * 1. เปิด Google Sheet ที่จะใช้เก็บข้อมูล (สร้างใหม่ก็ได้)
 * 2. เมนู Extensions > Apps Script
 * 3. ลบโค้ดเดิมทั้งหมด แล้ววางโค้ดนี้แทน
 * 4. เปลี่ยนค่า SHARED_SECRET ด้านล่างเป็นข้อความลับของคุณเอง (ต้องตรงกับใน index.html)
 * 5. กด Deploy > New deployment
 *    - Select type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. คัดลอก Web app URL ที่ได้ ไปใส่ในตัวแปร SCRIPT_URL ของไฟล์ index.html
 *
 * หมายเหตุความปลอดภัย:
 * - URL ของ Web app นี้จำเป็นต้องเปิดเผยในหน้าเว็บอยู่แล้ว (เป็นธรรมชาติของ Apps Script)
 *   SHARED_SECRET ด้านล่างช่วยกันไม่ให้คนอื่นยิง POST ปลอมเข้ามาสแปมชีตได้ง่าย ๆ
 * - ถ้า URL เดิมเคยหลุดไปแล้ว (เช่นเคยขึ้น public repo/commit history) แนะนำให้สร้าง
 *   deployment ใหม่ (จะได้ URL ใหม่) แล้วเลิกใช้ของเก่า
 */

var SHARED_SECRET = "ptr-xray-survey-8f3k2m9qL7"; // ต้องตรงกับ SUBMIT_TOKEN ใน index.html

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);

    // ตรวจ secret token ก่อนบันทึกทุกครั้ง
    if (d._token !== SHARED_SECRET) {
      return ContentService
        .createTextOutput(JSON.stringify({ status: "error", message: "invalid token" }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var sheet = getResponseSheet_();

    sheet.appendRow([
      new Date(),
      safe_(d.serviceType),
      safe_(d.gender),
      safe_(d.ageRange),
      safe_(d.education),
      safe_(d.personType),
      safe_(d.process_1), safe_(d.process_2), safe_(d.process_3), safe_(d.process_4),
      safe_(d.staff_1), safe_(d.staff_2), safe_(d.staff_3), safe_(d.staff_4), safe_(d.staff_5),
      safe_(d.facility_1), safe_(d.facility_2), safe_(d.facility_3), safe_(d.facility_4), safe_(d.facility_5),
      safe_(d.comment1), safe_(d.comment2), safe_(d.comment3)
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput("Survey backend is running.")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * ป้องกัน Formula Injection: ถ้าข้อความขึ้นต้นด้วย = + - @
 * Google Sheets จะตีความเป็นสูตร จึงเติม ' นำหน้าเพื่อบังคับให้เป็นข้อความล้วน
 */
function safe_(value) {
  var v = (value === undefined || value === null) ? "" : String(value);
  if (/^[=+\-@]/.test(v)) {
    return "'" + v;
  }
  return v;
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
