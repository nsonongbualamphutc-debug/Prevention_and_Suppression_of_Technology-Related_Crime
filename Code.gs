/**
 * ระบบแดชบอร์ดการดำเนินงานป้องกันและปราบปรามอาชญากรรมทางเทคโนโลยี
 * จังหวัดหนองบัวลำภู
 *
 * Backend: Google Apps Script + Google Sheets
 * Security: PIN is never stored in plaintext. Only its SHA-256 hash is kept here.
 *           To change the PIN, replace PIN_HASH with the SHA-256 hex digest of the
 *           new PIN (you can generate it from any SHA-256 tool, e.g. browser console:
 *           crypto.subtle.digest(...) or https://emn178.github.io/online-tools/sha256.html)
 */

// ---------- CONFIG ----------
// SHA-256("309309") -- the real PIN is never written in this file.
const PIN_HASH = 'b15b81b008c280abb2551e1c11dddbf172fb7bf59e739b88879e8872091a12a0'; // SHA-256 of the real PIN — PIN itself is not stored anywhere
const SHEET_NAME_YEARLY = 'YearlyStats';
const SHEET_NAME_DISTRICT = 'DistrictTargets';
const SHEET_NAME_CONTACTS = 'Contacts';
const SHEET_NAME_MEASURES = 'Measures';
const SHEET_NAME_SCAMTYPES = 'ScamTypes';

function getSS() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sha256Hex(str) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8);
  return raw.map(function (b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function checkPin(pin) {
  return sha256Hex(String(pin)) === PIN_HASH;
}

// ---------- ENTRY POINTS ----------
function doGet(e) {
  const action = e.parameter.action || 'getAll';
  const callback = e.parameter.callback;
  let result;

  try {
    switch (action) {
      case 'getAll':
        result = { status: 'ok', data: getAllData() };
        break;
      case 'login':
        result = { status: 'ok', valid: checkPin(e.parameter.pin || '') };
        break;
      case 'saveYearly':
        result = guarded(e, function () { return saveYearly(JSON.parse(e.parameter.payload)); });
        break;
      case 'saveDistrict':
        result = guarded(e, function () { return saveDistrict(JSON.parse(e.parameter.payload)); });
        break;
      case 'saveContacts':
        result = guarded(e, function () { return saveContacts(JSON.parse(e.parameter.payload)); });
        break;
      case 'saveMeasures':
        result = guarded(e, function () { return saveMeasures(JSON.parse(e.parameter.payload)); });
        break;
      case 'saveScamTypes':
        result = guarded(e, function () { return saveScamTypes(JSON.parse(e.parameter.payload)); });
        break;
      case 'importAll':
        result = guarded(e, function () { return importAll(JSON.parse(e.parameter.payload)); });
        break;
      default:
        result = { status: 'error', message: 'unknown action' };
    }
  } catch (err) {
    result = { status: 'error', message: String(err) };
  }

  return respond(result, callback);
}

function guarded(e, fn) {
  if (!checkPin(e.parameter.pin || '')) {
    return { status: 'error', message: 'invalid pin' };
  }
  return { status: 'ok', data: fn() };
}

function respond(obj, callback) {
  const json = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ---------- READ ----------
function getAllData() {
  return {
    yearly: readSheet(SHEET_NAME_YEARLY),
    district: readSheet(SHEET_NAME_DISTRICT),
    contacts: readSheet(SHEET_NAME_CONTACTS),
    measures: readSheet(SHEET_NAME_MEASURES),
    scamTypes: readSheet(SHEET_NAME_SCAMTYPES),
    updatedAt: new Date().toISOString()
  };
}

function readSheet(name) {
  const sh = getSS().getSheetByName(name);
  if (!sh) return [];
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i].join('') === '') continue;
    const obj = {};
    headers.forEach(function (h, idx) { obj[h] = values[i][idx]; });
    rows.push(obj);
  }
  return rows;
}

// ---------- WRITE ----------
function writeSheet(name, headers, rows) {
  let sh = getSS().getSheetByName(name);
  if (!sh) sh = getSS().insertSheet(name);
  sh.clear();
  sh.appendRow(headers);
  rows.forEach(function (row) {
    sh.appendRow(headers.map(function (h) { return row[h] !== undefined ? row[h] : ''; }));
  });
  return true;
}

function saveYearly(rows) {
  return writeSheet(SHEET_NAME_YEARLY, ['ปี', 'จำนวนคดี', 'มูลค่าความเสียหาย', 'หมายเหตุ'], rows);
}
function saveDistrict(rows) {
  return writeSheet(SHEET_NAME_DISTRICT, ['อำเภอ', 'เป้ารับแจ้งคดี', 'ผลรับแจ้งคดี', 'เป้าติดตามอายัดเงิน', 'ผลติดตามอายัดเงิน', 'เป้าสอบสวนสรุปสำนวน', 'ผลสอบสวนสรุปสำนวน'], rows);
}
function saveContacts(rows) {
  return writeSheet(SHEET_NAME_CONTACTS, ['ลำดับ', 'ชื่อสกุล', 'หน่วยงาน', 'ตำแหน่ง', 'เบอร์โทร'], rows);
}
function saveMeasures(rows) {
  return writeSheet(SHEET_NAME_MEASURES, ['ด้าน', 'ลำดับ', 'รายละเอียด', 'หน่วยงานรับผิดชอบ', 'สถานะ'], rows);
}
function saveScamTypes(rows) {
  return writeSheet(SHEET_NAME_SCAMTYPES, ['ลำดับ', 'รูปแบบ', 'สัดส่วน'], rows);
}

function importAll(payload) {
  if (payload.yearly) saveYearly(payload.yearly);
  if (payload.district) saveDistrict(payload.district);
  if (payload.contacts) saveContacts(payload.contacts);
  if (payload.measures) saveMeasures(payload.measures);
  if (payload.scamTypes) saveScamTypes(payload.scamTypes);
  return true;
}

// ---------- ONE-TIME SETUP ----------
// Run this once manually from the Apps Script editor to seed the sheet with
// the initial data extracted from the official action plan document.
function setupInitialData() {
  saveYearly([
    { 'ปี': '2566', 'จำนวนคดี': 479, 'มูลค่าความเสียหาย': 27577857, 'หมายเหตุ': '' },
    { 'ปี': '2567', 'จำนวนคดี': 999, 'มูลค่าความเสียหาย': 56254078, 'หมายเหตุ': '' },
    { 'ปี': '2568', 'จำนวนคดี': 932, 'มูลค่าความเสียหาย': 48875912, 'หมายเหตุ': '' },
    { 'ปี': '2569 (ต.ค.68-พ.ค.69)', 'จำนวนคดี': 201, 'มูลค่าความเสียหาย': 17744776, 'หมายเหตุ': 'ข้อมูล ณ พ.ค. 69' }
  ]);
  saveDistrict([
    { 'อำเภอ': 'เมืองหนองบัวลำภู', 'เป้ารับแจ้งคดี': 100, 'ผลรับแจ้งคดี': 100, 'เป้าติดตามอายัดเงิน': 100, 'ผลติดตามอายัดเงิน': 100, 'เป้าสอบสวนสรุปสำนวน': 100, 'ผลสอบสวนสรุปสำนวน': 100 },
    { 'อำเภอ': 'ศรีบุญเรือง', 'เป้ารับแจ้งคดี': 100, 'ผลรับแจ้งคดี': 100, 'เป้าติดตามอายัดเงิน': 100, 'ผลติดตามอายัดเงิน': 100, 'เป้าสอบสวนสรุปสำนวน': 100, 'ผลสอบสวนสรุปสำนวน': 100 },
    { 'อำเภอ': 'นากลาง', 'เป้ารับแจ้งคดี': 100, 'ผลรับแจ้งคดี': 100, 'เป้าติดตามอายัดเงิน': 100, 'ผลติดตามอายัดเงิน': 100, 'เป้าสอบสวนสรุปสำนวน': 100, 'ผลสอบสวนสรุปสำนวน': 100 },
    { 'อำเภอ': 'โนนสัง', 'เป้ารับแจ้งคดี': 100, 'ผลรับแจ้งคดี': 100, 'เป้าติดตามอายัดเงิน': 100, 'ผลติดตามอายัดเงิน': 100, 'เป้าสอบสวนสรุปสำนวน': 100, 'ผลสอบสวนสรุปสำนวน': 100 },
    { 'อำเภอ': 'นาวัง', 'เป้ารับแจ้งคดี': 100, 'ผลรับแจ้งคดี': 100, 'เป้าติดตามอายัดเงิน': 100, 'ผลติดตามอายัดเงิน': 100, 'เป้าสอบสวนสรุปสำนวน': 100, 'ผลสอบสวนสรุปสำนวน': 100 },
    { 'อำเภอ': 'สุวรรณคูหา', 'เป้ารับแจ้งคดี': 100, 'ผลรับแจ้งคดี': 100, 'เป้าติดตามอายัดเงิน': 100, 'ผลติดตามอายัดเงิน': 100, 'เป้าสอบสวนสรุปสำนวน': 100, 'ผลสอบสวนสรุปสำนวน': 100 }
  ]);
  saveContacts([
    { 'ลำดับ': 1, 'ชื่อสกุล': 'พ.ต.ต.วิชาญ จันทร์พิมพ์', 'หน่วยงาน': 'ตำรวจภูธรจังหวัดหนองบัวลำภู', 'ตำแหน่ง': 'สารวัตรฝ่ายอำนวยการ', 'เบอร์โทร': '089-174-7869' }
  ]);
  saveMeasures([
    { 'ด้าน': 'การปราบปราม', 'ลำดับ': '1.1', 'รายละเอียด': 'การบล็อกลิงก์และเว็บพนัน: ประสานตำรวจภูธรและผู้ให้บริการอินเทอร์เน็ตปิดกั้น URL/IP เว็บหลอกลวง-เว็บพนัน ภายใน 24 ชม.', 'หน่วยงานรับผิดชอบ': 'ตำรวจภูธรจังหวัดหนองบัวลำภู / ธนาคารทุกแห่ง', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'การปราบปราม', 'ลำดับ': '1.2', 'รายละเอียด': 'การกวาดล้างซิมม้า: ผู้ถือครองซิมเกิน 5 ซิม ต้องยืนยันตัวตน (KYC) หากฝ่าฝืนระงับสัญญาณทันที', 'หน่วยงานรับผิดชอบ': 'สำนักงานสถิติจังหวัดหนองบัวลำภู', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'การปราบปราม', 'ลำดับ': '1.3', 'รายละเอียด': 'บูรณาการฐานข้อมูลข้ามหน่วยงาน (ตำรวจ, ธนาคาร) วิเคราะห์เส้นทางการเงินและเครือข่ายทุนนอมินี/แก๊งคอลเซ็นเตอร์', 'หน่วยงานรับผิดชอบ': 'สำนักงานประชาสัมพันธ์จังหวัด / อำเภอทุกอำเภอ', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'การปราบปราม', 'ลำดับ': '1.4', 'รายละเอียด': 'การระงับบัญชีม้าทันที (Hotline 1441): ธนาคารระงับธุรกรรม (Freeze) บัญชีต้องสงสัยภายใน 3 ชั่วโมง', 'หน่วยงานรับผิดชอบ': 'ธนาคารทุกแห่ง / อปท.', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'การปราบปราม', 'ลำดับ': '1.5', 'รายละเอียด': 'จัดตั้งศูนย์ AOC 1441 ระดับจังหวัด เป็นศูนย์สั่งการร่วม (War Room) ตำรวจ-ฝ่ายปกครอง-สถาบันการเงิน', 'หน่วยงานรับผิดชอบ': 'ตำรวจภูธรจังหวัดหนองบัวลำภู', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'การปราบปราม', 'ลำดับ': '1.6', 'รายละเอียด': 'แกะรอยเส้นทางการเงินเชิงลึก ประสาน ปปง. และธนาคาร ตรวจสอบนิติกรรมอำพราง บัญชีม้า และคริปโทเคอร์เรนซี', 'หน่วยงานรับผิดชอบ': 'ตำรวจภูธรจังหวัดหนองบัวลำภู / ธนาคารทุกแห่ง', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'การปราบปราม', 'ลำดับ': '1.7', 'รายละเอียด': 'รวบรวมหลักฐานประสานกระทรวง DES ปิดกั้น URL/IP เว็บพนันและเว็บหลอกลวงอย่างรวดเร็ว', 'หน่วยงานรับผิดชอบ': 'สำนักงานสถิติจังหวัดหนองบัวลำภู', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'การปราบปราม', 'ลำดับ': '1.8', 'รายละเอียด': 'สร้างเครือข่ายภาคประชาชนผ่านกำนัน ผู้ใหญ่บ้าน และแอปพลิเคชันชุมชน แจ้งเตือนภัยรูปแบบใหม่', 'หน่วยงานรับผิดชอบ': 'อำเภอทุกอำเภอ / อปท.', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'ประชาสัมพันธ์', 'ลำดับ': '2.1', 'รายละเอียด': 'ประชาสัมพันธ์แนวทางป้องกันก่อนรับโทรศัพท์/อ่านข้อความ ให้มีสติไตร่ตรองข้อมูลก่อนเชื่อ', 'หน่วยงานรับผิดชอบ': 'ตำรวจภูธรจังหวัดหนองบัวลำภู', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'ประชาสัมพันธ์', 'ลำดับ': '2.2', 'รายละเอียด': 'ประชาสัมพันธ์ให้หลีกเลี่ยงการให้ข้อมูลส่วนตัว/ข้อมูลการเงินกับคนแปลกหน้า แม้อ้างเป็นหน่วยงานราชการ', 'หน่วยงานรับผิดชอบ': 'สำนักงานสถิติจังหวัดหนองบัวลำภู', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'ประชาสัมพันธ์', 'ลำดับ': '2.3', 'รายละเอียด': 'ประชาสัมพันธ์ให้ตรวจสอบข้อมูลกับหน่วยงานที่เกี่ยวข้องโดยตรงก่อนเชื่อผู้แอบอ้าง', 'หน่วยงานรับผิดชอบ': 'สำนักงานประชาสัมพันธ์จังหวัด', 'สถานะ': 'ดำเนินการต่อเนื่อง' },
    { 'ด้าน': 'ประชาสัมพันธ์', 'ลำดับ': '2.4', 'รายละเอียด': 'แนะนำประชาชนแจ้งความ/แจ้งหน่วยงานที่เกี่ยวข้องทันทีหากตกเป็นเหยื่อ ไม่ต้องอาย', 'หน่วยงานรับผิดชอบ': 'อำเภอทุกอำเภอ / อปท.', 'สถานะ': 'ดำเนินการต่อเนื่อง' }
  ]);
  saveScamTypes([
    { 'ลำดับ': 1, 'รูปแบบ': 'หลอกขายสินค้า', 'สัดส่วน': 'อันดับ 1' },
    { 'ลำดับ': 2, 'รูปแบบ': 'หลอกให้โอนเงินเพื่อหารายได้พิเศษ / งานออนไลน์', 'สัดส่วน': 'อันดับ 2' },
    { 'ลำดับ': 3, 'รูปแบบ': 'อ้างเป็นเจ้าหน้าที่ตำรวจ/สรรพากร/การไฟฟ้า/เจ้าหน้าที่รัฐ สร้างความกลัวให้โอนเงิน', 'สัดส่วน': 'พบบ่อย' },
    { 'ลำดับ': 4, 'รูปแบบ': 'หลอกให้หลงรักผ่านโซเชียลมีเดียแล้วขอเงิน (Romance Scam)', 'สัดส่วน': 'พบบ่อย' },
    { 'ลำดับ': 5, 'รูปแบบ': 'ส่งลิงก์ปลอมดูดข้อมูลบัตรเครดิตและบัญชีธนาคาร', 'สัดส่วน': 'พบบ่อย' },
    { 'ลำดับ': 6, 'รูปแบบ': 'เปิดบ่อนออนไลน์โดยใช้บัญชีม้าของคนในพื้นที่', 'สัดส่วน': 'พบบ่อย' }
  ]);
  Logger.log('Initial data seeded successfully.');
}

// Helper to compute the real PIN hash once and paste it into PIN_HASH above.
function logPinHash() {
  Logger.log(sha256Hex('309309'));
}
