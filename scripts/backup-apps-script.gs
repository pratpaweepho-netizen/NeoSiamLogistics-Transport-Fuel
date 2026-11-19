// ============================================================
// NeoSiam Fuel — สำรองข้อมูลอัตโนมัติทุกคืนลง Google Drive (ฟรี)
// ไฟล์ที่ได้เป็นรูปแบบเดียวกับปุ่ม "ดาวน์โหลดสำรองทั้งระบบ" ในแอป → ใช้ "กู้คืนจากไฟล์สำรอง" ได้เลย
// (ต้นแบบจากโปรเจ็คสรุปรอบบิล · ปรับใช้ 19 พ.ย. 2569)
//
// วิธีติดตั้ง (ทำครั้งเดียว):
//   1. เปิด Google Sheets สร้างชีตใหม่ ตั้งชื่อ "NeoSiam Fuel — สำรองข้อมูล"
//   2. เมนู Extensions (ส่วนขยาย) > Apps Script
//   3. ลบโค้ดเดิมในหน้านั้น แล้ววางโค้ดในไฟล์นี้ทั้งหมด > กด Save (รูปแผ่นดิสก์)
//   4. ช่องเลือกฟังก์ชันด้านบน เลือก setupDailyTrigger > กด Run
//   5. กด Review permissions > เลือกบัญชี Google ของบริษัท > Allow
//   6. เสร็จ: ได้ไฟล์แรกในโฟลเดอร์ "NeoSiamFuel-Backup" ใน Google Drive และแถวแรกในชีต "สำรอง-log"
//      หลังจากนี้ระบบสำรองเองทุกคืนตี 1 และเก็บไว้ 30 ไฟล์ล่าสุด
//
// หมายเหตุ: ไฟล์สำรองมีเลขบัญชีคนขับ — ใช้บัญชี Google ของบริษัท และอย่าแชร์โฟลเดอร์ให้คนนอก
//          ตอนนี้ฐานข้อมูลยังเปิดให้อ่านได้ สคริปต์จึงอ่านได้โดยไม่ต้องล็อกอิน ถ้าวันหน้าปิดฐานข้อมูล
//          (ระยะ 3 ที่พักไว้) ต้องเพิ่มการล็อกอินในสคริปต์นี้ ไม่อย่างนั้นการสำรองจะหยุดและชีตจะขึ้น ERROR
// ============================================================

const DB_URL    = 'https://fleet-fuel-management-system-default-rtdb.asia-southeast1.firebasedatabase.app/.json';
const APP       = 'NeoSiamLogistics & Transport Fuel'; // ต้องตรงกับในแอป (ไม่อย่างนั้นแอปจะไม่รับไฟล์)
const FOLDER    = 'NeoSiamFuel-Backup';
const LOG_SHEET = 'สำรอง-log';
const KEEP      = 30;   // เก็บไฟล์ล่าสุดกี่ไฟล์
const HOUR      = 1;    // สำรองตอนกี่โมง (0–23)

function backupNow() {
  const now = new Date();
  const stamp = Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd_HH-mm');
  try {
    const data = JSON.parse(JSON.stringify(fetchDb_()));
    const counts = countAll_(data);
    const file = { app: APP, version: 2, exportedAt: now.toISOString(), counts: counts, data: data };
    const json = JSON.stringify(file);
    const folder = getFolder_(FOLDER);
    // a Blob (up to 50 MB) — createFile(name, text) is limited to 10 MB and the database is already ~16 MB
    folder.createFile(Utilities.newBlob(json, 'application/json', 'NeoSiamFuel_backup_' + stamp + '.json'));
    pruneOld_(folder, KEEP);
    logRow_([Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/') + (now.getFullYear() + 543) + Utilities.formatDate(now, 'Asia/Bangkok', ' HH:mm'),
      'สำเร็จ', counts.trips || 0, counts.fuelRateStandards || 0, counts.drivers || 0, counts.vehicles || 0, Math.round(json.length / 1024) + ' KB', '']);
  } catch (e) {
    logRow_([Utilities.formatDate(now, 'Asia/Bangkok', 'dd/MM/') + (now.getFullYear() + 543) + Utilities.formatDate(now, 'Asia/Bangkok', ' HH:mm'),
      'ERROR', '', '', '', '', '', String(e && e.message || e)]);
    throw e; // Google ส่งอีเมลแจ้งเจ้าของสคริปต์เมื่อ trigger ล้มเหลว
  }
}

function fetchDb_() {
  for (let i = 0; i < 3; i++) {
    const res = UrlFetchApp.fetch(DB_URL, { muteHttpExceptions: true });
    if (res.getResponseCode() === 200) return JSON.parse(res.getContentText()) || {};
    Utilities.sleep(10000);
  }
  throw new Error('อ่านฐานข้อมูลไม่สำเร็จ (HTTP error) — ถ้าเพิ่งปิดสิทธิ์ฐานข้อมูล ต้องเพิ่มการล็อกอินในสคริปต์');
}

// นับแบบเดียวกับในแอป (src/backup.js): เที่ยวรถ/ลิตรขาล่องนับรายเที่ยว, ตารางอัตรานับรายแถว
function countAll_(root) {
  const out = {};
  const n = function (o) { return Object.keys(o || {}).filter(function (k) { return o[k] !== null && o[k] !== undefined; }).length; };
  Object.keys(root || {}).forEach(function (k) {
    const v = root[k];
    if (v === null || v === undefined) out[k] = 0;
    else if (typeof v !== 'object') out[k] = 1;
    else if (k === 'trips' || k === 'downhillFuel') out[k] = Object.keys(v).reduce(function (s, m) { return s + (v[m] && typeof v[m] === 'object' ? n(v[m]) : 0); }, 0);
    else if (k === 'fuelRateStandards') out[k] = n(v.rows);
    else out[k] = n(v);
  });
  return out;
}

function getFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function pruneOld_(folder, keep) {
  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  files.slice(keep).forEach(function (f) { f.setTrashed(true); });
}

function logRow_(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(['เวลาสำรอง', 'ผล', 'เที่ยวรถ', 'แถวอัตรา', 'คนขับ', 'ทะเบียนรถ', 'ขนาดไฟล์', 'ข้อผิดพลาด']);
    sh.getRange('1:1').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  sh.appendRow(row);
}

function setupDailyTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('backupNow').timeBased().everyDays(1).atHour(HOUR).inTimezone('Asia/Bangkok').create();
  backupNow(); // ทดสอบทันที 1 ครั้ง
}
