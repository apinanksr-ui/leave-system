// ============================================================
//  ระบบแจ้งลาออนไลน์ — Plan Consultants Co.,Ltd.
//  Google Apps Script Web App  (Code.gs)
//  วาง code นี้ใน Google Apps Script แล้ว Deploy เป็น Web App
// ============================================================

// ── ชื่อ Sheet ใน Google Sheets ──────────────────────────────
var SHEET_EMP      = '1_พนักงาน';
var SHEET_APPROVER = '2_ผู้อนุมัติ';
var SHEET_QUOTA    = '3_โควต้าการลา';
var SHEET_LEAVE    = '4_ใบลา';
var SHEET_ORG      = '5_โครงสร้างองค์กร';

// ── แปลงชื่อคอลัมน์ภาษาไทย → ชื่อ key ภาษาอังกฤษ ──────────
var COL_MAP_EMP = {
  'รหัสพนักงาน*':        'empId',
  'รหัสพนักงาน':         'empId',
  'คำนำหน้า':            'title',
  'ชื่อ-นามสกุล*':       'name',
  'ชื่อ-นามสกุล':        'name',
  'อีเมล*':              'email',
  'อีเมล':               'email',
  'เบอร์โทรศัพท์*':      'phone',
  'เบอร์โทรศัพท์':       'phone',
  'ฝ่าย*':               'dept',
  'ฝ่าย':                'dept',
  'แผนก*':               'section',
  'แผนก':                'section',
  'ตำแหน่ง':             'position',
  'สิทธิ์ (role)*':      'role',
  'สิทธิ์ (role)':       'role',
  'รูปแบบอนุมัติ*':      'workflow',
  'รูปแบบอนุมัติ':       'workflow',
  'รหัสผู้อนุมัติ L1':   'l1Id',
  'รหัสผู้อนุมัติ L2*':  'l2Id',
  'รหัสผู้อนุมัติ L2':   'l2Id',
  'สถานะ*':              'status',
  'สถานะ':               'status',
  'รหัสผ่านเริ่มต้น*':   'password',
  'รหัสผ่านเริ่มต้น':    'password',
  'หมายเหตุ':            'note'
};

var COL_MAP_LEAVE = {
  'รหัสใบลา':            'leaveId',
  'เลขที่ใบลา':          'leaveId',   // ชื่อทางเลือกใน Sheet
  'รหัสพนักงาน':         'empId',
  'ชื่อพนักงาน':         'empName',
  'ฝ่าย':                'dept',
  'ตำแหน่ง':             'position',
  'ประเภทการลา':         'leaveType',
  'วันที่เริ่มต้น':      'startDate',
  'วันที่เริ่มลา':       'startDate',  // ชื่อทางเลือกใน Sheet
  'วันที่สิ้นสุด':       'endDate',
  'จำนวนวัน':            'days',
  'เหตุผล':              'reason',
  'สถานะ':               'status',
  'วันที่ยื่น':          'submitDate',
  'รหัสผู้อนุมัติ L1':   'l1AppId',
  'รหัสผู้อนุมัติ L2':   'l2AppId',
  'ผลการพิจารณา L1':     'l1Decision',
  'ผลการพิจารณา L2':     'l2Decision',
  'ความเห็น L1':         'l1Comment',
  'ความเห็น L2':         'l2Comment',
  'ไฟล์แนบ':             'attachmentUrl'
};

var COL_MAP_QUOTA = {
  'รหัสพนักงาน':         'empId',
  'ชื่อพนักงาน':         'name',
  'โควต้าลาป่วย':        'sickQuota',
  'ใช้ไปลาป่วย':         'sickUsed',
  'โควต้าลากิจ':         'businessQuota',
  'ใช้ไปลากิจ':          'businessUsed',
  'โควต้าลาพักผ่อน':     'vacationQuota',
  'ใช้ไปลาพักผ่อน':      'vacationUsed'
};

var COL_MAP_APPROVER = {
  'รหัสผู้อนุมัติ':      'id',
  'ชื่อ-นามสกุล':        'name',
  'ตำแหน่ง':             'position',
  'ระดับ (L1/L2)':       'level',
  'อีเมล':               'email',
  'เบอร์โทรศัพท์':       'phone',
  'ฝ่าย':                'dept',
  'สถานะ':               'status'
};

/** แปลง header โดยใช้ map ที่ให้มา ถ้าไม่พบให้ใช้ค่าเดิม (ตัด * ออกก่อน) */
function normalizeHeader(raw, map) {
  var trimmed = String(raw).trim().replace(/\*/g, '').trim();
  return map[trimmed] || trimmed;
}

/**
 * หา header row ใน sheet และ return ข้อมูลที่ต้องการ
 * คืนค่า: { rawHdr, normalizedHdr, headerRowIdx, all }
 */
function findHeaderRow(sh, colMap) {
  var all = sh.getDataRange().getValues();
  var headerRowIdx = 0;
  for (var r = 0; r < Math.min(5, all.length); r++) {
    var nonEmpty = all[r].filter(function(c) { return c !== '' && c !== null; }).length;
    if (nonEmpty >= 3) { headerRowIdx = r; break; }
  }
  var rawHdr = all[headerRowIdx];
  var map = colMap || {};
  var normalizedHdr = rawHdr.map(function(h) { return normalizeHeader(h, map); });
  return { rawHdr: rawHdr, normalizedHdr: normalizedHdr, headerRowIdx: headerRowIdx, all: all };
}

// ============================================================
//  Entry Point — รับ HTTP GET
// ============================================================
function doGet(e) {
  var params = e.parameter;
  var action = params.action || '';

  try {
    var result;
    switch (action) {
      case 'login':          result = login(params.id, params.password);    break;
      case 'getEmployees':   result = getEmployees();                        break;
      case 'getEmployee':    result = getEmployee(params.id);                break;
      case 'getApprovers':   result = getApprovers();                        break;
      case 'getLeaves':      result = getLeaves(params);                     break;
      case 'getQuotas':      result = getQuotas(params.empId);               break;
      case 'getOrgStructure':result = getOrgStructure();                     break;
      default:               result = { ok: false, error: 'Unknown action' };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
//  Entry Point — รับ HTTP POST
// ============================================================
function doPost(e) {
  var params;
  try {
    params = JSON.parse(e.postData.contents);
  } catch(x) {
    return jsonResponse({ ok: false, error: 'Invalid JSON body' });
  }

  var action = params.action || '';

  try {
    var result;
    switch (action) {
      case 'submitLeave':    result = submitLeave(params.data);              break;
      case 'updateLeave':    result = updateLeave(params.leaveId, params.data); break;
      case 'addEmployee':    result = addEmployee(params.data);              break;
      case 'updateEmployee': result = updateEmployee(params.empId, params.data); break;
      case 'deleteEmployee': result = deleteEmployee(params.empId);          break;
      case 'resetPassword':  result = resetPassword(params.empId, params.newPassword); break;
      case 'updateQuota':    result = updateQuota(params.empId, params.data); break;
      default:               result = { ok: false, error: 'Unknown action' };
    }
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
//  AUTH — Login
// ============================================================
function login(id, password) {
  if (!id || !password) return { ok: false, error: 'กรุณากรอก ID และรหัสผ่าน' };
  // ใช้ readSheet พร้อม COL_MAP_EMP เพื่อ normalize header เป็น English key
  var data = readSheet(SHEET_EMP, COL_MAP_EMP);
  var rows = data.rows;

  for (var i = 0; i < rows.length; i++) {
    var u = rows[i];
    if (!u.empId) continue;
    if (String(u.empId).trim().toUpperCase() === String(id).trim().toUpperCase()) {
      if (String(u.status).trim() === 'suspended')
        return { ok: false, error: 'บัญชีนี้ถูกระงับการใช้งาน' };
      if (String(u.password).trim() === String(password).trim())
        return { ok: true, user: u };
      else
        return { ok: false, error: 'รหัสผ่านไม่ถูกต้อง' };
    }
  }
  return { ok: false, error: 'ไม่พบรหัสผู้ใช้งาน' };
}

// ============================================================
//  EMPLOYEES
// ============================================================
function getEmployees() {
  var data = readSheet(SHEET_EMP, COL_MAP_EMP);
  // ไม่ส่ง password กลับ
  data.rows = data.rows.map(function(u) {
    var copy = Object.assign({}, u);
    delete copy.password;
    return copy;
  });
  return { ok: true, data: data.rows };
}

function getEmployee(id) {
  var data = readSheet(SHEET_EMP, COL_MAP_EMP);
  var user = data.rows.find(function(u) { return String(u.empId) === String(id); });
  if (!user) return { ok: false, error: 'ไม่พบพนักงาน' };
  var copy = Object.assign({}, user);
  delete copy.password;
  return { ok: true, data: copy };
}

function addEmployee(obj) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName(SHEET_EMP);
  var info = findHeaderRow(sh, COL_MAP_EMP);
  var row  = info.normalizedHdr.map(function(key) {
    return obj[key] !== undefined ? obj[key] : '';
  });
  sh.appendRow(row);
  return { ok: true, message: 'เพิ่มพนักงานเรียบร้อย' };
}

function updateEmployee(empId, obj) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName(SHEET_EMP);
  var info = findHeaderRow(sh, COL_MAP_EMP);
  var IDX  = info.normalizedHdr.indexOf('empId');
  for (var i = info.headerRowIdx + 1; i < info.all.length; i++) {
    if (String(info.all[i][IDX]).trim() === String(empId).trim()) {
      info.normalizedHdr.forEach(function(key, c) {
        if (obj[key] !== undefined && key !== 'empId') {
          sh.getRange(i + 1, c + 1).setValue(obj[key]);
        }
      });
      return { ok: true, message: 'อัปเดตข้อมูลเรียบร้อย' };
    }
  }
  return { ok: false, error: 'ไม่พบพนักงาน ID: ' + empId };
}

function deleteEmployee(empId) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName(SHEET_EMP);
  var info = findHeaderRow(sh, COL_MAP_EMP);
  var IDX  = info.normalizedHdr.indexOf('empId');
  for (var i = info.headerRowIdx + 1; i < info.all.length; i++) {
    if (String(info.all[i][IDX]).trim() === String(empId).trim()) {
      sh.deleteRow(i + 1);
      return { ok: true, message: 'ลบพนักงานเรียบร้อย' };
    }
  }
  return { ok: false, error: 'ไม่พบพนักงาน' };
}

function resetPassword(empId, newPassword) {
  return updateEmployee(empId, { password: newPassword });
}

// ============================================================
//  APPROVERS
// ============================================================
function getApprovers() {
  var data = readSheet(SHEET_APPROVER, COL_MAP_APPROVER);
  return { ok: true, data: data.rows };
}

// ============================================================
//  LEAVE REQUESTS
// ============================================================
function getLeaves(params) {
  var data = readSheet(SHEET_LEAVE, COL_MAP_LEAVE);
  var rows = data.rows;

  // กรองตาม parameter ที่ส่งมา
  if (params.empId)  rows = rows.filter(function(r) { return String(r.empId)  === String(params.empId);  });
  if (params.status) rows = rows.filter(function(r) { return String(r.status) === String(params.status); });
  if (params.month)  rows = rows.filter(function(r) {
    var d = new Date(r.startDate);
    return (d.getMonth() + 1) === parseInt(params.month) && d.getFullYear() === parseInt(params.year || new Date().getFullYear());
  });
  if (params.year && !params.month) rows = rows.filter(function(r) {
    return new Date(r.startDate).getFullYear() === parseInt(params.year);
  });

  return { ok: true, data: rows };
}

function submitLeave(obj) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName(SHEET_LEAVE);
  var info = findHeaderRow(sh, COL_MAP_LEAVE);

  // สร้าง leaveId อัตโนมัติ
  obj.leaveId    = 'LV' + String(Date.now()).slice(-6);
  obj.submitDate = new Date().toLocaleDateString('th-TH');
  obj.status     = 'pending_l1';
  obj.l1Decision = '';
  obj.l2Decision = '';
  obj.l1Comment  = '';
  obj.l2Comment  = '';

  // ── บันทึกไฟล์แนบลง Google Drive (ถ้ามี) ──────────────────
  if (obj.attachmentBase64 && obj.attachmentName) {
    try {
      var b64Data      = obj.attachmentBase64;
      var base64String = b64Data.indexOf(',') !== -1 ? b64Data.split(',')[1] : b64Data;
      var mimeType     = b64Data.indexOf(',') !== -1 ? b64Data.split(';')[0].replace('data:','') : 'application/octet-stream';
      var decoded      = Utilities.base64Decode(base64String);
      var blob         = Utilities.newBlob(decoded, mimeType, obj.attachmentName);
      var folderName   = 'LeaveAttachments_PlanConsultants';
      var folders      = DriveApp.getFoldersByName(folderName);
      var folder       = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
      var file         = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      obj.attachmentUrl = file.getUrl();
    } catch(e) {
      obj.attachmentUrl = 'อัปโหลดไม่สำเร็จ: ' + e.message;
    }
    delete obj.attachmentBase64;
    delete obj.attachmentName;
  }

  // ใช้ normalizedHdr (English keys) map กับ obj แล้ว appendRow
  var row = info.normalizedHdr.map(function(key) {
    return obj[key] !== undefined ? obj[key] : '';
  });
  sh.appendRow(row);
  return { ok: true, leaveId: obj.leaveId, message: 'ส่งใบลาเรียบร้อย' };
}

function updateLeave(leaveId, obj) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName(SHEET_LEAVE);
  var info = findHeaderRow(sh, COL_MAP_LEAVE);
  var IDX  = info.normalizedHdr.indexOf('leaveId');

  for (var i = info.headerRowIdx + 1; i < info.all.length; i++) {
    if (String(info.all[i][IDX]).trim() === String(leaveId).trim()) {
      info.normalizedHdr.forEach(function(key, c) {
        if (obj[key] !== undefined) {
          sh.getRange(i + 1, c + 1).setValue(obj[key]);
        }
      });
      return { ok: true, message: 'อัปเดตใบลาเรียบร้อย' };
    }
  }
  return { ok: false, error: 'ไม่พบใบลา ID: ' + leaveId };
}

// ============================================================
//  LEAVE QUOTAS
// ============================================================
function getQuotas(empId) {
  var data = readSheet(SHEET_QUOTA, COL_MAP_QUOTA);
  var rows = data.rows;
  if (empId) rows = rows.filter(function(r) { return String(r.empId) === String(empId); });
  return { ok: true, data: rows };
}

function updateQuota(empId, obj) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName(SHEET_QUOTA);
  var info = findHeaderRow(sh, COL_MAP_QUOTA);
  var IDX  = info.normalizedHdr.indexOf('empId');
  for (var i = info.headerRowIdx + 1; i < info.all.length; i++) {
    if (String(info.all[i][IDX]).trim() === String(empId).trim()) {
      info.normalizedHdr.forEach(function(key, c) {
        if (obj[key] !== undefined) sh.getRange(i + 1, c + 1).setValue(obj[key]);
      });
      return { ok: true, message: 'อัปเดตโควต้าเรียบร้อย' };
    }
  }
  return { ok: false, error: 'ไม่พบข้อมูลโควต้าพนักงาน' };
}

// ============================================================
//  ORG STRUCTURE
// ============================================================
function getOrgStructure() {
  var data = readSheet(SHEET_ORG);
  return { ok: true, data: data.rows };
}

// ============================================================
//  HELPERS
// ============================================================

/** อ่าน Sheet แล้วแปลงเป็น array of objects พร้อม normalize header */
function readSheet(sheetName, colMap) {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName(sheetName);
  if (!sh) return { rows: [] };
  var vals = sh.getDataRange().getValues();
  // ข้ามแถว title/subtitle (มักเป็น merged cell หรือ row ที่ไม่ใช่ header)
  // หา header row = แถวแรกที่ไม่ว่างและมีคอลัมน์มากกว่า 2 ค่า
  var headerRowIdx = 0;
  for (var r = 0; r < Math.min(5, vals.length); r++) {
    var nonEmpty = vals[r].filter(function(c){ return c !== '' && c !== null; }).length;
    if (nonEmpty >= 3) { headerRowIdx = r; break; }
  }
  if (vals.length <= headerRowIdx + 1) return { rows: [] };
  var rawHdr = vals[headerRowIdx];
  // normalize header
  var map = colMap || {};
  var hdr = rawHdr.map(function(h) { return normalizeHeader(h, map); });
  var rows = [];
  for (var i = headerRowIdx + 1; i < vals.length; i++) {
    var row = vals[i];
    if (row.every(function(c) { return c === '' || c === null; })) continue;
    // ข้ามแถวคำแนะนำ/ตัวอย่าง: ตรวจจาก "/" ในเซลล์ 2+ ช่อง (เช่น "นาย/นาง", "active/suspended")
    var slashCount = row.filter(function(c) { return String(c).indexOf('/') !== -1; }).length;
    if (slashCount >= 2) continue;
    rows.push(rowToObj(hdr, row));
  }
  return { rows: rows };
}

/** แปลง array row เป็น object โดยใช้ header เป็น key */
function rowToObj(hdr, row) {
  var obj = {};
  hdr.forEach(function(key, i) {
    var val = row[i];
    // แปลง Date object → string
    if (val instanceof Date) {
      val = val.toLocaleDateString('th-TH');
    }
    obj[key] = val;
  });
  return obj;
}

/** ส่ง JSON response พร้อม CORS header */
function jsonResponse(data) {
  var output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
  return output;
}
