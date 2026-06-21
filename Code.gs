// ============================================================
//  نظام إدارة الرواتب — Google Apps Script
//  الملف: Code.gs — انسخه في محرر Apps Script
// ============================================================

// ⚠️ ضع معرّف جدول البيانات من رابطه:
// https://docs.google.com/spreadsheets/d/[المعرّف]/edit
const SHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

// قائمة الإيميلات المسموح لها (فارغة = جميع حسابات Google)
const ALLOWED_EMAILS = [];

// ======================================================
//  doGet — تقديم واجهة المستخدم
// ======================================================
function doGet() {
  if (!isAuthorized()) {
    return HtmlService.createHtmlOutput(
      '<div style="font-family:Arial;text-align:center;padding:3rem;color:#c62828"><h2>⛔ غير مصرح بالوصول</h2></div>'
    );
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('نظام إدارة الرواتب')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ======================================================
//  doPost — REST API للوصول الخارجي
// ======================================================
function doPost(e) {
  try {
    if (!isAuthorized()) return jsonOut({ success: false, error: 'غير مصرح' });
    var body = JSON.parse(e.postData.contents);
    return jsonOut(dispatch(body.action, body.payload || {}));
  } catch (err) {
    return jsonOut({ success: false, error: err.toString() });
  }
}

// ======================================================
//  دوال مكشوفة لـ google.script.run
// ======================================================
function getAllData() {
  if (!isAuthorized()) throw new Error('غير مصرح');
  return dispatch('getAll', {});
}

function saveRecord(payload) {
  if (!isAuthorized()) throw new Error('غير مصرح');
  return dispatch('syncRecord', payload);
}

function saveSettingsData(payload) {
  if (!isAuthorized()) throw new Error('غير مصرح');
  return dispatch('saveSettings', payload);
}

function calcPayrollData(month) {
  if (!isAuthorized()) throw new Error('غير مصرح');
  return dispatch('calcPayroll', { month: month });
}

// ======================================================
//  Dispatcher
// ======================================================
function dispatch(action, payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  initSheets(ss);
  switch (action) {
    case 'getAll':       return handleGetAll(ss);
    case 'syncRecord':   return handleSyncRecord(ss, payload);
    case 'saveSettings': return handleSaveSettings(ss, payload);
    case 'calcPayroll':  return handleCalcPayroll(ss, payload);
    default: return { success: false, error: 'إجراء غير معروف: ' + action };
  }
}

// ======================================================
//  Authorization
// ======================================================
function isAuthorized() {
  if (ALLOWED_EMAILS.length === 0) return true;
  try {
    return ALLOWED_EMAILS.indexOf(Session.getEffectiveUser().getEmail()) !== -1;
  } catch(e) { return false; }
}

// ======================================================
//  getAll
// ======================================================
function handleGetAll(ss) {
  var result = {};
  ['employees','loans','deductions','arrears','bonuses','attendance'].forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) { result[name] = []; return; }
    var data = sh.getDataRange().getValues();
    var headers = data[0];
    result[name] = data.slice(1)
      .filter(function(r) { return r[0] !== '' && r[0] !== null; })
      .map(function(r) {
        var o = {};
        headers.forEach(function(h, i) { o[h] = (r[i] == null) ? '' : r[i]; });
        return o;
      });
  });
  var sh = ss.getSheetByName('settings');
  var settings = {};
  if (sh && sh.getLastRow() > 0) {
    sh.getDataRange().getValues().forEach(function(r) { if (r[0]) settings[String(r[0])] = r[1]; });
  }
  result.settings = settings;
  return { success: true, data: result };
}

// ======================================================
//  syncRecord
// ======================================================
function handleSyncRecord(ss, payload) {
  var sheetName = payload.sheet, record = payload.record, op = payload.operation;
  if (!sheetName || !record) return { success: false, error: 'بيانات ناقصة' };

  var sh = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();

  if (op === 'delete') {
    if (lastRow >= 2) {
      var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(record.id)) { sh.deleteRow(i + 2); break; }
      }
    }
    return { success: true };
  }

  var headers;
  if (lastRow === 0 || lastCol === 0) {
    headers = Object.keys(record);
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length).setBackground('#0f2942').setFontColor('#fff').setFontWeight('bold');
    lastRow = 1;
  } else {
    headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  }

  var values = headers.map(function(h) { return record[h] !== undefined ? record[h] : ''; });

  if (lastRow > 1) {
    var existing = sh.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]) === String(record.id)) {
        sh.getRange(i + 2, 1, 1, values.length).setValues([values]);
        return { success: true };
      }
    }
  }
  sh.appendRow(values);
  return { success: true };
}

// ======================================================
//  saveSettings
// ======================================================
function handleSaveSettings(ss, payload) {
  var sh = ss.getSheetByName('settings') || ss.insertSheet('settings');
  sh.clearContents();
  Object.keys(payload).forEach(function(k) { sh.appendRow([k, payload[k]]); });
  return { success: true };
}

// ======================================================
//  calcPayroll — احتساب مسيّر الرواتب
// ======================================================
function handleCalcPayroll(ss, payload) {
  var month = payload.month;
  if (!month) return { success: false, error: 'الشهر مطلوب' };

  var d        = handleGetAll(ss).data;
  var gosiRate = parseFloat(d.settings.gosi || 9) / 100;
  var workDays = parseInt(d.settings.workDays || 30);
  var currency = d.settings.currency || 'ريال';

  var results = d.employees
    .filter(function(e) { return e.status !== 'inactive'; })
    .map(function(emp) {
      var basic = +emp.basic||0, housing = +emp.housing||0,
          transport = +emp.transport||0, phone = +emp.phone||0, other = +emp.other||0;
      var gross   = basic + housing + transport + phone + other;
      var gosi    = +(basic * gosiRate).toFixed(2);
      var loanDed = d.loans
        .filter(function(l) { return String(l.empId)===String(emp.id) && l.status==='active'; })
        .reduce(function(s,l) { return s+(+l.monthly||0); }, 0);
      var extraDed = d.deductions
        .filter(function(x) { return String(x.empId)===String(emp.id) && (x.month===month||x.status==='recurring'); })
        .reduce(function(s,x) { return s+(+x.amount||0); }, 0);
      var bonuses = d.bonuses
        .filter(function(b) { return String(b.empId)===String(emp.id) && b.month===month; })
        .reduce(function(s,b) { return s+(+b.amount||0); }, 0);
      var arrears = d.arrears
        .filter(function(a) { return String(a.empId)===String(emp.id) && a.status==='pending'; })
        .reduce(function(s,a) { return s+(+a.amount||0); }, 0);
      var att       = d.attendance.find(function(a) { return String(a.empId)===String(emp.id) && a.month===month; }) || {};
      var dailyRate = basic / workDays;
      var absDed    = +(parseFloat(att.absent||0) * dailyRate).toFixed(2);
      var lateDed   = +(parseFloat(att.late||0) * (dailyRate/8)).toFixed(2);
      var net       = +(gross - gosi - loanDed - extraDed - absDed - lateDed + bonuses + arrears).toFixed(2);

      return { empId:emp.id, name:emp.name, dept:emp.dept,
               basic, gross, gosi, loanDed, extraDed, absDed, lateDed, bonuses, arrears, net };
    });

  return { success: true, month: month, currency: currency, data: results };
}

// ======================================================
//  initSheets — ينشئ الأوراق تلقائياً إن لم تكن موجودة
// ======================================================
function initSheets(ss) {
  var schemas = {
    employees:  ['id','name','natId','dept','pos','basic','housing','transport','phone','other','iban','joinDate','status'],
    loans:      ['id','empId','type','amount','remaining','monthly','date','reason','status'],
    deductions: ['id','empId','type','amount','month','desc','status'],
    arrears:    ['id','empId','type','amount','months','date','desc','status'],
    bonuses:    ['id','empId','type','amount','month','desc','status'],
    attendance: ['id','empId','month','present','absent','late','leaves','approved'],
    settings:   ['key','value']
  };
  Object.keys(schemas).forEach(function(name) {
    if (ss.getSheetByName(name)) return;
    var sh = ss.insertSheet(name);
    sh.appendRow(schemas[name]);
    sh.getRange(1,1,1,schemas[name].length).setBackground('#0f2942').setFontColor('#fff').setFontWeight('bold');
  });
}

// ======================================================
//  jsonOut
// ======================================================
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ======================================================
//  تشغيل مرة واحدة فقط من المحرر لإعداد جدول البيانات
// ======================================================
function setupSpreadsheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  initSheets(ss);
  var sh = ss.getSheetByName('settings');
  if (sh.getLastRow() < 2) {
    [['company','نظام الرواتب'],['currency','ريال'],['gosi',9],['workDays',30]]
      .forEach(function(r) { sh.appendRow(r); });
  }
  SpreadsheetApp.getUi().alert('✅ تم إعداد جدول البيانات بنجاح!');
}
