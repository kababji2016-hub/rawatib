// ============================================================
//  نظام إدارة الرواتب — Google Apps Script
//  الملف: Code.gs
// ============================================================

const SHEET_ID = '1cNlObaJBupKzCpq8AzluPSEv5AGda_URDaBsh6bTnQM';
const ALLOWED_EMAILS = []; // فارغة = جميع حسابات Google

// ============================================================
//  تعريف الأوراق — العربية أولاً، ثم الإنجليزية كاحتياط
//  map: اسم العمود العربي → اسم الحقل الإنجليزي في الواجهة
// ============================================================
var SHEET_DEFS = {
  employees: {
    aliases: ['موظفين', 'employees'],
    map: {
      'معرف':           'id',
      'الاسم':          'name',
      'رقم_الهوية':     'natId',
      'رقم الهوية':     'natId',
      'الجنسية':        'nationality',
      'المسمى':         'pos',
      'الفرع':          'dept',
      'الراتب_الأساسي': 'basic',
      'الراتب الأساسي': 'basic',
      'لراتب_الأساسي':  'basic',
      'الحالة':         'status',
      'بواسطة':         'addedBy'
    }
  },
  loans: {
    aliases: ['سلف_وقروض', 'loans'],
    map: {
      'معرف_السلفة':    'id',
      'معرف السلفة':    'id',
      'معرف_الموظف':    'empId',
      'معرف الموظف':    'empId',
      'اسم_الموظف':     'empName',
      'اسم الموظف':     'empName',
      'الفرع':          'dept',
      'النوع':          'type',
      'المبلغ_الأصلي':  'amount',
      'المبلغ الأصلي':  'amount',
      'المبلغ':         'amount',
      'المتبقي':        'remaining',
      'القسط_الشهري':   'monthly',
      'القسط الشهري':   'monthly',
      'التاريخ':        'date',
      'الحالة':         'status',
      'البيان':         'reason',
      'بواسطة':         'addedBy'
    }
  },
  deductions: {
    aliases: ['خصومات', 'deductions'],
    map: {
      'معرف_الخصم':     'id',
      'معرف الخصم':     'id',
      'معرف_الموظف':    'empId',
      'معرف الموظف':    'empId',
      'اسم_الموظف':     'empName',
      'اسم الموظف':     'empName',
      'الفرع':          'dept',
      'نوع_الخصم':      'type',
      'نوع الخصم':      'type',
      'المبلغ':         'amount',
      'الشهر':          'month',
      'البيان':         'desc',
      'الحالة':         'status',
      'بواسطة':         'addedBy'
    }
  },
  arrears:    { aliases: ['arrears'] }
};

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
//  doPost — REST API
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
function getAllData()          { if (!isAuthorized()) throw new Error('غير مصرح'); return dispatch('getAll', {}); }
function saveRecord(p)        { if (!isAuthorized()) throw new Error('غير مصرح'); return dispatch('syncRecord', p); }
function saveSettingsData(p)  { if (!isAuthorized()) throw new Error('غير مصرح'); return dispatch('saveSettings', p); }
function calcPayrollData(m)   { if (!isAuthorized()) throw new Error('غير مصرح'); return dispatch('calcPayroll', { month: m }); }

// ======================================================
//  Dispatcher
// ======================================================
function dispatch(action, payload) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
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
  try { return ALLOWED_EMAILS.indexOf(Session.getEffectiveUser().getEmail()) !== -1; }
  catch(e) { return false; }
}

// ======================================================
//  getAll — يقرأ من الأوراق العربية أولاً
// ======================================================
function handleGetAll(ss) {
  var result = {};

  Object.keys(SHEET_DEFS).forEach(function(key) {
    var def = SHEET_DEFS[key];
    var sh  = null;

    // ابحث عن أول ورقة بها بيانات
    for (var i = 0; i < def.aliases.length; i++) {
      var c = ss.getSheetByName(def.aliases[i]);
      if (c && c.getLastRow() >= 2) { sh = c; break; }
    }
    if (!sh) { result[key] = []; return; }

    var data    = sh.getDataRange().getValues();
    var rawHdrs = data[0];
    var map     = def.map || {};

    // حوّل أسماء الأعمدة العربية إلى إنجليزية
    var headers = rawHdrs.map(function(h) {
      var t = String(h).trim();
      return map[t] || t;
    });

    result[key] = data.slice(1)
      .filter(function(r) { return r[0] !== '' && r[0] !== null && r[0] !== undefined; })
      .map(function(r) {
        var o = {};
        headers.forEach(function(h, i) { o[h] = (r[i] == null) ? '' : r[i]; });
        return o;
      });
  });

  // الإعدادات
  var sh = ss.getSheetByName('settings');
  var settings = {};
  if (sh && sh.getLastRow() > 0) {
    sh.getDataRange().getValues().forEach(function(r) { if (r[0]) settings[String(r[0])] = r[1]; });
  }
  result.settings = settings;
  return { success: true, data: result };
}

// ======================================================
//  syncRecord — يحفظ في الورقة الصحيحة
// ======================================================
function handleSyncRecord(ss, payload) {
  var logicalSheet = payload.sheet;
  var record       = payload.record;
  var op           = payload.operation;
  if (!logicalSheet || !record) return { success: false, error: 'بيانات ناقصة' };

  // جد الورقة الفعلية
  var def = SHEET_DEFS[logicalSheet] || { aliases: [logicalSheet], map: {} };
  var sh  = null;
  for (var i = 0; i < def.aliases.length; i++) {
    var c = ss.getSheetByName(def.aliases[i]);
    if (c) { sh = c; break; }
  }
  if (!sh) {
    sh = ss.insertSheet(logicalSheet);
  }

  var lastRow = sh.getLastRow();
  var lastCol = sh.getLastColumn();
  var map     = def.map || {};

  // ---- حذف ----
  if (op === 'delete') {
    if (lastRow >= 2) {
      var ids = sh.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (String(ids[i][0]) === String(record.id)) { sh.deleteRow(i + 2); break; }
      }
    }
    return { success: true };
  }

  // ---- Upsert ----
  var rawHeaders;
  if (lastRow === 0 || lastCol === 0) {
    // ورقة فارغة — أنشئ عناوين إنجليزية
    rawHeaders = Object.keys(record);
    sh.appendRow(rawHeaders);
    sh.getRange(1,1,1,rawHeaders.length).setBackground('#0f2942').setFontColor('#fff').setFontWeight('bold');
    lastRow = 1;
  } else {
    rawHeaders = sh.getRange(1,1,1,lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
  }

  // بناء القيم: لكل عمود في الورقة، جد القيمة المقابلة من الـ record
  var values = rawHeaders.map(function(rawH) {
    var englishKey = map[rawH] || rawH; // العمود العربي → مفتاح إنجليزي
    var val = record[englishKey];
    if (val === undefined) val = record[rawH]; // جرّب المفتاح الخام
    return val !== undefined ? val : '';
  });

  // هل السجل موجود؟
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
//  calcPayroll — احتساب الرواتب
// ======================================================
function handleCalcPayroll(ss, payload) {
  var month = payload.month;
  if (!month) return { success: false, error: 'الشهر مطلوب' };

  var d        = handleGetAll(ss).data;
  var workDays = parseInt(d.settings.workDays  || 30);
  var currency = d.settings.currency || 'ريال';

  var results = d.employees
    .filter(function(e) { return String(e.status).toLowerCase() === 'active'; })
    .map(function(emp) {
      var basic     = +emp.basic     || 0;
      var housing   = +emp.housing   || 0;
      var transport = +emp.transport || 0;
      var phone     = +emp.phone     || 0;
      var other     = +emp.other     || 0;
      var gross     = basic + housing + transport + phone + other;

      var loanDed = d.loans
        .filter(function(l) {
          return String(l.empId) === String(emp.id) &&
                 String(l.status).toLowerCase() === 'active';
        })
        .reduce(function(s, l) { return s + (+l.monthly || 0); }, 0);

      var extraDed = d.deductions
        .filter(function(x) {
          if (String(x.empId) !== String(emp.id)) return false;
          if (x.status === 'recurring') return true;
          return String(x.month || '').substring(0, 7) === month;
        })
        .reduce(function(s, x) { return s + (+x.amount || 0); }, 0);

      var arrears = d.arrears
        .filter(function(a) {
          return String(a.empId) === String(emp.id) && a.status === 'pending';
        })
        .reduce(function(s, a) { return s + (+a.amount || 0); }, 0);

      var net = +(gross - loanDed - extraDed + arrears).toFixed(2);

      return {
        empId: emp.id, name: emp.name, dept: emp.dept,
        basic, gross, loanDed, extraDed, arrears, net
      };
    });

  return { success: true, month: month, currency: currency, data: results };
}

// ======================================================
//  jsonOut
// ======================================================
function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ======================================================
//  تشغيل مرة واحدة من المحرر لإنشاء أوراق attendance/bonuses/arrears
// ======================================================
function setupSpreadsheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var extra = {
    arrears:  ['id','empId','type','amount','months','date','desc','status'],
    settings: ['key','value']
  };
  Object.keys(extra).forEach(function(name) {
    if (ss.getSheetByName(name)) return;
    var sh = ss.insertSheet(name);
    sh.appendRow(extra[name]);
    sh.getRange(1,1,1,extra[name].length).setBackground('#0f2942').setFontColor('#fff').setFontWeight('bold');
  });
  var sh = ss.getSheetByName('settings');
  if (sh && sh.getLastRow() < 2) {
    [['company','نظام الرواتب'],['currency','ريال'],['gosi',9],['workDays',30]]
      .forEach(function(r) { sh.appendRow(r); });
  }
  SpreadsheetApp.getUi().alert('✅ تم الإعداد بنجاح!');
}
