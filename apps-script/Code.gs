/**
 * خادم صفحة تبادل الحصص — يُلصق في مشروع Google Apps Script (مستقل أو مرتبط بجدول).
 * الأعمدة: الطابع الزمني | التاريخ | المعلمة | المعلمة البديلة | الحصص | ملاحظات | المعرف | حصص البديلة (للتبادل)
 *  - «الحصص»: حصص المعلمة التي تأخذها البديلة.
 *  - «حصص البديلة»: حصص البديلة التي تأخذها المعلمة مقابلها (فارغة = تغطية فقط).
 */
const SHEET_NAME = 'التبادل';
const HEADERS = ['الطابع الزمني', 'التاريخ', 'المعلمة', 'المعلمة البديلة', 'الحصص', 'ملاحظات', 'المعرف', 'حصص البديلة (للتبادل)'];

/**
 * شغّليها مرة واحدة من المحرر (زر «تنفيذ») لمنح الأذونات وإنشاء الجدول.
 * رابط الجدول يظهر في «سجل التنفيذ».
 */
function setup() {
  const ss = getSpreadsheet_();
  getSheet_();
  Logger.log('تم التجهيز ✓ رابط جدول البيانات: ' + ss.getUrl());
}

/** يعمل سواء كان المشروع مرتبطًا بجدول أو مستقلًا (ينشئ جدولًا باسم «تبادل الحصص» أول مرة). */
function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  const ss = SpreadsheetApp.create('تبادل الحصص');
  ss.getSheets()[0].setName(SHEET_NAME);
  props.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function getSheet_() {
  const ss = getSpreadsheet_();
  const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    sh.getRange('B:B').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('@');
    sh.getRange('H:H').setNumberFormat('@');
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function fmtDate_(v) {
  // لو حوّل الجدول النص إلى تاريخ، نقرؤه بتوقيت المشروع نفسه حتى لا يتغيّر اليوم
  return v instanceof Date ? Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(v || '').trim();
}

function readRows_() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues().slice(1);
  return values.filter(r => r[6]).map(r => ({
    createdAt: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
    date: fmtDate_(r[1]),
    teacher: String(r[2]),
    substitute: String(r[3]),
    periods: parsePeriods_(r[4]),
    returnPeriods: parsePeriods_(r[7]),
    notes: String(r[5] || ''),
    id: String(r[6])
  }));
}

function parsePeriods_(v) {
  // يقبل «1,4» وأيضًا «1.4» لو حوّلها الجدول إلى رقم عشري
  return String(v == null ? '' : v).split(/[^0-9]+/).map(Number).filter(p => p >= 1 && p <= 7);
}

function cleanPeriods_(list) {
  return [...new Set((list || []).map(Number))].filter(p => p >= 1 && p <= 7).sort((a, b) => a - b);
}

/** كل حصة في السجل = من يغطّي (coverer) حصة من (owner) */
function assignments_(r) {
  return r.periods.map(p => ({ p: p, coverer: r.substitute, owner: r.teacher }))
    .concat(r.returnPeriods.map(p => ({ p: p, coverer: r.teacher, owner: r.substitute })));
}

function clean_(s, max) {
  return String(s || '').replace(/\s+/g, ' ').trim().slice(0, max);
}

/** GET ?date=YYYY-MM-DD  → قائمة التبادلات لذلك اليوم (أو كلها بدون تاريخ) */
function doGet(e) {
  const date = e && e.parameter && e.parameter.date;
  let rows = readRows_();
  if (date) rows = rows.filter(r => r.date === date);
  return json_({ ok: true, rows: rows });
}

/** يرجع رقم الصف في الجدول لسجل بهذا المعرف، أو 0 إن لم يوجد */
function findRow_(sh, id) {
  if (!id || sh.getLastRow() < 2) return 0;
  const ids = sh.getRange(2, 7, sh.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

/**
 * POST (نص JSON):
 *  إضافة: {teacher, substitute, date, periods:[1..7], returnPeriods:[1..7], swap, notes}
 *  تعديل: نفس الحقول + {action:'update', id}
 *  حذف:   {action:'delete', id}
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action || 'add';
    const sh = getSheet_();

    if (action === 'delete') {
      const row = findRow_(sh, body.id);
      if (!row) return json_({ ok: false, error: 'لم يُعثر على هذا السجل، ربما حُذف مسبقًا.' });
      sh.deleteRow(row);
      return json_({ ok: true });
    }

    const teacher = clean_(body.teacher, 80);
    const substitute = clean_(body.substitute, 80);
    const date = clean_(body.date, 10);
    const notes = clean_(body.notes, 300);
    const periods = cleanPeriods_(body.periods);
    const returnPeriods = cleanPeriods_(body.returnPeriods);

    if (!teacher || !substitute) return json_({ ok: false, error: 'الرجاء كتابة اسم المعلمة واسم المعلمة البديلة.' });
    if (teacher === substitute) return json_({ ok: false, error: 'لا يمكن أن تكون المعلمة البديلة هي نفس المعلمة.' });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: 'التاريخ غير صحيح.' });
    if (!periods.length) return json_({ ok: false, error: 'الرجاء اختيار حصة واحدة على الأقل.' });
    if (body.swap && !returnPeriods.length) return json_({ ok: false, error: 'في التبادل اختاري حصص المعلمة البديلة التي ستأخذينها.' });

    let row = 0;
    if (action === 'update') {
      row = findRow_(sh, body.id);
      if (!row) return json_({ ok: false, error: 'لم يُعثر على هذا السجل، ربما حُذف مسبقًا.' });
    }

    // منع التعارض في نفس اليوم (مع تجاهل السجل نفسه عند التعديل)
    const existing = [];
    readRows_().filter(r => r.date === date && r.id !== String(body.id || ''))
      .forEach(r => existing.push(...assignments_(r)));
    const busy = [];
    assignments_({ teacher: teacher, substitute: substitute, periods: periods, returnPeriods: returnPeriods })
      .forEach(a => existing.forEach(x => {
        if (x.p !== a.p) return;
        if (x.coverer === a.coverer) busy.push('المعلمة ' + a.coverer + ' تغطي حصة أخرى في الحصة ' + a.p);
        if (x.owner === a.owner) busy.push('الحصة ' + a.p + ' للمعلمة ' + a.owner + ' مسجّلة مسبقًا');
      }));
    if (busy.length) return json_({ ok: false, error: 'يوجد تعارض: ' + [...new Set(busy)].join('، ') });

    const id = action === 'update' ? String(body.id) : Utilities.getUuid();
    if (!row) row = sh.getLastRow() + 1;
    // نكتب الصف كنص صريح حتى لا يحوّل الجدول التاريخ أو الحصص تلقائيًا
    const range = sh.getRange(row, 1, 1, HEADERS.length);
    range.setNumberFormats([['yyyy-mm-dd hh:mm', '@', '@', '@', '@', '@', '@', '@']]);
    range.setValues([[new Date(), date, teacher, substitute, periods.join(','), notes, id, returnPeriods.join(',')]]);
    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: 'حدث خطأ في الخادم: ' + err });
  } finally {
    lock.releaseLock();
  }
}
