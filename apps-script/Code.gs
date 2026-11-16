/**
 * خادم صفحة تبادل الحصص — يُلصق في Google Apps Script المرتبط بجدول Google Sheets.
 * الأعمدة: الطابع الزمني | التاريخ | المعلمة | المعلمة البديلة | الحصص | ملاحظات | المعرف | حصص البديلة (للتبادل)
 *  - «الحصص»: حصص المعلمة التي تأخذها البديلة.
 *  - «حصص البديلة»: حصص البديلة التي تأخذها المعلمة مقابلها (فارغة = تغطية فقط).
 */
const SHEET_NAME = 'التبادل';
const HEADERS = ['الطابع الزمني', 'التاريخ', 'المعلمة', 'المعلمة البديلة', 'الحصص', 'ملاحظات', 'المعرف', 'حصص البديلة (للتبادل)'];
const TZ = 'Asia/Riyadh';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
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
  return v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '');
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
  return String(v == null ? '' : v).split(',').map(Number).filter(p => p >= 1 && p <= 7);
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

/** POST (نص JSON): {teacher, substitute, date, periods:[1..7], returnPeriods:[1..7], notes} */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
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

    // منع التعارض في نفس اليوم: معلمة تغطي حصتين في نفس الوقت، أو حصة تُعطى مرتين
    const existing = [];
    readRows_().filter(r => r.date === date).forEach(r => existing.push(...assignments_(r)));
    const busy = [];
    assignments_({ teacher: teacher, substitute: substitute, periods: periods, returnPeriods: returnPeriods })
      .forEach(a => existing.forEach(x => {
        if (x.p !== a.p) return;
        if (x.coverer === a.coverer) busy.push('المعلمة ' + a.coverer + ' تغطي حصة أخرى في الحصة ' + a.p);
        if (x.owner === a.owner) busy.push('الحصة ' + a.p + ' للمعلمة ' + a.owner + ' مسجّلة مسبقًا');
      }));
    if (busy.length) return json_({ ok: false, error: 'يوجد تعارض: ' + [...new Set(busy)].join('، ') });

    const id = Utilities.getUuid();
    getSheet_().appendRow([new Date(), date, teacher, substitute, periods.join(','), notes, id, returnPeriods.join(',')]);
    return json_({ ok: true, id: id });
  } catch (err) {
    return json_({ ok: false, error: 'حدث خطأ في الخادم: ' + err });
  } finally {
    lock.releaseLock();
  }
}
