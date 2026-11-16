/**
 * خادم صفحة تبادل الحصص — يُلصق في مشروع Google Apps Script (مستقل أو مرتبط بجدول).
 *
 * الأعمدة: الطابع الزمني | التاريخ | المعلمة | المعلمة البديلة | الحصص | ملاحظات | المعرف | حصص البديلة (للتبادل) | رمز الطلب | صف المعلمة | صف البديلة
 *  - «الحصص»: حصص المعلمة التي تأخذها البديلة.
 *  - «حصص البديلة»: حصص البديلة التي تأخذها المعلمة مقابلها (فارغة = تغطية فقط).
 *  - «رمز الطلب»: رمز سري يُعطى للمعلمة عند الحفظ ويُحفظ في جهازها، لتعدّل طلبها أو تلغيه.
 *  - «صف المعلمة»: صف حصص المعلمة التي تأخذها البديلة. «صف البديلة»: صف حصص البديلة التي تأخذها المعلمة (في التبادل).
 *
 * الرمز السري للإدارة لا يُكتب هنا، بل في:
 *   إعدادات المشروع ⚙️ ← خصائص النص البرمجي ← إضافة خاصية:  ADMIN_PIN = <الرمز>
 */
const SHEET_NAME = 'التبادل';
const HEADERS = ['الطابع الزمني', 'التاريخ', 'المعلمة', 'المعلمة البديلة', 'الحصص', 'ملاحظات', 'المعرف', 'حصص البديلة (للتبادل)', 'رمز الطلب', 'صف المعلمة', 'صف البديلة'];
const TEXT_FORMATS = [['yyyy-mm-dd hh:mm', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@']];

/**
 * شغّليها مرة واحدة من المحرر (زر «تنفيذ») لمنح الأذونات وإنشاء الجدول.
 * رابط الجدول يظهر في «سجل التنفيذ».
 */
function setup() {
  const ss = getSpreadsheet_();
  getSheet_();
  Logger.log('تم التجهيز ✓ رابط جدول البيانات: ' + ss.getUrl());
  Logger.log(PropertiesService.getScriptProperties().getProperty('ADMIN_PIN')
    ? 'الرمز السري للإدارة مضبوط ✓'
    : '⚠️ لم يُضبط الرمز السري بعد: إعدادات المشروع ⚙️ ← خصائص النص البرمجي ← ADMIN_PIN');
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
  } else if (sh.getRange(1, HEADERS.length).getValue() === '') {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]); // جدول قديم: نضيف العمود الجديد
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

/** كل السجلات مع رقم الصف والرمز (للاستخدام الداخلي فقط) */
function readAll_(sh) {
  const values = sh.getDataRange().getValues();
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (!r[6]) continue;
    out.push({
      row: i + 1,
      token: String(r[8] || ''),
      createdAt: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      date: fmtDate_(r[1]),
      teacher: String(r[2]),
      substitute: String(r[3]),
      periods: parsePeriods_(r[4]),
      returnPeriods: parsePeriods_(r[7]),
      notes: String(r[5] || ''),
      id: String(r[6]),
      cls: String(r[9] || ''),
      retCls: String(r[10] || '')
    });
  }
  return out;
}

/** ما يُرسل للصفحة: بدون الرمز ورقم الصف */
function publicRow_(r) {
  return { createdAt: r.createdAt, date: r.date, teacher: r.teacher, substitute: r.substitute,
           periods: r.periods, returnPeriods: r.returnPeriods, notes: r.notes, id: r.id,
           cls: r.cls, retCls: r.retCls };
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

/**
 * التحقق من رمز الإدارة، مع إيقاف مؤقت بعد محاولات خاطئة كثيرة.
 * يرجع: 'ok' | 'wrong' | 'locked' | 'unset'
 */
function checkPin_(pin) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PIN');
  if (!real) return 'unset';
  const cache = CacheService.getScriptCache();
  const fails = Number(cache.get('pin_fails') || 0);
  if (fails >= 10) return 'locked';
  if (String(pin || '') === String(real)) return 'ok';
  cache.put('pin_fails', String(fails + 1), 600);
  return 'wrong';
}

const PIN_ERRORS = {
  wrong: 'الرمز السري غير صحيح.',
  locked: 'محاولات خاطئة كثيرة، حاولي بعد ١٠ دقائق.',
  unset: 'لم يُضبط الرمز السري في إعدادات المشروع بعد.'
};

/** الرابط المباشر لا يعرض أي بيانات */
function doGet() {
  return json_({ ok: true, message: 'خادم تبادل الحصص يعمل ✓' });
}

/**
 * POST (نص JSON) — الحقل action:
 *  add     {teacher, substitute, date, periods, cls, returnPeriods, retCls, swap, notes}  → {ok, id, token}
 *  mine    {items:[{id, token}]}                                            → {ok, rows}
 *  update  نفس حقول add + {id, token | pin}
 *  delete  {id, token | pin}
 *  login   {pin}
 *  list    {pin, date}                                                       → {ok, rows}
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const action = body.action || 'add';
    const sh = getSheet_();

    if (action === 'login') {
      const st = checkPin_(body.pin);
      return json_(st === 'ok' ? { ok: true } : { ok: false, error: PIN_ERRORS[st] });
    }

    if (action === 'list') {
      const st = checkPin_(body.pin);
      if (st !== 'ok') return json_({ ok: false, error: PIN_ERRORS[st], auth: true });
      const date = clean_(body.date, 10);
      return json_({ ok: true, rows: readAll_(sh).filter(r => !date || r.date === date).map(publicRow_) });
    }

    if (action === 'mine') {
      const want = {};
      (body.items || []).slice(0, 200).forEach(it => { if (it && it.id && it.token) want[String(it.id)] = String(it.token); });
      const rows = readAll_(sh).filter(r => r.token && want[r.id] === r.token).map(publicRow_);
      return json_({ ok: true, rows: rows });
    }

    // التعديل والحذف: صاحبة الطلب (بالرمز المحفوظ في جهازها) أو الإدارة (بالرمز السري)
    let target = null;
    if (action === 'update' || action === 'delete') {
      target = readAll_(sh).find(r => r.id === String(body.id || ''));
      if (!target) return json_({ ok: false, error: 'لم يُعثر على هذا الطلب، ربما حُذف مسبقًا.' });
      const ownerOk = target.token && body.token && String(body.token) === target.token;
      if (!ownerOk) {
        const st = body.pin ? checkPin_(body.pin) : 'wrong';
        if (st !== 'ok') return json_({ ok: false, error: body.pin ? PIN_ERRORS[st] : 'لا يمكنك تعديل هذا الطلب.' });
      }
      if (action === 'delete') {
        sh.deleteRow(target.row);
        return json_({ ok: true });
      }
    } else if (action !== 'add') {
      return json_({ ok: false, error: 'طلب غير معروف.' });
    }

    const teacher = clean_(body.teacher, 80);
    const substitute = clean_(body.substitute, 80);
    const date = clean_(body.date, 10);
    const notes = clean_(body.notes, 300);
    const periods = cleanPeriods_(body.periods);
    const returnPeriods = cleanPeriods_(body.returnPeriods);
    const cls = clean_(body.cls, 30);
    const retCls = returnPeriods.length ? clean_(body.retCls, 30) : '';

    if (!teacher || !substitute) return json_({ ok: false, error: 'الرجاء كتابة اسم المعلمة واسم المعلمة البديلة.' });
    if (teacher === substitute) return json_({ ok: false, error: 'لا يمكن أن تكون المعلمة البديلة هي نفس المعلمة.' });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return json_({ ok: false, error: 'التاريخ غير صحيح.' });
    if (!periods.length) return json_({ ok: false, error: 'الرجاء اختيار حصة واحدة على الأقل.' });
    if (body.swap && !returnPeriods.length) return json_({ ok: false, error: 'في التبادل اختاري حصص المعلمة البديلة التي ستأخذينها.' });
    if (!cls) return json_({ ok: false, error: 'الرجاء كتابة الصف.' });
    if (returnPeriods.length && !retCls) return json_({ ok: false, error: 'الرجاء كتابة صف المعلمة البديلة.' });

    // منع التعارض في نفس اليوم (مع تجاهل الطلب نفسه عند التعديل)
    const existing = [];
    readAll_(sh).filter(r => r.date === date && (!target || r.id !== target.id))
      .forEach(r => existing.push(...assignments_(r)));
    const busy = [];
    assignments_({ teacher: teacher, substitute: substitute, periods: periods, returnPeriods: returnPeriods })
      .forEach(a => existing.forEach(x => {
        if (x.p !== a.p) return;
        if (x.coverer === a.coverer) busy.push('المعلمة ' + a.coverer + ' تغطي حصة أخرى في الحصة ' + a.p);
        if (x.owner === a.owner) busy.push('الحصة ' + a.p + ' للمعلمة ' + a.owner + ' مسجّلة مسبقًا');
      }));
    if (busy.length) return json_({ ok: false, error: 'يوجد تعارض: ' + [...new Set(busy)].join('، ') });

    const id = target ? target.id : Utilities.getUuid();
    const token = target ? (target.token || Utilities.getUuid()) : Utilities.getUuid();
    const row = target ? target.row : sh.getLastRow() + 1;
    // نكتب الصف كنص صريح حتى لا يحوّل الجدول التاريخ أو الحصص تلقائيًا
    const range = sh.getRange(row, 1, 1, HEADERS.length);
    range.setNumberFormats(TEXT_FORMATS);
    range.setValues([[new Date(), date, teacher, substitute, periods.join(','), notes, id, returnPeriods.join(','), token, cls, retCls]]);
    return json_({ ok: true, id: id, token: token });
  } catch (err) {
    return json_({ ok: false, error: 'حدث خطأ في الخادم: ' + err });
  } finally {
    lock.releaseLock();
  }
}
