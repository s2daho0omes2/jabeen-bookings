/**
 * مرآة حية لحجوزات مشروع جبين في Google Sheets.
 * المصدر الوحيد للحقيقة هو قاعدة Firebase الخاصة بالموقع؛ هذا السكربت يقرأ منها فقط ويعيد بناء الأوراق.
 * - الموقع يستدعي doPost بعد كل إجراء (حجز / اعتماد / رفض / إلغاء) فتتحدث الأوراق فوراً.
 * - مؤقت زمني يعيد المزامنة كل 5 دقائق كشبكة أمان.
 */
var FIREBASE = 'https://jabeen-bookings-default-rtdb.europe-west1.firebasedatabase.app/jabeenBookings';
var STATUS_AR = { available: 'متاحة', pending: 'مبدئي', approved: 'معتمد', rejected: 'مرفوض', cancelled: 'ملغي' };
var ROLE_AR = { maker: 'مُدخل', checker: 'مُراجع', admin: 'مدير' };

/* جدول التخصيص الخاص بالمستخدم: ورقة لكل حي، العمود A رقم القطعة والعمود H الحالة (قائمة منسدلة) */
var ALLOC_SHEET_ID = '1Jobrngw85gUFGG9YlnYW2_rpIikwzAtEluM5lQrM5gA';
var ALLOC_TABS = { 'العزيزية': 'العزيزية', 'النخيل': 'النخيل' };
var ALLOC_STATUS = { available: 'متاح', pending: 'قيد الحجز', approved: 'محجوز' };
var ALLOC_COL = 8; // H

function doPost(e) { return respond_(syncAll()); }
function doGet(e)  { return respond_(syncAll()); }
function respond_(r) { return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON); }

/** شغّلها مرة واحدة من المحرر لإنشاء مؤقت كل 5 دقائق */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncAll').timeBased().everyMinutes(5).create();
  return syncAll();
}

function fetchJson_(url) {
  var opts = { muteHttpExceptions: true };
  // قاعدة البيانات مقفلة: السكربت يقرأ بصلاحية حساب Google المالك لمشروع Firebase
  if (url.indexOf(FIREBASE) === 0) opts.headers = { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() };
  var r = UrlFetchApp.fetch(url, opts);
  if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode() + ' ' + url);
  return JSON.parse(r.getContentText());
}

function syncAll() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'busy' };
  try {
    var bookings = fetchJson_(FIREBASE + '/bookings.json') || {};
    var audit = fetchJson_(FIREBASE + '/audit.json') || {};
    var mo = fetchJson_(FIREBASE + '/master.json') || {};
    var master = Object.keys(mo).sort().map(function (k) { return mo[k]; });
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // ---------- الحجوزات ----------
    var list = Object.keys(bookings).map(function (k) { var b = bookings[k]; b._id = k; return b; });
    list.sort(function (a, b) { return String(a.refNo || '').localeCompare(String(b.refNo || '')); });

    // جدول التخصيص أولاً (الأسرع والأهم): حالة كل قطعة في العمود H
    var active0 = {};
    list.forEach(function (b) { if (b.status === 'pending' || b.status === 'approved') active0[b.unitId] = b; });
    var allocChanged = -1;
    try { allocChanged = syncAllocation_(master, active0); } catch (err) { console.error('allocation sheet: ' + err); }
    var bHead = ['كود الحجز', 'الوحدة', 'الحي', 'رقم القطعة', 'البلوك', 'النموذج', 'الشركة', 'السعر', 'اسم العميل', 'رقم الهوية', 'الرقم الوظيفي', 'الجوال', 'طريقة الدفع', 'تاريخ الحجز', 'الحالة', 'المُدخل', 'وقت الإدخال', 'المُراجع', 'وقت المراجعة', 'سبب الرفض/الإلغاء', 'ملاحظات', 'المصدر'];
    var bRows = list.map(function (b) {
      return [b.refNo || '', b.unitId || '', b.district || '', b.plot || '', b.block || '', b.model || '', b.company || '', b.price || '',
        b.name || '', b.natId ? "'" + b.natId : '', b.empNo || '', b.phone ? "'" + b.phone : '', b.payMethod || '', b.bookingDate || '',
        STATUS_AR[b.status] || b.status || '', b.maker || '', ts_(b.createdAt), b.checker || '', ts_(b.checkedAt), b.rejectReason || '', b.notes || '', b.source || ''];
    });
    writeSheet_(ss, 'الحجوزات', bHead, bRows);

    // ---------- الوحدات: كل الفلل (506) وعلى كل صف بيانات حجزها إن وُجد ----------
    var active = {};
    list.forEach(function (b) { if (b.status === 'pending' || b.status === 'approved') active[b.unitId] = b; });
    var uHead = ['رقم الوحدة', 'الحي', 'رقم القطعة', 'البلوك', 'النموذج', 'التخصيص', 'غرف', 'حمامات', 'مساحة الأرض', 'مساحة البناء', 'السعر',
      'الحالة', 'كود الحجز', 'اسم العميل', 'رقم الهوية', 'الرقم الوظيفي', 'الجوال', 'شركة الحجز', 'طريقة الدفع', 'تاريخ الحجز', 'المُدخل', 'وقت الإدخال', 'المُراجع', 'وقت المراجعة', 'ملاحظات'];
    var uRows = master.map(function (u) {
      var b = active[u.id];
      return [u.id, u.district, u.plot, u.block, u.model, u.company, u.bed, u.bath, u.plotArea, u.builtArea, u.price,
        b ? STATUS_AR[b.status] : 'متاحة',
        b ? b.refNo || '' : '', b ? b.name || '' : '', b && b.natId ? "'" + b.natId : '', b ? b.empNo || '' : '', b && b.phone ? "'" + b.phone : '',
        b ? b.company || '' : '', b ? b.payMethod || '' : '', b ? b.bookingDate || '' : '', b ? b.maker || '' : '', b ? ts_(b.createdAt) : '',
        b ? b.checker || '' : '', b ? ts_(b.checkedAt) : '', b ? b.notes || '' : ''];
    });
    writeSheet_(ss, 'الوحدات', uHead, uRows);
    if (uRows.length) {
      var ush = ss.getSheetByName('الوحدات'), stRange = ush.getRange(2, 12, uRows.length, 1);
      ush.clearConditionalFormatRules();
      ush.setConditionalFormatRules([
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('متاحة').setBackground('#e6f3ec').setFontColor('#1e6b45').setRanges([stRange]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('معتمد').setBackground('#fdecea').setFontColor('#b3261e').setRanges([stRange]).build(),
        SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('مبدئي').setBackground('#fff5d6').setFontColor('#9a6b00').setRanges([stRange]).build()
      ]);
    }

    // ---------- سجل الحركات ----------
    var aList = Object.keys(audit).map(function (k) { return audit[k]; }).sort(function (a, b) { return (b.at || 0) - (a.at || 0); });
    writeSheet_(ss, 'سجل الحركات', ['الوقت', 'من', 'الدور', 'الحركة', 'الوحدة', 'تفاصيل'],
      aList.map(function (a) { return [ts_(a.at), a.who || '', ROLE_AR[a.role] || a.role || '', a.action || '', a.unitId || '', a.details || '']; }));

    // ---------- ملخص ----------
    var companies = {};
    master.forEach(function (u) {
      var c = u.company || '';
      companies[c] = companies[c] || { total: 0, available: 0, pending: 0, approved: 0 };
      companies[c].total++;
      var b = active[u.id];
      companies[c][b ? b.status : 'available']++;
    });
    var sRows = Object.keys(companies).map(function (c) { var s = companies[c]; return [c, s.total, s.available, s.pending, s.approved]; });
    var tot = sRows.reduce(function (t, r) { return [t[0], t[1] + r[1], t[2] + r[2], t[3] + r[3], t[4] + r[4]]; }, ['الإجمالي', 0, 0, 0, 0]);
    sRows.push(tot);
    sRows.push([]);
    sRows.push(['آخر مزامنة', Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm:ss')]);
    writeSheet_(ss, 'ملخص', ['التخصيص', 'إجمالي الوحدات', 'متاحة', 'مبدئي', 'معتمد'], sRows);

    // ترتيب الأوراق وحذف الورقة الافتراضية الفارغة
    var ORDER = ['الوحدات', 'الحجوزات', 'سجل الحركات', 'ملخص'];
    ORDER.forEach(function (n, i) { ss.setActiveSheet(ss.getSheetByName(n)); ss.moveActiveSheet(i + 1); });
    ss.getSheets().forEach(function (sh) { if (ORDER.indexOf(sh.getName()) < 0 && sh.getLastRow() === 0) ss.deleteSheet(sh); });
    ss.setActiveSheet(ss.getSheetByName('الوحدات'));

    return { ok: true, bookings: list.length, units: master.length, allocChanged: allocChanged, at: new Date().toISOString() };
  } finally { lock.releaseLock(); }
}

/** يحدّث العمود H في جدول التخصيص لكل قطعة حسب حالتها في الموقع. يعيد عدد الخلايا التي تغيّرت. */
function syncAllocation_(master, active) {
  var ss = SpreadsheetApp.openById(ALLOC_SHEET_ID);
  var byKey = {};
  master.forEach(function (u) { byKey[u.district + '|' + parseInt(u.plot, 10)] = u; });
  var changed = 0;
  Object.keys(ALLOC_TABS).forEach(function (district) {
    var sh = ss.getSheetByName(ALLOC_TABS[district]);
    if (!sh) return;
    var last = sh.getLastRow();
    if (last < 2) return;
    var plots = sh.getRange(2, 1, last - 1, 1).getValues();
    var cur = sh.getRange(2, ALLOC_COL, last - 1, 1).getValues();
    var out = cur.map(function (r, i) {
      var p = parseInt(plots[i][0], 10);
      var u = p ? byKey[district + '|' + p] : null;
      if (!u) return [r[0]];
      var b = active[u.id];
      return [b ? ALLOC_STATUS[b.status] : ALLOC_STATUS.available];
    });
    var diff = out.filter(function (r, i) { return r[0] !== cur[i][0]; }).length;
    if (diff) {
      // تأكد أن القائمة المنسدلة تقبل القيم الثلاث قبل الكتابة
      var rng = sh.getRange(2, ALLOC_COL, out.length, 1);
      rng.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['متاح', 'قيد الحجز', 'محجوز'], true).setAllowInvalid(false).build());
      rng.setValues(out);
      changed += diff;
    }
    // تنسيق شرطي أصفر لـ "قيد الحجز" إن لم يكن موجوداً
    var rules = sh.getConditionalFormatRules();
    var has = rules.some(function (r) { var c = r.getBooleanCondition(); return c && c.getCriteriaValues().indexOf('قيد الحجز') >= 0; });
    if (!has) {
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('قيد الحجز').setBackground('#ffe59a').setFontColor('#6b4e00').setRanges([sh.getRange('H2:H')]).build());
      sh.setConditionalFormatRules(rules);
    }
  });
  return changed;
}

function writeSheet_(ss, name, head, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  var data = [head].concat(rows.length ? rows : []);
  var width = head.length;
  data = data.map(function (r) { r = r.slice(0, width); while (r.length < width) r.push(''); return r; });
  sh.getRange(1, 1, data.length, width).setValues(data);
  sh.setRightToLeft(true);
  sh.setFrozenRows(1);
  var hr = sh.getRange(1, 1, 1, width);
  hr.setFontWeight('bold').setBackground('#1f4d3a').setFontColor('#ffffff');
  if (sh.getFilter()) sh.getFilter().remove();
  if (data.length > 1) sh.getRange(1, 1, data.length, width).createFilter();
  sh.autoResizeColumns(1, width);
}

function ts_(ms) { return ms ? Utilities.formatDate(new Date(ms), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm') : ''; }
