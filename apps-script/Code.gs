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

/* الإقرار والتعهد الإلكتروني: الرموز تُخزَّن مشفّرة خارج jabeenBookings فلا يصل إليها أي موظف */
var DB_ROOT = FIREBASE.replace(/\/jabeenBookings$/, '');
var SECRETS = DB_ROOT + '/pledgeSecrets';
var FB_API_KEY = 'AIzaSyBs6zMzOL4ggegfLYIPgSb_RDqYXNf1cuo';
var SUPER_ADMIN = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || ''; // بريد المدير للرسائل التجريبية والتقارير؛ يُضبط في Script Properties
var PLEDGE_TTL_MS = 5 * 60 * 1000;   // صلاحية الرمز
var PLEDGE_RESEND_MS = 60 * 1000;    // أقل فاصل بين إرسالين
var PLEDGE_MAX_ATTEMPTS = 5;
var MAIL_NAME = 'حجوزات مشروع جبين';
var MAIL_FROM = 'noreply@jabeen-bookings.online'; // عبر Resend بعد توثيق الدومين
var MAIL_REPLY = 'garawias@jyic.com.sa';            // الردود تصل هنا
var WORK_START = 7, WORK_END = 17;                   // ساعات العمل (توقيت الرياض) لتقرير الإجراءات خارج الدوام
var UNCONFIRMED_HOURS = 48;

/** إرسال موحّد: عبر Resend من دومين المشروع إذا وُجد المفتاح RESEND_KEY في Script Properties، وإلا من حساب السكربت (احتياط) */
function sendMail_(to, subject, html) {
  var key = PropertiesService.getScriptProperties().getProperty('RESEND_KEY');
  if (key) {
    var r = UrlFetchApp.fetch('https://api.resend.com/emails', { method: 'post', contentType: 'application/json', headers: { Authorization: 'Bearer ' + key },
      payload: JSON.stringify({ from: MAIL_NAME + ' <' + MAIL_FROM + '>', to: [to], subject: subject, html: html, reply_to: MAIL_REPLY }), muteHttpExceptions: true });
    if (r.getResponseCode() >= 300) throw new Error('تعذر إرسال البريد (' + r.getResponseCode() + '): ' + r.getContentText().slice(0, 200));
    return;
  }
  MailApp.sendEmail({ to: to, subject: subject, htmlBody: html, name: MAIL_NAME, replyTo: MAIL_REPLY });
}

function doPost(e) {
  var body = null;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || ''); } catch (err) { body = null; }
  if (body && body.action) {
    try { return respond_(body.action === 'notify' ? notifyRoute_(body) : body.action === 'backupExport' ? backupExport_(body) : body.action.indexOf('file') === 0 ? fileRoute_(body) : pledgeRoute_(body)); }
    catch (err) { return respond_({ ok: false, error: String(err.message || err) }); }
  }
  return respond_(syncAll());
}
function doGet(e)  { return respond_(syncAll()); }
function respond_(r) { return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON); }

/** شغّلها مرة واحدة من المحرر لإنشاء مؤقت كل 5 دقائق */
function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'syncAll') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncAll').timeBased().everyMinutes(5).create();
  return syncAll();
}

/* ---------- Firebase REST بصلاحية المالك ---------- */
function fbHeaders_() { return { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }; }
function fbGet_(url) { var r = UrlFetchApp.fetch(url + '.json', { headers: fbHeaders_(), muteHttpExceptions: true }); if (r.getResponseCode() !== 200) throw new Error('DB read ' + r.getResponseCode()); var t = r.getContentText(); return t === 'null' ? null : JSON.parse(t); }
function fbWrite_(url, obj, method) { var r = UrlFetchApp.fetch(url + '.json', { method: method || 'put', contentType: 'application/json', headers: fbHeaders_(), payload: JSON.stringify(obj), muteHttpExceptions: true }); if (r.getResponseCode() !== 200) throw new Error('DB write ' + r.getResponseCode()); }
function fbDelete_(url) { var r = UrlFetchApp.fetch(url + '.json', { method: 'delete', headers: fbHeaders_(), muteHttpExceptions: true }); if (r.getResponseCode() !== 200) throw new Error('DB delete ' + r.getResponseCode()); }
function fbPush_(url, obj) { var r = UrlFetchApp.fetch(url + '.json', { method: 'post', contentType: 'application/json', headers: fbHeaders_(), payload: JSON.stringify(obj), muteHttpExceptions: true }); if (r.getResponseCode() !== 200) throw new Error('DB push ' + r.getResponseCode()); }
function emailKey_(e) { return String(e || '').toLowerCase().replace(/\./g, ','); }
function sha256_(str) { return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8).map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join(''); }
function tsAr_(ms) { return ms ? Utilities.formatDate(new Date(ms), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm') : ''; }
function fmtNum_(n) { return n == null || n === '' ? '' : Number(n).toLocaleString('en-US'); }
function escH_(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }

/** يتحقق من موظف مسجَّل دخوله في الموقع عبر رمز Firebase، ويعيد {email,name,role} */
function employee_(idToken) {
  if (!idToken) throw new Error('يجب تسجيل الدخول');
  var r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FB_API_KEY, { method: 'post', contentType: 'application/json', payload: JSON.stringify({ idToken: idToken }), muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('جلسة الدخول غير صالحة، أعد تسجيل الدخول');
  var u = (JSON.parse(r.getContentText()).users || [])[0];
  if (!u || !u.emailVerified) throw new Error('البريد غير موثّق');
  var email = String(u.email).toLowerCase();
  var rec = fbGet_(FIREBASE + '/users/' + emailKey_(email)) || {};
  if (rec.super) rec.role = 'admin'; // المدير الرئيسي = علامة super في سجل المستخدمين
  if (['maker', 'checker', 'admin'].indexOf(rec.role) < 0) throw new Error('هذا الحساب غير مصرح له');
  return { email: email, name: rec.name || email, role: rec.role };
}
function audit_(emp, action, unitId, bid, details) {
  try { fbPush_(FIREBASE + '/audit', { at: Date.now(), who: emp ? emp.name : 'النظام', whoEmail: emp ? emp.email : '', role: 'system', action: action, unitId: unitId || '', bid: bid || '', details: details || '' }); } catch (e) {}
}

/* ---------- مسارات التعهد ---------- */
function pledgeRoute_(body) {
  var emp = employee_(body.idToken);
  var bid = String(body.bid || '').replace(/[^A-Za-z0-9_-]/g, '');
  if (!bid) throw new Error('رقم الحجز مفقود');
  var booking = fbGet_(FIREBASE + '/bookings/' + bid);
  if (!booking) throw new Error('الحجز غير موجود');
  if (booking.status !== 'pending' && booking.status !== 'approved') throw new Error('الحجز غير نشط');
  var cfg = fbGet_(FIREBASE + '/config/pledge') || {};
  var existing = fbGet_(FIREBASE + '/pledges/' + bid);
  if (existing && existing.state === 'confirmed' && body.action !== 'pledgeStatus') throw new Error('تم تأكيد التعهد لهذا الحجز مسبقاً');
  if (body.action === 'pledgeEmail') return pledgeEmail_(emp, bid, booking, body.email);
  if (!cfg.enabled || !cfg.text) throw new Error('طلب التعهد غير مفعّل في الإعدادات');
  if (body.action === 'pledgeSend') return pledgeSend_(emp, bid, booking, cfg, existing);
  if (body.action === 'pledgeVerify') return pledgeVerify_(emp, bid, booking, cfg, body.code);
  throw new Error('إجراء غير معروف');
}

function pledgeEmail_(emp, bid, booking, email) {
  email = String(email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('البريد الإلكتروني غير صحيح');
  if (emp.role !== 'admin' && String(booking.makerEmail || '').toLowerCase() !== emp.email) throw new Error('تعديل البريد متاح لمُدخل الحجز أو المدير فقط');
  fbWrite_(FIREBASE + '/bookings/' + bid, { email: email }, 'patch');
  try { fbDelete_(SECRETS + '/' + bid); } catch (e) {}
  try { fbDelete_(FIREBASE + '/pledges/' + bid); } catch (e) {}
  audit_(emp, 'تعديل بريد العميل', booking.unitId, bid, (booking.refNo || '') + ' — ' + email);
  return { ok: true, email: email };
}

function pledgeSend_(emp, bid, booking, cfg, existing) {
  var email = String(booking.email || '').trim().toLowerCase();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('لا يوجد بريد إلكتروني صحيح للعميل');
  var prev = fbGet_(SECRETS + '/' + bid);
  if (prev && prev.sentAt && Date.now() - prev.sentAt < PLEDGE_RESEND_MS) throw new Error('انتظر دقيقة قبل إعادة الإرسال');
  var code = String(Math.floor(100000 + Math.random() * 900000));
  var salt = Utilities.getUuid();
  var now = Date.now(), exp = now + PLEDGE_TTL_MS;
  var version = Number(cfg.version || 1), textHash = sha256_(cfg.text), ctx = pledgeCtx_(booking);
  fbWrite_(SECRETS + '/' + bid, { hash: sha256_(salt + '|' + code + '|' + bid + '|' + email + '|' + version + '|' + textHash + '|' + ctx), salt: salt, email: email, version: version, textHash: textHash, ctx: ctx, unitId: booking.unitId, sentAt: now, expiresAt: exp, attempts: 0, used: false });
  sendPledgeMail_(booking, cfg, email, code, exp);
  fbWrite_(FIREBASE + '/pledges/' + bid, { state: 'sent', email: email, version: version, sentAt: now, expiresAt: exp, sentBy: emp.name, sentByEmail: emp.email, refNo: booking.refNo || '', unitId: booking.unitId });
  audit_(emp, 'إرسال التعهد ورمز التحقق', booking.unitId, bid, (booking.refNo || '') + ' — ' + email);
  return { ok: true, expiresAt: exp };
}

var PLEDGE_CITY = 'ينبع الصناعية';
/** يملأ الحقول بين قوسين في نص التعهد من بيانات الحجز */
function fillPledge_(t, b) {
  var m = { 'الاسم الكامل': b.name, 'رقم الهوية': b.natId, 'رقم الجوال': b.phone, 'رقم الوحدة': b.unitId, 'النموذج': b.model, 'اسم الحي': b.district, 'المدينة': PLEDGE_CITY, 'السعر': fmtNum_(b.price) };
  return String(t || '').replace(/\(([^()]+)\)/g, function (x, k) { k = k.trim(); return m[k] != null && m[k] !== '' ? String(m[k]) : x; });
}
/** يثبّت نص التعهد المعتمد في الإعدادات إن لم يُحفظ نص بعد (يُشغَّل مرة من المحرر) */
function seedPledgeText() {
  var cfg = fbGet_(FIREBASE + '/config/pledge') || {};
  if (cfg.text) return 'already set (v' + cfg.version + ')';
  var text = 'أُقرّ أنا المواطن (الاسم الكامل)، حامل الهوية الوطنية رقم (رقم الهوية) ورقم جوال (رقم الجوال)، بأنني أتقدم بطلب حجز الوحدة السكنية رقم (رقم الوحدة)، نموذج (النموذج)، الواقعة في حي (اسم الحي) بمدينة (المدينة)، بسعر مبدئي قدره (السعر) ريال سعودي.\n\nوأُقرّ بعلمي وقبولي بأن بيانات الوحدة وسعرها مبدئية وقابلة للتعديل حتى إبرام عقد البيع، وأن البيانات والسعر النهائيين هما ما يتم الاتفاق عليه في العقد.\n\nكما أُقرّ بأنه لم يسبق لي تملّك أي وحدة سكنية والإفراغ باسمي من الهيئة الملكية في أيٍّ من مدنها، وأتعهد بصحة البيانات والإقرارات المقدمة مني، وأتحمل مسؤولية صحتها.';
  fbWrite_(FIREBASE + '/config/pledge', { enabled: false, requireForApproval: false, reportEmail: '', text: text, version: 1, updatedAt: Date.now(), updatedBy: 'system' });
  return 'seeded v1';
}

/** سياق الحجز الذي يرتبط به الرمز: أي تغيير فيه يُبطل الرمز */
function pledgeCtx_(b) { return [b.unitId, b.price, b.payMethod, b.natId, b.name].map(function (v) { return String(v == null ? '' : v); }).join('|'); }

function pledgeVerify_(emp, bid, booking, cfg, code) {
  code = String(code || '').replace(/\D/g, '');
  var sec = fbGet_(SECRETS + '/' + bid);
  if (!sec || sec.used) throw new Error('لا يوجد رمز فعّال لهذا الحجز، أرسل التعهد أولاً');
  if (Date.now() > sec.expiresAt) throw new Error('انتهت صلاحية الرمز (5 دقائق)، أعد الإرسال');
  if (sec.attempts >= PLEDGE_MAX_ATTEMPTS) throw new Error('تجاوزت عدد المحاولات، أعد إرسال التعهد');
  var email = String(booking.email || '').toLowerCase(), version = Number(cfg.version || 1), textHash = sha256_(cfg.text), ctx = pledgeCtx_(booking);
  if (sec.email !== email || sec.version !== version || sec.textHash !== textHash || sec.ctx !== ctx) {
    fbDelete_(SECRETS + '/' + bid);
    throw new Error('تغيّرت بيانات الحجز أو نص التعهد بعد إرسال الرمز، يلزم إرسال جديد');
  }
  if (code.length !== 6 || sha256_(sec.salt + '|' + code + '|' + bid + '|' + email + '|' + version + '|' + textHash + '|' + ctx) !== sec.hash) {
    fbWrite_(SECRETS + '/' + bid, { attempts: (sec.attempts || 0) + 1 }, 'patch');
    audit_(emp, 'رمز تعهد خاطئ', booking.unitId, bid, booking.refNo || '');
    var left = PLEDGE_MAX_ATTEMPTS - (sec.attempts || 0) - 1;
    throw new Error('الرمز غير صحيح' + (left > 0 ? ' (متبقٍ ' + left + ' محاولات)' : '، أعد إرسال التعهد'));
  }
  var now = Date.now();
  fbWrite_(SECRETS + '/' + bid, { used: true, usedAt: now }, 'patch');
  var rec = { state: 'confirmed', statement: 'تم تأكيد الإقرار والتعهد بواسطة الموظف بالرمز المرسل إلى بريد العميل',
    email: email, name: booking.name || '', natId: booking.natId || '', refNo: booking.refNo || '', unitId: booking.unitId, district: booking.district || '', block: booking.block || '', model: booking.model || '', price: booking.price || '', company: booking.company || '',
    payMethod: booking.payMethod || '', version: version, textHash: textHash, text: fillPledge_(cfg.text, booking), template: cfg.text, sentAt: sec.sentAt, confirmedAt: now, by: emp.name, byEmail: emp.email, byRole: emp.role };
  fbWrite_(FIREBASE + '/pledges/' + bid, rec);
  audit_(emp, 'تأكيد الإقرار والتعهد', booking.unitId, bid, (booking.refNo || '') + ' — ' + email);
  try { sendConfirmMail_(booking, cfg, email, now); } catch (e) { audit_(emp, 'تعذر إرسال رسالة التأكيد', booking.unitId, bid, String(e)); }
  return { ok: true, confirmedAt: now };
}

/* ---------- الرسائل ---------- */
function bookingTable_(b) {
  var rows = [['كود الحجز', b.refNo], ['رقم الوحدة', b.unitId], ['الحي', b.district], ['رقم القطعة', b.plot], ['البلوك', b.block], ['النموذج', b.model], ['الشركة', b.company], ['السعر', fmtNum_(b.price) + ' ر.س'], ['اسم العميل', b.name], ['رقم الهوية', b.natId], ['تاريخ الحجز', b.bookingDate]];
  return '<table dir="rtl" style="border-collapse:collapse;font-size:14px;width:100%;max-width:520px">' + rows.map(function (r) { return '<tr><td style="padding:6px 10px;border:1px solid #ddd;background:#f3f6f4;font-weight:bold;width:34%">' + escH_(r[0]) + '</td><td style="padding:6px 10px;border:1px solid #ddd">' + escH_(r[1]) + '</td></tr>'; }).join('') + '</table>';
}
function mailShell_(title, inner) {
  return '<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;max-width:640px;margin:auto;color:#222;line-height:1.8"><div style="background:#1f4d3a;color:#fff;padding:14px 18px;border-radius:12px 12px 0 0;font-size:18px;font-weight:bold">' + escH_(title) + '</div><div style="border:1px solid #ddd;border-top:0;padding:18px;border-radius:0 0 12px 12px">' + inner + '<p style="color:#777;font-size:12px;margin-top:22px">رسالة آلية من نظام حجوزات مشروع جبين السكني — ينبع الصناعية.</p></div></div>';
}
function sendPledgeMail_(b, cfg, email, code, exp) {
  var inner = '<p>الأستاذ/ة <b>' + escH_(b.name) + '</b>، تحية طيبة.</p>' +
    '<p>تم تسجيل حجز باسمك برقم <b>' + escH_(b.refNo) + '</b>. هذه بيانات الوحدة المحجوزة:</p>' + bookingTable_(b) +
    '<h3 style="margin:22px 0 6px">نص الإقرار والتعهد (الإصدار ' + escH_(cfg.version || 1) + ')</h3>' +
    '<div style="white-space:pre-wrap;background:#fafafa;border:1px solid #e5e5e5;border-radius:10px;padding:14px">' + escH_(fillPledge_(cfg.text, b)) + '</div>' +
    '<h3 style="margin:22px 0 6px">رمز التحقق</h3>' +
    '<div style="font-size:34px;letter-spacing:10px;font-weight:bold;text-align:center;direction:ltr;background:#eef6f1;border:2px dashed #1f4d3a;border-radius:12px;padding:12px">' + code + '</div>' +
    '<p style="color:#b3261e;font-weight:bold">الرمز صالح حتى ' + tsAr_(exp) + ' (5 دقائق) ويُستخدم مرة واحدة.</p>' +
    '<p style="background:#fff6d6;border-radius:10px;padding:12px;font-weight:bold">بمشاركتك هذا الرمز مع موظف الحجز، فإنك تؤكد اطلاعك على بيانات الحجز والتعهد وموافقتك عليهما، وتفوّض الموظف بتسجيل موافقتك.</p>' +
    '<p>إذا لم توافق على ما ورد أعلاه، فلا تشارك الرمز مع أحد.</p>';
  sendMail_(email, 'الإقرار والتعهد ورمز التحقق — حجز ' + (b.refNo || ''), mailShell_('الإقرار والتعهد — مشروع جبين السكني', inner));
}
function sendConfirmMail_(b, cfg, email, when) {
  var inner = '<p>الأستاذ/ة <b>' + escH_(b.name) + '</b>، تحية طيبة.</p>' +
    '<p>تم تسجيل موافقتك على الإقرار والتعهد الخاص بالحجز رقم <b>' + escH_(b.refNo) + '</b> بتاريخ ' + tsAr_(when) + '.</p>' + bookingTable_(b) +
    '<h3 style="margin:22px 0 6px">نسخة الإقرار والتعهد المؤكَّد (الإصدار ' + escH_(cfg.version || 1) + ')</h3>' +
    '<div style="white-space:pre-wrap;background:#fafafa;border:1px solid #e5e5e5;border-radius:10px;padding:14px">' + escH_(fillPledge_(cfg.text, b)) + '</div>' +
    '<p>احتفظ بهذه الرسالة كمرجع.</p>';
  sendMail_(email, 'تأكيد الإقرار والتعهد — حجز ' + (b.refNo || ''), mailShell_('تم تأكيد الإقرار والتعهد — مشروع جبين السكني', inner));
}

/* ---------- الرسائل الآلية (مباركة الحجز، رسوم الجدية، مهلة التوثيق، تذكير التوقيع، تذكير الموظفين بسكني) ----------
 * كل رسالة لها مفتاح تفعيل في config/mail (الإعدادات في الموقع) وكلها موقوفة افتراضياً.
 * كل إرسال يُسجَّل في mailLog/{bid} ويُحتسب في الحد اليومي MAIL_DAILY_CAP (الخطة المجانية في Resend: 100/يوم).
 */
var MAIL_DAILY_CAP = 90;
var SITE_URL = 'https://jabeen-bookings.online';
var DAY_MS = 24 * 3600 * 1000;
function mailCfg_() { return fbGet_(FIREBASE + '/config/mail') || {}; }
function riyadhDate_(ms) { return Utilities.formatDate(new Date(ms || Date.now()), 'Asia/Riyadh', 'yyyy-MM-dd'); }
function bookingDate_(b) { return (b.bookingDate && /^\d{4}-\d{2}-\d{2}$/.test(b.bookingDate)) ? b.bookingDate : (b.createdAt ? riyadhDate_(b.createdAt) : ''); }
function daysSince_(ds) { if (!ds) return null; return Math.max(0, Math.round((Date.parse(riyadhDate_() + 'T00:00:00Z') - Date.parse(ds + 'T00:00:00Z')) / DAY_MS)); }
function addDays_(ds, n) { return Utilities.formatDate(new Date(Date.parse(ds + 'T00:00:00Z') + n * DAY_MS), 'UTC', 'yyyy-MM-dd'); }
function daysAr_(n) { if (n === 0) return 'اليوم'; if (n === 1) return 'يوم واحد'; if (n === 2) return 'يومين'; if (n <= 10) return n + ' أيام'; return n + ' يوماً'; }
function isActive_(b) { return !!b && (b.status === 'pending' || b.status === 'approved'); }
function validEmail_(e) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(e || '')); }
function quotaTake_() {
  var d = riyadhDate_(), n = Number(fbGet_(FIREBASE + '/mailQuota/' + d) || 0);
  if (n >= MAIL_DAILY_CAP) return false;
  fbWrite_(FIREBASE + '/mailQuota/' + d, n + 1); return true;
}
function logOf_(log, bid) { var l = (log && log[bid]) || {}; return Object.keys(l).map(function (k) { return l[k]; }); }
function sentAt_(entries, type) { var t = 0; entries.forEach(function (e) { if (e && e.type === type && e.status === 'sent' && e.at > t) t = e.at; }); return t; }
/** يرسل ويسجّل. يعيد 'sent' | 'quota' | 'failed' */
function sendLogged_(bid, type, to, subject, html, emp) {
  if (!quotaTake_()) return 'quota';
  var rec = { at: Date.now(), type: type, to: to, subject: subject, status: 'sent', by: emp ? emp.name : 'المؤقّت' };
  try { sendMail_(to, subject, html); }
  catch (e) { rec.status = 'failed'; rec.error = String(e.message || e).slice(0, 200); }
  try { fbPush_(FIREBASE + '/mailLog/' + bid, rec); } catch (e) {}
  return rec.status;
}
/** مهلة توقيع عرض السعر وتوثيق العقد تبدأ من تاريخ التسجيل في سكني (sak.at)، وليس من تاريخ الحجز */
function signDays_(cfg) { return Number(cfg.signDays || 30); }
function sakaniDate_(sak) { return sak && sak.at ? riyadhDate_(sak.at) : ''; }
function signDeadline_(cfg, sak) { var ds = sakaniDate_(sak); return ds ? addDays_(ds, signDays_(cfg)) : ''; }
function nextStepsHtml_(cfg, b) {
  var n = signDays_(cfg);
  return '<h3 style="margin:22px 0 6px">الخطوات التالية</h3><ol style="padding-right:20px;margin:0">' +
    '<li>التسجيل في منصة <b>سكني</b> وربط الحجز بحسابك (يتولى موظف الحجز تسجيله ويصلك إشعار).</li>' +
    '<li><b>توثيق الحجز</b> بدفع <b>رسوم الجدية</b> من خلال منصة سكني.</li>' +
    '<li>توقيع <b>عرض السعر</b> و<b>توثيق العقد</b> عبر منصة سكني خلال <b>' + n + ' يوماً</b> من تاريخ التسجيل في سكني.</li></ol>' +
    '<p style="background:#fff6d6;border-radius:10px;padding:12px;font-weight:bold">في حال عدم توقيع عرض السعر وتوثيق العقد خلال ' + n + ' يوماً من التسجيل في سكني، يُلغى الحجز وتُتاح الوحدة لغيرك.</p>';
}
function welcomeMail_(cfg, b, emp) {
  var inner = '<p>الأستاذ/ة <b>' + escH_(b.name) + '</b>، تحية طيبة.</p>' +
    '<p style="font-size:17px;font-weight:bold;color:#1f4d3a">مبارك لك حجز وحدتك السكنية في مشروع جبين 🎉</p>' +
    '<p>تم تسجيل حجزك برقم <b>' + escH_(b.refNo) + '</b>' + (b.status === 'pending' ? ' (حجز مبدئي بانتظار الاعتماد)' : '') + '. هذه بيانات الوحدة:</p>' + bookingTable_(b) + nextStepsHtml_(cfg, b);
  return sendLogged_(b.bid, 'welcome', b.email, 'مبارك حجزك — حجز ' + (b.refNo || ''), mailShell_('مبارك الحجز — مشروع جبين السكني', inner), emp);
}
function earnestMail_(cfg, b, sak, emp) {
  var ds = sakaniDate_(sak) || riyadhDate_();
  var inner = '<p>الأستاذ/ة <b>' + escH_(b.name) + '</b>، تحية طيبة.</p>' +
    '<p>تم إدخال حجزك رقم <b>' + escH_(b.refNo) + '</b> للوحدة <b>' + escH_(b.unitId) + '</b> في منصة <b>سكني</b> بتاريخ <b>' + escH_(ds) + '</b>.</p>' +
    '<div style="font-size:18px;font-weight:bold;text-align:center;background:#eef6f1;color:#1f4d3a;border:2px dashed #1f4d3a;border-radius:12px;padding:12px">الخطوة المطلوبة الآن: توثيق الحجز بدفع رسوم الجدية عبر منصة سكني</div>' +
    '<ol style="padding-right:20px;margin:14px 0 0"><li>الدخول إلى حسابك في منصة سكني.</li><li>فتح الحجز الخاص بالوحدة <b>' + escH_(b.unitId) + '</b>.</li><li>سداد <b>رسوم الجدية</b> بالطريقة المتاحة في المنصة.</li></ol>' +
    '<p style="background:#fff6d6;border-radius:10px;padding:12px;font-weight:bold;margin-top:14px">لا يُعد الحجز موثقاً في سكني إلا بعد سداد رسوم الجدية.</p>' +
    '<p>بعد السداد تأتي خطوة توقيع عرض السعر وتوثيق العقد خلال <b>' + signDays_(cfg) + ' يوماً</b> من تاريخ التسجيل في سكني.</p>' + bookingTable_(b);
  return sendLogged_(b.bid, 'earnest', b.email, 'توثيق الحجز بدفع رسوم الجدية عبر سكني — حجز ' + (b.refNo || ''), mailShell_('توثيق الحجز في سكني — مشروع جبين السكني', inner), emp);
}
function contractDocMail_(cfg, b, sak, emp) {
  var n = signDays_(cfg), ds = sakaniDate_(sak) || riyadhDate_(), dl = addDays_(ds, n);
  var inner = '<p>الأستاذ/ة <b>' + escH_(b.name) + '</b>، تحية طيبة.</p>' +
    '<p>تم تسجيل حجزك رقم <b>' + escH_(b.refNo) + '</b> للوحدة <b>' + escH_(b.unitId) + '</b> في منصة <b>سكني</b> بتاريخ <b>' + escH_(ds) + '</b>.</p>' +
    '<p>الخطوة التالية: الدخول إلى حسابك في منصة سكني، ومراجعة بيانات الحجز، ثم <b>توقيع عرض السعر</b> و<b>توثيق العقد</b>.</p>' +
    '<div style="font-size:18px;font-weight:bold;text-align:center;background:#eef6f1;color:#1f4d3a;border:2px dashed #1f4d3a;border-radius:12px;padding:12px">مهلة التوقيع والتوثيق: ' + n + ' يوماً من تاريخ التسجيل في سكني<div style="font-size:14px;font-weight:normal;margin-top:4px">آخر موعد: ' + escH_(dl) + '</div></div>' +
    '<p style="background:#fff6d6;border-radius:10px;padding:12px;font-weight:bold;margin-top:14px">في حال عدم توقيع عرض السعر وتوثيق العقد قبل هذا الموعد، يُلغى الحجز وتُتاح الوحدة لغيرك.</p>' + bookingTable_(b);
  return sendLogged_(b.bid, 'contractDoc', b.email, 'توثيق العقد عبر منصة سكني — حجز ' + (b.refNo || ''), mailShell_('حجزك مسجَّل في سكني — مشروع جبين السكني', inner), emp);
}
function signRemindMail_(cfg, b, remaining, deadline, sakDate, emp) {
  var urgent = remaining <= 2;
  var inner = '<p>الأستاذ/ة <b>' + escH_(b.name) + '</b>، تحية طيبة.</p>' +
    '<p>نذكّرك بضرورة <b>توقيع عرض السعر وتوثيق العقد</b> عبر منصة سكني للوحدة <b>' + escH_(b.unitId) + '</b> (حجز رقم <b>' + escH_(b.refNo) + '</b>)، فالمهلة <b>' + signDays_(cfg) + ' يوماً</b> من تاريخ تسجيل حجزك في سكني' + (sakDate ? ' (' + escH_(sakDate) + ')' : '') + '.</p>' +
    '<div style="font-size:20px;font-weight:bold;text-align:center;background:' + (urgent ? '#fdecea;color:#b3261e;border:2px solid #b3261e' : '#eef6f1;color:#1f4d3a;border:2px dashed #1f4d3a') + ';border-radius:12px;padding:12px">' +
    (remaining === 0 ? 'اليوم آخر يوم في مهلة التوقيع والتوثيق' : 'المتبقي من المهلة: ' + daysAr_(remaining)) + '<div style="font-size:14px;font-weight:normal;margin-top:4px">آخر موعد: ' + escH_(deadline) + '</div></div>' +
    '<p style="background:#fff6d6;border-radius:10px;padding:12px;font-weight:bold;margin-top:14px">في حال عدم توقيع عرض السعر وتوثيق العقد خلال المدة المتبقية سيتم إلغاء الحجز وإتاحة الوحدة لغيرك.</p>' + bookingTable_(b);
  return sendLogged_(b.bid, 'signRemind', b.email, 'تذكير: توقيع عرض السعر وتوثيق العقد — حجز ' + (b.refNo || '') + ' (المتبقي ' + daysAr_(remaining) + ')', mailShell_('تذكير بتوقيع عرض السعر وتوثيق العقد — مشروع جبين السكني', inner), emp);
}
/** من الموقع فور الحدث: {action:'notify', event:'welcome'|'sakani', bid, idToken} */
function notifyRoute_(body) {
  var emp = employee_(body.idToken);
  var bid = String(body.bid || '').replace(/[^A-Za-z0-9_-]/g, '');
  var b = fbGet_(FIREBASE + '/bookings/' + bid);
  if (!b || !isActive_(b)) throw new Error('الحجز غير نشط');
  b.bid = bid;
  var cfg = mailCfg_(), entries = logOf_({ x: fbGet_(FIREBASE + '/mailLog/' + bid) }, 'x');
  if (!validEmail_(b.email)) return { ok: true, sent: false, reason: 'لا يوجد بريد للعميل' };
  if (body.event === 'welcome') {
    if (!cfg.welcome) return { ok: true, sent: false, reason: 'الرسالة موقوفة' };
    if (sentAt_(entries, 'welcome')) return { ok: true, sent: false, reason: 'أُرسلت سابقاً' };
    return { ok: true, sent: welcomeMail_(cfg, b, emp) === 'sent' };
  }
  if (body.event === 'sakani') {
    var sak = fbGet_(FIREBASE + '/customers/' + b.natId + '/sakani/' + bid);
    if (!sak) return { ok: true, sent: false, reason: 'لم يُسجَّل في سكني بعد' };
    var sent = [];
    if (cfg.earnest && !sentAt_(entries, 'earnest') && earnestMail_(cfg, b, sak, emp) === 'sent') sent.push('earnest');
    if (cfg.contractDoc && Number(cfg.contractDelay || 0) === 0 && !sentAt_(entries, 'contractDoc') && contractDocMail_(cfg, b, sak, emp) === 'sent') sent.push('contractDoc');
    return { ok: true, sent: sent.length > 0, types: sent };
  }
  throw new Error('حدث غير معروف');
}
/** شغّلها مرة واحدة من المحرر: مؤقّت كل ساعة للرسائل الآلية */
function setupMailTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'mailScheduler') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('mailScheduler').timeBased().everyHours(1).create();
  return mailScheduler();
}
/** يعمل كل ساعة: يلتقط ما فات (مباركة/سكني)، يرسل تذكيرات التوقيع، ويرسل ملخص سكني لمدخلي الحجوزات مرة يومياً بعد الساعة 9 صباحاً */
function mailScheduler() {
  var cfg = mailCfg_();
  if (!cfg.welcome && !cfg.earnest && !cfg.contractDoc && !cfg.signRemind && !cfg.sakaniStaff) return { ok: true, skipped: 'all off' };
  var bookings = fbGet_(FIREBASE + '/bookings') || {}, customers = fbGet_(FIREBASE + '/customers') || {}, log = fbGet_(FIREBASE + '/mailLog') || {};
  var since = cfg.since || {}, out = { welcome: 0, earnest: 0, contractDoc: 0, signRemind: 0, staff: 0, quota: false };
  var contractDelay = Math.max(0, Number(cfg.contractDelay || 0)) * DAY_MS; // رسالة مهلة التوقيع تُرسل بعد رسالة رسوم الجدية بهذه المدة
  var h = riyadhHour_(Date.now()), quiet = h < 8 || h >= 21; // لا رسائل للعملاء ليلاً
  var waitingSakani = {};
  Object.keys(bookings).forEach(function (bid) {
    var b = bookings[bid]; if (!isActive_(b)) return; b.bid = bid;
    var cust = customers[b.natId] || {}, sak = cust.sakani && cust.sakani[bid], signed = cust.signed && cust.signed[bid];
    if (!sak) { var mk = String(b.makerEmail || '').toLowerCase(); if (validEmail_(mk)) (waitingSakani[mk] = waitingSakani[mk] || []).push(b); }
    if (quiet || out.quota || !validEmail_(b.email)) return;
    var entries = logOf_(log, bid), r;
    if (cfg.welcome && b.createdAt >= (since.welcome || 0) && !sentAt_(entries, 'welcome')) { r = welcomeMail_(cfg, b); if (r === 'quota') out.quota = true; else if (r === 'sent') out.welcome++; }
    if (cfg.earnest && sak && sak.at >= (since.earnest || 0) && !sentAt_(entries, 'earnest')) { r = earnestMail_(cfg, b, sak); if (r === 'quota') out.quota = true; else if (r === 'sent') out.earnest++; }
    if (cfg.contractDoc && sak && sak.at >= (since.contractDoc || 0) && Date.now() - sak.at >= contractDelay && !sentAt_(entries, 'contractDoc')) { r = contractDocMail_(cfg, b, sak); if (r === 'quota') out.quota = true; else if (r === 'sent') out.contractDoc++; }
    if (cfg.signRemind && !signed && sak) { // المهلة تبدأ من تاريخ التسجيل في سكني
      var ds = sakaniDate_(sak), passed = daysSince_(ds), remaining = signDays_(cfg) - passed, deadline = signDeadline_(cfg, sak);
      var before = Number(cfg.remindBefore || 7), every = Math.max(1, Number(cfg.remindEvery || 3)), last = sentAt_(entries, 'signRemind');
      if (remaining >= 0 && remaining <= before && (!last || Date.now() - last >= every * DAY_MS - 3600 * 1000)) { r = signRemindMail_(cfg, b, remaining, deadline, ds); if (r === 'quota') out.quota = true; else if (r === 'sent') out.signRemind++; }
    }
  });
  // ملخص يومي لكل مُدخل: عملاؤه الذين لم يُسجَّلوا في سكني
  if (cfg.sakaniStaff && h >= 9 && !out.quota) {
    var staffLog = log._staff || {}, today = riyadhDate_();
    var sentToday = Object.keys(staffLog).some(function (k) { return staffLog[k] && staffLog[k].status === 'sent' && riyadhDate_(staffLog[k].at) === today; });
    if (!sentToday) Object.keys(waitingSakani).forEach(function (mk) {
      var list = waitingSakani[mk].sort(function (a, b) { return (a.createdAt || 0) - (b.createdAt || 0); });
      var rows = list.map(function (b) { return '<tr>' + [b.refNo || '', b.unitId, b.name || '', b.natId || '', b.phone || '', bookingDate_(b), daysAr_(daysSince_(bookingDate_(b)) || 0), STATUS_AR[b.status] || b.status].map(function (c) { return '<td style="border:1px solid #ddd;padding:5px 8px">' + escH_(c) + '</td>'; }).join('') + '</tr>'; }).join('');
      var inner = '<p>لديك <b>' + list.length + '</b> ' + (list.length === 1 ? 'حجز نشط لم يُسجَّل' : 'حجوزات نشطة لم تُسجَّل') + ' بعد في منصة سكني:</p>' +
        '<table dir="rtl" style="border-collapse:collapse;font-size:13px;width:100%"><tr>' + ['كود الحجز', 'الوحدة', 'العميل', 'الهوية', 'الجوال', 'تاريخ الحجز', 'منذ', 'الحالة'].map(function (x) { return '<th style="border:1px solid #ddd;padding:5px 8px;background:#f3f6f4">' + x + '</th>'; }).join('') + '</tr>' + rows + '</table>' +
        '<p style="margin-top:16px">بعد تسجيل كل حجز في سكني، افتح بطاقة العميل في <a href="' + SITE_URL + '">نظام الحجوزات</a> (تبويب العملاء) واضغط «تم عكس الحجز في سكني ✓» ليتوقف هذا التذكير ويصل العميل إشعار توثيق العقد.</p>';
      var st = sendLogged_('_staff', 'sakaniStaff', mk, 'تذكير: ' + list.length + ' من عملائك بانتظار التسجيل في سكني', mailShell_('تذكير يومي — التسجيل في سكني', inner));
      if (st === 'quota') out.quota = true; else if (st === 'sent') out.staff++;
    });
  }
  if (out.quota) audit_(null, 'توقف الإرسال الآلي', '', '', 'تجاوز الحد اليومي للرسائل (' + MAIL_DAILY_CAP + ') — يُستكمل غداً');
  return { ok: true, result: out };
}
/** اختبار من المحرر: يرسل نماذج الرسائل الأربع إلى بريد المالك دون المساس بأي حجز */
function testAutoMails() {
  var cfg = { signDays: 30, remindBefore: 7, remindEvery: 3 }, sak = { at: Date.now() - 27 * DAY_MS };
  var b = { bid: '_test', refNo: 'JB-TEST-000-00000', unitId: 'AZ-000', district: 'العزيزية', plot: '000', block: '00', model: 'B1', company: 'اختبار', price: 1590000, name: 'عميل تجريبي', natId: '1000000000', bookingDate: riyadhDate_(), email: SUPER_ADMIN, maker: 'موظف تجريبي', status: 'approved', createdAt: Date.now() };
  var inner = welcomeMail_(cfg, b) + ' / ' + earnestMail_(cfg, b, sak) + ' / ' + contractDocMail_(cfg, b, sak) + ' / ' + signRemindMail_(cfg, b, 3, signDeadline_(cfg, sak), sakaniDate_(sak));
  try { fbDelete_(FIREBASE + '/mailLog/_test'); } catch (e) {}
  return inner + ' → ' + SUPER_ADMIN;
}

/* ---------- تقرير الحوكمة اليومي للإدارة ---------- */
function setupReportTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyPledgeReport') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyPledgeReport').timeBased().atHour(8).everyDays(1).inTimezone('Asia/Riyadh').create();
  return dailyPledgeReport();
}
function riyadhHour_(ms) { return Number(Utilities.formatDate(new Date(ms), 'Asia/Riyadh', 'H')); }
function riyadhDow_(ms) { return Utilities.formatDate(new Date(ms), 'Asia/Riyadh', 'u'); } // 5=الجمعة 6=السبت
function dailyPledgeReport() {
  var cfg = fbGet_(FIREBASE + '/config/pledge') || {};
  var to = cfg.reportEmail || SUPER_ADMIN;
  var bookings = fbGet_(FIREBASE + '/bookings') || {}, pledges = fbGet_(FIREBASE + '/pledges') || {}, audit = fbGet_(FIREBASE + '/audit') || {};
  var now = Date.now(), day = 24 * 3600 * 1000;
  // 1) حجوزات نشطة بلا تعهد مؤكد منذ أكثر من 48 ساعة (عند تفعيل الميزة)
  var late = [];
  if (cfg.enabled) Object.keys(bookings).forEach(function (k) { var b = bookings[k]; if (!b || (b.status !== 'pending' && b.status !== 'approved')) return;
    var p = pledges[k]; if (p && p.state === 'confirmed') return;
    var since = p && p.sentAt ? p.sentAt : b.createdAt; if (!since || now - since < UNCONFIRMED_HOURS * 3600 * 1000) return;
    late.push([b.refNo || '', b.unitId, b.name || '', b.maker || '', p ? 'أُرسل ' + tsAr_(p.sentAt) + ' ولم يُؤكَّد' : 'لم يُرسل', STATUS_AR[b.status] || b.status]); });
  // 2) رموز خاطئة متكررة خلال 7 أيام (3 فأكثر لنفس الحجز)
  var wrong = {}, entries = Object.keys(audit).map(function (k) { return audit[k]; });
  entries.forEach(function (a) { if (a.action === 'رمز تعهد خاطئ' && now - a.at < 7 * day) { var key = (a.bid || '') + '|' + (a.whoEmail || a.who || ''); wrong[key] = wrong[key] || { n: 0, a: a }; wrong[key].n++; } });
  var repeated = Object.keys(wrong).filter(function (k) { return wrong[k].n >= 3; }).map(function (k) { var w = wrong[k]; return [w.a.details || '', w.a.unitId || '', w.a.who || '', w.n]; });
  // 3) إجراءات خلال آخر 24 ساعة خارج ساعات العمل أو في الجمعة/السبت
  var SENSITIVE = ['إنشاء حجز مبدئي', 'اعتماد حجز', 'إلغاء حجز', 'رفض حجز', 'إرسال التعهد ورمز التحقق', 'تأكيد الإقرار والتعهد', 'تعديل بريد العميل', 'إعدادات التعهد'];
  var offHours = entries.filter(function (a) { if (now - a.at > day || SENSITIVE.indexOf(a.action) < 0) return false; var h = riyadhHour_(a.at), d = riyadhDow_(a.at); return d === '5' || d === '6' || h < WORK_START || h >= WORK_END; })
    .sort(function (a, b) { return b.at - a.at; }).map(function (a) { return [tsAr_(a.at), a.who || '', a.action, a.unitId || '', a.details || '']; });
  function tbl(head, rows) { if (!rows.length) return '<p style="color:#1e6b45">لا شيء.</p>'; return '<table dir="rtl" style="border-collapse:collapse;font-size:13px;width:100%"><tr>' + head.map(function (h) { return '<th style="border:1px solid #ddd;padding:5px 8px;background:#f3f6f4">' + escH_(h) + '</th>'; }).join('') + '</tr>' + rows.map(function (r) { return '<tr>' + r.map(function (c) { return '<td style="border:1px solid #ddd;padding:5px 8px">' + escH_(c) + '</td>'; }).join('') + '</tr>'; }).join('') + '</table>'; }
  var confirmedToday = Object.keys(pledges).filter(function (k) { var p = pledges[k]; return p && p.state === 'confirmed' && now - p.confirmedAt < day; }).length;
  var inner = '<p>تقرير ' + tsAr_(now).slice(0, 10) + ' — التعهدات المؤكدة خلال 24 ساعة: <b>' + confirmedToday + '</b>' + (cfg.enabled ? '' : ' (ميزة التعهد موقوفة حالياً)') + '</p>' +
    '<h3>1) حجوزات نشطة بلا تعهد مؤكد منذ أكثر من ' + UNCONFIRMED_HOURS + ' ساعة (' + late.length + ')</h3>' + tbl(['كود الحجز', 'الوحدة', 'العميل', 'المُدخل', 'حالة التعهد', 'حالة الحجز'], late) +
    '<h3>2) رموز خاطئة متكررة خلال 7 أيام (' + repeated.length + ')</h3>' + tbl(['الحجز', 'الوحدة', 'الموظف', 'عدد المحاولات'], repeated) +
    '<h3>3) إجراءات خارج ساعات العمل (' + WORK_START + ':00–' + WORK_END + ':00) أو في عطلة نهاية الأسبوع خلال 24 ساعة (' + offHours.length + ')</h3>' + tbl(['الوقت', 'الموظف', 'الإجراء', 'الوحدة', 'التفاصيل'], offHours);
  sendMail_(to, 'تقرير حوكمة التعهدات — ' + tsAr_(now).slice(0, 10), mailShell_('تقرير حوكمة التعهدات اليومي', inner));
  return { ok: true, to: to, late: late.length, repeated: repeated.length, offHours: offHours.length };
}

/** اختبار من المحرر: يرسل رسالتي التعهد والتأكيد إلى بريد المالك لحجز تجريبي دون المساس بأي حجز حقيقي */
function testPledgeMail() {
  var cfg = fbGet_(FIREBASE + '/config/pledge') || { text: '(نص التعهد لم يُحفظ بعد في الإعدادات)', version: 0 };
  var b = { refNo: 'JB-TEST-000-00000', unitId: 'AZ-000', district: 'العزيزية', plot: '000', block: '00', model: 'B1', company: 'اختبار', price: 1590000, name: 'عميل تجريبي', natId: '1000000000', bookingDate: '2026-11-03' };
  sendPledgeMail_(b, cfg, SUPER_ADMIN, '123456', Date.now() + PLEDGE_TTL_MS);
  sendConfirmMail_(b, cfg, SUPER_ADMIN, Date.now());
  return 'sent to ' + SUPER_ADMIN + ' via ' + (PropertiesService.getScriptProperties().getProperty('RESEND_KEY') ? 'Resend (' + MAIL_FROM + ')' : 'Gmail fallback');
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
    var pledges = fetchJson_(FIREBASE + '/pledges.json') || {};
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
    var bHead = ['كود الحجز', 'الوحدة', 'الحي', 'رقم القطعة', 'البلوك', 'النموذج', 'الشركة', 'السعر', 'اسم العميل', 'رقم الهوية', 'الرقم الوظيفي', 'الجوال', 'طريقة الدفع', 'تاريخ الحجز', 'الحالة', 'المُدخل', 'وقت الإدخال', 'المُراجع', 'وقت المراجعة', 'سبب الرفض/الإلغاء', 'ملاحظات', 'المصدر', 'البريد', 'التعهد', 'وقت التعهد'];
    var bRows = list.map(function (b) {
      return [b.refNo || '', b.unitId || '', b.district || '', b.plot || '', b.block || '', b.model || '', b.company || '', b.price || '',
        b.name || '', b.natId ? "'" + b.natId : '', b.empNo || '', b.phone ? "'" + b.phone : '', b.payMethod || '', b.bookingDate || '',
        STATUS_AR[b.status] || b.status || '', b.maker || '', ts_(b.createdAt), b.checker || '', ts_(b.checkedAt), b.rejectReason || '', b.notes || '', b.source || '',
        b.email || '', (pledges[b._id] && pledges[b._id].state === 'confirmed') ? 'مؤكَّد' : '', (pledges[b._id] && pledges[b._id].confirmedAt) ? ts_(pledges[b._id].confirmedAt) : ''];
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

/* ---------- مستندات العملاء: تُحفظ في Google Drive لحساب مالك السكربت، وتُسجَّل بياناتها في Firebase ---------- */
var FILE_MAX = 10 * 1024 * 1024;                 // 10 ميجابايت
var FILE_ROOT_NAME = 'jabeen-bookings-attachments';
var FILE_MIME_OK = /^(application\/pdf|image\/(jpeg|png|webp|heic|heif))$/;
function fileRoute_(body) {
  var emp = employee_(body.idToken);
  var natId = String(body.natId || '').replace(/\D/g, '');
  if (!/^[12]\d{9}$/.test(natId)) throw new Error('رقم هوية غير صالح');
  if (body.action === 'fileUpload') return fileUpload_(emp, natId, body);
  if (body.action === 'fileGet') return fileGet_(emp, natId, String(body.fid || ''));
  if (body.action === 'fileDelete') return fileDelete_(emp, natId, String(body.fid || ''));
  throw new Error('إجراء غير معروف');
}
function folderByName_(parent, name) { var it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function fileRoot_() { // المجلد الجذري للمستندات يُحفظ برقمه في Script Properties فيبقى صالحاً بعد نقل الملكية لحساب آخر
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('FILE_ROOT_ID');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var root = folderByName_(DriveApp.getRootFolder(), FILE_ROOT_NAME); props.setProperty('FILE_ROOT_ID', root.getId()); return root;
}
function fileFolder_(natId) { return folderByName_(fileRoot_(), natId); }
function fbPushKey_(url, obj) {
  var r = UrlFetchApp.fetch(url + '.json', { method: 'post', contentType: 'application/json', headers: fbHeaders_(), payload: JSON.stringify(obj), muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('DB push ' + r.getResponseCode());
  return JSON.parse(r.getContentText()).name;
}
function fileUpload_(emp, natId, body) {
  var bytes = Utilities.base64Decode(String(body.data || ''));
  if (!bytes.length) throw new Error('الملف فارغ');
  if (bytes.length > FILE_MAX) throw new Error('حجم الملف يتجاوز 10 ميجابايت');
  var mime = String(body.mime || '');
  if (!FILE_MIME_OK.test(mime)) throw new Error('نوع الملف غير مسموح: PDF أو صورة (JPG/PNG) فقط');
  var name = String(body.name || 'file').replace(/[\\\/:*?"<>|]/g, '_').slice(0, 120);
  var label = String(body.label || 'أخرى').slice(0, 60);
  if (!/^[12]\d{9}$/.test(natId)) throw new Error('رقم هوية غير صالح');
  var f = fileFolder_(natId).createFile(Utilities.newBlob(bytes, mime, name));
  f.setDescription(label + ' · ' + natId + ' · رفعه ' + emp.email);
  var rec = { name: name, mime: mime, size: bytes.length, label: label, driveId: f.getId(), at: Date.now(), by: emp.name, byEmail: emp.email };
  var fid = fbPushKey_(FIREBASE + '/customers/' + natId + '/files', rec);
  audit_(emp, 'رفع مستند عميل', '', '', label + ' · ' + name + ' · ' + natId);
  return { ok: true, fid: fid, file: rec };
}
function fileGet_(emp, natId, fid) {
  if (!fid) throw new Error('معرّف المستند مفقود');
  var rec = fbGet_(FIREBASE + '/customers/' + natId + '/files/' + fid);
  if (!rec || !rec.driveId) throw new Error('المستند غير موجود');
  var blob = DriveApp.getFileById(rec.driveId).getBlob();
  audit_(emp, 'فتح مستند عميل', '', '', (rec.label || '') + ' · ' + (rec.name || '') + ' · ' + natId);
  return { ok: true, name: rec.name, mime: rec.mime || blob.getContentType(), data: Utilities.base64Encode(blob.getBytes()) };
}
function fileDelete_(emp, natId, fid) {
  if (!fid) throw new Error('معرّف المستند مفقود');
  var rec = fbGet_(FIREBASE + '/customers/' + natId + '/files/' + fid);
  if (!rec) throw new Error('المستند غير موجود');
  if (emp.role !== 'admin' && rec.byEmail !== emp.email) throw new Error('حذف المستند للمدير أو لمن رفعه فقط');
  try { DriveApp.getFileById(rec.driveId).setTrashed(true); } catch (e) {}
  fbDelete_(FIREBASE + '/customers/' + natId + '/files/' + fid);
  audit_(emp, 'حذف مستند عميل', '', '', (rec.label || '') + ' · ' + (rec.name || '') + ' · ' + natId);
  return { ok: true };
}
/** اختبار من المحرر: يتأكد أن صلاحية Google Drive تعمل (ينشئ مجلداً تجريبياً ثم ينقله إلى المهملات) */
function testDriveFolder() {
  var f = fileFolder_('1000000000'); var id = f.getId(); f.setTrashed(true);
  return 'Drive OK · ' + id;
}
/** بعد نقل الملكية: شغّلها مرة واحدة من الحساب الجديد لإنشاء المؤقتات باسمه (المزامنة كل 5 دقائق + التقرير اليومي) */
function setupAllTriggers() {
  setupTrigger(); setupReportTrigger(); setupMailTrigger(); setupBackupTrigger();
  return ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); }).join(', ') + ' · المالك: ' + Session.getEffectiveUser().getEmail();
}

/* ---------- النسخ الاحتياطي الكامل ----------
 * dailyBackup: ينسخ قاعدة Firebase كاملة (jabeenBookings + pledgeSecrets) يومياً إلى Drive مالك السكربت
 * في مجلد jabeen-bookings-backups بصيغة JSON مضغوطة، ويحتفظ بآخر BACKUP_KEEP_DAYS يوماً.
 * المستندات المرفوعة محفوظة أصلاً في Drive، وكود الموقع في GitHub.
 * الاسترجاع: listBackups() لعرض النسخ، ثم restoreFromBackup('اسم الملف') من المحرر — أو استيراد JSON يدوياً من لوحة Firebase.
 * backupExport: يعطي النسخة الحالية لسكربت النسخ على اللابتوب برمز سري BACKUP_TOKEN في Script Properties.
 */
var BACKUP_FOLDER_NAME = 'jabeen-bookings-backups';
var BACKUP_KEEP_DAYS = 60;
function backupFolder_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('BACKUP_FOLDER_ID');
  if (id) { try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  var root = folderByName_(DriveApp.getRootFolder(), BACKUP_FOLDER_NAME); props.setProperty('BACKUP_FOLDER_ID', root.getId()); return root;
}
function dbSnapshot_() {
  var r = UrlFetchApp.fetch(DB_ROOT + '/.json', { headers: fbHeaders_(), muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('DB read ' + r.getResponseCode());
  var txt = r.getContentText(), data = JSON.parse(txt);
  if (!data || !data.jabeenBookings) throw new Error('النسخة فارغة أو بلا jabeenBookings — لم تُحفظ');
  return { txt: txt, data: data };
}
function dailyBackup() {
  var at = Date.now(), status = { at: at, ok: false };
  try {
    var snap = dbSnapshot_(), jb = snap.data.jabeenBookings;
    var name = 'jabeen-db-' + Utilities.formatDate(new Date(at), 'Asia/Riyadh', 'yyyy-MM-dd_HHmm') + '.json';
    var gz = Utilities.gzip(Utilities.newBlob(snap.txt, 'application/json', name), name + '.gz');
    var file = backupFolder_().createFile(gz);
    var extras = backupExtras_(at);
    var cut = at - BACKUP_KEEP_DAYS * 86400000, removed = 0, it = backupFolder_().getFiles();
    while (it.hasNext()) { var f = it.next(); if (/^jabeen-(db|rules|sheet|site)-/.test(f.getName()) && f.getDateCreated().getTime() < cut) { f.setTrashed(true); removed++; } }
    status = { at: at, ok: true, file: file.getName(), bytes: snap.txt.length, gzBytes: file.getSize(), bookings: Object.keys(jb.bookings || {}).length, units: Object.keys(jb.master || {}).length, customers: Object.keys(jb.customers || {}).length, removed: removed, extras: extras.ok, extrasFailed: extras.failed };
  } catch (e) {
    status.error = String(e.message || e);
    try { if (SUPER_ADMIN) sendMail_(SUPER_ADMIN, '⚠️ فشل النسخ الاحتياطي لنظام حجوزات جبين', '<div dir="rtl">فشل النسخ الاحتياطي اليومي: ' + escH_(status.error) + '<br>' + tsAr_(at) + '</div>'); } catch (e2) {}
  }
  try { fbWrite_(FIREBASE + '/backupStatus', status); } catch (e) {}
  return status;
}
function setupBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'dailyBackup') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('dailyBackup').timeBased().atHour(2).everyDays(1).inTimezone('Asia/Riyadh').create();
  return dailyBackup();
}
/** يعرض آخر النسخ في سجل التنفيذ */
function listBackups() {
  var out = [], it = backupFolder_().getFiles();
  while (it.hasNext()) { var f = it.next(); out.push([f.getDateCreated().getTime(), f.getName(), f.getSize()]); }
  out.sort(function (a, b) { return b[0] - a[0]; });
  var s = out.slice(0, 30).map(function (x) { return x[1] + '  ·  ' + Math.round(x[2] / 1024) + ' KB'; }).join('\n');
  Logger.log(s); return s;
}
/** استرجاع كامل: يستبدل محتوى قاعدة البيانات بالنسخة المحددة. يأخذ نسخة من الوضع الحالي أولاً احتياطاً.
 *  شغّلها من المحرر بعد كتابة اسم الملف في RESTORE_FILE أدناه. */
var RESTORE_FILE = ''; // مثال: 'jabeen-db-2026-11-10_0200.json.gz'
function restoreFromBackup(fileName) {
  fileName = fileName || RESTORE_FILE;
  if (!fileName) throw new Error('اكتب اسم ملف النسخة في RESTORE_FILE (من listBackups) ثم أعد التشغيل');
  var it = backupFolder_().getFilesByName(fileName); if (!it.hasNext()) throw new Error('الملف غير موجود: ' + fileName);
  var blob = it.next().getBlob(), txt = /\.gz$/.test(fileName) ? Utilities.ungzip(blob.setContentType('application/x-gzip')).getDataAsString('UTF-8') : blob.getDataAsString('UTF-8');
  var data = JSON.parse(txt); if (!data || !data.jabeenBookings) throw new Error('الملف لا يحتوي jabeenBookings');
  var before = dailyBackup(); if (!before.ok && !/فارغة/.test(before.error || '')) throw new Error('تعذر حفظ نسخة من الوضع الحالي قبل الاسترجاع: ' + before.error);
  var r = UrlFetchApp.fetch(DB_ROOT + '/.json', { method: 'put', contentType: 'application/json', headers: fbHeaders_(), payload: txt, muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('فشل الاسترجاع ' + r.getResponseCode() + ': ' + r.getContentText().slice(0, 200));
  audit_(null, 'استرجاع قاعدة البيانات من نسخة احتياطية', '', '', fileName);
  return 'تم الاسترجاع من ' + fileName + ' · الحجوزات: ' + Object.keys(data.jabeenBookings.bookings || {}).length;
}
/** للابتوب: يعيد النسخة الحالية كاملة بشرط الرمز السري */
function backupExport_(body) {
  var tok = PropertiesService.getScriptProperties().getProperty('BACKUP_TOKEN');
  if (!tok || String(body.token || '') !== tok) throw new Error('رمز النسخ غير صحيح');
  var snap = dbSnapshot_();
  return { ok: true, at: Date.now(), data: snap.data };
}
/** شغّلها مرة من المحرر: تنشئ الرمز السري للابتوب وتعرضه في سجل التنفيذ */
function setupBackupToken() {
  var props = PropertiesService.getScriptProperties(), tok = props.getProperty('BACKUP_TOKEN');
  if (!tok) { tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, ''); props.setProperty('BACKUP_TOKEN', tok); }
  Logger.log(tok); return tok;
}
/* مكونات النظام الأخرى مع النسخة اليومية: قواعد الأمان، الجدولان بصيغة Excel، وكود الموقع والسكربت (أسبوعياً يوم الجمعة أو إذا لم توجد نسخة) */
var SITE_REPO_ZIP = 'https://codeload.github.com/s2daho0omes2/jabeen-bookings/zip/refs/heads/main';
function backupExtras_(at) {
  var folder = backupFolder_(), stamp = Utilities.formatDate(new Date(at), 'Asia/Riyadh', 'yyyy-MM-dd_HHmm'), ok = [], failed = [];
  var auth = { headers: fbHeaders_(), muteHttpExceptions: true };
  function save(label, fn) { try { fn(); ok.push(label); } catch (e) { failed.push(label + ': ' + String(e.message || e).slice(0, 120)); } }
  save('rules', function () {
    var r = UrlFetchApp.fetch(DB_ROOT + '/.settings/rules.json', auth);
    if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
    folder.createFile('jabeen-rules-' + stamp + '.json', r.getContentText(), 'application/json');
  });
  function sheet(label, id) {
    save(label, function () {
      var r = UrlFetchApp.fetch('https://docs.google.com/spreadsheets/d/' + id + '/export?format=xlsx', auth);
      if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
      folder.createFile(r.getBlob().setName('jabeen-sheet-' + label + '-' + stamp + '.xlsx'));
    });
  }
  try { sheet('bookings', SpreadsheetApp.getActiveSpreadsheet().getId()); } catch (e) { failed.push('bookings sheet: ' + e.message); }
  sheet('allocation', ALLOC_SHEET_ID);
  var weekly = Utilities.formatDate(new Date(at), 'Asia/Riyadh', 'u') === '5' || !folder.searchFiles("title contains 'jabeen-site-' and trashed = false").hasNext();
  if (weekly) save('site', function () {
    var r = UrlFetchApp.fetch(SITE_REPO_ZIP, { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
    folder.createFile(r.getBlob().setName('jabeen-site-' + stamp + '.zip'));
  });
  ensureRestoreGuide_(folder);
  return { ok: ok.join(','), failed: failed.join(' | ') };
}
function ensureRestoreGuide_(folder) {
  var name = '00-دليل-الاسترجاع.txt';
  if (folder.getFilesByName(name).hasNext()) return;
  folder.createFile(name, [
    'دليل استرجاع نظام حجوزات مشروع جبين',
    '',
    'محتوى هذا المجلد (يُحدَّث يومياً الساعة 2 فجراً، ويُحتفظ بآخر 60 يوماً):',
    '- jabeen-db-*.json.gz : قاعدة البيانات كاملة (الحجوزات، العملاء، التعهدات، المستخدمون، الإعدادات، سجل التدقيق).',
    '- jabeen-rules-*.json : قواعد الأمان لقاعدة Firebase.',
    '- jabeen-sheet-bookings-*.xlsx و jabeen-sheet-allocation-*.xlsx : نسخة من الجدولين.',
    '- jabeen-site-*.zip : كود الموقع كاملاً مع كود السكربت (مجلد apps-script) — أسبوعياً.',
    'المستندات المرفوعة للعملاء محفوظة في مجلد jabeen-bookings-attachments في نفس Drive.',
    '',
    'استرجاع قاعدة البيانات:',
    '1) افتح السكربت: من جدول «حجوزات جبين» > الإضافات > Apps Script.',
    '2) شغّل listBackups لعرض النسخ، ثم اكتب اسم الملف في RESTORE_FILE وشغّل restoreFromBackup.',
    '   أو يدوياً: فك ضغط الملف، ثم في لوحة Firebase > Realtime Database > ⋮ > Import JSON على الجذر.',
    '',
    'استرجاع قواعد الأمان: لوحة Firebase > Realtime Database > Rules > الصق محتوى jabeen-rules-*.json > Publish.',
    'استرجاع الموقع: ارفع محتوى jabeen-site-*.zip إلى مستودع GitHub (s2daho0omes2/jabeen-bookings) مع تفعيل GitHub Pages والدومين jabeen-bookings.online.',
    'استرجاع السكربت: انسخ apps-script/Code.gs و appsscript.json من ملف الموقع إلى محرر Apps Script، ثم شغّل setupAllTriggers وانشر نسخة جديدة.',
    'الإعدادات السرية (Script Properties): RESEND_KEY من حساب Resend، ADMIN_EMAIL، BACKUP_TOKEN (تُنشأ بـ setupBackupToken).'
  ].join('\n'), 'text/plain');
}
