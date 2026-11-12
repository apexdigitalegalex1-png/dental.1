// ===== نظام العيادة - المرحلة 1: الأساس =====
// الصق الملف ده في Apps Script جوه Google Sheet، وشغّل setup() مرة واحدة.

const T = {
  Branches: ['id', 'name', 'address'],
  Rooms: ['id', 'branch', 'name'],
  Roles: ['id', 'name', 'perms', 'scope'],
  Employees: ['id', 'name', 'role', 'branch', 'fingerprintId', 'baseSalary', 'hourlyRate', 'commissionPct', 'phone', 'status', 'username', 'shiftStart', 'shiftEnd', 'weeklyOff'],
  Users: ['username', 'hash', 'role'],
  Procedures: ['id', 'name', 'duration', 'price', 'materialsCost', 'markupPct'],
  Schedules: ['id', 'doctor', 'branch', 'room', 'day', 'from', 'to', 'validFrom', 'validTo'],
  Patients: ['id', 'name', 'phone', 'age', 'allergies', 'chronic', 'source', 'notes'],
  Appointments: ['id', 'patient', 'doctor', 'branch', 'room', 'date', 'time', 'procedure', 'duration', 'status'],
  Attendance: ['id', 'fpId', 'employee', 'branch', 'date', 'timeIn', 'timeOut', 'minutes'],
  Adjustments: ['id', 'employee', 'month', 'type', 'amount', 'note'],
  Settings: ['id', 'label', 'value']
};
const BRN = ['Appointments', 'Schedules', 'Rooms', 'Attendance']; // جداول مرتبطة بفرع: اللي نطاقه "فرعه فقط" مايشوفش غير فرعه
const COSTS = ['materialsCost', 'markupPct']; // مخفية إلا لمن عنده صلاحية costs
const DAYS = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
// الصلاحية المطلوبة للتعديل على كل جدول
const PERM = { Employees: 'employees', Branches: 'branches', Rooms: 'branches', Roles: 'roles',
  Procedures: 'procedures', Schedules: 'schedules', Patients: 'patients', Appointments: 'appointments',
  Attendance: 'attendance', Adjustments: 'payroll', Settings: 'payroll' };

const ss = () => SpreadsheetApp.getActive();

function setup() {
  Object.keys(T).forEach(n => {
    const s = ss().getSheetByName(n) || ss().insertSheet(n);
    s.getRange(1, 1, s.getMaxRows(), T[n].length).setNumberFormat('@'); // نص عشان الأصفار والأوقات ما تتغيرش
    s.getRange(1, 1, 1, T[n].length).setValues([T[n]]); s.setFrozenRows(1);
  });
  const b = ss().getSheetByName('Branches');
  if (b.getLastRow() < 2) {
    ['العجمي', 'سموحة', 'محرم بك', 'كليوبترا', 'الزمالك', 'التجمع', 'المعادي', 'أكتوبر']
      .forEach((n, i) => b.appendRow(['B' + (i + 1), n, i < 4 ? 'الإسكندرية' : 'القاهرة']));
  }
  const r = ss().getSheetByName('Roles');
  if (r.getLastRow() < 2) {
    [['مدير', '*'], ['استقبال', 'patients,appointments'], ['دكتور', ''], ['تمريض', ''], ['عمال', '']]
      .forEach((x, i) => r.appendRow(['R' + (i + 1), x[0], x[1]]));
  }
  // أدوار الاستقبال + نطاق الرؤية (فرعه فقط / كل الفروع)
  [['استقبال فرع', 'patients,appointments'], ['استقبال ميديا', 'patients,appointments'], ['استقبال فون', 'patients,appointments']]
    .forEach(x => { if (!rows('Roles').some(r => r.name === x[0])) upsert('Roles', { id: 'R' + Utilities.getUuid().slice(0, 5), name: x[0], perms: x[1] }, 'id'); });
  const ALLB = ['مدير', 'مسؤول مخازن', 'استقبال ميديا', 'استقبال فون'];
  rows('Roles').forEach(r => { if (!r.scope) upsert('Roles', Object.assign({}, r, { scope: ALLB.includes(r.name) ? 'كل الفروع' : 'فرعه فقط' }), 'id'); });
  const st = ss().getSheetByName('Settings'); // قواعد المرتبات (تقدر تعدلها من الموقع)
  if (st.getLastRow() < 2) {
    [['graceMin', 'سماح التأخير (دقيقة)', 15], ['otGraceMin', 'سماح قبل احتساب الإضافي (دقيقة)', 30], ['daysPerMonth', 'أيام الشهر لحساب اليومية', 30],
     ['lateMultiplier', 'معامل خصم التأخير (1 = بنفس قيمة الوقت)', 1], ['absenceMultiplier', 'معامل خصم يوم الغياب', 1], ['overtimeMultiplier', 'معامل ساعة الإضافي', 1.5]]
      .forEach(x => st.appendRow(x));
  }
  const pr = ss().getSheetByName('Procedures');
  if (pr.getLastRow() < 2) {
    [['كشف', 30], ['تنظيف', 45], ['حشو', 60], ['خلع', 30], ['علاج عصب', 90]]
      .forEach((x, i) => pr.appendRow(['P' + (i + 1), x[0], x[1], 0, 0, 0]));
  }
  if (ss().getSheetByName('Users').getLastRow() < 2) {
    upsert('Users', { username: 'admin', hash: hash('admin123'), role: 'مدير' }, 'username');
  }
}

function salt() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('salt');
  if (!s) { s = Utilities.getUuid(); p.setProperty('salt', s); }
  return s;
}
const hash = pw => Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt() + pw));

function rows(n) {
  const v = ss().getSheetByName(n).getDataRange().getValues();
  const h = v.shift();
  return v.map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}

function upsert(n, o, key) {
  const s = ss().getSheetByName(n), h = T[n], v = s.getDataRange().getValues();
  const row = h.map(k => (o[k] === undefined ? '' : o[k]));
  const i = v.findIndex((r, j) => j > 0 && String(r[h.indexOf(key)]) === String(o[key]));
  if (i > 0) s.getRange(i + 1, 1, 1, h.length).setValues([row]); else s.appendRow(row);
}

function doPost(e) {
  let res;
  try { res = { ok: true, data: handle(JSON.parse(e.postData.contents)) }; }
  catch (err) { res = { ok: false, error: String(err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}

function handle(q) {
  if (q.action === 'login') {
    const u = rows('Users').find(x => x.username === q.username);
    if (!u || u.hash !== hash(q.password || '')) throw new Error('اسم الدخول أو كلمة السر غلط');
    const emp = rows('Employees').find(e => e.username === u.username);
    if (emp && emp.status === 'موقوف') throw new Error('الحساب موقوف');
    // الموظف ممكن يكون له أكتر من دور (مفصولين بفاصلة): الصلاحيات بتتجمع، ولو أي دور نطاقه "كل الفروع" يشوف كل الفروع
    const names = String(emp ? emp.role : u.role).split(',').map(s => s.trim());
    const rs = rows('Roles').filter(r => names.includes(r.name));
    const perms = [...new Set([].concat(...rs.map(r => String(r.perms || '').split(',').map(s => s.trim()).filter(Boolean))))];
    const all = perms.includes('*') || rs.some(r => r.scope === 'كل الفروع');
    const token = Utilities.getUuid();
    const user = { username: u.username, role: names.join(' + '), perms, all, branch: emp ? emp.branch : '' };
    CacheService.getScriptCache().put('t_' + token, JSON.stringify(user), 21600);
    return { token, user };
  }

  const u = JSON.parse(CacheService.getScriptCache().get('t_' + q.token) || 'null');
  if (!u) throw new Error('انتهت الجلسة، سجّل دخول تاني');
  const can = p => u.perms.includes('*') || u.perms.includes(p);
  if (q.action === 'payroll') {
    if (!can('payroll')) throw new Error('مفيش صلاحية');
    return payroll(q.month, q.branch, u);
  }
  if (q.action === 'importAttendance') {
    if (!can('attendance')) throw new Error('مفيش صلاحية');
    const lk = LockService.getScriptLock(); lk.waitLock(30000);
    try { return importAttendance(q, u); } finally { lk.releaseLock(); }
  }
  if (q.action === 'seedDemo' || q.action === 'clearDemo') { // للمدير فقط: تعبئة/مسح البيانات التجريبية من الموقع
    if (!u.perms.includes('*')) throw new Error('للمدير فقط');
    const lk = LockService.getScriptLock(); lk.waitLock(30000);
    try { if (q.action === 'seedDemo') seedDemoData(); else clearDemoData(); } finally { lk.releaseLock(); }
    return true;
  }
  const n = q.table;
  if (!T[n] || n === 'Users') throw new Error('جدول غير معروف');

  if (q.action === 'list') {
    let r = rows(n);
    if (['Patients', 'Appointments', 'Attendance', 'Adjustments', 'Settings'].includes(n) && !can(PERM[n])) throw new Error('مفيش صلاحية');
    if (n === 'Attendance' && q.month) r = r.filter(x => String(x.date).startsWith(q.month));
    if (n === 'Employees' && !can('employees')) r = r.map(x => ({ id: x.id, name: x.name, role: x.role, branch: x.branch }));
    if (BRN.includes(n) && !u.all) r = r.filter(x => x.branch === u.branch);
    if (n === 'Procedures' && !can('costs')) r = r.map(x => { COSTS.forEach(k => delete x[k]); return x; });
    return r;
  }
  if (!can(PERM[n])) throw new Error('مفيش صلاحية');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (q.action === 'save') {
      const o = q.row;
      if (!o.id) o.id = Utilities.getUuid().slice(0, 8);
      if (n === 'Attendance') { // يكمّل الكود والمدة ويثبّت رقم السجل
        const e = rows('Employees').find(x => x.name === o.employee);
        if (!e || e.fingerprintId === '') throw new Error('الموظف غير موجود أو ملوش كود بصمة');
        o.fpId = e.fingerprintId; o.id = o.fpId + '|' + o.date + '|' + o.branch;
        o.minutes = o.timeOut ? Math.max(0, toMin(o.timeOut) - toMin(o.timeIn)) : 0;
      }
      if (n === 'Procedures' && !can('costs')) { // اللي مالوش صلاحية ما يمسحش التكاليف
        const old = rows(n).find(x => String(x.id) === String(o.id)) || {};
        COSTS.forEach(k => { o[k] = old[k] === undefined ? '' : old[k]; });
      }
      if (BRN.includes(n) && !u.all) {
        const prev = rows(n).find(x => String(x.id) === String(o.id));
        if (o.branch !== u.branch || (prev && prev.branch !== u.branch)) throw new Error('تقدر تتعامل مع فرعك (' + u.branch + ') بس');
      }
      if (n === 'Appointments') checkAppt(o);
      upsert(n, o, 'id');
      if (n === 'Employees' && o.username && q.password) {
        if (!can('users')) throw new Error('مفيش صلاحية لإنشاء مستخدمين');
        upsert('Users', { username: o.username, hash: hash(q.password), role: o.role }, 'username');
      }
      return o;
    }
    if (q.action === 'remove') {
      const s = ss().getSheetByName(n), v = s.getDataRange().getValues();
      const i = v.findIndex((r, j) => j > 0 && String(r[0]) === String(q.id));
      if (i > 0 && !u.all && BRN.includes(n) && v[i][T[n].indexOf('branch')] !== u.branch) throw new Error('تقدر تتعامل مع فرعك بس');
      if (i > 0) s.deleteRow(i + 1);
      return true;
    }
  } finally { lock.releaseLock(); }
  throw new Error('أمر غير معروف');
}

const toMin = t => { const p = String(t).split(':'); return Number(p[0]) * 60 + Number(p[1]); };

// يتأكد إن الدكتور شغال في الفرع والغرفة في الوقت ده، وإن مفيش تعارض مع حجز تاني
function checkAppt(o) {
  if (o.status === 'ملغي') return;
  const day = DAYS[new Date(o.date + 'T00:00:00').getDay()];
  const st = toMin(o.time), en = st + Number(o.duration || 30);
  const ok = rows('Schedules').some(s => s.doctor === o.doctor && s.branch === o.branch && s.room === o.room &&
    s.day === day && String(s.validFrom) <= o.date && (!s.validTo || String(s.validTo) >= o.date) &&
    toMin(s.from) <= st && toMin(s.to) >= en);
  if (!ok) throw new Error('الدكتور مش متاح في الفرع/الغرفة دي في الميعاد ده حسب الجدول');
  rows('Appointments').forEach(a => {
    if (String(a.id) === String(o.id) || a.status === 'ملغي' || a.date !== o.date) return;
    const s2 = toMin(a.time), e2 = s2 + Number(a.duration || 30);
    if (st < e2 && s2 < en && (a.doctor === o.doctor || (a.branch === o.branch && a.room === o.room)))
      throw new Error('الميعاد ده متعارض مع حجز تاني (' + a.time + ')');
  });
}

// ===================== بيانات تجريبية =====================
// شغّل seedDemoData() من محرر Apps Script لتعبئة كل الأقسام ببيانات عشوائية.
// وشغّل clearDemoData() لمسح البيانات التجريبية بس (اللي أرقامها بتبدأ بـ D-).
// حسابات التجربة (كلمة السر demo123): demo_doctor / demo_reception / demo_store

const rnd = a => a[Math.floor(Math.random() * a.length)];
const pad = n => String(n).padStart(2, '0');
const hm = m => pad(Math.floor(m / 60)) + ':' + pad(m % 60);
const ymd = d => Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
const digits = n => String(Math.floor(Math.random() * Math.pow(10, n))).padStart(n, '0');

function bulk(n, objs) {
  if (!objs.length) return;
  const s = ss().getSheetByName(n);
  s.getRange(s.getLastRow() + 1, 1, objs.length, T[n].length)
    .setValues(objs.map(o => T[n].map(k => (o[k] === undefined ? '' : o[k]))));
}

function clearDemoData() {
  ['Rooms', 'Employees', 'Schedules', 'Patients', 'Appointments', 'Procedures', 'Attendance', 'Adjustments', 'Users'].forEach(n => {
    const s = ss().getSheetByName(n); if (!s) return;
    const v = s.getDataRange().getValues();
    const keep = v.slice(1).filter(r => !String(r[0]).startsWith(n === 'Users' ? 'demo_' : 'D-'));
    if (v.length > 1) s.getRange(2, 1, v.length - 1, T[n].length).clearContent();
    if (keep.length) s.getRange(2, 1, keep.length, T[n].length).setValues(keep);
  });
}

function seedDemoData() {
  setup(); clearDemoData();
  const br = rows('Branches').map(b => b.name);
  if (!rows('Roles').some(r => r.name === 'مسؤول مخازن'))
    upsert('Roles', { id: 'R6', name: 'مسؤول مخازن', perms: 'procedures,costs', scope: 'كل الفروع' }, 'id');

  // الإجراءات والأسعار (سعر المريض + سعر الأدوات جملة + نسبة الربح)
  [['P1', 'كشف', 30, 300, 20, 30], ['P2', 'تنظيف', 45, 600, 80, 40], ['P3', 'حشو', 60, 500, 100, 30],
   ['P4', 'خلع', 30, 400, 60, 30], ['P5', 'علاج عصب', 90, 1500, 300, 35], ['D-P6', 'تبييض', 60, 2500, 400, 40],
   ['D-P7', 'تركيب تاج', 60, 3000, 700, 35], ['D-P8', 'أشعة بانوراما', 15, 250, 30, 30],
   ['D-P9', 'تقويم - متابعة', 30, 500, 50, 30], ['D-P10', 'زراعة سن', 120, 15000, 5000, 25]]
    .forEach(x => upsert('Procedures', { id: x[0], name: x[1], duration: x[2], price: x[3], materialsCost: x[4], markupPct: x[5] }, 'id'));

  // الغرف: 3 في كل فرع
  const rooms = [];
  br.forEach((b, bi) => [1, 2, 3].forEach(i => rooms.push({ id: 'D-R' + bi + i, branch: b, name: 'غرفة ' + i })));
  bulk('Rooms', rooms);

  // الموظفين
  const emp = []; let k = 0;
  const add = (name, role, branch, base, hr, com) => {
    k++;
    const o = { id: 'D-E' + k, name, role, branch, fingerprintId: String(100 + k), baseSalary: base, hourlyRate: hr,
      commissionPct: com, phone: '010' + digits(8), status: 'نشط', username: '', shiftStart: '09:00', shiftEnd: '17:00', weeklyOff: 'الجمعة' };
    emp.push(o); return o;
  };
  const docs = ['د. أحمد سمير', 'د. منى عادل', 'د. محمد حسن', 'د. ياسمين فؤاد', 'د. كريم عبد الله', 'د. سارة مجدي',
    'د. عمر طارق', 'د. هبة نصر', 'د. خالد إبراهيم', 'د. نهى شريف'].map((n, i) => add(n, 'دكتور', br[i % br.length], 3000, 0, rnd([30, 35, 40, 45])));
  const rec = ['منال فتحي', 'رشا السيد', 'إيمان حسن', 'دعاء محمود', 'شيماء علي', 'نادية سالم', 'هالة جمال', 'أسماء رضا']
    .map((n, i) => add(n, 'استقبال فرع', br[i % br.length], 4500, 0, 0));
  ['سمر عادل', 'ولاء يوسف', 'مروة سعيد', 'جيهان أحمد', 'نجلاء فرج', 'رحاب طه', 'إسراء ماهر', 'أميرة زكي']
    .forEach((n, i) => add(n, 'تمريض', br[i % br.length], 4000, 0, 0));
  ['عم سيد', 'عم رمضان', 'أم محمد', 'عم عبده'].forEach((n, i) => add(n, 'عمال', br[(i * 2) % br.length], 0, 25, 0));
  const store = add('أستاذ هاني', 'مسؤول مخازن', br[0], 6000, 0, 0);
  const media = add('ليلى حمدي', 'استقبال ميديا', br[0], 4500, 0, 0), phone = add('بسمة رأفت', 'استقبال فون', br[4], 4500, 0, 0);
  add('هبة الله ناصر', 'استقبال فرع,استقبال ميديا,استقبال فون', br[1], 5000, 0, 0); // موظفة بتعمل الأدوار التلاتة
  media.username = 'demo_media'; phone.username = 'demo_phone';
  docs[0].username = 'demo_doctor'; rec[0].username = 'demo_reception'; store.username = 'demo_store';
  bulk('Employees', emp);
  const h = hash('demo123');
  bulk('Users', [{ username: 'demo_doctor', hash: h, role: 'دكتور' }, { username: 'demo_reception', hash: h, role: 'استقبال فرع' },
    { username: 'demo_media', hash: h, role: 'استقبال ميديا' }, { username: 'demo_phone', hash: h, role: 'استقبال فون' },
    { username: 'demo_store', hash: h, role: 'مسؤول مخازن' }]);

  // جداول الدكاترة
  const today = new Date();
  const dayOf = off => ymd(new Date(today.getFullYear(), today.getMonth(), today.getDate() + off));
  const pairs = [['السبت', 'الاثنين', 'الأربعاء'], ['الأحد', 'الثلاثاء', 'الخميس'], ['السبت', 'الثلاثاء'], ['الأحد', 'الأربعاء']];
  const sch = []; let sn = 0;
  const addS = (doc, branch, room, day, from, to, vf, vt) =>
    sch.push({ id: 'D-S' + (++sn), doctor: doc, branch, room: 'غرفة ' + room, day, from, to, validFrom: vf, validTo: vt || '' });
  docs.forEach((d, i) => pairs[i % 4].forEach(day =>
    addS(d.name, br[i % br.length], i < 8 ? 1 : 2, day, i % 2 ? '14:00' : '10:00', i % 2 ? '20:00' : '16:00', dayOf(-14))));
  addS(docs[0].name, br[4], 3, 'الخميس', '12:00', '18:00', dayOf(-14));                    // دكتور بيتنقل بين فرعين
  addS(docs[1].name, br[5], 2, 'السبت', '10:00', '15:00', dayOf(-60), dayOf(-15));         // جدول قديم منتهي
  bulk('Schedules', sch);

  // المرضى
  const fn = ['محمد', 'أحمد', 'محمود', 'مصطفى', 'علي', 'عمر', 'يوسف', 'كريم', 'هدى', 'منى', 'سارة', 'نور', 'ياسمين', 'فاطمة', 'آية', 'مريم', 'دينا', 'رنا'];
  const ln = ['السيد', 'حسن', 'إبراهيم', 'عبد الرحمن', 'فتحي', 'المصري', 'سالم', 'عطية', 'البدري', 'نصار'];
  const seen = {}, pats = [];
  while (pats.length < 40) {
    const name = rnd(fn) + ' ' + rnd(ln); if (seen[name]) continue; seen[name] = 1;
    pats.push({ id: 'D-PT' + (pats.length + 1), name, phone: '01' + rnd(['0', '1', '2', '5']) + digits(8), age: 6 + Math.floor(Math.random() * 65),
      allergies: rnd(['', '', '', 'بنسلين', 'أسبرين', 'مخدر موضعي', 'لاتكس']), chronic: rnd(['', '', '', 'سكر', 'ضغط', 'قلب', 'سكر وضغط']),
      source: rnd(['فيسبوك', 'انستجرام', 'صديق أو قريب', 'جوجل', 'إعلان', 'غير ذلك']), notes: '' });
  }
  bulk('Patients', pats);

  // المواعيد: متوزعة على جداول الدكاترة بدون تعارض (من أسبوع فات لـ 12 يوم جاي)
  const procs = rows('Procedures'), appts = []; let an = 0;
  for (let off = -7; off <= 12; off++) {
    const d = new Date(today.getFullYear(), today.getMonth(), today.getDate() + off), ds = ymd(d), day = DAYS[d.getDay()];
    sch.forEach(s => {
      if (s.day !== day || s.validFrom > ds || (s.validTo && s.validTo < ds)) return;
      let t = toMin(s.from) + rnd([0, 30]);
      const cnt = 2 + Math.floor(Math.random() * 3);
      for (let j = 0; j < cnt; j++) {
        const p = rnd(procs), dur = Number(p.duration);
        if (t + dur > toMin(s.to)) break;
        const pt = rnd(pats);
        const status = off < 0 ? rnd(['اتعمل', 'اتعمل', 'اتعمل', 'ملغي']) : off === 0 ? rnd(['محجوز', 'تم الحضور']) : 'محجوز';
        appts.push({ id: 'D-A' + (++an), patient: pt.name + ' - ' + pt.phone, doctor: s.doctor, branch: s.branch, room: s.room,
          date: ds, time: hm(t), procedure: p.name, duration: dur, status });
        t += dur + rnd([0, 15]);
      }
    });
  }
  bulk('Appointments', appts);

  // حضور وانصراف: الشهر اللي فات والشهر ده لحد النهارده (تأخير وغياب وإضافي عشوائي)
  const att = [];
  [-1, 0].forEach(mo => {
    const last = mo === 0 ? today : new Date(today.getFullYear(), today.getMonth(), 0);
    for (let d = new Date(today.getFullYear(), today.getMonth() + mo, 1); d <= last; d.setDate(d.getDate() + 1)) {
      if (DAYS[d.getDay()] === 'الجمعة') continue;
      emp.forEach(e => {
        if (Math.random() < 0.07) return; // غياب
        const tin = 540 + (Math.random() < 0.25 ? 20 + Math.floor(Math.random() * 40) : Math.floor(Math.random() * 12));
        const tout = 1020 + rnd([-20, 0, 0, 10, 45, 90]);
        att.push({ id: 'D-' + e.fingerprintId + '|' + ymd(d) + '|' + e.branch, fpId: e.fingerprintId, employee: e.name, branch: e.branch,
          date: ymd(d), timeIn: hm(tin), timeOut: hm(tout), minutes: tout - tin });
      });
    }
  });
  bulk('Attendance', att);
  const mon = ymd(today).slice(0, 7);
  bulk('Adjustments', [{ id: 'D-J1', employee: rec[0].name, month: mon, type: 'مكافأة', amount: 500, note: 'مكافأة تميز' },
    { id: 'D-J2', employee: docs[2].name, month: mon, type: 'سلفة', amount: 1000, note: 'سلفة' },
    { id: 'D-J3', employee: rec[1].name, month: mon, type: 'خصم', amount: 200, note: 'جزاء' }]);
  Logger.log('تم: %s موظف، %s جدول، %s مريض، %s ميعاد', emp.length, sch.length, pats.length, appts.length);
}

// ===================== الحضور والمرتبات =====================
function settings() {
  const o = {};
  rows('Settings').forEach(r => { if (r.value !== '' && !isNaN(Number(r.value))) o[r.id] = Number(r.value); });
  return Object.assign({ graceMin: 15, otGraceMin: 30, daysPerMonth: 30, lateMultiplier: 1, absenceMultiplier: 1, overtimeMultiplier: 1.5 }, o);
}

// الملف بيتجمّع في المتصفح: كل يوم لكل موظف = أول بصمة وآخر بصمة. هنا بنربطه بالموظف حسب كود البصمة ونحفظه.
function importAttendance(q, u) {
  if (!u.all && q.branch !== u.branch) throw new Error('تقدر ترفع حضور فرعك بس');
  const byFp = {};
  rows('Employees').forEach(e => { if (e.fingerprintId !== '') byFp[String(e.fingerprintId).trim()] = e; });
  const s = ss().getSheetByName('Attendance'), idx = {};
  s.getDataRange().getValues().forEach((r, i) => { if (i) idx[r[0]] = i + 1; });
  const add = [], unknown = {}; let upd = 0;
  q.days.forEach(d => {
    const e = byFp[String(d.fp).trim()];
    if (!e) { unknown[d.fp] = 1; return; }
    const tout = d.tout && toMin(d.tout) - toMin(d.tin) >= 30 ? d.tout : ''; // بصمة واحدة = انصراف ناقص
    const id = e.fingerprintId + '|' + d.date + '|' + q.branch;
    const row = [id, e.fingerprintId, e.name, q.branch, d.date, d.tin, tout, tout ? toMin(tout) - toMin(d.tin) : 0];
    if (idx[id]) { s.getRange(idx[id], 1, 1, 8).setValues([row]); upd++; } else add.push(row);
  });
  if (add.length) s.getRange(s.getLastRow() + 1, 1, add.length, 8).setValues(add);
  return { added: add.length, updated: upd, unknown: Object.keys(unknown) };
}

function payroll(month, branch, u) {
  const S = settings(), Y = Number(month.slice(0, 4)), M = Number(month.slice(5, 7));
  const dim = new Date(Y, M, 0).getDate(), today = ymd(new Date());
  const att = rows('Attendance').filter(a => String(a.date).startsWith(month));
  const adj = rows('Adjustments').filter(a => a.month === month);
  const price = {}; rows('Procedures').forEach(p => { price[p.name] = Number(p.price) || 0; });
  const done = rows('Appointments').filter(a => String(a.date).startsWith(month) && a.status === 'اتعمل');
  return rows('Employees').filter(e => e.status !== 'موقوف' && (!branch || e.branch === branch) && (u.all || e.branch === u.branch)).map(e => {
    const base = Number(e.baseSalary) || 0, hr = Number(e.hourlyRate) || 0, com = Number(e.commissionPct) || 0;
    const s0 = e.shiftStart || '09:00', s1 = e.shiftEnd || '17:00', off = e.weeklyOff || 'الجمعة';
    const shiftMin = Math.max(0, toMin(s1) - toMin(s0)), days = {};
    att.filter(a => e.fingerprintId !== '' && String(a.fpId) === String(e.fingerprintId)).forEach(a => {
      const d = days[a.date] = days[a.date] || { tin: '', tout: '', min: 0 };
      if (!d.tin || a.timeIn < d.tin) d.tin = a.timeIn;
      if (a.timeOut && a.timeOut > d.tout) d.tout = a.timeOut;
      d.min += Number(a.minutes) || 0;
    });
    let worked = 0, absent = 0, late = 0, ot = 0, mins = 0;
    Object.keys(days).forEach(k => { mins += days[k].min; });
    for (let d = 1; d <= dim; d++) {
      const ds = month + '-' + pad(d);
      if (DAYS[new Date(Y, M - 1, d).getDay()] === off || ds > today) continue;
      const x = days[ds];
      if (!x) { absent++; continue; }
      worked++;
      late += Math.max(0, toMin(x.tin) - toMin(s0) - S.graceMin);
      if (x.tout) ot += Math.max(0, toMin(x.tout) - toMin(s1) - S.otGraceMin);
    }
    const monthly = base > 0, daily = base / S.daysPerMonth, perMin = shiftMin ? daily / shiftMin : 0;
    const salary = monthly ? base : mins / 60 * hr;                       // اللي أساسيه صفر بيتحاسب بالساعة
    const lateDed = monthly ? late * perMin * S.lateMultiplier : 0;
    const absDed = monthly ? absent * daily * S.absenceMultiplier : 0;
    const otPay = monthly ? ot / 60 * (hr || (shiftMin ? daily / (shiftMin / 60) : 0)) * S.overtimeMultiplier : 0;
    const commission = com ? done.filter(a => a.doctor === e.name).reduce((t, a) => t + (price[a.procedure] || 0), 0) * com / 100 : 0;
    let bonus = 0, adjDed = 0;
    adj.filter(a => a.employee === e.name).forEach(a => { if (a.type === 'مكافأة') bonus += Number(a.amount) || 0; else adjDed += Number(a.amount) || 0; });
    const r = Math.round;
    return { id: e.id, name: e.name, role: e.role, branch: e.branch, base, days: worked, absent, lateMin: late, otMin: ot, hours: Math.round(mins / 6) / 10,
      salary: r(salary), otPay: r(otPay), commission: r(commission), bonus: r(bonus), lateDed: r(lateDed), absDed: r(absDed), adjDed: r(adjDed),
      net: r(salary + otPay + commission + bonus - lateDed - absDed - adjDed) };
  });
}
