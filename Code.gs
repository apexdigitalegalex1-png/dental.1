// ===== نظام العيادة - المرحلة 1: الأساس =====
// الصق الملف ده في Apps Script جوه Google Sheet، وشغّل setup() مرة واحدة.

const T = {
  Branches: ['id', 'name', 'address'],
  Rooms: ['id', 'branch', 'name'],
  Roles: ['id', 'name', 'perms'],
  Employees: ['id', 'name', 'role', 'branch', 'fingerprintId', 'baseSalary', 'hourlyRate', 'commissionPct', 'phone', 'status', 'username'],
  Users: ['username', 'hash', 'role'],
  Procedures: ['id', 'name', 'duration', 'price', 'materialsCost', 'markupPct'],
  Schedules: ['id', 'doctor', 'branch', 'room', 'day', 'from', 'to', 'validFrom', 'validTo'],
  Patients: ['id', 'name', 'phone', 'age', 'allergies', 'chronic', 'source', 'notes'],
  Appointments: ['id', 'patient', 'doctor', 'branch', 'room', 'date', 'time', 'procedure', 'duration', 'status']
};
const COSTS = ['materialsCost', 'markupPct']; // مخفية إلا لمن عنده صلاحية costs
const DAYS = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];
// الصلاحية المطلوبة للتعديل على كل جدول
const PERM = { Employees: 'employees', Branches: 'branches', Rooms: 'branches', Roles: 'roles',
  Procedures: 'procedures', Schedules: 'schedules', Patients: 'patients', Appointments: 'appointments' };

const ss = () => SpreadsheetApp.getActive();

function setup() {
  Object.keys(T).forEach(n => {
    const s = ss().getSheetByName(n) || ss().insertSheet(n);
    s.getRange(1, 1, s.getMaxRows(), T[n].length).setNumberFormat('@'); // نص عشان الأصفار والأوقات ما تتغيرش
    if (s.getLastRow() === 0) { s.appendRow(T[n]); s.setFrozenRows(1); }
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
    const role = rows('Roles').find(r => r.name === u.role);
    const perms = String((role && role.perms) || '').split(',').map(s => s.trim()).filter(Boolean);
    const token = Utilities.getUuid();
    const user = { username: u.username, role: u.role, perms };
    CacheService.getScriptCache().put('t_' + token, JSON.stringify(user), 21600);
    return { token, user };
  }

  const u = JSON.parse(CacheService.getScriptCache().get('t_' + q.token) || 'null');
  if (!u) throw new Error('انتهت الجلسة، سجّل دخول تاني');
  const can = p => u.perms.includes('*') || u.perms.includes(p);
  const n = q.table;
  if (!T[n] || n === 'Users') throw new Error('جدول غير معروف');

  if (q.action === 'list') {
    let r = rows(n);
    if ((n === 'Patients' || n === 'Appointments') && !can(PERM[n])) throw new Error('مفيش صلاحية');
    if (n === 'Employees' && !can('employees')) r = r.map(x => ({ id: x.id, name: x.name, role: x.role, branch: x.branch }));
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
      if (n === 'Procedures' && !can('costs')) { // اللي مالوش صلاحية ما يمسحش التكاليف
        const old = rows(n).find(x => String(x.id) === String(o.id)) || {};
        COSTS.forEach(k => { o[k] = old[k] === undefined ? '' : old[k]; });
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
