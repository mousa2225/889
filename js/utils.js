// ============================
// Theme
// ============================
function togTh() {
  isDk = !isDk;
  document.documentElement.className = isDk ? '' : 'lt';
  document.body.style.background = 'var(--bg)';
  document.querySelectorAll('.tb i').forEach(function(i) {
    i.className = isDk ? 'fa-solid fa-sun' : 'fa-solid fa-moon';
  });
  localStorage.setItem('th', isDk ? 'd' : 'l');
  document.getElementById('exM').classList.remove('on');
}
(function() {
  if (localStorage.getItem('th') === 'l') { isDk = false; document.documentElement.className = 'lt'; }
})();

// ============================
// String Helpers
// ============================
function en(n) {
  if (n == null) return '';
  return String(n)
    .replace(/[٠-٩]/g, function(d) { return '٠١٢٣٤٥٦٧٨٩'.indexOf(d); })
    .replace(/[۰-۹]/g, function(d) { return '۰۱۲۳۴۵۶٧۸۹'.indexOf(d); });
}
function esc(t) {
  if (!t) return '';
  var d = document.createElement('div');
  d.textContent = t;
  return d.innerHTML;
}
function escA(s) {
  if (!s) return '';
  return esc(s).replace(/'/g, '&#39;').replace(/"/g, '&quot;');
}

// ============================
// Status / Permission Helpers
// ============================
function getSt(c) {
  if (c.status) return c.status;
  return c.refunded ? 'refunded' : 'uploaded';
}
// Helper: هل السجل مرفوض
function isRejected(c) { return getSt(c) === 'rejected'; }
function gP() {
  if (!cu) return { canAdd: true, canDelete: true, canEdit: true };
  if (cu.role === 'admin') return { canAdd: true, canDelete: true, canEdit: true };
  return cu.permissions || { canAdd: true, canDelete: true, canEdit: true };
}
function cRf(p, d, c) {
  if (d <= 0 || p <= 0) return 0;
  return p - ((p / d) * c);
}
function shTH() {
  return hideTm && cu && cu.role !== 'admin';
}

// ============================
// Date Formatting
// ============================
function fmtDt(ts) {
  if (!ts || !ts.seconds) return '—';
  var d = new Date(ts.seconds * 1000);
  return d.getDate() + ' ' + MA[d.getMonth()] + ' ' + d.getFullYear();
}
function fmtDtT(ts) {
  if (!ts || !ts.seconds) return '—';
  var d = new Date(ts.seconds * 1000);
  return d.getDate() + ' ' + MA[d.getMonth()] + ' ' + d.getFullYear() + ' ' +
    String(d.getHours()).padStart(2, '0') + ':' + String(d.getMinutes()).padStart(2, '0');
}
function fmtDtS(ts) {
  if (!ts || !ts.seconds) return '—';
  var d = new Date(ts.seconds * 1000);
  return d.getDate() + ' ' + MA[d.getMonth()];
}

// ============================
// Date String Normalizer & Parser
//
// يدعم جميع الأشكال:
//   "8 May"       → "8 May 2026"  (يضيف السنة الحالية وقت الإدخال)
//   "May 8"       → "May 8 2026"
//   "8 مايو"      → "8 May 2026"
//   "8 May 2025"  → لا يغير شيء
//   "8-5-2026"    → DD-MM-YYYY
//   "8/5/2026"    → DD/MM/YYYY
// ============================
var AR_MONTHS_MAP = {
  'يناير':'January','فبراير':'February','مارس':'March',
  'أبريل':'April','ابريل':'April',
  'مايو':'May','يونيو':'June','يوليو':'July',
  'أغسطس':'August','اغسطس':'August',
  'سبتمبر':'September','أكتوبر':'October','اكتوبر':'October',
  'نوفمبر':'November','ديسمبر':'December','دیسمبر':'December'
};
var EN_MONTH_IDX = {
  'january':0,'february':1,'march':2,'april':3,'may':4,'june':5,
  'july':6,'august':7,'september':8,'october':9,'november':10,'december':11,
  'jan':0,'feb':1,'mar':2,'apr':3,'jun':5,'jul':6,'aug':7,
  'sep':8,'sept':8,'oct':9,'nov':10,'dec':11
};

// تحويل نص التاريخ: أشهر عربية → إنجليزية + إضافة السنة إن غابت
function normDateStr(s) {
  if (!s || !s.trim()) return s;
  s = s.trim();
  for (var ar in AR_MONTHS_MAP) {
    if (s.indexOf(ar) !== -1)
      s = s.replace(new RegExp(ar, 'g'), AR_MONTHS_MAP[ar]);
  }
  // إضافة السنة الحالية وقت الإدخال إن لم تكن موجودة
  if (!/\d{4}/.test(s)) {
    s = s + ' ' + new Date().getFullYear();
  }
  return s;
}

// تحليل نص التاريخ → Date object
// يعمل مع أي ترتيب وأي شكل
function parseSubDate(s) {
  if (!s || !s.trim()) return null;

  // خطوة 1: تطبيع النص
  var ns = normDateStr(s.trim());

  // خطوة 2: نمط DD-MM-YYYY أو DD/MM/YYYY
  var numMatch = ns.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
  if (numMatch) {
    var d = new Date(parseInt(numMatch[3]), parseInt(numMatch[2]) - 1, parseInt(numMatch[1]));
    if (!isNaN(d.getTime())) return d;
  }

  // خطوة 3: token-based — يوم + اسم شهر + سنة بأي ترتيب
  var tokens = ns.replace(/,/g, ' ').split(/[\s\-\/]+/).filter(Boolean);
  var day = null, month = null, year = null;
  tokens.forEach(function(tok) {
    var tl = tok.toLowerCase();
    if (/^\d{4}$/.test(tok)) {
      year = parseInt(tok);
    } else if (EN_MONTH_IDX[tl] !== undefined) {
      month = EN_MONTH_IDX[tl];
    } else if (/^\d{1,2}$/.test(tok) && day === null) {
      day = parseInt(tok);
    }
  });
  if (day !== null && month !== null && year !== null) {
    var d2 = new Date(year, month, day);
    if (!isNaN(d2.getTime())) return d2;
  }

  // خطوة 4: native parse كاحتياط
  var d3 = new Date(ns);
  if (!isNaN(d3.getTime())) return d3;

  return null;
}

// ============================
// UI Helpers
// ============================
function hAE() {
  ['lEr','rEr','fEr'].forEach(function(id) {
    var e = document.getElementById(id); if (e) e.style.display = 'none';
  });
  ['lU','lP','rU','rP2','fU','fP'].forEach(function(id) {
    var e = document.getElementById(id); if (e) e.classList.remove('er');
  });
}
function sE(id, m) {
  var e = document.getElementById(id);
  if (e) { e.style.display = 'flex'; if (m) e.querySelector('span').textContent = m; }
}
function sO(id, m) {
  var e = document.getElementById(id);
  if (e) { e.style.display = 'block'; if (m) e.querySelector('span').textContent = m; }
}
function hO(id) {
  var e = document.getElementById(id); if (e) e.style.display = 'none';
}
function sCf(m, cb) {
  document.getElementById('cfMsg').textContent = m;
  document.getElementById('cfM').classList.add('on');
  cfCb = cb;
}
function cfOk() {
  document.getElementById('cfM').classList.remove('on');
  if (cfCb) { var fn = cfCb; cfCb = null; fn(); }
}
function cfNo() {
  document.getElementById('cfM').classList.remove('on');
  cfCb = null;
}

// ============================
// Toast Notifications
// ============================
function toast(m, tp) {
  var b = document.getElementById('tBox'), t = document.createElement('div');
  var co = { s:'rgba(34,197,94,.12)', e:'rgba(220,53,69,.12)', i:'rgba(0,212,170,.12)', w:'rgba(245,158,11,.12)' };
  var bs = { s:'rgba(34,197,94,.25)', e:'rgba(220,53,69,.25)', i:'rgba(0,212,170,.25)', w:'rgba(245,158,11,.25)' };
  var ic = { s:'fa-solid fa-circle-check', e:'fa-solid fa-circle-xmark', i:'fa-solid fa-circle-info', w:'fa-solid fa-circle-exclamation' };
  var cl = { s:'#22c55e', e:'#dc3545', i:'#00d4aa', w:'#f59e0b' };
  var k = tp || 'i';
  t.className = 'tI';
  t.style.cssText = 'padding:9px 16px;border-radius:10px;font-size:12px;font-weight:600;display:flex;align-items:center;gap:7px;background:' + co[k] + ';border:1px solid ' + bs[k] + ';color:var(--tx);backdrop-filter:blur(8px);box-shadow:0 4px 20px rgba(0,0,0,.2);min-width:200px;max-width:360px;pointer-events:auto';
  t.innerHTML = '<i class="' + ic[k] + '" style="color:' + cl[k] + ';font-size:14px;flex-shrink:0"></i><span>' + esc(m) + '</span>';
  b.appendChild(t);
  setTimeout(function() {
    t.className = 'tO';
    setTimeout(function() { if (t.parentNode) t.parentNode.removeChild(t); }, 200);
  }, 2800);
}

// ============================
// Server Clock Sync
// ============================
function syncST() {
  if (!db) return;
  var t1 = Date.now();
  db.collection('_ts').doc('p').set({ t: firebase.firestore.FieldValue.serverTimestamp() }, { merge: true })
    .then(function() { return db.collection('_ts').doc('p').get(); })
    .then(function(s) {
      if (!s.exists) return;
      var v = s.data().t;
      srvOff = (v.seconds * 1000 + (v.nanoseconds || 0) / 1e6) - ((t1 + Date.now()) / 2);
    }).catch(function() { srvOff = 0; });
}
function gSN() { return new Date(Date.now() + srvOff); }
function sSC() {
  if (clkInt) clearInterval(clkInt);
  clkInt = setInterval(function() {
    var e = document.getElementById('srvDisp'); if (!e) return;
    var n = gSN();
    e.textContent = String(n.getHours()).padStart(2,'0') + ':' + String(n.getMinutes()).padStart(2,'00') + ':' + String(n.getSeconds()).padStart(2,'0');
  }, 1000);
}

// ============================
// Date Filter UI
// ============================
function initDF() {
  var ms = document.getElementById('dtM'), ys = document.getElementById('dtMY'), yss = document.getElementById('dtY');
  for (var i = 0; i < 12; i++) {
    var o = document.createElement('option'); o.value = i; o.textContent = MA[i]; ms.appendChild(o);
  }
  var cy = new Date().getFullYear();
  for (var y = cy - 2; y <= cy + 1; y++) {
    var o1 = document.createElement('option'); o1.value = y; o1.textContent = y; ys.appendChild(o1);
    var o2 = document.createElement('option'); o2.value = y; o2.textContent = y; yss.appendChild(o2);
  }
  var sel1 = document.getElementById('dtMY'), sel2 = document.getElementById('dtY');
  if (sel1) sel1.value = cy;
  if (sel2) sel2.value = cy;
}
function cDT() {
  var v = document.getElementById('dtType').value;
  document.getElementById('dtDateBox').style.display  = v === 'date'  ? 'block' : 'none';
  document.getElementById('dtRangeBox').style.display = v === 'range' ? 'block' : 'none';
  document.getElementById('dtRangeBox2').style.display= v === 'range' ? 'block' : 'none';
  document.getElementById('dtMonthBox').style.display = v === 'month' ? 'block' : 'none';
  document.getElementById('dtYearBox').style.display  = v === 'year'  ? 'block' : 'none';
}
function aDF() {
  var tp = document.getElementById('dtType').value, inf = '', fldLbl = '';
  dtFlFld = document.getElementById('dtField').value;
  if      (dtFlFld === 'refund')       fldLbl = ' (تاريخ الاسترداد)';
  else if (dtFlFld === 'subscription') fldLbl = ' (تاريخ الاشتراك)';
  else if (dtFlFld === 'cancel')       fldLbl = ' (تاريخ الإلغاء)';
  else                                 fldLbl = ' (تاريخ الإضافة)';

  if (tp === 'date') {
    var dv = document.getElementById('dtDate').value;
    if (!dv) { toast('اختر تاريخ','e'); return; }
    dtFlA = 'date:' + dv;
    var p = dv.split('-'); inf = 'يوم: ' + parseInt(p[2]) + ' ' + MA[parseInt(p[1])-1] + ' ' + p[0] + fldLbl;
  } else if (tp === 'range') {
    var fr = document.getElementById('dtFrom').value, to = document.getElementById('dtTo').value;
    if (!fr || !to) { toast('اختر التاريخين','e'); return; }
    dtFlA = 'range:' + fr + '|' + to;
    var a = fr.split('-'), b = to.split('-');
    inf = 'من ' + parseInt(a[2]) + ' ' + MA[parseInt(a[1])-1] + ' إلى ' + parseInt(b[2]) + ' ' + MA[parseInt(b[1])-1] + fldLbl;
  } else if (tp === 'month') {
    var m = parseInt(document.getElementById('dtM').value), y = parseInt(document.getElementById('dtMY').value);
    dtFlA = 'month:' + y + '-' + m;
    inf = 'شهر: ' + MA[m] + ' ' + y + fldLbl;
  } else {
    var yr = parseInt(document.getElementById('dtY').value);
    dtFlA = 'year:' + yr; inf = 'سنة: ' + yr + fldLbl;
  }
  document.getElementById('dtInfo').innerHTML = '<span style="color:var(--dm)"><i class="fa-solid fa-calendar-days text-warm"></i> ' + inf + '</span>';
  document.getElementById('dtInfo').style.display = 'flex';
  document.getElementById('dtRstBtn').style.display = 'inline';
  cfl = 'all';
  document.querySelectorAll('#flS .ft').forEach(function(x) { x.classList.remove('on'); });
  rnT();
}
function rDF() {
  dtFlA = false; dtFlFld = 'created';
  document.getElementById('dtField').value = 'created';
  document.getElementById('dtInfo').style.display = 'none';
  document.getElementById('dtRstBtn').style.display = 'none';
  cfl = 'all';
  document.querySelectorAll('#flS .ft').forEach(function(x) { x.classList.remove('on'); });
  document.querySelector('#flS .ft').classList.add('on');
  rnT();
}

// ============================
// Date Filter Matcher
// يدعم: تاريخ الإضافة، الاسترداد، الاشتراك، الإلغاء
// فلتر الشهر مُصلح: substring(6)
// ============================
function mDF(c) {
  if (!dtFlA) return true;

  var d;
  if (dtFlFld === 'refund') {
    var ts = c.refundDate;
    if (!ts || !ts.seconds) return false;
    d = new Date(ts.seconds * 1000);
  } else if (dtFlFld === 'subscription') {
    var pd = parseSubDate(c.subscriptionDate);
    if (!pd) return false;
    d = pd;
  } else if (dtFlFld === 'cancel') {
    var pd = parseSubDate(c.cancelDate);
    if (!pd) return false;
    d = pd;
  } else {
    var ts = c.createdAt;
    if (!ts || !ts.seconds) return false;
    d = new Date(ts.seconds * 1000);
  }

  if (dtFlA.startsWith('date:')) {
    var p = dtFlA.substring(5).split('-');
    var f = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2])); f.setHours(0,0,0,0);
    var t = new Date(f); t.setHours(23,59,59,999);
    return d >= f && d <= t;
  }
  if (dtFlA.startsWith('range:')) {
    var parts = dtFlA.substring(6).split('|');
    var fp = parts[0].split('-'), tp2 = parts[1].split('-');
    var f2 = new Date(parseInt(fp[0]), parseInt(fp[1])-1, parseInt(fp[2])); f2.setHours(0,0,0,0);
    var t2 = new Date(parseInt(tp2[0]), parseInt(tp2[1])-1, parseInt(tp2[2])); t2.setHours(23,59,59,999);
    return d >= f2 && d <= t2;
  }
  if (dtFlA.startsWith('month:')) {
    var mv = dtFlA.substring(6).split('-'); // FIXED: (6) not (7)
    return d.getFullYear() === parseInt(mv[0]) && d.getMonth() === parseInt(mv[1]);
  }
  if (dtFlA.startsWith('year:')) return d.getFullYear() === parseInt(dtFlA.substring(5));
  return true;
}
