// ============================
// Admin Panel Open/Close
// ============================
function openAP() { document.getElementById('adM').classList.add('on'); ldPd(); chkP(); }
function clAP()   { document.getElementById('adM').classList.remove('on'); }
function swAT(t, btn) {
  ['adP','adA','adK','adU','adS'].forEach(function(i) { document.getElementById(i).style.display = 'none'; });
  ['abP','abA','abK','abU','abS'].forEach(function(i) { document.getElementById(i).classList.remove('on'); });
  btn.classList.add('on');
  document.getElementById({ p:'adP', a:'adA', k:'adK', u:'adU', s:'adS' }[t]).style.display = 'block';
  if (t==='p') ldPd();
  if (t==='a') ldAp();
  if (t==='k') ldPR();
  if (t==='u') ldUr();
  if (t==='s') ldSt();
}

// ============================
// Pending Requests Tab
// ============================
function ldPd() {
  var up = cls.filter(function(c) { return getSt(c) === 'uploaded'; });
  var pn = cls.filter(function(c) { return getSt(c) === 'pending'; });
  var tot = up.length + pn.length;
  document.getElementById('pCt').textContent = '(' + tot + ')';
  var el = document.getElementById('adP');
  if (!tot) {
    el.innerHTML = '<div class="text-center py-8" style="color:var(--mt)"><i class="fa-solid fa-circle-check text-2xl mb-2 block text-success"></i><p class="font-bold text-sm">لا توجد طلبات معلقة</p></div>';
    return;
  }
  var h = '';
  if (up.length) {
    h += '<div class="text-xs font-bold mb-2" style="color:var(--dm)"><i class="fa-solid fa-cloud-arrow-up" style="color:#94a3b8"></i> تم الرفع (' + up.length + ')</div>';
    up.forEach(function(c) {
      h += '<div class="rounded-lg p-3 mb-2" style="background:var(--cs);border:1px solid var(--brd);border-right:3px solid #94a3b8"><div class="flex items-start justify-between gap-3"><div class="flex-1"><div class="flex items-center gap-1.5 mb-1"><span class="font-bold text-sm">'+esc(c.name)+'</span>'+(c.previouslyRefunded?'<span class="pb-t"><i class="fa-solid fa-triangle-exclamation"></i> سبق</span>':'')+'</div><div class="text-[11px]" style="color:var(--dm)"><div class="num"><i class="fa-solid fa-phone ml-1" style="color:var(--mt)"></i>'+esc(c.phone||c.mobile||'—')+'</div><div><i class="fa-solid fa-utensils ml-1" style="color:var(--mt)"></i>'+esc(c.packageType)+'</div></div></div><div class="text-left shrink-0"><div class="text-[10px] mb-1.5" style="color:var(--dm)">المبلغ: <span class="text-accent font-bold num">'+en((c.refundAmount||0).toFixed(2))+'</span></div><button type="button" onclick="toPn(\''+c.id+'\')" class="bp bp-warm bp-sm px-3"><i class="fa-solid fa-arrow-left ml-1"></i>تحويل</button></div></div></div>';
    });
  }
  if (pn.length) {
    h += '<div class="text-xs font-bold mb-2 mt-4" style="color:var(--dm)"><i class="fa-solid fa-clock text-warm"></i> قيد المراجعة (' + pn.length + ')</div>';
    pn.forEach(function(c) {
      h += '<div class="rounded-lg p-3 mb-2" style="background:var(--cs);border:1px solid var(--brd);border-right:3px solid #e97a2a"><div class="flex items-start justify-between gap-3"><div class="flex-1"><div class="flex items-center gap-1.5 mb-1"><span class="font-bold text-sm">'+esc(c.name)+'</span>'+(c.previouslyRefunded?'<span class="pb-t"><i class="fa-solid fa-triangle-exclamation"></i> سبق</span>':'')+'</div><div class="text-[11px]" style="color:var(--dm)"><div class="num"><i class="fa-solid fa-phone ml-1" style="color:var(--mt)"></i>'+esc(c.phone||c.mobile||'—')+'</div><div><i class="fa-solid fa-utensils ml-1" style="color:var(--mt)"></i>'+esc(c.packageType)+'</div></div></div><div class="text-left shrink-0"><div class="text-[10px] mb-1.5" style="color:var(--dm)">المبلغ: <span class="text-accent font-bold num">'+en((c.refundAmount||0).toFixed(2))+'</span></div><button type="button" onclick="apR(\''+c.id+'\')" class="bp bp-sm px-3"><i class="fa-solid fa-check ml-1"></i>موافقة</button></div></div></div>';
    });
  }
  el.innerHTML = h;
}

// ============================
// Pending Accounts Tab
// ============================
function ldAp() {
  db.collection('users').where('approved','==',false).get().then(function(s) {
    var n = s.size, bg = document.getElementById('apCt');
    if (bg) { bg.textContent = '(' + n + ')'; bg.style.display = n > 0 ? 'inline' : 'none'; }
    var el = document.getElementById('adA');
    if (!n) { el.innerHTML = '<div class="text-center py-8" style="color:var(--mt)"><i class="fa-solid fa-user-check text-2xl mb-2 block text-success"></i><p class="font-bold text-sm">لا توجد حسابات معلقة</p></div>'; return; }
    el.innerHTML = s.docs.map(function(d) {
      var u = d.data();
      return '<div class="rounded-lg p-3 mb-2.5 flex items-center justify-between gap-3" style="background:var(--cs);border:1px solid var(--brd)"><div class="flex items-center gap-2.5"><div class="w-9 h-9 rounded-full flex items-center justify-center font-bold text-sm" style="background:rgba(233,122,42,.1);color:#e97a2a">'+u.username.charAt(0).toUpperCase()+'</div><div><div class="font-bold text-sm">'+esc(u.username)+'</div><div class="text-[10px] mt-0.5" style="color:var(--mt)">'+fmtDt(u.createdAt)+'</div></div></div><div class="flex items-center gap-1.5"><button type="button" onclick="apU(\''+d.id+'\')" class="bp bp-sm px-3"><i class="fa-solid fa-check ml-1"></i>قبول</button><button type="button" onclick="rjU(\''+d.id+'\',\''+escA(u.username)+'\')" class="bd py-1.5 px-2.5"><i class="fa-solid fa-xmark ml-1"></i>رفض</button></div></div>';
    }).join('');
  }).catch(function() { toast('خطأ','e'); });
}
function apU(id) {
  db.collection('users').doc(id).update({ approved: true })
    .then(function() { toast('تم قبول الحساب','s'); ldAp(); ldUr(); chkP(); })
    .catch(function() { toast('خطأ','e'); });
}
function rjU(id, un) {
  sCf('حذف حساب "' + un.replace(/&#39;/g,"'") + '"؟', function() {
    db.collection('users').doc(id).delete()
      .then(function() { toast('تم رفض الحساب','i'); ldAp(); ldUr(); chkP(); })
      .catch(function() { toast('خطأ','e'); });
  });
}

// ============================
// Password Reset Requests
// ============================
function ldPR() {
  var el = document.getElementById('adK');
  el.innerHTML = '<div class="text-center py-6"><span style="display:inline-block;width:18px;height:18px;border:2px solid var(--brd);border-top-color:#00d4aa;border-radius:50%;animation:spin .6s linear infinite"></span></div>';
  db.collection('passwordResetRequests').where('status','==','pending').get().then(function(s) {
    var docs = s.docs.sort(function(a,b) { return (b.data().requestedAt?b.data().requestedAt.seconds:0)-(a.data().requestedAt?a.data().requestedAt.seconds:0); });
    var n = docs.length, bg = document.getElementById('pwCt');
    if (bg) { bg.textContent = '(' + n + ')'; bg.style.display = n > 0 ? 'inline' : 'none'; }
    if (!n) { el.innerHTML = '<div class="text-center py-8" style="color:var(--mt)"><i class="fa-solid fa-key text-2xl mb-2 block text-success"></i><p class="font-bold text-sm">لا توجد طلبات</p></div>'; return; }
    el.innerHTML = docs.map(function(d) {
      var r = d.data(), ms = '';
      for (var i = 0; i < r.newPassword.length; i++) ms += '●';
      return '<div class="rounded-xl p-4 mb-3" style="background:var(--cs);border:1px solid var(--brd);border-right:4px solid #f59e0b"><div class="flex items-center justify-between gap-4 flex-wrap"><div class="flex items-center gap-3"><div class="w-11 h-11 rounded-xl flex items-center justify-center" style="background:rgba(245,158,11,.1);color:#f59e0b"><i class="fa-solid fa-key text-sm"></i></div><div><div class="font-bold text-sm mb-1">'+esc(r.username)+'</div><div class="text-[10px]" style="color:var(--mt)">'+fmtDtT(r.requestedAt)+'</div></div></div><div class="flex items-center gap-2"><button type="button" onclick="apPR(\''+d.id+'\',\''+escA(r.username)+'\',\''+escA(r.newPassword)+'\')" class="bp bp-sm px-4 py-1.5"><i class="fa-solid fa-check ml-1"></i>موافقة</button><button type="button" onclick="rjPR(\''+d.id+'\')" class="bd py-1.5 px-3 text-xs"><i class="fa-solid fa-xmark ml-1"></i>رفض</button></div></div><div class="mt-3 pt-3 flex items-center gap-2" style="border-top:1px solid var(--brd)"><span class="text-[11px] font-bold" style="color:var(--dm)"><i class="fa-solid fa-lock ml-1"></i>كلمة المرور:</span><span class="lk-cell px-2.5 py-1 rounded-lg" onclick="tPV(this,\''+escA(r.newPassword)+'\')" style="background:rgba(245,158,11,.08);border:1px solid rgba(245,158,11,.15);color:#f59e0b;font-weight:800;letter-spacing:2px;font-size:13px;cursor:pointer;min-width:80px;text-align:center;display:inline-block" data-m="'+ms+'" data-s="0">'+ms+'</span><span class="text-[9px]" style="color:var(--mt)">(اضغط لعرض)</span></div></div>';
    }).join('');
  }).catch(function() { el.innerHTML = '<div class="text-center py-8" style="color:#dc3545"><p class="font-bold text-sm">خطأ في التحميل</p></div>'; });
}
function tPV(el, pw) {
  if (el.dataset.s === '1') { el.textContent = el.dataset.m; el.dataset.s = '0'; el.style.letterSpacing = '2px'; }
  else { el.textContent = pw; el.dataset.s = '1'; el.style.letterSpacing = '0'; }
}
function apPR(id, un, np) {
  db.collection('users').where('username','==',un).limit(1).get().then(function(s) {
    if (s.empty) { toast('غير موجود','e'); return; }
    return db.collection('users').doc(s.docs[0].id).update({ password: np });
  }).then(function() {
    return db.collection('passwordResetRequests').doc(id).update({ status:'approved', handledAt: firebase.firestore.FieldValue.serverTimestamp() });
  }).then(function() { toast('تم تغيير كلمة المرور','s'); ldPR(); chkP(); })
    .catch(function() { toast('خطأ','e'); });
}
function rjPR(id) {
  sCf('هل تريد رفض الطلب؟', function() {
    db.collection('passwordResetRequests').doc(id).update({ status:'rejected', handledAt: firebase.firestore.FieldValue.serverTimestamp() })
      .then(function() { toast('تم رفض الطلب','i'); ldPR(); chkP(); })
      .catch(function() { toast('خطأ','e'); });
  });
}

// ============================
// Users Management Tab
// ============================
function ldUr() {
  db.collection('users').orderBy('createdAt','asc').get().then(function(s) {
    var el = document.getElementById('adU'), docs = s.docs.filter(function(d) { return d.data().approved !== false; });
    if (!docs.length) { el.innerHTML = '<div class="text-center py-8" style="color:var(--mt)">لا يوجد مستخدمين</div>'; return; }
    el.innerHTML = docs.map(function(d) {
      var u = d.data(), me = cu && d.id === cu.id, pm = u.permissions || { canAdd:true, canDelete:true, canEdit:true }, ad = u.role === 'admin';
      var ph = '';
      if (!ad) {
        ph = '<div class="pm-box"><div class="pm-title"><i class="fa-solid fa-sliders"></i> الصلاحيات</div><div class="pm-row"><span>إضافة</span><label class="tgg"><input type="checkbox" '+(pm.canAdd?'checked':'')+' onchange="tgPm(\''+d.id+'\',\'canAdd\',this.checked)"><span class="tgg-sl"></span></label></div><div class="pm-row"><span>حذف</span><label class="tgg"><input type="checkbox" '+(pm.canDelete?'checked':'')+' onchange="tgPm(\''+d.id+'\',\'canDelete\',this.checked)"><span class="tgg-sl"></span></label></div><div class="pm-row"><span>تعديل</span><label class="tgg"><input type="checkbox" '+(pm.canEdit?'checked':'')+' onchange="tgPm(\''+d.id+'\',\'canEdit\',this.checked)"><span class="tgg-sl"></span></label></div></div>';
      }
      var pw = !me ? '<div class="flex items-center gap-2 mt-3"><input type="text" id="ri_'+d.id+'" class="fi fi-sm" style="flex:1" placeholder="كلمة مرور جديدة"><button type="button" onclick="cnRP(\''+d.id+'\')" class="bp bp-sm px-3"><i class="fa-solid fa-key ml-1"></i>تغيير</button></div>' : '';
      var dl = !me ? '<button type="button" onclick="dlU(\''+d.id+'\',\''+escA(u.username)+'\')" class="bd mt-3 w-full py-1.5 justify-center"><i class="fa-solid fa-trash-can ml-1"></i> حذف المستخدم</button>' : '';
      return '<div class="rounded-lg p-4 mb-3" style="background:var(--cs);border:1px solid var(--brd)"><div class="flex items-center gap-3 mb-1"><div class="w-10 h-10 rounded-full flex items-center justify-center font-bold text-base" style="background:'+(ad?'rgba(0,212,170,.1);color:#00d4aa':'rgba(100,116,139,.1);color:var(--mt)')+'">'+u.username.charAt(0).toUpperCase()+'</div><div class="flex-1"><div class="flex items-center gap-2"><span class="font-bold text-sm">'+esc(u.username)+'</span>'+(me?'<span class="text-[9px] font-bold px-2 py-0.5 rounded-full" style="background:rgba(0,212,170,.1);color:#00d4aa">أنت</span>':'')+'</div><div class="text-[10px]" style="color:var(--mt)">'+(ad?'مدير':'مستخدم')+' — '+fmtDt(u.createdAt)+'</div></div></div>'+ph+pw+dl+'</div>';
    }).join('');
  }).catch(function() { toast('خطأ','e'); });
}
function tgPm(id, p, v) {
  db.collection('users').doc(id).get().then(function(doc) {
    if (!doc.exists) return;
    var c = doc.data().permissions || { canAdd:true, canDelete:true, canEdit:true };
    c[p] = v;
    return db.collection('users').doc(id).update({ permissions: c });
  }).then(function() { toast('تم التحديث','s'); })
    .catch(function() { toast('خطأ','e'); });
}
function cnRP(id) {
  var np = document.getElementById('ri_'+id).value.trim();
  if (!np || np.length < 3) { toast('قصيرة','e'); return; }
  db.collection('users').doc(id).update({ password: np })
    .then(function() { toast('تم التغيير','s'); ldUr(); })
    .catch(function() { toast('خطأ','e'); });
}
function dlU(id, un) {
  sCf('حذف المستخدم "' + un.replace(/&#39;/g,"'") + '"؟', function() {
    db.collection('users').doc(id).delete()
      .then(function() { toast('تم الحذف','i'); ldUr(); chkP(); })
      .catch(function() { toast('خطأ','e'); });
  });
}

// ============================
// Settings Tab
// ============================
function tgAT(v) {
  db.collection('settings').doc('appSettings').set({ autoTransitionEnabled: v }, { merge: true })
    .then(function() { if (v) stAT(); toast(v ? 'تم تفعيل التحويل التلقائي' : 'تم تعطيله','s'); ldSt(); })
    .catch(function() { toast('خطأ','e'); });
}
function uATT(v) {
  if (!v || !v.match(/^\d{2}:\d{2}$/)) { toast('صيغة خاطئة','e'); return; }
  db.collection('settings').doc('appSettings').set({ autoTransitionTime: v }, { merge: true })
    .then(function() { return db.collection('settings').doc('appSettings').update({ lastAutoTransition: firebase.firestore.FieldValue.delete() }); })
    .then(function() { toast('تم تحديد الوقت: ' + v,'s'); ldSt(); setTimeout(chkAT, 800); })
    .catch(function() { toast('خطأ','e'); });
}
function tgHT(v) {
  hideTm = v;
  db.collection('settings').doc('appSettings').set({ hideTimeInLogs: v }, { merge: true })
    .catch(function() { toast('خطأ','e'); });
  rnT();
  var i = document.getElementById('htInfo');
  if (i) i.innerHTML = '<i class="fa-solid fa-shield-halved"></i> الحالة: <strong>' + (v ? 'مفعّل' : 'معطّل') + '</strong>';
  toast(v ? 'تم تفعيل إخفاء الوقت' : 'تم تعطيله','s');
}
function ldSt() {
  db.collection('settings').doc('appSettings').get().then(function(doc) {
    var s = doc.exists ? doc.data() : {};
    var ae = s.autoTransitionEnabled || false, at = s.autoTransitionTime || '18:00', la = s.lastAutoTransition ? fmtDtT(s.lastAutoTransition) : '—';
    var sv = document.getElementById('srvDisp') ? document.getElementById('srvDisp').textContent : '—';
    document.getElementById('adS').innerHTML =
      '<div class="rounded-lg p-4 mb-3" style="background:var(--cs);border:1px solid var(--brd)"><div class="flex items-center justify-between gap-4"><div class="flex-1"><div class="flex items-center gap-2 mb-1.5"><div class="w-8 h-8 rounded-lg flex items-center justify-center" style="background:rgba(0,212,170,.1)"><i class="fa-solid fa-eye-slash text-accent text-xs"></i></div><h3 class="font-bold text-sm">إخفاء الوقت من السجلات</h3></div><p class="text-[11px] mr-10" style="color:var(--dm)">يتم إخفاء الوقت فقط من تاريخ الاسترداد أمام المستخدمين</p><div class="mt-2 p-2 rounded-lg text-[10px] flex items-center gap-1.5" style="background:rgba(0,212,170,.05);border:1px solid rgba(0,212,170,.1);color:#00d4aa" id="htInfo"><i class="fa-solid fa-shield-halved"></i> الحالة: <strong>'+(hideTm?'مفعّل':'معطّل')+'</strong></div></div><label class="tgg"><input type="checkbox" '+(hideTm?'checked':'')+' onchange="tgHT(this.checked)"><span class="tgg-sl"></span></label></div></div>' +
      '<div class="rounded-lg p-4" style="background:var(--cs);border:1px solid var(--brd)"><div class="flex items-center gap-3 mb-3"><label class="tgg"><input type="checkbox" '+(ae?'checked':'')+' onchange="tgAT(this.checked)"><span class="tgg-sl"></span></label><span class="text-xs font-bold" style="color:var(--dm)">تحويل تلقائي "تم الرفع" ← "قيد المراجعة"</span></div><div class="flex items-center gap-3 flex-wrap"><div><label class="text-[10px] font-bold" style="color:var(--dm)">وقت التحويل اليومي</label><input type="time" class="fi fi-sm" style="width:120px" value="'+at+'" onchange="uATT(this.value)"></div><button type="button" onclick="dMT()" class="bp bp-warm bp-sm px-4 py-1.5"><i class="fa-solid fa-bolt ml-1"></i>تحويل يدوي الآن</button><div class="mr-auto text-[10px]" style="color:var(--mt)"><div><i class="fa-solid fa-server ml-1"></i>السيرفر: <span class="cv font-bold" style="color:var(--dm)">'+sv+'</span></div><div><i class="fa-solid fa-clock-rotate-left ml-1"></i>آخر تحويل: <span class="font-bold" style="color:var(--dm)">'+la+'</span></div></div></div></div>';
  }).catch(function() {});
}
