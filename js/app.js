// ============================
// Show App After Login
// ============================
var curEditId = null;

function shApp() {
  document.getElementById('authScr').style.display = 'none';
  document.getElementById('app').style.display = 'block';
  document.getElementById('uNm').textContent = cu.username;
  var a = cu.role === 'admin';
  var b = document.getElementById('uBg');
  b.textContent = a ? 'مدير' : 'مستخدم';
  b.className = a ? 'rb rb-a' : 'rb rb-u';
  document.getElementById('adB').style.display = a ? 'inline-flex' : 'none';
  var exAEl = document.getElementById('exA'); if (exAEl) exAEl.style.display = a ? 'flex' : 'none';
  var exSEl = document.getElementById('exS'); if (exSEl) exSEl.style.display = a ? 'block' : 'none';
  var dtFlEl = document.getElementById('dtFl'); if (dtFlEl) dtFlEl.style.display = a ? 'block' : 'none';
  document.getElementById('srvClk').style.display = a ? 'flex' : 'none';
  document.getElementById('addBtn').style.display = gP().canAdd ? 'inline-flex' : 'none';
  var rpBtn = document.getElementById('rpBtn'); if (rpBtn) rpBtn.style.display = a ? 'inline-flex' : 'none';
  var srBtn = document.getElementById('startRfBtn'); if (srBtn) srBtn.style.display = a ? 'flex' : 'none';
  setTimeout(initExMenu, 100);

  db.collection('settings').doc('appSettings').onSnapshot(function(d) {
    hideTm = d.exists ? (d.data().hideTimeInLogs || false) : false;
    rnT();
  }, function() {});

  if (a) { syncST(); sSC(); stAT(); }

  if (!a && cu && cu.id) {
    if (unsPm) { unsPm(); unsPm = null; }
    unsPm = db.collection('users').doc(cu.id).onSnapshot(function(doc) {
      if (!doc.exists) return;
      var d = doc.data();
      if (d.approved === false) {
        setTimeout(function() { doOut(); toast('تم إيقاف حسابك من قبل المدير','e'); }, 200);
        return;
      }
      cu.permissions = d.permissions || { canAdd:true, canDelete:true, canEdit:true };
      sessionStorage.setItem('cu', JSON.stringify(cu));
      document.getElementById('addBtn').style.display = gP().canAdd ? 'inline-flex' : 'none';
      rnT();
    }, function() {});
  }
}

// ============================
// Auto-Transition
// ============================
function stAT() {
  if (autoInt) clearInterval(autoInt);
  autoInt = setInterval(chkAT, 30000);
  setTimeout(chkAT, 5000);
}
function chkAT() {
  if (!cu || cu.role !== 'admin' || !db) return;
  db.collection('settings').doc('appSettings').get().then(function(doc) {
    if (!doc.exists) return;
    var s = doc.data();
    if (!s.autoTransitionEnabled || !s.autoTransitionTime) return;
    var n = gSN(), pt = s.autoTransitionTime.split(':'), th = parseInt(pt[0]), tm = parseInt(pt[1]);
    if (isNaN(th) || isNaN(tm)) return;
    if (n.getHours() * 60 + n.getMinutes() < th * 60 + tm) return;
    if (s.lastAutoTransition && s.lastAutoTransition.seconds) {
      var ld = new Date(s.lastAutoTransition.seconds * 1000);
      if (ld.getFullYear()===n.getFullYear()&&ld.getMonth()===n.getMonth()&&ld.getDate()===n.getDate()) return;
    }
    return db.collection('cancellations').where('status','==','uploaded').get();
  }).then(function(snap) {
    if (!snap || snap.empty) return;
    var b = db.batch();
    snap.docs.forEach(function(d) { b.update(d.ref, { status:'pending', refunded:false }); });
    return b.commit().then(function() { return snap.size; });
  }).then(function(cnt) {
    if (cnt && cnt > 0) {
      toast('تحويل تلقائي: ' + cnt + ' طلب','s');
      return db.collection('settings').doc('appSettings').update({ lastAutoTransition: firebase.firestore.FieldValue.serverTimestamp() });
    }
  }).catch(function(e) { console.error('[AT]', e); });
}
function dMT() {
  if (!cu || cu.role !== 'admin') { toast('للمدير فقط','e'); return; }
  toast('جاري التحويل...','i');
  db.collection('cancellations').where('status','==','uploaded').get().then(function(snap) {
    if (snap.empty) { toast('لا توجد طلبات "تم الرفع"','w'); return; }
    var b = db.batch();
    snap.docs.forEach(function(d) { b.update(d.ref, { status:'pending', refunded:false }); });
    return b.commit().then(function() {
      return db.collection('settings').doc('appSettings').update({ lastAutoTransition: firebase.firestore.FieldValue.serverTimestamp() })
        .then(function() { return snap.size; });
    });
  }).then(function(cnt) {
    if (cnt) { toast('تم تحويل ' + cnt + ' طلب','s'); ldPd(); chkP(); }
  }).catch(function(e) { console.error(e); toast('خطأ','e'); });
}

// ============================
// Realtime Listener
// ============================
function stL() {
  uns = db.collection('cancellations').orderBy('createdAt','desc').onSnapshot(function(s) {
    cls = s.docs.map(function(d) { return { id: d.id, ...d.data() }; });
    rnT(); upSt();
  }, function(e) {
    toast(e.code === 'permission-denied' ? 'خطأ في الصلاحيات' : 'خطأ اتصال','e');
  });
}

// ============================
// Pending Badge
// ============================
function chkP() {
  if (!cu || cu.role !== 'admin') return;
  db.collection('users').where('approved','==',false).get().then(function(s) {
    var t = s.size;
    return db.collection('passwordResetRequests').where('status','==','pending').get().then(function(p) {
      t += p.size;
      t += cls.filter(function(c) { return getSt(c) === 'uploaded'; }).length;
      var bg = document.getElementById('pendBdg');
      if (bg) { bg.textContent = t; bg.style.display = t > 0 ? 'flex' : 'none'; }
    });
  }).catch(function() {});
}

// ============================
// Statistics
// ============================
function upSt() {
  var t = cls.length, tr = cls.reduce(function(s,c) { return s+(c.refundAmount||0); }, 0);
  var rf = cls.filter(function(c) { return getSt(c)==='refunded'; }), ra = rf.reduce(function(s,c) { return s+(c.refundAmount||0); }, 0);
  var pn = cls.filter(function(c) { return getSt(c)==='pending'; }), pa = pn.reduce(function(s,c) { return s+(c.refundAmount||0); }, 0);
  var up = cls.filter(function(c) { return getSt(c)==='uploaded'; });
  document.getElementById('sT').textContent  = en(t);
  // بادج ابدأ الاسترداد
  var srBdg = document.getElementById('pendRfBdg');
  if (srBdg) { var pnCnt=cls.filter(function(c){return getSt(c)==='pending';}).length; srBdg.textContent=pnCnt; srBdg.style.display=pnCnt>0?'flex':'none'; }
  document.getElementById('sU').textContent  = en(up.length);
  document.getElementById('sTR').textContent = en(tr.toFixed(2));
  document.getElementById('sR').textContent  = en(rf.length);
  document.getElementById('sRA').textContent = en(ra.toFixed(2)) + ' ريال';
  document.getElementById('sP').textContent  = en(pn.length);
  document.getElementById('sPA').textContent = en(pa.toFixed(2)) + ' ريال';
}

// ============================
// Dynamic previouslyRefunded
// ============================
function isDynPrevRef(c) {
  var ph = (c.phone||c.mobile||'').trim(), sub = (c.subscriptionDate||'').trim();
  if (!ph || !sub) return false;
  return cls.some(function(x) {
    return x.id !== c.id && (x.phone||x.mobile||'').trim() === ph &&
           (x.subscriptionDate||'').trim() === sub && getSt(x) === 'refunded';
  });
}

// ============================
// Table Render
// ============================
function sFt(f, btn) {
  cfl = f;
  document.querySelectorAll('#flS .ft').forEach(function(x) { x.classList.remove('on'); });
  btn.classList.add('on');
  rnT();
}
function rnT() {
  var tb = document.getElementById('tB'), sr = document.getElementById('sI').value.trim().toLowerCase();
  var a = cu && cu.role === 'admin', pm = gP(), ht = shTH();
  var fl = cls.filter(function(c) {
    if (!mDF(c)) return false;
    if (sr && !c.name.toLowerCase().includes(sr) && !(c.phone||'').toLowerCase().includes(sr) && !(c.mobile||'').toLowerCase().includes(sr)) return false;
    if (cfl === 'uploaded' && getSt(c) !== 'uploaded') return false;
    if (cfl === 'pending'  && getSt(c) !== 'pending')  return false;
    if (cfl === 'refunded' && getSt(c) !== 'refunded') return false;
    if (cfl === 'prev' && !isDynPrevRef(c)) return false;
    if (cfl === 'rejected' && getSt(c) !== 'rejected') return false;
    return true;
  });
  var dupPhones = {};
  cls.forEach(function(c) { var p=(c.phone||c.mobile||'').trim(); if(p) dupPhones[p]=(dupPhones[p]||0)+1; });

  var hd = cls.length > 0;
  document.getElementById('empS').style.display  = hd ? 'none'  : 'block';
  document.getElementById('tbS').style.display   = hd ? 'block' : 'none';
  document.getElementById('sS').style.display    = hd ? 'grid'  : 'none';
  document.getElementById('flS').style.display   = hd ? 'flex'  : 'none';
  document.getElementById('exB').style.display   = hd ? 'inline-flex' : 'none';
  if (!fl.length) {
    tb.innerHTML = '<tr><td colspan="20" class="text-center py-8" style="color:var(--mt)"><i class="fa-solid fa-magnifying-glass text-lg mb-1 block"></i>'+(cls.length?'لا توجد نتائج':'لا توجد بيانات')+'</td></tr>';
    return;
  }
  tb.innerHTML = fl.map(function(c,i) {
    return (a && c.id === curEditId) ? renderEditRow(c,i) : renderNormalRow(c,i,a,pm,ht,dupPhones);
  }).join('');
}

// ============================
// Normal Row
// ============================
function renderNormalRow(c, i, a, pm, ht, dupPhones) {
  var st = getSt(c), isDirect = c.refundType === 'direct';
  var rawPh = (c.phone||c.mobile||'').trim(), ph = esc(en(rawPh||'—'));
  var isDup = rawPh && dupPhones[rawPh] > 1, prevRef = isDynPrevRef(c);
  var cd = fmtDtS(c.createdAt), rd = '';
  if (st==='refunded'&&c.refundDate&&c.refundDate.seconds)
    rd = '<span class="num" style="font-size:11px;color:#22c55e;font-weight:700">'+(ht?fmtDt(c.refundDate):fmtDtT(c.refundDate))+'</span>';
  else rd = '<span style="color:var(--mt);font-size:12px">—</span>';

  // الخانات الثلاث: لا تعديل إذا الحالة 'refunded' حتى للمدير
  var ce = (a && st !== 'refunded' && st !== 'rejected') || (st==='uploaded' && pm.canEdit), pc,dc,cc,nc;
  if (!isDirect&&ce) {
    pc = '<input type="number" class="ism num" value="'+en(c.packagePrice||0)+'" step="0.01" dir="ltr" onchange="upF(\''+c.id+'\',\'packagePrice\',this.value)">';
    dc = '<input type="number" class="ism num" value="'+en(c.packageDays||0)+'" dir="ltr" onchange="upF(\''+c.id+'\',\'packageDays\',this.value)">';
    cc = '<input type="number" class="ism num" value="'+en(c.consumedDays||0)+'" dir="ltr" onchange="upF(\''+c.id+'\',\'consumedDays\',this.value)">';
  } else if (!isDirect) {
    pc = '<span class="ro num">'+en(c.packagePrice||0)+'</span>';
    dc = '<span class="ro num">'+en(c.packageDays||0)+'</span>';
    cc = '<span class="ro num">'+en(c.consumedDays||0)+'</span>';
  } else { pc=dc=cc='<span style="color:var(--mt);font-size:11px">—</span>'; }
  nc = ce ? '<input type="text" class="ism w" value="'+escA(c.notes||'')+'" placeholder="—" onchange="upF(\''+c.id+'\',\'notes\',this.value)">'
          : '<span class="ro" style="max-width:80px;display:inline-block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+escA(c.notes||'')+'">'+esc(c.notes||'—')+'</span>';

  var rfAmt = '<span class="rv num">'+en((c.refundAmount||0).toFixed(2))+'</span>';
  var aeBadge = c.adminEdited ? '<div style="font-size:8px;font-weight:700;color:#a78bfa;margin-top:2px;display:flex;align-items:center;gap:2px;white-space:nowrap"><i class="fa-solid fa-pen-to-square" style="font-size:7px"></i> تم التعديل من المالية</div>' : '';
  var rfCell = '<td><div style="display:flex;flex-direction:column;align-items:center">'+rfAmt+aeBadge+'</div></td>';

  var rc = '';
  if (a&&st==='refunded') rc='<td><input type="text" class="ism rf" value="'+escA(c.referenceNumber||'')+'" placeholder="أدخل الرقم" dir="ltr" onchange="upRf(\''+c.id+'\',this.value)"></td>';
  else if (c.referenceNumber) rc='<td><span class="num" style="color:#f59e0b;font-weight:700;font-size:12px">'+esc(c.referenceNumber)+'</span></td>';
  else rc='<td><span style="color:var(--mt);font-size:12px">—</span></td>';

  var ac = '';
  if (a) {
    var cid = c.id;
    var revertBtn = (st==='refunded' || st==='rejected')
      ? '<button type="button" data-id="'+cid+'" data-act="revert" style="width:28px;height:28px;border-radius:7px;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:11px;border:1px solid rgba(245,158,11,.25);background:rgba(245,158,11,.07);color:#f59e0b" title="تراجع عن الاسترداد"><i class="fa-solid fa-rotate-right"></i></button>'
      : '';
    var clearAEBtn = c.adminEdited
      ? '<button type="button" data-id="'+cid+'" data-act="clearae" style="width:28px;height:28px;border-radius:7px;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:10px;border:1px solid rgba(167,139,250,.2);background:rgba(167,139,250,.06);color:#a78bfa" title="إزالة علامة التعديل"><i class="fa-solid fa-eraser"></i></button>'
      : '';
    ac='<div style="display:flex;gap:4px;align-items:center">'
      +'<button type="button" data-id="'+cid+'" data-act="del" class="bd" title="حذف"><i class="fa-solid fa-trash-can"></i></button>'
      +'<button type="button" data-id="'+cid+'" data-act="edit" style="width:28px;height:28px;border-radius:7px;cursor:pointer;display:flex;align-items:center;justify-content:center;font-size:12px;border:1px solid rgba(96,165,250,.25);background:rgba(96,165,250,.07);color:#60a5fa" title="تعديل"><i class="fa-solid fa-pen-to-square"></i></button>'
      +revertBtn+clearAEBtn
      +'</div>';
  }
  else if (st==='uploaded'&&pm.canDelete) ac='<button type="button" class="bd" onclick="dlC(\''+c.id+'\')"><i class="fa-solid fa-trash-can"></i></button>';
  else ac='<span style="color:var(--mt);font-size:10px"><i class="fa-solid fa-lock"></i></span>';

  var stHtml;
  if (a && st === 'pending') {
    stHtml = '<td><div style="display:flex;gap:3px;align-items:center">'
      +'<button type="button" data-id="'+c.id+'" data-act="approve" class="sb sb-click sb-d" title="موافقة"><i class="'+STI['refunded']+' ml-0.5" style="font-size:9px"></i> '+STL['refunded']+'</button>'
      +'<button type="button" data-id="'+c.id+'" data-act="reject" style="padding:3px 7px;border-radius:14px;font-size:9.5px;font-weight:700;border:1px solid rgba(220,53,69,.3);background:rgba(220,53,69,.06);color:#dc3545;cursor:pointer;white-space:nowrap" title="رفض"><i class="fa-solid fa-circle-xmark" style="font-size:9px"></i> رفض</button>'
      +'</div></td>';
  } else if (st === 'rejected') {
    stHtml = '<td><button type="button" data-id="'+c.id+'" data-act="viewreject" class="sb sb-r" style="cursor:pointer" title="اعرف السبب"><i class="'+STI[st]+' ml-0.5" style="font-size:9px"></i> '+STL[st]+' <i class="fa-solid fa-eye" style="font-size:8px;opacity:.7"></i></button></td>';
  } else {
    stHtml = '<td><span class="sb '+STC[st]+'"><i class="'+STI[st]+' ml-0.5" style="font-size:9px"></i> '+STL[st]+'</span></td>';
  }

  var phCell = isDup
    ? '<td><span class="num" dir="ltr" style="font-size:12px;color:#dc3545;font-weight:800">'+ph+' <span style="display:inline-flex;align-items:center;gap:2px;padding:1px 5px;border-radius:7px;font-size:8px;font-weight:800;background:rgba(220,53,69,.15);color:#dc3545;border:1px solid rgba(220,53,69,.3)"><i class="fa-solid fa-copy" style="font-size:7px"></i> مكرر</span></span></td>'
    : '<td><span class="num" dir="ltr" style="font-size:12px">'+ph+'</span></td>';

  var rtCell = isDirect
    ? '<td><span style="display:inline-flex;align-items:center;gap:3px;padding:2px 7px;border-radius:8px;font-size:9px;font-weight:800;background:rgba(233,122,42,.1);color:#e97a2a;border:1px solid rgba(233,122,42,.2)"><i class="fa-solid fa-money-bill-wave" style="font-size:7px"></i> مباشر</span></td>'
    : '<td><span style="display:inline-flex;align-items:center;gap:3px;padding:2px 7px;border-radius:8px;font-size:9px;font-weight:800;background:rgba(0,212,170,.08);color:#00d4aa;border:1px solid rgba(0,212,170,.18)"><i class="fa-solid fa-rotate-left" style="font-size:7px"></i> اشتراك</span></td>';

  return '<tr class="rn '+(prevRef?'pr':'')+(st==='rejected'?' rj':'')+'">'+
    '<td class="font-bold text-[13px] num">'+(i+1)+'</td>'+
    '<td>'+esc(c.name)+(prevRef?' <span class="pb-t"><i class="fa-solid fa-triangle-exclamation"></i> سبق</span>':'')+'</td>'+
    phCell+rtCell+
    '<td class="text-[11px] whitespace-nowrap" style="color:var(--dm)">'+esc(c.subscriptionDate||'—')+'</td>'+
    '<td class="text-[11px] whitespace-nowrap" style="color:var(--dm)">'+esc(c.packageType||'—')+'</td>'+
    '<td class="num">'+pc+'</td><td class="num">'+dc+'</td><td class="num">'+cc+'</td>'+
    rfCell+
    '<td class="text-[11px] whitespace-nowrap" style="color:var(--dm)">'+esc(c.cancelDate||'—')+'</td>'+
    '<td class="text-[11px] whitespace-nowrap" style="color:var(--dm)">'+esc(c.cancellationPeriod||'—')+'</td>'+
    '<td class="text-[11px]" style="color:var(--dm);max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+escA(c.cancelReason||'')+'">'+esc(c.cancelReason||'—')+'</td>'+
    '<td class="text-[11px] whitespace-nowrap" style="color:var(--mt)">'+cd+'</td>'+
    '<td>'+rd+'</td>'+
    '<td class="text-[11px]"><span style="color:rgba(0,212,170,.6)">'+esc(c.addedByUsername||'—')+'</span></td>'+
    stHtml+rc+
    '<td>'+nc+'</td>'+
    '<td>'+ac+'</td></tr>';
}

// ============================
// Edit Row
// ============================
function renderEditRow(c, i) {
  var isDirect = c.refundType === 'direct', id = c.id;
  var fiW = 'class="fi-sm" style="border-color:rgba(96,165,250,.4);background:rgba(96,165,250,.06)"';
  var fiO = 'class="fi-sm" style="opacity:.4;pointer-events:none"';

  var rfInput = isDirect
    ? '<input id="eRA" type="number" step="0.01" dir="ltr" value="'+en(c.refundAmount||0)+'" class="fi-sm" style="font-weight:800;color:#e97a2a;border-color:rgba(233,122,42,.4);background:rgba(233,122,42,.06);width:80px">'
    : '<div><span id="eRfDisp" class="rv num" style="font-size:12px">'+en((c.refundAmount||0).toFixed(2))+'</span><div style="font-size:7px;color:var(--mt)">تلقائي</div></div>';

  return '<tr style="background:rgba(96,165,250,.05);outline:2px solid rgba(96,165,250,.35);outline-offset:-1px">'+
    '<td class="font-bold text-[13px] num">'+(i+1)+'</td>'+
    '<td><input id="eN" '+fiW+' value="'+escA(c.name||'')+'"></td>'+
    '<td><input id="ePh" '+fiW+' dir="ltr" value="'+escA(c.phone||c.mobile||'')+'"></td>'+
    '<td style="text-align:center">'+(isDirect?'<span style="font-size:8px;color:#e97a2a;font-weight:700">مباشر</span>':'<span style="font-size:8px;color:#00d4aa;font-weight:700">اشتراك</span>')+'</td>'+
    '<td><input id="eSub" '+fiW+' value="'+escA(c.subscriptionDate||'')+'"></td>'+
    '<td><input id="ePk" '+fiW+' value="'+escA(c.packageType||'')+'"></td>'+
    '<td><input id="ePr" type="number" step="0.01" dir="ltr" value="'+en(c.packagePrice||0)+'" '+(isDirect?fiO:fiW)+' oninput="upECalc()"></td>'+
    '<td><input id="eDy" type="number" dir="ltr" value="'+en(c.packageDays||0)+'" '+(isDirect?fiO:fiW)+' oninput="upECalc()"></td>'+
    '<td><input id="eCo" type="number" dir="ltr" value="'+en(c.consumedDays||0)+'" '+(isDirect?fiO:fiW)+' oninput="upECalc()"></td>'+
    '<td>'+rfInput+'</td>'+
    '<td><input id="eDt" '+fiW+' value="'+escA(c.cancelDate||'')+'"></td>'+
    '<td><input id="eDr" '+fiW+' value="'+escA(c.cancellationPeriod||'')+'"></td>'+
    '<td><input id="eRs" '+fiW+' value="'+escA(c.cancelReason||'')+'"></td>'+
    '<td class="text-[11px]" style="color:var(--mt)">'+fmtDtS(c.createdAt)+'</td>'+
    '<td></td><td></td><td></td><td></td>'+
    '<td><input id="eNo" '+fiW+' value="'+escA(c.notes||'')+'"></td>'+
    '<td><div style="display:flex;gap:3px">'+
      '<button type="button" onclick="saveEdit(\''+id+'\')" style="padding:3px 8px;border-radius:6px;font-size:10px;font-weight:800;border:none;cursor:pointer;background:rgba(0,212,170,.15);color:#00d4aa;font-family:Tajawal,sans-serif"><i class="fa-solid fa-check text-[9px]"></i> حفظ</button>'+
      '<button type="button" onclick="cancelEdit()" style="padding:3px 7px;border-radius:6px;font-size:10px;font-weight:700;border:none;cursor:pointer;background:var(--ib);color:var(--mt);font-family:Tajawal,sans-serif"><i class="fa-solid fa-xmark text-[9px]"></i></button>'+
    '</div></td></tr>';
}
function upECalc() {
  var pr=parseFloat(document.getElementById('ePr').value)||0;
  var dy=parseFloat(document.getElementById('eDy').value)||0;
  var co=parseFloat(document.getElementById('eCo').value)||0;
  var d=document.getElementById('eRfDisp'); if(d) d.textContent=cRf(pr,dy,co).toFixed(2);
}

// ============================
// Edit Start / Cancel / Save (OPTIMISTIC — لحظي)
// ============================
function startEdit(id) { curEditId = id; rnT(); }
function cancelEdit() { curEditId = null; rnT(); }

function saveEdit(id) {
  var c = cls.find(function(x){ return x.id===id; });
  if (!c) return;
  var isDirect = c.refundType === 'direct';
  var gv = function(eid){ var el=document.getElementById(eid); return el?el.value.trim():''; };

  var pr=parseFloat(gv('ePr'))||0, dy=parseFloat(gv('eDy'))||0, co=parseFloat(gv('eCo'))||0;
  var ra = isDirect ? (parseFloat(gv('eRA'))||0) : cRf(pr,dy,co);

  if (!gv('eN'))  { toast('الاسم مطلوب','e'); return; }
  if (!gv('ePh')) { toast('الرقم مطلوب','e'); return; }

  var u = {
    name: gv('eN'), phone: gv('ePh'), mobile: gv('ePh'),
    subscriptionDate: normDateStr(gv('eSub')), packageType: gv('ePk'),
    packagePrice: pr, packageDays: dy, consumedDays: co, refundAmount: ra,
    cancelDate: normDateStr(gv('eDt')), cancellationPeriod: gv('eDr'),
    cancelReason: gv('eRs'), notes: gv('eNo'), adminEdited: true
  };

  // ← تحديث محلي فوري (optimistic) — لا ننتظر Firestore
  Object.assign(c, u);
  curEditId = null;
  rnT();
  toast('تم الحفظ','s');

  // ثم نرسل للخادم في الخلفية
  db.collection('cancellations').doc(id).update(u)
    .catch(function(e) { toast('خطأ في المزامنة مع الخادم','e'); console.error(e); });
}

// ============================
// Field Updates
// ============================
function upF(id, f, v) {
  var u = {}, isAdmin = cu&&cu.role==='admin';
  if (f==='packagePrice'||f==='packageDays'||f==='consumedDays') {
    u[f] = parseFloat(v)||0;
    var c = cls.find(function(x){ return x.id===id; });
    if (c) {
      var newRa = cRf(
        f==='packagePrice'?parseFloat(v)||0:c.packagePrice||0,
        f==='packageDays'?parseFloat(v)||0:c.packageDays||0,
        f==='consumedDays'?parseFloat(v)||0:c.consumedDays||0
      );
      u.refundAmount = newRa;
      if (isAdmin) {
        u.adminEdited = true;
        // حفظ المبلغ الأصلي أول مرة فقط
        if (!c.originalRefundAmount && c.originalRefundAmount !== 0) {
          u.originalRefundAmount = c.refundAmount || 0;
        }
      }
    }
  } else { u[f]=v; }
  db.collection('cancellations').doc(id).update(u).catch(function(){ toast('خطأ في التحديث','e'); });
}
function upRf(id,v) {
  db.collection('cancellations').doc(id).update({referenceNumber:v})
    .then(function(){ toast('تم تحديث الرقم المرجعي','s'); })
    .catch(function(){ toast('خطأ','e'); });
}
function dlC(id) {
  sCf('هل أنت متأكد من حذف هذا العميل نهائياً؟', function() {
    db.collection('cancellations').doc(id).delete()
      .then(function(){ toast('تم الحذف بنجاح','s'); chkP(); })
      .catch(function(){ toast('خطأ في الحذف','e'); });
  });
}
function toPn(id) {
  if (!cu||cu.role!=='admin'){ toast('للمدير فقط','e'); return; }
  db.collection('cancellations').doc(id).update({status:'pending',refunded:false})
    .then(function(){ toast('تم التحويل إلى قيد المراجعة','s'); ldPd(); chkP(); })
    .catch(function(){ toast('خطأ','e'); });
}
function apR(id) {
  if (!cu||cu.role!=='admin'){ toast('للمدير فقط','e'); return; }
  db.collection('cancellations').doc(id).update({
    status:'refunded', refunded:true,
    refundDate: firebase.firestore.FieldValue.serverTimestamp()
  }).then(function(){ toast('تمت الموافقة على الاسترداد','s'); ldPd(); chkP(); rnT(); })
    .catch(function(){ toast('خطأ','e'); });
}

// ============================
// Add Client Modal
// ============================
function swRT(t) {
  var isSub = t==='sub';
  document.getElementById('subSec').style.display = isSub?'block':'none';
  document.getElementById('dirSec').style.display = isSub?'none':'block';
  var b1=document.getElementById('rtB1'), b2=document.getElementById('rtB2');
  if (isSub) { b1.style.cssText='background:rgba(0,212,170,.12);color:#00d4aa;border:1px solid rgba(0,212,170,.25)'; b2.style.cssText='background:transparent;color:var(--mt);border:1px solid transparent'; }
  else       { b1.style.cssText='background:transparent;color:var(--mt);border:1px solid transparent'; b2.style.cssText='background:rgba(233,122,42,.12);color:#e97a2a;border:1px solid rgba(233,122,42,.25)'; }
  document.getElementById('pW').classList.remove('on');
}
function openCM() {
  document.getElementById('cmM').classList.add('on');
  ['tmI','fN','fPh','fSub','fPk','fDt','fDr','fPr','fDy','fCo','fRs','fNo','fDN','fDPh','fDSub','fDPk','fDRs','fDRA'].forEach(function(id){ var el=document.getElementById(id); if(el) el.value=''; });
  document.getElementById('pRf').textContent='0.00';
  document.getElementById('pW').classList.remove('on');
  swRT('sub');
}
function clCM() { document.getElementById('cmM').classList.remove('on'); }
function chkPD()  { _chkDup(document.getElementById('fN').value.trim(), document.getElementById('fPh').value.trim()); }
function chkPD2() { _chkDup(document.getElementById('fDN').value.trim(), document.getElementById('fDPh').value.trim()); }
function _chkDup(n,p) {
  if (!n&&!p){ document.getElementById('pW').classList.remove('on'); return; }
  var d=cls.find(function(c){ return (n&&c.name===n)||(p&&(c.phone===p||c.mobile===p)); });
  if (d){ document.getElementById('pW').classList.add('on'); document.getElementById('pWT').textContent='تحذير: العميل موجود مسبقاً ('+esc(d.name)+')'; }
  else  { document.getElementById('pW').classList.remove('on'); }
}
function uPv() {
  var p=parseFloat(document.getElementById('fPr').value)||0, d=parseFloat(document.getElementById('fDy').value)||0, c=parseFloat(document.getElementById('fCo').value)||0;
  document.getElementById('pRf').textContent=cRf(p,d,c).toFixed(2);
}
function prsTm() {
  var t=document.getElementById('tmI').value; if(!t.trim()) return;
  var n='',ph='',pk='',dt='',rs='',co='',dr='',sub='',pamt='';
  t.split('\n').forEach(function(l){ l=l.trim(); if(!l) return; var idx=l.indexOf(':'); if(idx<0) return; var k=l.substring(0,idx).trim(),v=l.substring(idx+1).trim(),kn=k.replace(/[أإآ]/g,'ا').replace(/ة/g,'ه');
    if(kn.includes('اسم')&&kn.includes('عميل'))n=v;
    else if(kn.includes('رقم')&&!kn.includes('مرجع'))ph=v;
    else if(kn.includes('مستهلك'))co=v;
    else if(kn.includes('سبب')&&!kn.includes('تعذر')&&!kn.includes('تسليم'))rs=v;
    else if(kn.includes('اشتراك'))sub=v;
    else if(kn.includes('تاريخ'))dt=v;
    else if(kn.includes('مده'))dr=v;
    else if(kn.includes('باقه')||kn.includes('نوع'))pk=v;
    else if(kn.includes('مبلغ')&&!kn.includes('مسترد')&&!kn.includes('مسترجع'))pamt=v;
  });
  if(n) document.getElementById('fN').value=n;
  if(ph) document.getElementById('fPh').value=ph;
  if(pk) document.getElementById('fPk').value=pk;
  if(rs) document.getElementById('fRs').value=rs;
  if(dr) document.getElementById('fDr').value=dr;
  if(dt) document.getElementById('fDt').value=normDateStr(dt);
  if(sub) document.getElementById('fSub').value=normDateStr(sub);
  if(co){var cv=parseFloat(en(co));if(!isNaN(cv)&&cv>=0)document.getElementById('fCo').value=cv;}
  if(pk){var nums=pk.match(/[\d.]+/g);if(nums&&nums.length>=1)document.getElementById('fDy').value=nums[nums.length-1];}
  if(pamt){var av=parseFloat(en(pamt).replace(/[^\d.]/g,''));if(!isNaN(av)&&av>0)document.getElementById('fPr').value=av;}
  uPv(); chkPD();
}
function addC() {
  var isDirect=document.getElementById('dirSec').style.display!=='none', no=document.getElementById('fNo').value.trim();
  if (isDirect) {
    var n=document.getElementById('fDN').value.trim(),ph=document.getElementById('fDPh').value.trim();
    var dsub=document.getElementById('fDSub').value.trim(),dpk=document.getElementById('fDPk').value.trim();
    var drs=document.getElementById('fDRs').value.trim(),draStr=document.getElementById('fDRA').value.trim(),dra=parseFloat(draStr)||0;
    if(!n){toast('أدخل اسم العميل','e');return;} if(!ph){toast('أدخل الرقم','e');return;}
    if(!drs){toast('أدخل سبب الاسترداد','e');return;} if(!draStr||dra<=0){toast('أدخل المبلغ المسترد','e');return;}
    var ns=normDateStr(dsub), prevRef=cls.some(function(x){return (x.phone||x.mobile||'').trim()===ph&&(x.subscriptionDate||'').trim()===ns&&getSt(x)==='refunded';});
    db.collection('cancellations').add({name:n,phone:ph,mobile:ph,refundType:'direct',subscriptionDate:ns,packageType:dpk,cancelDate:'',cancellationPeriod:'',packagePrice:0,packageDays:0,consumedDays:0,refundAmount:dra,cancelReason:drs,notes:no,status:'uploaded',refunded:false,previouslyRefunded:prevRef,addedByUsername:cu.username,createdAt:firebase.firestore.FieldValue.serverTimestamp()})
      .then(function(){toast('تم إضافة العميل (مباشر)'+(prevRef?' ⚠ سبق الاسترداد':''),'s');clCM();}).catch(function(e){toast('خطأ في الإضافة','e');console.error(e);});
    return;
  }
  var n=document.getElementById('fN').value.trim(),ph=document.getElementById('fPh').value.trim();
  var pk=document.getElementById('fPk').value.trim(),dt=document.getElementById('fDt').value.trim();
  var dr=document.getElementById('fDr').value.trim(),sub=document.getElementById('fSub').value.trim();
  var rs=document.getElementById('fRs').value.trim(),prStr=document.getElementById('fPr').value.trim();
  var dyStr=document.getElementById('fDy').value.trim(),coStr=document.getElementById('fCo').value.trim();
  var pr=parseFloat(prStr)||0,dy=parseFloat(dyStr)||0,co=parseFloat(coStr)||0;
  if(!n){toast('أدخل اسم العميل','e');return;} if(!ph){toast('أدخل الرقم','e');return;}
  if(!pk){toast('أدخل نوع الباقة','e');return;} if(!prStr||pr<=0){toast('أدخل المبلغ المدفوع','e');return;}
  if(!dyStr||dy<=0){toast('أدخل أيام الباقة','e');return;} if(coStr===''){toast('أدخل الأيام المستهلكة (0 إن لم تُستهلك أيام)','e');return;}
  var ns=normDateStr(sub), prevRef=cls.some(function(x){return (x.phone||x.mobile||'').trim()===ph&&(x.subscriptionDate||'').trim()===ns&&getSt(x)==='refunded';});
  db.collection('cancellations').add({name:n,phone:ph,mobile:ph,refundType:'subscription',packageType:pk,cancelDate:normDateStr(dt),cancellationPeriod:dr,subscriptionDate:ns,packagePrice:pr,packageDays:dy,consumedDays:co,refundAmount:cRf(pr,dy,co),cancelReason:rs,notes:no,status:'uploaded',refunded:false,previouslyRefunded:prevRef,addedByUsername:cu.username,createdAt:firebase.firestore.FieldValue.serverTimestamp()})
    .then(function(){toast('تم إضافة العميل'+(prevRef?' ⚠ سبق الاسترداد':''),'s');clCM();}).catch(function(e){toast('خطأ في الإضافة','e');console.error(e);});
}

// ============================
// Table Click Delegation — يتعامل مع كل أزرار الجدول
// ============================
document.addEventListener('click', function(e) {
  var btn = e.target.closest('[data-act]');
  if (!btn) return;
  var id  = btn.getAttribute('data-id');
  var act = btn.getAttribute('data-act');
  if (!id || !act) return;
  if      (act === 'del')        dlC(id);
  else if (act === 'edit')       startEdit(id);
  else if (act === 'revert')     revertRefund(id);
  else if (act === 'clearae')    clearAdminEdited(id);
  else if (act === 'approve')    apR(id);
  else if (act === 'reject')     openReject(id);
  else if (act === 'viewreject') viewRejectReason(id);
});

// ============================
// Export Menu
// ============================
function togEm() { document.getElementById('exM').classList.toggle('on'); }
document.addEventListener('click', function(e) {
  if (!e.target.closest('#exM')&&!e.target.closest('#exB')) document.getElementById('exM').classList.remove('on');
});
document.addEventListener('keydown', function(e) {
  if (e.key==='Escape'){ clCM(); clAP(); cfNo(); cancelEdit(); }
});

// ============================
// ابدأ الاسترداد — تنقل بين قيد المراجعة
// ============================
var qvList  = [];
var qvIdx   = 0;
var qvId    = null;
var qvDirty = false; // true فقط إذا المستخدم غيّر قيمة فعلاً

function openQV(startId) {
  qvList = cls.filter(function(c){ return getSt(c) === 'pending'; })
    .sort(function(a,b){
      return (a.createdAt ? a.createdAt.seconds : 0) - (b.createdAt ? b.createdAt.seconds : 0);
    });
  if (!qvList.length) { toast('لا يوجد عملاء قيد المراجعة','w'); return; }
  if (startId) {
    var f = qvList.findIndex(function(c){ return c.id === startId; });
    qvIdx = f >= 0 ? f : 0;
  } else { qvIdx = 0; }
  document.getElementById('qvM').classList.add('on');
  qvRender();
}

function clQV() {
  document.getElementById('qvM').classList.remove('on');
  qvId = null; qvList = []; qvIdx = 0; qvDirty = false;
}

function qvNav(dir) {
  if (qvDirty) qvSave(true); // حفظ فقط لو في تغيير حقيقي
  qvIdx = (qvIdx + dir + qvList.length) % qvList.length;
  qvRender();
}

function qvMarkDirty() {
  qvDirty = true; // يُستدعى من oninput في الحقول
}

function qvRender() {
  var c = qvList[qvIdx];
  if (!c) return;
  qvId    = c.id;
  qvDirty = false; // ← إعادة ضبط عند كل عرض جديد

  var isDirect = c.refundType === 'direct';
  var prevRef  = isDynPrevRef(c);
  var rawPh    = (c.phone||c.mobile||'').trim();
  var dupPhones = {};
  cls.forEach(function(x){ var p=(x.phone||x.mobile||'').trim(); if(p) dupPhones[p]=(dupPhones[p]||0)+1; });
  var isDup = rawPh && dupPhones[rawPh] > 1;

  document.getElementById('qvCounter').textContent = (qvIdx+1) + ' / ' + qvList.length;
  document.getElementById('qvPrevBtn').style.opacity = qvList.length > 1 ? '1' : '0.3';
  document.getElementById('qvNextBtn').style.opacity = qvList.length > 1 ? '1' : '0.3';

  var typeEl = document.getElementById('qvType');
  typeEl.textContent = isDirect ? '💵 مباشر' : '🔄 اشتراك';
  typeEl.style.cssText = isDirect
    ? 'background:rgba(233,122,42,.1);color:#e97a2a;border:1px solid rgba(233,122,42,.2);padding:3px 10px;border-radius:8px;font-size:10px;font-weight:800'
    : 'background:rgba(0,212,170,.1);color:#00d4aa;border:1px solid rgba(0,212,170,.2);padding:3px 10px;border-radius:8px;font-size:10px;font-weight:800';

  document.getElementById('qvPrev').style.display = prevRef ? 'flex' : 'none';
  document.getElementById('qvDup').style.display  = isDup   ? 'flex' : 'none';
  // إشارة تعديل المالية — تظهر فقط إذا السجل فيه adminEdited مسبقاً
  document.getElementById('qvAEBadge').style.display = c.adminEdited ? 'flex' : 'none';

  document.getElementById('qvName').textContent  = c.name||'—';
  document.getElementById('qvPhone').textContent = en(rawPh||'—');
  document.getElementById('qvSub').textContent   = c.subscriptionDate||'—';
  document.getElementById('qvCdt').textContent   = c.cancelDate||'—';
  document.getElementById('qvDr').textContent    = c.cancellationPeriod||'—';
  document.getElementById('qvPk').textContent    = c.packageType||'—';
  document.getElementById('qvRs').textContent    = c.cancelReason||'—';
  document.getElementById('qvCrd').textContent   = fmtDtT(c.createdAt);
  document.getElementById('qvBy').textContent    = c.addedByUsername||'—';

  var subFields = document.getElementById('qvSubFields');
  var refBox    = document.getElementById('qvRefBox');
  if (isDirect) {
    subFields.style.display = 'none';
    document.getElementById('qvRA').readOnly = false;
    document.getElementById('qvRA').style.color = '#e97a2a';
    refBox.style.borderColor = 'rgba(233,122,42,.3)';
  } else {
    subFields.style.display = 'block';
    // ← تعيين القيم بدقة بدون en() لتجنب مشاكل المقارنة
    document.getElementById('qvPr').value = (c.packagePrice||0).toString();
    document.getElementById('qvDy').value = (c.packageDays||0).toString();
    document.getElementById('qvCo').value = (c.consumedDays||0).toString();
    document.getElementById('qvRA').readOnly = true;
    document.getElementById('qvRA').style.color = '#00d4aa';
    refBox.style.borderColor = 'rgba(0,212,170,.2)';
  }
  document.getElementById('qvRA').value = (c.refundAmount||0).toFixed(2);
}

function qvCalc() {
  var pr=parseFloat(document.getElementById('qvPr').value)||0;
  var dy=parseFloat(document.getElementById('qvDy').value)||0;
  var co=parseFloat(document.getElementById('qvCo').value)||0;
  document.getElementById('qvRA').value = cRf(pr,dy,co).toFixed(2);
  qvDirty = true; // المستخدم غيّر قيمة
}

// حفظ — لا يُنفَّذ إلا إذا qvDirty = true
function qvSave(silent) {
  if (!qvId || !qvDirty) {
    if (!silent) toast('لا توجد تعديلات','i');
    return;
  }
  var c = cls.find(function(x){ return x.id===qvId; });
  if (!c) return;
  var isDirect = c.refundType === 'direct';
  var u;
  if (isDirect) {
    u = { refundAmount: parseFloat(document.getElementById('qvRA').value)||0, adminEdited: true };
  } else {
    var pr=parseFloat(document.getElementById('qvPr').value)||0;
    var dy=parseFloat(document.getElementById('qvDy').value)||0;
    var co=parseFloat(document.getElementById('qvCo').value)||0;
    u = { packagePrice:pr, packageDays:dy, consumedDays:co, refundAmount:cRf(pr,dy,co), adminEdited:true };
  }
  Object.assign(c, u);
  qvDirty = false;
  db.collection('cancellations').doc(qvId).update(u)
    .catch(function(e){ toast('خطأ في المزامنة','e'); console.error(e); });
  if (!silent) toast('تم الحفظ','s');
  // تحديث شارة adminEdited في الموديل
  document.getElementById('qvAEBadge').style.display = 'flex';
}

// موافقة من الموديل
function qvApprove() {
  if (!qvId) return;
  if (qvDirty) qvSave(true);
  var id = qvId;
  db.collection('cancellations').doc(id).update({
    status:'refunded', refunded:true,
    refundDate: firebase.firestore.FieldValue.serverTimestamp()
  }).then(function(){
    toast('تم الاسترداد ✓','s');
    ldPd(); chkP();
    qvList = qvList.filter(function(c){ return c.id !== id; });
    if (!qvList.length) { clQV(); toast('تم الانتهاء من جميع الطلبات 🎉','s'); return; }
    if (qvIdx >= qvList.length) qvIdx = 0;
    qvId = null; qvDirty = false;
    qvRender();
  }).catch(function(){ toast('خطأ','e'); });
}

// ============================
// التراجع عن حالة "تم الاسترداد" → قيد المراجعة
// ============================
function revertRefund(id) {
  if (!cu || cu.role !== 'admin') { toast('للمدير فقط','e'); return; }
  var c = cls.find(function(x){ return x.id===id; });
  var isRej = c && getSt(c) === 'rejected';
  var msg = isRej
    ? 'هل تريد التراجع عن الرفض وإرجاع الطلب إلى "قيد المراجعة"؟'
    : 'هل تريد التراجع عن هذا الاسترداد وإرجاعه إلى "قيد المراجعة"؟';
  sCf(msg, function() {
    var u = { status: 'pending', refunded: false };
    if (!isRej) u.refundDate = firebase.firestore.FieldValue.delete();
    db.collection('cancellations').doc(id).update(u)
      .then(function(){
        toast('تم التراجع — العميل الآن قيد المراجعة','s');
        ldPd(); chkP();
      }).catch(function(e){ toast('خطأ','e'); console.error(e); });
  });
}

// ============================
// إزالة علامة "تم التعديل من المالية"
// ============================
function clearAdminEdited(id) {
  if (!cu || cu.role !== 'admin') return;
  var c = cls.find(function(x){ return x.id===id; });
  if (!c) return;
  c.adminEdited = false;
  db.collection('cancellations').doc(id).update({ adminEdited: false })
    .then(function(){ toast('تم إزالة علامة التعديل','s'); rnT(); })
    .catch(function(){ toast('خطأ','e'); });
}

// ============================
// رفض من شاشة التمرير السريع
// ============================
function qvReject() {
  if (!qvId) return;
  if (qvDirty) qvSave(true);
  openReject(qvId);
  // بعد الرفض: أزل من القائمة وانتقل للتالي
  var origConfirm = window._qvRejectConfirm;
  window._qvRejectConfirm = qvId;
}

// ============================
// رفض الطلب
// ============================
var rjTargetId = null;

function openReject(id) {
  rjTargetId = id;
  document.getElementById('rjView').style.display = 'none';
  document.getElementById('rjEdit').style.display = 'block';
  document.getElementById('rjReason').value = '';
  document.getElementById('rjM').classList.add('on');
  setTimeout(function(){ document.getElementById('rjReason').focus(); }, 100);
}

function viewRejectReason(id) {
  var c = cls.find(function(x){ return x.id===id; });
  if (!c) return;
  document.getElementById('rjViewTxt').textContent = c.rejectReason || 'لم يُذكر سبب';
  document.getElementById('rjView').style.display = 'block';
  document.getElementById('rjEdit').style.display = 'none';
  document.getElementById('rjM').classList.add('on');
}

function clRJ() {
  document.getElementById('rjM').classList.remove('on');
  rjTargetId = null;
}

function confirmReject() {
  if (!rjTargetId) return;
  var reason = document.getElementById('rjReason').value.trim();
  if (!reason) { toast('أدخل سبب الرفض','e'); return; }
  var id = rjTargetId;
  db.collection('cancellations').doc(id).update({
    status: 'rejected', refunded: false, rejectReason: reason
  }).then(function(){
    toast('تم رفض الطلب','s');
    clRJ(); ldPd(); chkP();
    // لو الرفض جاء من شاشة التمرير، انتقل للتالي
    if (window._qvRejectConfirm === id && qvList.length) {
      qvList = qvList.filter(function(c){ return c.id !== id; });
      if (!qvList.length) { clQV(); toast('تم الانتهاء من جميع الطلبات 🎉','s'); }
      else { if (qvIdx >= qvList.length) qvIdx = 0; qvId = null; qvDirty = false; qvRender(); }
      window._qvRejectConfirm = null;
    }
  }).catch(function(e){ toast('خطأ','e'); console.error(e); });
}
