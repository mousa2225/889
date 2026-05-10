// ============================
// Show App After Login
// ============================
function shApp() {
  document.getElementById('authScr').style.display = 'none';
  document.getElementById('app').style.display = 'block';
  document.getElementById('uNm').textContent = cu.username;
  var a = cu.role === 'admin';
  var b = document.getElementById('uBg');
  b.textContent = a ? 'مدير' : 'مستخدم';
  b.className = a ? 'rb rb-a' : 'rb rb-u';
  document.getElementById('adB').style.display = a ? 'inline-flex' : 'none';
  document.getElementById('exA').style.display = a ? 'flex' : 'none';
  var exSEl = document.getElementById('exS'); if (exSEl) exSEl.style.display = a ? 'block' : 'none';
  var dtFlEl = document.getElementById('dtFl'); if (dtFlEl) dtFlEl.style.display = a ? 'block' : 'none';
  document.getElementById('srvClk').style.display = a ? 'flex' : 'none';
  document.getElementById('addBtn').style.display = gP().canAdd ? 'inline-flex' : 'none';

  db.collection('settings').doc('appSettings').onSnapshot(function(d) {
    hideTm = d.exists ? (d.data().hideTimeInLogs || false) : false;
    rnT();
  }, function() {});

  if (a) { syncST(); sSC(); stAT(); }

  // Real-time Permissions Listener
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
      if (ld.getFullYear() === n.getFullYear() && ld.getMonth() === n.getMonth() && ld.getDate() === n.getDate()) return;
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
  document.getElementById('sU').textContent  = en(up.length);
  document.getElementById('sTR').textContent = en(tr.toFixed(2));
  document.getElementById('sR').textContent  = en(rf.length);
  document.getElementById('sRA').textContent = en(ra.toFixed(2)) + ' ريال';
  document.getElementById('sP').textContent  = en(pn.length);
  document.getElementById('sPA').textContent = en(pa.toFixed(2)) + ' ريال';
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
    if (cfl === 'prev' && !c.previouslyRefunded) return false;
    return true;
  });

  // خريطة الأرقام المكررة
  var dupPhones = {};
  cls.forEach(function(c) {
    var p = (c.phone || c.mobile || '').trim();
    if (p) dupPhones[p] = (dupPhones[p] || 0) + 1;
  });

  var hd = cls.length > 0;
  document.getElementById('empS').style.display  = hd ? 'none'  : 'block';
  document.getElementById('tbS').style.display   = hd ? 'block' : 'none';
  document.getElementById('sS').style.display    = hd ? 'grid'  : 'none';
  document.getElementById('flS').style.display   = hd ? 'flex'  : 'none';
  document.getElementById('exB').style.display   = hd ? 'inline-flex' : 'none';
  if (!fl.length) {
    tb.innerHTML = '<tr><td colspan="20" class="text-center py-8" style="color:var(--mt)"><i class="fa-solid fa-magnifying-glass text-lg mb-1 block"></i>' + (cls.length ? 'لا توجد نتائج' : 'لا توجد بيانات') + '</td></tr>';
    return;
  }
  tb.innerHTML = fl.map(function(c, i) {
    var st = getSt(c), isDirect = c.refundType === 'direct';
    var rawPh = (c.phone || c.mobile || '').trim();
    var ph = esc(en(rawPh || '—'));
    var isDup = rawPh && dupPhones[rawPh] > 1;
    var cd = fmtDtS(c.createdAt), rd = '';
    if (st === 'refunded' && c.refundDate && c.refundDate.seconds)
      rd = '<span class="num" style="font-size:11px;color:#22c55e;font-weight:700">' + (ht ? fmtDt(c.refundDate) : fmtDtT(c.refundDate)) + '</span>';
    else
      rd = '<span style="color:var(--mt);font-size:12px">—</span>';

    // خانات قابلة للتعديل
    var ce = a || (st === 'uploaded' && pm.canEdit);
    var pc, dc, cc, nc;
    if (!isDirect && ce) {
      pc = '<input type="number" class="ism num" value="'+en(c.packagePrice||0)+'" step="0.01" dir="ltr" onchange="upF(\''+c.id+'\',\'packagePrice\',this.value)">';
      dc = '<input type="number" class="ism num" value="'+en(c.packageDays||0)+'" dir="ltr" onchange="upF(\''+c.id+'\',\'packageDays\',this.value)">';
      cc = '<input type="number" class="ism num" value="'+en(c.consumedDays||0)+'" dir="ltr" onchange="upF(\''+c.id+'\',\'consumedDays\',this.value)">';
    } else if (!isDirect) {
      pc = '<span class="ro num">'+en(c.packagePrice||0)+'</span>';
      dc = '<span class="ro num">'+en(c.packageDays||0)+'</span>';
      cc = '<span class="ro num">'+en(c.consumedDays||0)+'</span>';
    } else {
      pc = '<span style="color:var(--mt);font-size:11px">—</span>';
      dc = '<span style="color:var(--mt);font-size:11px">—</span>';
      cc = '<span style="color:var(--mt);font-size:11px">—</span>';
    }
    if (ce) {
      nc = '<input type="text" class="ism w" value="'+escA(c.notes||'')+'" placeholder="—" onchange="upF(\''+c.id+'\',\'notes\',this.value)">';
    } else {
      nc = '<span class="ro" style="max-width:100px;display:inline-block;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+escA(c.notes||'')+'">'+esc(c.notes||'—')+'</span>';
    }

    // المبلغ المسترد + إشارة التعديل من المالية
    var rfAmt = '<span class="rv num">'+en((c.refundAmount||0).toFixed(2))+'</span>';
    var adminEditBadge = c.adminEdited
      ? '<div style="font-size:8px;font-weight:700;color:#a78bfa;margin-top:2px;display:flex;align-items:center;gap:2px;white-space:nowrap"><i class="fa-solid fa-pen-to-square" style="font-size:7px"></i> تم التعديل من المالية</div>'
      : '';
    var rfCell = '<td><div style="display:flex;flex-direction:column;align-items:center">' + rfAmt + adminEditBadge + '</div></td>';

    var rc = '';
    if (a && st === 'refunded') rc = '<td><input type="text" class="ism rf" value="'+escA(c.referenceNumber||'')+'" placeholder="أدخل الرقم" dir="ltr" onchange="upRf(\''+c.id+'\',this.value)"></td>';
    else if (c.referenceNumber) rc = '<td><span class="num" style="color:#f59e0b;font-weight:700;font-size:12px">'+esc(c.referenceNumber)+'</span></td>';
    else rc = '<td><span style="color:var(--mt);font-size:12px">—</span></td>';

    var ac = '';
    if (a) ac = '<button type="button" class="bd" onclick="dlC(\''+c.id+'\')"><i class="fa-solid fa-trash-can"></i></button>';
    else if (st === 'uploaded' && pm.canDelete) ac = '<button type="button" class="bd" onclick="dlC(\''+c.id+'\')"><i class="fa-solid fa-trash-can"></i></button>';
    else ac = '<span style="color:var(--mt);font-size:10px"><i class="fa-solid fa-lock"></i></span>';

    var stHtml;
    if (a && st === 'pending')
      stHtml = '<td><button type="button" class="sb sb-click sb-p" onclick="apR(\''+c.id+'\')" title="اضغط للموافقة"><i class="'+STI[st]+' ml-0.5" style="font-size:9px"></i> '+STL[st]+'</button></td>';
    else
      stHtml = '<td><span class="sb '+STC[st]+'"><i class="'+STI[st]+' ml-0.5" style="font-size:9px"></i> '+STL[st]+'</span></td>';

    // خلية الرقم
    var phCell = isDup
      ? '<td><span class="num" dir="ltr" style="font-size:12px;color:#dc3545;font-weight:800">' + ph +
          ' <span style="display:inline-flex;align-items:center;gap:2px;padding:1px 5px;border-radius:7px;font-size:8px;font-weight:800;background:rgba(220,53,69,.15);color:#dc3545;border:1px solid rgba(220,53,69,.3)"><i class="fa-solid fa-copy" style="font-size:7px"></i> مكرر</span></span></td>'
      : '<td><span class="num" dir="ltr" style="font-size:12px">'+ph+'</span></td>';

    // عمود نوع الاسترداد
    var rtCell = isDirect
      ? '<td><span style="display:inline-flex;align-items:center;gap:3px;padding:2px 7px;border-radius:8px;font-size:9px;font-weight:800;background:rgba(233,122,42,.1);color:#e97a2a;border:1px solid rgba(233,122,42,.2)"><i class="fa-solid fa-money-bill-wave" style="font-size:7px"></i> مباشر</span></td>'
      : '<td><span style="display:inline-flex;align-items:center;gap:3px;padding:2px 7px;border-radius:8px;font-size:9px;font-weight:800;background:rgba(0,212,170,.08);color:#00d4aa;border:1px solid rgba(0,212,170,.18)"><i class="fa-solid fa-rotate-left" style="font-size:7px"></i> اشتراك</span></td>';

    // عمود نوع الباقة — يظهر القيمة لكلا النوعين (أو — إن فارغة)
    var pkCell = '<td class="text-[11px] whitespace-nowrap" style="color:var(--dm)">'+esc(c.packageType||'—')+'</td>';

    return '<tr class="rn '+(c.previouslyRefunded?'pr':'')+'">' +
      '<td class="font-bold text-[13px] num">'+(i+1)+'</td>' +
      '<td>'+esc(c.name)+(c.previouslyRefunded?' <span class="pb-t"><i class="fa-solid fa-triangle-exclamation"></i> سبق</span>':'')+'</td>' +
      phCell +
      rtCell +
      '<td class="text-[11px] whitespace-nowrap" style="color:var(--dm)">'+esc(c.subscriptionDate||'—')+'</td>' +
      pkCell +
      '<td class="num">'+pc+'</td><td class="num">'+dc+'</td><td class="num">'+cc+'</td>' +
      rfCell +
      '<td class="text-[11px] whitespace-nowrap" style="color:var(--dm)">'+esc(c.cancelDate||'—')+'</td>' +
      '<td class="text-[11px] whitespace-nowrap" style="color:var(--dm)">'+esc(c.cancellationPeriod||'—')+'</td>' +
      '<td class="text-[11px]" style="color:var(--dm);max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="'+escA(c.cancelReason||'')+'">'+esc(c.cancelReason||'—')+'</td>' +
      '<td class="text-[11px] whitespace-nowrap" style="color:var(--mt)">'+cd+'</td>' +
      '<td>'+rd+'</td>' +
      '<td class="text-[11px]"><span style="color:rgba(0,212,170,.6)">'+esc(c.addedByUsername||'—')+'</span></td>' +
      stHtml + rc +
      '<td>'+nc+'</td>' +
      '<td>'+ac+'</td></tr>';
  }).join('');
}

// ============================
// Field Updates
// لو المدير عدّل الخانات الثلاث → adminEdited: true
// ============================
function upF(id, f, v) {
  var u = {}, isAdmin = cu && cu.role === 'admin';
  if (f === 'packagePrice' || f === 'packageDays' || f === 'consumedDays') {
    u[f] = parseFloat(v) || 0;
    var c = cls.find(function(x) { return x.id === id; });
    if (c) u.refundAmount = cRf(
      f==='packagePrice'  ? parseFloat(v)||0 : c.packagePrice||0,
      f==='packageDays'   ? parseFloat(v)||0 : c.packageDays||0,
      f==='consumedDays'  ? parseFloat(v)||0 : c.consumedDays||0
    );
    if (isAdmin) u.adminEdited = true;
  } else { u[f] = v; }
  db.collection('cancellations').doc(id).update(u).catch(function() { toast('خطأ في التحديث','e'); });
}
function upRf(id, v) {
  db.collection('cancellations').doc(id).update({ referenceNumber: v })
    .then(function() { toast('تم تحديث الرقم المرجعي','s'); })
    .catch(function() { toast('خطأ','e'); });
}
function dlC(id) {
  sCf('هل أنت متأكد من حذف هذا العميل نهائياً؟', function() {
    db.collection('cancellations').doc(id).delete()
      .then(function() { toast('تم الحذف بنجاح','s'); chkP(); })
      .catch(function() { toast('خطأ في الحذف','e'); });
  });
}
function toPn(id) {
  if (!cu || cu.role !== 'admin') { toast('للمدير فقط','e'); return; }
  db.collection('cancellations').doc(id).update({ status:'pending', refunded:false })
    .then(function() { toast('تم التحويل إلى قيد المراجعة','s'); ldPd(); chkP(); })
    .catch(function() { toast('خطأ','e'); });
}
function apR(id) {
  if (!cu || cu.role !== 'admin') { toast('للمدير فقط','e'); return; }
  var c = cls.find(function(x) { return x.id === id; });
  db.collection('cancellations').doc(id).update({
    status: 'refunded', refunded: true,
    refundDate: firebase.firestore.FieldValue.serverTimestamp(),
    previouslyRefunded: c ? c.previouslyRefunded || false : false
  }).then(function() { toast('تمت الموافقة على الاسترداد','s'); ldPd(); chkP(); rnT(); })
    .catch(function() { toast('خطأ','e'); });
}

// ============================
// Add Client Modal — Toggle Type
// ============================
function swRT(t) {
  var isSub = t === 'sub';
  document.getElementById('subSec').style.display = isSub ? 'block' : 'none';
  document.getElementById('dirSec').style.display = isSub ? 'none'  : 'block';
  var b1 = document.getElementById('rtB1'), b2 = document.getElementById('rtB2');
  if (isSub) {
    b1.style.cssText = 'background:rgba(0,212,170,.12);color:#00d4aa;border:1px solid rgba(0,212,170,.25)';
    b2.style.cssText = 'background:transparent;color:var(--mt);border:1px solid transparent';
  } else {
    b1.style.cssText = 'background:transparent;color:var(--mt);border:1px solid transparent';
    b2.style.cssText = 'background:rgba(233,122,42,.12);color:#e97a2a;border:1px solid rgba(233,122,42,.25)';
  }
  document.getElementById('pW').classList.remove('on');
}
function openCM() {
  document.getElementById('cmM').classList.add('on');
  ['tmI','fN','fPh','fSub','fPk','fDt','fDr','fPr','fDy','fCo','fRs','fNo',
   'fDN','fDPh','fDSub','fDPk','fDRs','fDRA'].forEach(function(id) {
    var el = document.getElementById(id); if (el) el.value = '';
  });
  document.getElementById('pRf').textContent = '0.00';
  document.getElementById('pW').classList.remove('on');
  swRT('sub');
}
function clCM() { document.getElementById('cmM').classList.remove('on'); }

function chkPD() {
  var n = document.getElementById('fN').value.trim(), p = document.getElementById('fPh').value.trim();
  _chkDup(n, p);
}
function chkPD2() {
  var n = document.getElementById('fDN').value.trim(), p = document.getElementById('fDPh').value.trim();
  _chkDup(n, p);
}
function _chkDup(n, p) {
  if (!n && !p) { document.getElementById('pW').classList.remove('on'); return; }
  var d = cls.find(function(c) {
    if (n && c.name === n) return true;
    if (p && (c.phone === p || c.mobile === p)) return true;
    return false;
  });
  if (d) {
    document.getElementById('pW').classList.add('on');
    document.getElementById('pWT').textContent = 'تحذير: العميل موجود مسبقاً (' + esc(d.name) + ')';
  } else { document.getElementById('pW').classList.remove('on'); }
}
function uPv() {
  var p = parseFloat(document.getElementById('fPr').value)||0;
  var d = parseFloat(document.getElementById('fDy').value)||0;
  var c = parseFloat(document.getElementById('fCo').value)||0;
  document.getElementById('pRf').textContent = cRf(p,d,c).toFixed(2);
}

// ============================
// Template Parser
// ============================
function prsTm() {
  var t = document.getElementById('tmI').value;
  if (!t.trim()) return;
  var n='', ph='', pk='', dt='', rs='', co='', dr='', sub='', pamt='';
  t.split('\n').forEach(function(l) {
    l = l.trim(); if (!l) return;
    var idx = l.indexOf(':'); if (idx < 0) return;
    var k = l.substring(0, idx).trim(), v = l.substring(idx + 1).trim();
    var kn = k.replace(/[أإآ]/g,'ا').replace(/ة/g,'ه');
    if      (kn.includes('اسم') && kn.includes('عميل'))                               { n   = v; }
    else if (kn.includes('رقم') && !kn.includes('مرجع'))                              { ph  = v; }
    else if (kn.includes('مستهلك'))                                                    { co  = v; }
    else if (kn.includes('سبب') && !kn.includes('تعذر') && !kn.includes('تسليم'))    { rs  = v; }
    else if (kn.includes('اشتراك'))                                                    { sub = v; }
    else if (kn.includes('تاريخ'))                                                     { dt  = v; }
    else if (kn.includes('مده'))                                                       { dr  = v; }
    else if (kn.includes('باقه') || kn.includes('نوع'))                               { pk  = v; }
    else if (kn.includes('مبلغ') && !kn.includes('مسترد') && !kn.includes('مسترجع')) { pamt = v; }
  });
  if (n)   document.getElementById('fN').value   = n;
  if (ph)  document.getElementById('fPh').value  = ph;
  if (pk)  document.getElementById('fPk').value  = pk;
  if (rs)  document.getElementById('fRs').value  = rs;
  if (dr)  document.getElementById('fDr').value  = dr;
  if (dt)  document.getElementById('fDt').value  = normDateStr(dt);
  if (sub) document.getElementById('fSub').value = normDateStr(sub);
  if (co)  { var cv = parseFloat(en(co)); if (!isNaN(cv) && cv >= 0) document.getElementById('fCo').value = cv; }
  if (pk)  { var nums = pk.match(/[\d.]+/g); if (nums && nums.length >= 1) document.getElementById('fDy').value = nums[nums.length-1]; }
  if (pamt){ var av = parseFloat(en(pamt).replace(/[^\d.]/g,'')); if (!isNaN(av) && av > 0) document.getElementById('fPr').value = av; }
  uPv(); chkPD();
}

// ============================
// Add Client — يدعم النوعين
// ============================
function addC() {
  var isDirect = document.getElementById('dirSec').style.display !== 'none';
  var no = document.getElementById('fNo').value.trim();

  // ── استرداد مباشر ──
  if (isDirect) {
    var n      = document.getElementById('fDN').value.trim();
    var ph     = document.getElementById('fDPh').value.trim();
    var dsub   = document.getElementById('fDSub').value.trim();
    var dpk    = document.getElementById('fDPk').value.trim();
    var drs    = document.getElementById('fDRs').value.trim();
    var draStr = document.getElementById('fDRA').value.trim();
    var dra    = parseFloat(draStr) || 0;

    if (!n)                  { toast('أدخل اسم العميل','e'); return; }
    if (!ph)                 { toast('أدخل الرقم','e'); return; }
    if (!drs)                { toast('أدخل سبب الاسترداد','e'); return; }
    if (!draStr || dra <= 0) { toast('أدخل المبلغ المسترد','e'); return; }

    var prevRef = cls.some(function(x) {
      var xph = (x.phone || x.mobile || '').trim();
      return xph && xph === ph && getSt(x) === 'refunded';
    });

    db.collection('cancellations').add({
      name: n, phone: ph, mobile: ph,
      refundType: 'direct',
      subscriptionDate: normDateStr(dsub),
      packageType: dpk,   // اختياري — قد يكون فارغ
      cancelDate: '', cancellationPeriod: '',
      packagePrice: 0, packageDays: 0, consumedDays: 0,
      refundAmount: dra, cancelReason: drs, notes: no,
      status: 'uploaded', refunded: false,
      previouslyRefunded: prevRef,
      addedByUsername: cu.username,
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    }).then(function() {
      toast('تم إضافة العميل (مباشر)' + (prevRef ? ' ⚠ سبق الاسترداد' : ''), 's');
      clCM();
    }).catch(function(e) { toast('خطأ في الإضافة','e'); console.error(e); });
    return;
  }

  // ── استرداد اشتراك العميل ──
  var n  = document.getElementById('fN').value.trim();
  var ph = document.getElementById('fPh').value.trim();
  var pk    = document.getElementById('fPk').value.trim();
  var dt    = document.getElementById('fDt').value.trim();
  var dr    = document.getElementById('fDr').value.trim();
  var sub   = document.getElementById('fSub').value.trim();
  var rs    = document.getElementById('fRs').value.trim();
  var prStr = document.getElementById('fPr').value.trim();
  var dyStr = document.getElementById('fDy').value.trim();
  var coStr = document.getElementById('fCo').value.trim();
  var pr = parseFloat(prStr)||0, dy = parseFloat(dyStr)||0, co = parseFloat(coStr)||0;

  if (!n)                { toast('أدخل اسم العميل','e'); return; }
  if (!ph)               { toast('أدخل الرقم','e'); return; }
  if (!pk)               { toast('أدخل نوع الباقة','e'); return; }
  if (!prStr || pr <= 0) { toast('أدخل المبلغ المدفوع','e'); return; }
  if (!dyStr || dy <= 0) { toast('أدخل أيام الباقة','e'); return; }
  if (coStr === '')       { toast('أدخل الأيام المستهلكة (0 إن لم تُستهلك أيام)','e'); return; }

  var prevRef = cls.some(function(x) {
    var xph = (x.phone || x.mobile || '').trim();
    return xph && xph === ph && getSt(x) === 'refunded';
  });

  db.collection('cancellations').add({
    name: n, phone: ph, mobile: ph,
    refundType: 'subscription',
    packageType: pk, cancelDate: normDateStr(dt),
    cancellationPeriod: dr, subscriptionDate: normDateStr(sub),
    packagePrice: pr, packageDays: dy, consumedDays: co,
    refundAmount: cRf(pr, dy, co),
    cancelReason: rs, notes: no,
    status: 'uploaded', refunded: false,
    previouslyRefunded: prevRef,
    addedByUsername: cu.username,
    createdAt: firebase.firestore.FieldValue.serverTimestamp()
  }).then(function() {
    toast('تم إضافة العميل' + (prevRef ? ' ⚠ سبق الاسترداد' : ''), 's');
    clCM();
  }).catch(function(e) { toast('خطأ في الإضافة','e'); console.error(e); });
}

// ============================
// Export Menu
// ============================
function togEm() { document.getElementById('exM').classList.toggle('on'); }
document.addEventListener('click', function(e) {
  if (!e.target.closest('#exM') && !e.target.closest('#exB'))
    document.getElementById('exM').classList.remove('on');
});
document.addEventListener('keydown', function(e) {
  if (e.key === 'Escape') { clCM(); clAP(); cfNo(); }
});
