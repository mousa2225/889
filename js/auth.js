// ============================
// Auth Tab Switch
// ============================
function swT(t) {
  hAE(); hO('fOk'); hO('rOk');
  document.getElementById('tabL').classList.toggle('on', t === 'l');
  document.getElementById('tabR').classList.toggle('on', t === 'r');
  document.getElementById('lF').style.display = t === 'l' ? 'block' : 'none';
  document.getElementById('rF').style.display = t === 'r' ? 'block' : 'none';
  if (t === 'l') hideFg();
}
function showFg() {
  document.getElementById('lNorm').style.display = 'none';
  document.getElementById('lFg').style.display = 'block';
  hO('fOk'); hAE();
}
function hideFg() {
  document.getElementById('lNorm').style.display = 'block';
  document.getElementById('lFg').style.display = 'none';
}

// ============================
// Login
// ============================
function doL() {
  hAE();
  var u = document.getElementById('lU').value.trim(), p = document.getElementById('lP').value;
  if (!u) { document.getElementById('lU').classList.add('er'); sE('lEr','أدخل اسم المستخدم'); return; }
  if (!p) { document.getElementById('lP').classList.add('er'); sE('lEr','أدخل كلمة المرور'); return; }
  var btn = document.getElementById('lBt');
  btn.disabled = true;
  btn.innerHTML = '<span style="display:inline-block;width:14px;height:14px;border:2px solid transparent;border-top-color:#0b1120;border-radius:50%;animation:spin .6s linear infinite"></span>';
  db.collection('users').where('username','==',u).limit(1).get().then(function(s) {
    if (s.empty) { sE('lEr','اسم المستخدم غير موجود'); }
    else {
      var d = s.docs[0].data();
      if (d.password !== p) { sE('lEr','كلمة المرور خاطئة'); }
      else if (d.approved === false) { sE('lEr','حسابك بانتظار موافقة الإدارة'); }
      else {
        cu = { id: s.docs[0].id, username: d.username, role: d.role, permissions: d.permissions || { canAdd:true, canDelete:true, canEdit:true } };
        sessionStorage.setItem('cu', JSON.stringify(cu));
        shApp(); stL(); chkP();
        toast('مرحباً ' + cu.username, 's');
      }
    }
  }).catch(function(e) {
    if (e.code === 'permission-denied') sE('lEr','خطأ في الصلاحيات');
    else sE('lEr','خطأ في الاتصال');
  }).finally(function() {
    btn.disabled = false;
    btn.innerHTML = '<i class="fa-solid fa-right-to-bracket text-sm"></i> دخول';
  });
}

// ============================
// Register
// ============================
function doR() {
  hAE(); hO('rOk');
  var u = document.getElementById('rU').value.trim(), p = document.getElementById('rP').value, p2 = document.getElementById('rP2').value;
  if (!u) { document.getElementById('rU').classList.add('er'); sE('rEr','أدخل اسم المستخدم'); return; }
  if (!p || p.length < 3) { document.getElementById('rP').classList.add('er'); sE('rEr','كلمة المرور قصيرة'); return; }
  if (p !== p2) { document.getElementById('rP2').classList.add('er'); sE('rEr','كلمة المرور غير متطابقة'); return; }
  db.collection('users').where('username','==',u).limit(1).get().then(function(ex) {
    if (!ex.empty) { sE('rEr','اسم المستخدم موجود بالفعل'); return null; }
    return db.collection('users').get();
  }).then(function(all) {
    if (!all) return null;
    var f = all.empty;
    return db.collection('users').add({
      username: u, password: p, role: f ? 'admin' : 'user', approved: f,
      permissions: { canAdd:true, canDelete:true, canEdit:true },
      createdAt: firebase.firestore.FieldValue.serverTimestamp()
    }).then(function(r) { return { id: r.id, rl: f ? 'admin' : 'user', ap: f }; });
  }).then(function(res) {
    if (!res) return;
    if (res.ap) {
      cu = { id: res.id, username: u, role: res.rl, permissions: { canAdd:true, canDelete:true, canEdit:true } };
      sessionStorage.setItem('cu', JSON.stringify(cu));
      shApp(); stL(); chkP();
      toast('تم إنشاء حساب المدير','s');
    } else {
      document.getElementById('rU').value = ''; document.getElementById('rP').value = ''; document.getElementById('rP2').value = '';
      sO('rOk','تم تسجيل طلبك — بانتظار موافقة الإدارة');
      toast('تم تسجيل الطلب','s');
    }
  }).catch(function(e) {
    if (e.code === 'permission-denied') sE('rEr','خطأ في الصلاحيات');
    else sE('rEr','خطأ: ' + e.message);
  });
}

// ============================
// Forgot Password Request
// ============================
function doFg() {
  hAE(); hO('fOk');
  var u = document.getElementById('fU').value.trim(), p = document.getElementById('fP').value.trim();
  if (!u) { document.getElementById('fU').classList.add('er'); sE('fEr','أدخل اسم المستخدم'); return; }
  if (!p || p.length < 3) { document.getElementById('fP').classList.add('er'); sE('fEr','كلمة المرور قصيرة'); return; }
  db.collection('users').where('username','==',u).limit(1).get().then(function(s) {
    if (s.empty) { document.getElementById('fU').classList.add('er'); sE('fEr','اسم المستخدم غير موجود'); return Promise.reject('nf'); }
    return db.collection('passwordResetRequests').where('username','==',u).where('status','==','pending').limit(1).get();
  }).then(function(ex) {
    if (!ex.empty) { sE('fEr','يوجد طلب معلق بالفعل'); return; }
    return db.collection('passwordResetRequests').add({
      username: u, newPassword: p, status: 'pending',
      requestedAt: firebase.firestore.FieldValue.serverTimestamp()
    }).then(function() {
      sO('fOk','تم رفع طلب تغيير كلمة المرور');
      document.getElementById('fP').value = ''; document.getElementById('fU').value = '';
      toast('تم رفع الطلب','s');
    });
  }).catch(function(e) {
    if (e !== 'nf') {
      if (e.code === 'permission-denied') sE('fEr','خطأ في الصلاحيات');
      else sE('fEr','خطأ: ' + e.message);
    }
  });
}

// ============================
// Logout
// ============================
function doOut() {
  if (uns)   { uns();   uns   = null; }
  if (unsPm) { unsPm(); unsPm = null; }
  if (autoInt) { clearInterval(autoInt); autoInt = null; }
  if (clkInt)  { clearInterval(clkInt);  clkInt  = null; }
  cu = null; cls = []; hideTm = false;
  sessionStorage.removeItem('cu');
  document.getElementById('app').style.display = 'none';
  document.getElementById('authScr').style.display = 'flex';
  document.getElementById('lU').value = ''; document.getElementById('lP').value = '';
  hAE();
}
