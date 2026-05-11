// ============================
// إظهار / إخفاء أقسام التصدير حسب الصلاحية
// ============================
function initExMenu() {
  var isAdmin = cu && cu.role === 'admin';
  var adminSec = document.getElementById('exAdminSec');
  if (adminSec) adminSec.style.display = isAdmin ? 'block' : 'none';
  var userSec  = document.getElementById('exUserSec');
  if (userSec)  userSec.style.display  = 'block';
}

// ============================
// مساعد: تنسيق تاريخ لاسم الملف
// ============================
function fmtFN(dt) {
  if (!dt) return '';
  var d = new Date(dt);
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}
function todayStr() { return fmtFN(new Date()); }

// ============================
// مساعد: بناء بيانات الصف لكل عميل
// ============================
function buildRow(c, i) {
  return {
    '#': i + 1,
    'اسم العميل': c.name || '',
    'الرقم': c.phone || c.mobile || '',
    'نوع الاسترداد': c.refundType === 'direct' ? 'مباشر' : 'اشتراك',
    'تاريخ الاشتراك': c.subscriptionDate || '',
    'نوع الباقة': c.packageType || '',
    'المبلغ المدفوع': c.packagePrice || 0,
    'أيام الباقة': c.packageDays || 0,
    'أيام مستهلكة': c.consumedDays || 0,
    'المبلغ المسترد': c.refundAmount || 0,
    'تاريخ الإلغاء': c.cancelDate || '',
    'مدة الإلغاء': c.cancellationPeriod || '',
    'سبب الإلغاء/الاسترداد': c.cancelReason || '',
    'تاريخ الإضافة': c.createdAt ? fmtDtT(c.createdAt) : '',
    'تاريخ الاسترداد': c.refundDate ? fmtDtT(c.refundDate) : '',
    'أضاف بواسطة': c.addedByUsername || '',
    'الحالة': STL[getSt(c)] || '',
    'الرقم المرجعي': c.referenceNumber || '',
    'ملاحظات': (isDynPrevRef(c) ? 'سبق الاسترداد — ' : '') + (c.notes || '')
  };
}

// ============================
// محرك التصدير الاحترافي
// ============================
function doExport(data, title, subtitle, filename) {
  if (!data || !data.length) { toast('لا توجد بيانات للتصدير', 'w'); return; }

  var hd = ['#','اسم العميل','الرقم','نوع الاسترداد','تاريخ الاشتراك','نوع الباقة',
            'المبلغ المدفوع','أيام الباقة','أيام مستهلكة','المبلغ المسترد',
            'تاريخ الإلغاء','مدة الإلغاء','سبب الإلغاء/الاسترداد',
            'تاريخ الإضافة','تاريخ الاسترداد','أضاف بواسطة','الحالة','الرقم المرجعي','ملاحظات','سبب الرفض'];
  var cw = [5,22,16,14,16,22,14,11,11,16,14,12,22,18,18,16,14,14,18,18];

  try {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'V-SHAPE';
    wb.created = new Date();
    var ws = wb.addWorksheet('البيانات', { views:[{ rightToLeft:true }] });
    ws.columns = hd.map(function(h, i) { return { header:'', key:h, width: cw[i]||15 }; });
    ws.spliceRows(1,1);

    // ── صف العنوان الرئيسي ──
    var titleRow = ws.addRow(Array(hd.length).fill(''));
    titleRow.height = 40;
    ws.mergeCells(1, 1, 1, hd.length);
    titleRow.getCell(1).value = title;
    titleRow.getCell(1).font  = { name:'Tajawal', size:16, bold:true, color:{ argb:'FFFFFFFF' } };
    titleRow.getCell(1).fill  = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF0D9488' } };
    titleRow.getCell(1).alignment = { horizontal:'center', vertical:'middle' };

    // ── صف العنوان الفرعي ──
    if (subtitle) {
      var subRow = ws.addRow(Array(hd.length).fill(''));
      subRow.height = 20;
      ws.mergeCells(2, 1, 2, hd.length);
      subRow.getCell(1).value = subtitle;
      subRow.getCell(1).font  = { name:'Tajawal', size:10, color:{ argb:'FF64748B' } };
      subRow.getCell(1).fill  = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFF0FDF4' } };
      subRow.getCell(1).alignment = { horizontal:'center', vertical:'middle' };
    }

    // ── صف الإحصائيات ──
    var total  = data.length;
    var rfCnt  = data.filter(function(c){ return getSt(c)==='refunded'; }).length;
    var rfAmt  = data.filter(function(c){ return getSt(c)==='refunded'; }).reduce(function(s,c){ return s+(c.refundAmount||0); },0);
    var totAmt = data.reduce(function(s,c){ return s+(c.refundAmount||0); },0);

    var statRow = ws.addRow(Array(hd.length).fill(''));
    statRow.height = 22;
    ws.mergeCells(subtitle ? 3:2, 1, subtitle ? 3:2, 4);
    ws.mergeCells(subtitle ? 3:2, 5, subtitle ? 3:2, 8);
    ws.mergeCells(subtitle ? 3:2, 9, subtitle ? 3:2, 12);
    ws.mergeCells(subtitle ? 3:2, 13, subtitle ? 3:2, hd.length);
    var sBase = subtitle ? 3 : 2;
    var statStyle = function(cell, val, lbl, color) {
      cell.value = lbl + ': ' + val;
      cell.font  = { name:'Tajawal', size:10, bold:true, color:{ argb: color } };
      cell.fill  = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFF8FAFC' } };
      cell.alignment = { horizontal:'center', vertical:'middle' };
    };
    statStyle(ws.getRow(sBase).getCell(1),  total + ' سجل',         'الإجمالي', 'FF0F766E');
    statStyle(ws.getRow(sBase).getCell(5),  totAmt.toFixed(2) + ' ريال', 'إجمالي المبالغ', 'FFB45309');
    statStyle(ws.getRow(sBase).getCell(9),  rfCnt + ' سجل',         'تم الاسترداد', 'FF16A34A');
    statStyle(ws.getRow(sBase).getCell(13), rfAmt.toFixed(2) + ' ريال',  'المستردة فعلياً', 'FF16A34A');

    // ── صف فراغ ──
    ws.addRow([]).height = 6;

    // ── صف الرؤوس ──
    var hdRow = ws.addRow(hd);
    hdRow.height = 28;
    hdRow.eachCell(function(cell) {
      cell.font      = { name:'Tajawal', size:11, bold:true, color:{ argb:'FFFFFFFF' } };
      cell.fill      = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF0F766E' } };
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border    = { bottom:{ style:'medium', color:{ argb:'FF0D9488' } } };
    });

    // ── صفوف البيانات ──
    data.forEach(function(c, idx) {
      var rowData = buildRow(c, idx);
      var row = ws.addRow(Object.values(rowData));
      row.height = 22;
      var isPrev = isDynPrevRef(c);
      var isSub  = c.refundType !== 'direct';
      row.eachCell(function(cell, col) {
        cell.font      = { name:'Tajawal', size:11 };
        cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
        cell.border    = {
          top:   { style:'thin', color:{ argb:'FFE2E8F0' } },
          bottom:{ style:'thin', color:{ argb:'FFE2E8F0' } },
          left:  { style:'thin', color:{ argb:'FFE2E8F0' } },
          right: { style:'thin', color:{ argb:'FFE2E8F0' } }
        };
        // تلوين الصفوف
        var isRej = rowData['الحالة'] === 'مرفوض';
        if (isRej) {
          cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF5F5' } };
          cell.font = { name:'Tajawal', size:11, color:{ argb:'FF991B1B' } };
        } else if (isPrev) {
          cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFFF1F2' } };
          cell.font = { name:'Tajawal', size:11, color:{ argb:'FF9F1239' } };
        } else if (idx % 2 === 1) {
          cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFF8FAFC' } };
        }
        // عمود المبلغ المسترد (col 10)
        if (col === 10) {
          cell.font = { name:'Tajawal', size:11, bold:true, color:{ argb: isPrev ? 'FF9F1239' : 'FF0F766E' } };
          cell.numFmt = '#,##0.00';
          if (!isPrev) cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: idx%2===0 ? 'FFF0FDF4' : 'FFE6F9F5' } };
        }
        // عمود المبلغ المدفوع (col 7)
        if (col === 7) { cell.numFmt = '#,##0.00'; }
        // عمود الحالة (col 17)
        if (col === 17) {
          var st = rowData['الحالة'];
          var stColor = st==='تم الاسترداد' ? 'FF16A34A' : st==='قيد المراجعة' ? 'FFE97A2A' : 'FF64748B';
          cell.font = { name:'Tajawal', size:11, bold:true, color:{ argb: stColor } };
        }
        // عمود نوع الاسترداد (col 4)
        if (col === 4) {
          cell.font = { name:'Tajawal', size:11, bold:true,
            color:{ argb: rowData['نوع الاسترداد']==='مباشر' ? 'FFE97A2A' : 'FF0F766E' } };
        }
        // الرقم المرجعي (col 18)
        if (col === 18 && rowData['الرقم المرجعي']) {
          cell.font = { name:'Tajawal', size:11, bold:true, color:{ argb:'FFB45309' } };
        }
      });
    });

    // ── تجميد الصف الأول من البيانات ──
    var freezeRow = (subtitle ? 3 : 2) + 2; // بعد الإحصائيات + فراغ + رؤوس
    ws.views = [{ state:'frozen', ySplit: freezeRow, rightToLeft:true }];

    // ── تصدير ──
    wb.xlsx.writeBuffer().then(function(buf) {
      var bl = new Blob([buf], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var u = URL.createObjectURL(bl), a = document.createElement('a');
      a.href = u; a.download = filename + '.xlsx'; a.click();
      URL.revokeObjectURL(u);
      toast('تم التصدير بنجاح — ' + data.length + ' سجل', 's');
    }).catch(function(e) { console.error(e); toast('خطأ في التصدير: ' + e.message, 'e'); });

  } catch(e) { console.error('Export error:', e); toast('خطأ في التصدير', 'e'); }
}

// ============================
// تصدير المستخدم: عملاء اليوم بحالة "تم الرفع" فقط
// ============================
function expTodayUploaded() {
  var today = new Date(); today.setHours(0,0,0,0);
  var todayEnd = new Date(); todayEnd.setHours(23,59,59,999);
  var td = todayStr();
  var data = cls.filter(function(c) {
    if (getSt(c) !== 'uploaded') return false;
    if (!c.createdAt || !c.createdAt.seconds) return false;
    var d = new Date(c.createdAt.seconds * 1000);
    return d >= today && d <= todayEnd;
  });
  if (!data.length) { toast('لا توجد بيانات اليوم','w'); return; }
  doUserExport(data, td);
}

// ============================
// تصدير المستخدم — بسيط بدون معلومات حساسة
// ============================
function doUserExport(data, dateStr) {
  var hd = ['#','اسم العميل','الرقم','تاريخ الاشتراك','نوع الباقة',
            'المبلغ المدفوع','أيام الباقة','أيام مستهلكة','المبلغ المسترد',
            'تاريخ الإلغاء','مدة الإلغاء','سبب الإلغاء/الاسترداد','ملاحظات'];
  var cw = [5,24,16,16,24,14,11,11,16,14,12,24,18];
  try {
    var wb = new ExcelJS.Workbook();
    wb.creator = 'V-SHAPE';
    var ws = wb.addWorksheet('البيانات', { views:[{ rightToLeft:true }] });
    ws.columns = hd.map(function(h,i){ return { header:'', key:h, width:cw[i]||15 }; });
    ws.spliceRows(1,1);

    // ── عنوان: التاريخ فقط ──
    var tr = ws.addRow(Array(hd.length).fill(''));
    tr.height = 36;
    ws.mergeCells(1, 1, 1, hd.length);
    tr.getCell(1).value = dateStr;
    tr.getCell(1).font  = { name:'Tajawal', size:16, bold:true, color:{ argb:'FFFFFFFF' } };
    tr.getCell(1).fill  = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF0D9488' } };
    tr.getCell(1).alignment = { horizontal:'center', vertical:'middle' };

    // ── صف فراغ ──
    ws.addRow([]).height = 6;

    // ── رؤوس ──
    var hr = ws.addRow(hd); hr.height = 26;
    hr.eachCell(function(cell) {
      cell.font      = { name:'Tajawal', size:11, bold:true, color:{ argb:'FFFFFFFF' } };
      cell.fill      = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF0F766E' } };
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border    = { bottom:{ style:'medium', color:{ argb:'FF0D9488' } } };
    });

    // ── بيانات ──
    data.forEach(function(c, idx) {
      var row = ws.addRow([
        idx+1,
        c.name||'',
        c.phone||c.mobile||'',
        c.subscriptionDate||'',
        c.packageType||'',
        c.packagePrice||0,
        c.packageDays||0,
        c.consumedDays||0,
        c.refundAmount||0,
        c.cancelDate||'',
        c.cancellationPeriod||'',
        c.cancelReason||'',
        c.notes||''
      ]);
      row.height = 22;
      row.eachCell(function(cell, col) {
        cell.font      = { name:'Tajawal', size:11 };
        cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
        cell.border    = { top:{style:'thin',color:{argb:'FFE2E8F0'}}, bottom:{style:'thin',color:{argb:'FFE2E8F0'}}, left:{style:'thin',color:{argb:'FFE2E8F0'}}, right:{style:'thin',color:{argb:'FFE2E8F0'}} };
        if (idx % 2 === 1) cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFF8FAFC' } };
        if (col === 9) { // المبلغ المسترد
          cell.font = { name:'Tajawal', size:11, bold:true, color:{ argb:'FF0F766E' } };
          cell.numFmt = '#,##0.00';
          cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb: idx%2===0?'FFF0FDF4':'FFE6F9F5' } };
        }
        if (col === 6) cell.numFmt = '#,##0.00';
      });
    });

    ws.views = [{ state:'frozen', ySplit:3, rightToLeft:true }];

    wb.xlsx.writeBuffer().then(function(buf) {
      var bl = new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      var u = URL.createObjectURL(bl), a = document.createElement('a');
      a.href=u; a.download='استرداد_'+dateStr+'.xlsx'; a.click(); URL.revokeObjectURL(u);
      toast('تم التصدير — '+data.length+' سجل','s');
    }).catch(function(e){ toast('خطأ: '+e.message,'e'); });
  } catch(e){ toast('خطأ في التصدير','e'); console.error(e); }
}

// ============================
// تصدير المدير: بالخيارات الكاملة
// ============================
function expAdmin() {
  var st   = document.getElementById('exSt').value;
  var fld  = document.getElementById('exDateFld').value;
  var frV  = document.getElementById('exFr').value;
  var toV  = document.getElementById('exTo').value;

  if (!frV || !toV) { toast('اختر الفترة كاملة', 'e'); return; }

  var fr = new Date(frV); fr.setHours(0,0,0,0);
  var to = new Date(toV); to.setHours(23,59,59,999);

  var data = cls.filter(function(c) {
    // فلتر الحالة
    if (st !== 'all' && getSt(c) !== st) return false;
    // فلتر التاريخ حسب الحقل المختار
    var d;
    if (fld === 'refund') {
      if (!c.refundDate || !c.refundDate.seconds) return false;
      d = new Date(c.refundDate.seconds * 1000);
    } else if (fld === 'subscription') {
      d = parseSubDate(c.subscriptionDate);
      if (!d) return false;
    } else if (fld === 'cancel') {
      d = parseSubDate(c.cancelDate);
      if (!d) return false;
    } else {
      if (!c.createdAt || !c.createdAt.seconds) return false;
      d = new Date(c.createdAt.seconds * 1000);
    }
    return d >= fr && d <= to;
  });

  // بناء العنوان الديناميكي
  var stLbl = { all:'جميع الحالات', uploaded:'تم الرفع', pending:'قيد المراجعة', refunded:'تم الاسترداد' };
  var fldLbl = { created:'تاريخ الإضافة', subscription:'تاريخ الاشتراك', cancel:'تاريخ الإلغاء', refund:'تاريخ الاسترداد' };
  var title = 'عملاء ' + stLbl[st] + ' — ' + fldLbl[fld];
  var subtitle = 'من ' + frV + ' إلى ' + toV + ' | تصدير بواسطة: ' + (cu ? cu.username : '');
  var filename = stLbl[st] + '_' + fldLbl[fld] + '_' + frV + '_إلى_' + toV;

  doExport(data, title, subtitle, filename);
}

// ============================
// تصدير الكل (للمدير)
// ============================
function expAll() {
  if (!cls.length) { toast('لا توجد بيانات', 'w'); return; }
  doExport(cls,
    'جميع السجلات — V-SHAPE',
    'تصدير شامل بدون فلتر | ' + todayStr(),
    'الكل_' + todayStr()
  );
}

// ============================
// دوال قديمة للتوافق (تُحوَّل للنظام الجديد)
// ============================
function expD() { expTodayUploaded(); }
function expByDt() {
  var dv = document.getElementById('exDt').value;
  if (!dv) { toast('اختر تاريخ','e'); return; }
  document.getElementById('exFr').value = dv;
  document.getElementById('exTo').value = dv;
  document.getElementById('exSt').value = 'all';
  document.getElementById('exDateFld').value = 'created';
  expAdmin();
}
function expA() { expAll(); }
