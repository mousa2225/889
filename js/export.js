// ============================
// Export Data Helpers
// ============================
function gED(data) {
  return data.map(function(c, i) {
    return {
      '#': i + 1,
      'اسم العميل': c.name,
      'الرقم': c.phone || c.mobile || '',
      'تاريخ الاشتراك': c.subscriptionDate || '',
      'نوع الباقة': c.packageType,
      'المبلغ المدفوع': c.packagePrice || 0,
      'أيام الباقة': c.packageDays || 0,
      'أيام مستهلكة': c.consumedDays || 0,
      'المبلغ المسترد': c.refundAmount || 0,
      'تاريخ الإلغاء': c.cancelDate || '',
      'مدة الإلغاء': c.cancellationPeriod || '',
      'سبب الإلغاء': c.cancelReason || '',
      'تاريخ الإضافة': c.createdAt ? fmtDtT(c.createdAt) : '',
      'تاريخ الاسترداد': c.refundDate ? fmtDtT(c.refundDate) : '',
      'أضاف بواسطة': c.addedByUsername || '',
      'الحالة': STL[getSt(c)] || '',
      'الرقم المرجعي': c.referenceNumber || '',
      'ملاحظات': c.previouslyRefunded ? 'سبق الاسترداد' : (c.notes || '')
    };
  });
}

function dExp(data, fn) {
  try {
    var wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('استرداد العملاء');
    var hd = ['#','اسم العميل','الرقم','تاريخ الاشتراك','نوع الباقة','المبلغ المدفوع','أيام الباقة','أيام مستهلكة','المبلغ المسترد','تاريخ الإلغاء','مدة الإلغاء','سبب الإلغاء','تاريخ الإضافة','تاريخ الاسترداد','أضاف بواسطة','الحالة','الرقم المرجعي','ملاحظات'];
    var cw = [5,22,16,16,22,14,12,12,16,14,14,20,18,18,18,14,14,16];
    ws.columns = hd.map(function(h, i) { return { header:'', key:h, width:cw[i]||15 }; });
    ws.spliceRows(1, 1);
    // Title row
    var tr = ws.addRow(['نظام استرداد العملاء V-SHAPE','','','','','','','','','','','','','','','','','']);
    ws.mergeCells('A1:R1'); tr.height = 36;
    tr.getCell(1).value = 'نظام استرداد العملاء V-SHAPE';
    tr.getCell(1).font  = { name:'Tajawal', size:16, bold:true, color:{ argb:'FFFFFFFF' } };
    tr.getCell(1).fill  = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF0D9488' } };
    tr.getCell(1).alignment = { horizontal:'center', vertical:'middle' };
    ws.getRow(1).height = 36;
    ws.addRow([]).height = 6;
    // Header row
    var hr = ws.addRow(hd); hr.height = 28;
    hr.eachCell(function(cell) {
      cell.font      = { name:'Tajawal', size:11, bold:true, color:{ argb:'FFFFFFFF' } };
      cell.fill      = { type:'pattern', pattern:'solid', fgColor:{ argb:'FF0F766E' } };
      cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
      cell.border    = { top:{style:'thin',color:{argb:'FF0D9488'}}, bottom:{style:'thin',color:{argb:'FF0D9488'}}, left:{style:'thin',color:{argb:'FF14B8A6'}}, right:{style:'thin',color:{argb:'FF14B8A6'}} };
    });
    // Data rows
    data.forEach(function(c, idx) {
      var row = ws.addRow(Object.values(c)); row.height = 24;
      var ip = c['ملاحظات'] && c['ملاحظات'].includes('سبق');
      row.eachCell(function(cell, col) {
        cell.font      = { name:'Tajawal', size:11 };
        cell.alignment = { horizontal:'center', vertical:'middle', wrapText:true };
        cell.border    = { top:{style:'thin',color:{argb:'FFE2E8F0'}}, bottom:{style:'thin',color:{argb:'FFE2E8F0'}}, left:{style:'thin',color:{argb:'FFE2E8F0'}}, right:{style:'thin',color:{argb:'FFE2E8F0'}} };
        if (ip) { cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFFEE2E2' } }; cell.font = { name:'Tajawal', size:11, color:{ argb:'FF991B1B' } }; }
        else if (idx % 2 === 1) cell.fill = { type:'pattern', pattern:'solid', fgColor:{ argb:'FFF8FAFC' } };
        if (col === 8)  { cell.font = { name:'Tajawal', size:11, bold:true, color:{ argb:'FF0F766E' } }; cell.numFmt = '#,##0.00'; }
        if (col === 15) {
          var st = c['الحالة'];
          if      (st === 'تم الاسترداد') cell.font = { name:'Tajawal', size:11, bold:true, color:{ argb:'FF16A34A' } };
          else if (st === 'قيد المراجعة') cell.font = { name:'Tajawal', size:11, bold:true, color:{ argb:'FFE97A2A' } };
          else cell.font = { name:'Tajawal', size:11, color:{ argb:'FF64748B' } };
        }
        if (col === 16 && c['الرقم المرجعي']) cell.font = { name:'Tajawal', size:11, bold:true, color:{ argb:'FFB45309' } };
      });
    });
    wb.xlsx.writeBuffer().then(function(buf) {
      var bl = new Blob([buf], { type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var u = URL.createObjectURL(bl), a = document.createElement('a');
      a.href = u; a.download = fn + '.xlsx'; a.click();
      URL.revokeObjectURL(u);
      toast('تم التصدير بنجاح','s');
    }).catch(function(e) { console.error(e); toast('خطأ في التصدير: ' + e.message,'e'); });
  } catch(e) { console.error('Export error:', e); toast('خطأ في التصدير','e'); }
}

function expD() {
  var n = new Date(), td = n.getFullYear() + '-' + String(n.getMonth()+1).padStart(2,'0') + '-' + String(n.getDate()).padStart(2,'0');
  var d = gED(cls.filter(function(c) {
    if (!c.createdAt || !c.createdAt.seconds) return false;
    var x = new Date(c.createdAt.seconds * 1000);
    return x.getFullYear() === n.getFullYear() && x.getMonth() === n.getMonth() && x.getDate() === n.getDate();
  }));
  if (!d.length) { toast('لا توجد بيانات اليوم','w'); return; }
  dExp(d, 'استرداد_' + td);
}
function expByDt() {
  var dv = document.getElementById('exDt').value;
  if (!dv) { toast('اختر تاريخ','e'); return; }
  var d = gED(cls.filter(function(c) {
    if (!c.createdAt || !c.createdAt.seconds) return false;
    var x = new Date(c.createdAt.seconds * 1000), p = dv.split('-');
    return x.getFullYear() === parseInt(p[0]) && x.getMonth() === parseInt(p[1])-1 && x.getDate() === parseInt(p[2]);
  }));
  if (!d.length) { toast('لا توجد بيانات','w'); return; }
  dExp(d, 'استرداد_' + dv);
}
function expA() {
  if (!cls.length) { toast('لا توجد بيانات','w'); return; }
  dExp(gED(cls), 'استرداد_الكل');
}
