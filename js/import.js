// ============================
// استيراد بيانات Excel للمدير
// ============================

function openImport() {
  if (!cu || cu.role !== 'admin') return;
  document.getElementById('impM').classList.add('on');
  document.getElementById('impFile').value = '';
  document.getElementById('impPreview').innerHTML = '';
  document.getElementById('impInfo').textContent = '';
  document.getElementById('impDate').value = new Date().toISOString().split('T')[0];
  document.getElementById('impStatus').value = 'refunded';
  impRows = [];
}
function clImport() { document.getElementById('impM').classList.remove('on'); }

var impRows = [];

// قراءة الملف
function impReadFile() {
  var file = document.getElementById('impFile').files[0];
  if (!file) return;
  var reader = new FileReader();
  reader.onload = function(e) {
    try {
      var wb   = XLSX.read(e.target.result, { type:'array' });
      var ws   = wb.Sheets[wb.SheetNames[0]];
      var data = XLSX.utils.sheet_to_json(ws, { defval:'' });
      if (!data.length) { toast('الملف فارغ','e'); return; }
      impRows = data;
      impShowPreview(data);
    } catch(err) { toast('خطأ في قراءة الملف','e'); console.error(err); }
  };
  reader.readAsArrayBuffer(file);
}

// عرض معاينة
function impShowPreview(data) {
  var sample = data.slice(0, 5);
  var cols   = Object.keys(data[0]);

  // خريطة الأعمدة
  var mapDiv = document.getElementById('impColMap');
  var fields = [
    { key:'name',            label:'اسم العميل *',         req:true  },
    { key:'phone',           label:'الرقم *',              req:true  },
    { key:'subscriptionDate',label:'تاريخ الاشتراك',       req:false },
    { key:'packageType',     label:'نوع الباقة',           req:false },
    { key:'packageDays',     label:'أيام الباقة',          req:false },
    { key:'consumedDays',    label:'أيام مستهلكة',         req:false },
    { key:'packagePrice',    label:'المبلغ المدفوع',       req:false },
    { key:'refundAmount',    label:'المبلغ المسترد *',     req:true  },
    { key:'cancelReason',    label:'سبب الإلغاء/الاسترداد',req:false },
    { key:'referenceNumber', label:'الرقم المرجعي',        req:false },
    { key:'notes',           label:'ملاحظات',              req:false },
  ];

  var html = '<div style="font-size:10px;font-weight:700;color:var(--mt);margin-bottom:8px">تعيين الأعمدة</div>'
    + '<div style="display:grid;grid-template-columns:1fr 1fr;gap:8px">';
  fields.forEach(function(f) {
    html += '<div style="display:flex;align-items:center;gap:6px">'
      + '<label style="font-size:10px;font-weight:700;color:var(--dm);min-width:120px">'
      + f.label + '</label>'
      + '<select id="imp_'+f.key+'" class="fi fi-sm" style="flex:1">'
      + '<option value="">— لا يوجد —</option>'
      + cols.map(function(c){
          // محاولة تلقائية للمطابقة
          var autoMatch = autoMapCol(c, f.key);
          return '<option value="'+c+'"'+(autoMatch?' selected':'')+'>'+c+'</option>';
        }).join('')
      + '</select></div>';
  });
  html += '</div>';
  mapDiv.innerHTML = html;

  // معاينة أول 5 صفوف
  var prev = '<div style="overflow-x:auto;margin-top:12px"><table style="width:100%;border-collapse:collapse;font-size:10px">'
    + '<thead><tr>'
    + cols.map(function(c){ return '<th style="padding:4px 8px;background:rgba(0,212,170,.08);color:#00d4aa;font-weight:700;white-space:nowrap">'+c+'</th>'; }).join('')
    + '</tr></thead><tbody>'
    + sample.map(function(row, i){
        return '<tr style="background:'+(i%2===0?'transparent':'rgba(255,255,255,.03)')+'">'
          + cols.map(function(c){ return '<td style="padding:4px 8px;border-bottom:1px solid var(--brd);white-space:nowrap">'+escText(String(row[c]||''))+'</td>'; }).join('')
          + '</tr>';
      }).join('')
    + '</tbody></table></div>';

  document.getElementById('impPreview').innerHTML = prev;
  document.getElementById('impInfo').textContent = 'إجمالي الصفوف: ' + data.length;
}

function escText(t){ var d=document.createElement('div'); d.textContent=t; return d.innerHTML; }

// محاولة مطابقة تلقائية
function autoMapCol(colName, fieldKey) {
  var c = colName.toLowerCase().replace(/\s/g,'');
  var maps = {
    name:            ['اسم','name','عميل','client'],
    phone:           ['رقم','phone','mobile','جوال','هاتف','tel'],
    subscriptionDate:['اشتراك','subscription','تاريخ الاشتراك'],
    packageType:     ['باقة','package','نوع'],
    packageDays:     ['ايام الباقة','packagedays','أيام الباقة','days'],
    consumedDays:    ['مستهلك','consumed','مستهلكة'],
    packagePrice:    ['مدفوع','paid','price','سعر','المبلغ المدفوع'],
    refundAmount:    ['مسترد','refund','استرداد','المبلغ المسترد'],
    cancelReason:    ['سبب','reason','cancel'],
    referenceNumber: ['مرجعي','reference','ref'],
    notes:           ['ملاحظ','notes','note'],
  };
  var keys = maps[fieldKey] || [];
  return keys.some(function(k){ return c.includes(k.toLowerCase().replace(/\s/g,'')); });
}

function getImpVal(row, fieldKey) {
  var sel = document.getElementById('imp_'+fieldKey);
  if (!sel || !sel.value) return '';
  return row[sel.value] !== undefined ? String(row[sel.value]) : '';
}

// تنفيذ الاستيراد
function doImport() {
  if (!impRows.length) { toast('لا توجد بيانات','e'); return; }

  var dateVal = document.getElementById('impDate').value;
  var status  = document.getElementById('impStatus').value;
  if (!dateVal) { toast('حدد تاريخ الإضافة','e'); return; }

  // تحقق من الأعمدة الإلزامية
  var nameCol   = document.getElementById('imp_name').value;
  var phoneCol  = document.getElementById('imp_phone').value;
  var refundCol = document.getElementById('imp_refundAmount').value;
  if (!nameCol)   { toast('يجب تعيين عمود اسم العميل','e'); return; }
  if (!phoneCol)  { toast('يجب تعيين عمود الرقم','e'); return; }
  if (!refundCol) { toast('يجب تعيين عمود المبلغ المسترد','e'); return; }

  var importDate = new Date(dateVal);
  importDate.setHours(12,0,0,0);
  var ts = firebase.firestore.Timestamp.fromDate(importDate);

  var batch = db.batch();
  var count = 0, errors = 0;

  impRows.forEach(function(row) {
    var name   = getImpVal(row,'name').trim();
    var phone  = getImpVal(row,'phone').trim();
    var refund = parseFloat(String(getImpVal(row,'refundAmount')).replace(/[^\d.]/g,'')) || 0;

    if (!name || !phone) { errors++; return; }

    var sub    = normDateStr(getImpVal(row,'subscriptionDate').trim()) || '';
    var pk     = getImpVal(row,'packageType').trim();
    var dy     = parseFloat(getImpVal(row,'packageDays'))  || 0;
    var co     = parseFloat(getImpVal(row,'consumedDays')) || 0;
    var pr     = parseFloat(String(getImpVal(row,'packagePrice')).replace(/[^\d.]/g,'')) || 0;
    var rs     = getImpVal(row,'cancelReason').trim();
    var ref    = getImpVal(row,'referenceNumber').trim();
    var notes  = getImpVal(row,'notes').trim();

    // هل تم استرداده سابقاً بنفس الرقم والاشتراك
    var normSub = normDateStr(sub) || sub;
    var prevRef = cls.some(function(x){
      var xph=(x.phone||x.mobile||'').trim();
      return xph && xph===phone && (x.subscriptionDate||'').trim()===normSub && getSt(x)==='refunded';
    });

    var doc = {
      name: name, phone: phone, mobile: phone,
      refundType: pr>0&&dy>0 ? 'subscription' : 'direct',
      subscriptionDate: normSub,
      packageType: pk, packagePrice: pr, packageDays: dy, consumedDays: co,
      refundAmount: refund, cancelReason: rs,
      referenceNumber: ref, notes: notes,
      status: status, refunded: status==='refunded',
      previouslyRefunded: prevRef,
      addedByUsername: cu.username + ' (استيراد)',
      importedAt: ts,
      createdAt: ts
    };
    if (status==='refunded') doc.refundDate = ts;

    batch.set(db.collection('cancellations').doc(), doc);
    count++;
  });

  if (!count) { toast('لا توجد صفوف صالحة للاستيراد','e'); return; }

  var btn = document.getElementById('impDoBtn');
  btn.disabled = true;
  btn.textContent = 'جاري الاستيراد...';

  batch.commit().then(function(){
    var msg = 'تم استيراد '+count+' سجل'; if(errors) msg += ' (تم تخطي '+errors+')'; toast(msg, 's');
    clImport();
    chkP();
  }).catch(function(e){
    toast('خطأ في الاستيراد','e');
    console.error(e);
    btn.disabled = false;
    btn.textContent = 'استيراد';
  });
}
