// ============================
// التقارير المالية
// ============================
var rpCharts = {}; // لتخزين Chart instances لتدميرها عند إعادة الرسم

function openRP() {
  if (!cu || cu.role !== 'admin') return;
  document.getElementById('rpM').classList.add('on');
  rpSetDefaults();
  rpRender();
}
function clRP() { document.getElementById('rpM').classList.remove('on'); }

function rpSetDefaults() {
  var now = new Date();
  var y = now.getFullYear(), m = now.getMonth();
  var firstDay = new Date(y, m, 1), lastDay = new Date(y, m + 1, 0);
  var fmt = function(d) { return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0'); };
  document.getElementById('rpFr').value = fmt(firstDay);
  document.getElementById('rpTo').value = fmt(lastDay);
}

function rpGetData() {
  var fr = document.getElementById('rpFr').value, to = document.getElementById('rpTo').value;
  if (!fr || !to) return cls;
  var f = new Date(fr); f.setHours(0,0,0,0);
  var t = new Date(to); t.setHours(23,59,59,999);
  return cls.filter(function(c) {
    if (!c.createdAt || !c.createdAt.seconds) return false;
    var d = new Date(c.createdAt.seconds * 1000);
    return d >= f && d <= t;
  });
}

function rpRender() {
  var data = rpGetData();
  rpDestCharts();
  rpKPI(data);
  rpChartStatus(data);
  rpChartType(data);
  rpChartDaily(data);
  rpChartReasons(data);
  rpChartUsers(data);
  rpChartDiff(data);
  rpTableDiff(data);
  rpTableTopClients(data);
}

function rpDestCharts() {
  Object.keys(rpCharts).forEach(function(k) { if (rpCharts[k]) { rpCharts[k].destroy(); delete rpCharts[k]; } });
}

// ── KPI Cards ──
function rpKPI(data) {
  var total    = data.length;
  var refunded = data.filter(function(c){ return getSt(c)==='refunded'; });
  var pending  = data.filter(function(c){ return getSt(c)==='pending'; });
  var uploaded = data.filter(function(c){ return getSt(c)==='uploaded'; });
  var rejected = data.filter(function(c){ return getSt(c)==='rejected'; });
  var direct   = data.filter(function(c){ return c.refundType==='direct'; });
  var sub      = data.filter(function(c){ return c.refundType!=='direct'; });
  var totalAmt = data.reduce(function(s,c){ return s+(c.refundAmount||0); },0);
  var refAmt   = refunded.reduce(function(s,c){ return s+(c.refundAmount||0); },0);
  var pendAmt  = pending.reduce(function(s,c){ return s+(c.refundAmount||0); },0);
  var upAmt    = uploaded.reduce(function(s,c){ return s+(c.refundAmount||0); },0);
  var prevRef  = data.filter(function(c){ return isDynPrevRef(c); });
  // فروقات التعديل
  var diffRecs = data.filter(function(c){ return c.adminEdited && c.originalRefundAmount !== undefined && c.originalRefundAmount !== null; });
  var totalDiff= diffRecs.reduce(function(s,c){ return s + Math.abs((c.originalRefundAmount||0)-(c.refundAmount||0)); }, 0);

  var kv = {
    rpK1:  total,
    rpK2:  en(totalAmt.toFixed(2))+' ريال',
    rpK3:  refunded.length,
    rpK4:  en(refAmt.toFixed(2))+' ريال',
    rpK5:  pending.length,
    rpK6:  en(pendAmt.toFixed(2))+' ريال',
    rpK7:  uploaded.length,
    rpK8:  prevRef.length,
    rpK9:  sub.length,
    rpK10: direct.length,
    rpK11: en(upAmt.toFixed(2))+' ريال',
    rpK12: rejected.length,
    rpK13: diffRecs.length + ' سجل | ' + en(totalDiff.toFixed(2)) + ' ريال'
  };
  Object.keys(kv).forEach(function(id){ var el=document.getElementById(id); if(el) el.textContent=kv[id]; });
}

// ── Chart: توزيع الحالات ──
function rpChartStatus(data) {
  var rf=data.filter(function(c){return getSt(c)==='refunded';}).length;
  var pn=data.filter(function(c){return getSt(c)==='pending';}).length;
  var up=data.filter(function(c){return getSt(c)==='uploaded';}).length;
  var ctx=document.getElementById('rpCStatus'); if(!ctx) return;
  var isDk = document.documentElement.className !== 'lt';
  var rj=data.filter(function(c){return getSt(c)==='rejected';}).length;
  rpCharts.status = new Chart(ctx, {
    type:'doughnut',
    data:{
      labels:['تم الاسترداد','قيد المراجعة','تم الرفع','مرفوض'],
      datasets:[{data:[rf,pn,up,rj],backgroundColor:['#22c55e','#e97a2a','#64748b','#dc3545'],borderWidth:0,hoverOffset:6}]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'bottom',labels:{color:isDk?'#94a3b8':'#475569',font:{family:'Tajawal',size:11}}},tooltip:{callbacks:{label:function(c){return c.label+': '+c.parsed+' طلب';}}}}
    }
  });
}

// ── Chart: نوع الاسترداد ──
function rpChartType(data) {
  var sub=data.filter(function(c){return c.refundType!=='direct';});
  var dir=data.filter(function(c){return c.refundType==='direct';});
  var subAmt=sub.reduce(function(s,c){return s+(c.refundAmount||0);},0);
  var dirAmt=dir.reduce(function(s,c){return s+(c.refundAmount||0);},0);
  var ctx=document.getElementById('rpCType'); if(!ctx) return;
  var isDk = document.documentElement.className !== 'lt';
  rpCharts.type = new Chart(ctx, {
    type:'doughnut',
    data:{
      labels:['استرداد اشتراك','استرداد مباشر'],
      datasets:[{data:[Math.round(subAmt*100)/100, Math.round(dirAmt*100)/100],backgroundColor:['#00d4aa','#e97a2a'],borderWidth:0,hoverOffset:6}]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{position:'bottom',labels:{color:isDk?'#94a3b8':'#475569',font:{family:'Tajawal',size:11}}},tooltip:{callbacks:{label:function(c){return c.label+': '+c.parsed.toFixed(2)+' ريال';}}}}
    }
  });
}

// ── Chart: الاسترداد اليومي ──
function rpChartDaily(data) {
  var ctx=document.getElementById('rpCDaily'); if(!ctx) return;
  var isDk = document.documentElement.className !== 'lt';
  var days={};
  data.forEach(function(c){
    if(!c.createdAt||!c.createdAt.seconds) return;
    var d=new Date(c.createdAt.seconds*1000);
    var key=d.getDate()+' '+MA[d.getMonth()];
    if(!days[key]) days[key]={count:0,amt:0};
    days[key].count++;
    days[key].amt+=(c.refundAmount||0);
  });
  // Sort by date
  var fr=document.getElementById('rpFr').value, to=document.getElementById('rpTo').value;
  var labels=[], counts=[], amts=[];
  if (fr&&to) {
    var cur=new Date(fr); var end=new Date(to);
    while(cur<=end){
      var key=cur.getDate()+' '+MA[cur.getMonth()];
      labels.push(key);
      counts.push(days[key]?days[key].count:0);
      amts.push(days[key]?Math.round(days[key].amt*100)/100:0);
      cur.setDate(cur.getDate()+1);
    }
  } else {
    labels=Object.keys(days); counts=labels.map(function(k){return days[k].count;}); amts=labels.map(function(k){return Math.round(days[k].amt*100)/100;});
  }
  rpCharts.daily = new Chart(ctx, {
    type:'bar',
    data:{
      labels:labels,
      datasets:[
        {label:'المبلغ المسترد (ريال)',data:amts,backgroundColor:'rgba(0,212,170,.7)',borderRadius:4,yAxisID:'y'},
        {label:'عدد الطلبات',data:counts,type:'line',borderColor:'#e97a2a',backgroundColor:'rgba(233,122,42,.15)',pointRadius:3,tension:.3,yAxisID:'y1'}
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{color:isDk?'#94a3b8':'#475569',font:{family:'Tajawal',size:10}}}},
      scales:{
        x:{ticks:{color:isDk?'#64748b':'#94a3b8',font:{family:'Tajawal',size:9}},grid:{color:isDk?'rgba(255,255,255,.04)':'rgba(0,0,0,.04)'}},
        y:{position:'right',ticks:{color:'#00d4aa',font:{family:'Tajawal',size:9}},grid:{color:isDk?'rgba(255,255,255,.04)':'rgba(0,0,0,.04)'}},
        y1:{position:'left',ticks:{color:'#e97a2a',font:{family:'Tajawal',size:9}},grid:{display:false}}
      }
    }
  });
}

// ── Chart: أسباب الإلغاء ──
function rpChartReasons(data) {
  var ctx=document.getElementById('rpCReasons'); if(!ctx) return;
  var isDk = document.documentElement.className !== 'lt';
  var reasons={};
  data.forEach(function(c){
    var r=(c.cancelReason||'').trim()||'غير محدد';
    if(r.length>20) r=r.substring(0,20)+'...';
    reasons[r]=(reasons[r]||0)+1;
  });
  var sorted=Object.entries(reasons).sort(function(a,b){return b[1]-a[1];}).slice(0,8);
  rpCharts.reasons = new Chart(ctx, {
    type:'bar',
    data:{
      labels:sorted.map(function(x){return x[0];}),
      datasets:[{label:'عدد الحالات',data:sorted.map(function(x){return x[1];}),
        backgroundColor:['#00d4aa','#e97a2a','#60a5fa','#a78bfa','#f59e0b','#22c55e','#ec4899','#64748b'],borderRadius:4}]
    },
    options:{
      indexAxis:'y',responsive:true,maintainAspectRatio:false,
      plugins:{legend:{display:false}},
      scales:{
        x:{ticks:{color:isDk?'#64748b':'#94a3b8',font:{family:'Tajawal',size:9}},grid:{color:isDk?'rgba(255,255,255,.04)':'rgba(0,0,0,.04)'}},
        y:{ticks:{color:isDk?'#94a3b8':'#475569',font:{family:'Tajawal',size:9}},grid:{display:false}}
      }
    }
  });
}

// ── Chart: أداء المستخدمين ──
function rpChartUsers(data) {
  var ctx=document.getElementById('rpCUsers'); if(!ctx) return;
  var isDk = document.documentElement.className !== 'lt';
  var users={};
  data.forEach(function(c){
    var u=c.addedByUsername||'غير محدد';
    if(!users[u]) users[u]={count:0,amt:0,refunded:0};
    users[u].count++;
    users[u].amt+=(c.refundAmount||0);
    if(getSt(c)==='refunded') users[u].refunded++;
  });
  var sorted=Object.entries(users).sort(function(a,b){return b[1].count-a[1].count;});
  rpCharts.users = new Chart(ctx, {
    type:'bar',
    data:{
      labels:sorted.map(function(x){return x[0];}),
      datasets:[
        {label:'إجمالي الطلبات',data:sorted.map(function(x){return x[1].count;}),backgroundColor:'rgba(96,165,250,.7)',borderRadius:4},
        {label:'تم الاسترداد',data:sorted.map(function(x){return x[1].refunded;}),backgroundColor:'rgba(34,197,94,.7)',borderRadius:4}
      ]
    },
    options:{
      responsive:true,maintainAspectRatio:false,
      plugins:{legend:{labels:{color:isDk?'#94a3b8':'#475569',font:{family:'Tajawal',size:10}}}},
      scales:{
        x:{ticks:{color:isDk?'#94a3b8':'#475569',font:{family:'Tajawal',size:10}},grid:{display:false}},
        y:{ticks:{color:isDk?'#64748b':'#94a3b8',font:{family:'Tajawal',size:9}},grid:{color:isDk?'rgba(255,255,255,.04)':'rgba(0,0,0,.04)'}}
      }
    }
  });
}

// ── Chart: فروقات التعديلات المالية ──
function rpChartDiff(data) {
  var ctx = document.getElementById('rpCDiff'); if (!ctx) return;
  var isDk = document.documentElement.className !== 'lt';
  var recs = data.filter(function(c){
    return c.adminEdited && c.originalRefundAmount !== undefined && c.originalRefundAmount !== null
      && Math.abs((c.originalRefundAmount||0)-(c.refundAmount||0)) > 0.001;
  }).slice(0, 10);
  if (!recs.length) {
    ctx.parentElement.innerHTML = '<div style="text-align:center;padding:20px;color:var(--mt);font-size:12px"><i class="fa-solid fa-check-circle" style="color:#22c55e;font-size:20px;display:block;margin-bottom:8px"></i>لا توجد فروقات تعديل</div>';
    return;
  }
  var labels = recs.map(function(c){ return c.name ? c.name.substring(0,12) : c.id.substring(0,8); });
  var origVals = recs.map(function(c){ return Math.round((c.originalRefundAmount||0)*100)/100; });
  var newVals  = recs.map(function(c){ return Math.round((c.refundAmount||0)*100)/100; });
  var diffs    = recs.map(function(c){ return Math.round(Math.abs((c.originalRefundAmount||0)-(c.refundAmount||0))*100)/100; });
  rpCharts.diff = new Chart(ctx, {
    type:'bar',
    data:{
      labels: labels,
      datasets:[
        {label:'المبلغ الأصلي', data:origVals, backgroundColor:'rgba(96,165,250,.7)', borderRadius:4},
        {label:'بعد التعديل',   data:newVals,  backgroundColor:'rgba(0,212,170,.7)',  borderRadius:4},
        {label:'الفرق',         data:diffs,    backgroundColor:'rgba(245,158,11,.7)', borderRadius:4}
      ]
    },
    options:{
      responsive:true, maintainAspectRatio:false,
      plugins:{legend:{labels:{color:isDk?'#94a3b8':'#475569',font:{family:'Tajawal',size:10}}}},
      scales:{
        x:{ticks:{color:isDk?'#94a3b8':'#475569',font:{family:'Tajawal',size:9}},grid:{display:false}},
        y:{ticks:{color:isDk?'#64748b':'#94a3b8',font:{family:'Tajawal',size:9}},grid:{color:isDk?'rgba(255,255,255,.04)':'rgba(0,0,0,.04)'}}
      }
    }
  });
}

// ── جدول فروقات التعديلات ──
function rpTableDiff(data) {
  var tbody = document.getElementById('rpDiffBody'); if (!tbody) return;
  var recs = data.filter(function(c){
    return c.adminEdited && c.originalRefundAmount !== undefined && c.originalRefundAmount !== null
      && Math.abs((c.originalRefundAmount||0)-(c.refundAmount||0)) > 0.001;
  });
  if (!recs.length) {
    tbody.innerHTML='<tr><td colspan="6" style="text-align:center;padding:16px;color:var(--mt);font-size:12px"><i class="fa-solid fa-check-circle" style="color:#22c55e"></i> لا توجد فروقات</td></tr>';
    return;
  }
  tbody.innerHTML = recs.map(function(c){
    var orig = c.originalRefundAmount||0, curr = c.refundAmount||0, diff = curr-orig;
    var dcolor = diff < 0 ? '#22c55e' : '#dc3545';
    var dsign  = diff < 0 ? '▼ ' : '▲ ';
    return '<tr style="border-bottom:1px solid var(--brd)">'
      +'<td style="padding:6px 8px;font-size:11px;font-weight:600">'+esc(c.name)+'</td>'
      +'<td style="padding:6px 8px;font-size:11px" dir="ltr">'+esc(c.phone||c.mobile||'—')+'</td>'
      +'<td style="padding:6px 8px;font-size:12px;font-weight:700;color:#94a3b8;text-decoration:line-through">'+en(orig.toFixed(2))+'</td>'
      +'<td style="padding:6px 8px;font-size:12px;font-weight:800;color:#00d4aa">'+en(curr.toFixed(2))+'</td>'
      +'<td style="padding:6px 8px;font-size:12px;font-weight:800;color:'+dcolor+'">'+dsign+en(Math.abs(diff).toFixed(2))+'</td>'
      +'<td style="padding:6px 8px;font-size:11px;color:var(--dm)">'+esc(c.addedByUsername||'—')+'</td>'
      +'</tr>';
  }).join('');
}

// ── جدول العملاء المكررين ──
function rpTableTopClients(data) {
  var tbody=document.getElementById('rpTBody'); if(!tbody) return;
  var prev=data.filter(function(c){ return isDynPrevRef(c); });
  if (!prev.length) {
    tbody.innerHTML='<tr><td colspan="5" style="text-align:center;padding:16px;color:var(--mt);font-size:12px">لا يوجد</td></tr>';
    return;
  }
  tbody.innerHTML=prev.slice(0,10).map(function(c){
    var st=getSt(c);
    return '<tr style="border-bottom:1px solid var(--brd)">'+
      '<td style="padding:6px 8px;font-size:11px">'+esc(c.name)+'</td>'+
      '<td style="padding:6px 8px;font-size:11px;direction:ltr">'+esc(c.phone||c.mobile||'—')+'</td>'+
      '<td style="padding:6px 8px;font-size:11px">'+esc(c.subscriptionDate||'—')+'</td>'+
      '<td style="padding:6px 8px;font-size:11px;font-weight:700;color:#00d4aa">'+en((c.refundAmount||0).toFixed(2))+'</td>'+
      '<td style="padding:6px 8px"><span class="sb '+STC[st]+'" style="font-size:9px"><i class="'+STI[st]+'" style="font-size:8px"></i> '+STL[st]+'</span></td>'+
    '</tr>';
  }).join('');
}

// ── تصدير التقرير ──
function rpExport() {
  var data = rpGetData();
  if (!data.length) { toast('لا توجد بيانات في هذه الفترة','w'); return; }
  var fr=document.getElementById('rpFr').value, to=document.getElementById('rpTo').value;
  var fn = 'تقرير_مالي_'+(fr||'')+'_إلى_'+(to||'');

  // إحصائيات
  var rf=data.filter(function(c){return getSt(c)==='refunded';}), pn=data.filter(function(c){return getSt(c)==='pending';}), up=data.filter(function(c){return getSt(c)==='uploaded';});
  var totalAmt=data.reduce(function(s,c){return s+(c.refundAmount||0);},0);
  var refAmt=rf.reduce(function(s,c){return s+(c.refundAmount||0);},0);

  try {
    var wb=new ExcelJS.Workbook(), ws=wb.addWorksheet('التقرير المالي');
    // Title
    ws.mergeCells('A1:H1');
    var tr=ws.getRow(1); tr.height=36;
    tr.getCell(1).value='التقرير المالي الشامل — V-SHAPE';
    tr.getCell(1).font={name:'Tajawal',size:16,bold:true,color:{argb:'FFFFFFFF'}};
    tr.getCell(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0D9488'}};
    tr.getCell(1).alignment={horizontal:'center',vertical:'middle'};
    // Period
    ws.mergeCells('A2:H2');
    var tr2=ws.getRow(2); tr2.height=22;
    tr2.getCell(1).value='الفترة: '+(fr||'بداية')+' إلى '+(to||'نهاية');
    tr2.getCell(1).font={name:'Tajawal',size:11,color:{argb:'FF64748B'}};
    tr2.getCell(1).alignment={horizontal:'center',vertical:'middle'};
    ws.addRow([]);
    // Summary
    var sh=ws.addRow(['الإجمالي','تم الاسترداد','قيد المراجعة','تم الرفع','إجمالي المبالغ','المبالغ المستردة فعلياً','','']);
    sh.height=24;
    sh.eachCell(function(cell){cell.font={name:'Tajawal',size:11,bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0F766E'}};cell.alignment={horizontal:'center',vertical:'middle'};});
    var sv=ws.addRow([data.length,rf.length,pn.length,up.length,totalAmt.toFixed(2),refAmt.toFixed(2),'','']);
    sv.height=22;
    sv.eachCell(function(cell,col){
      cell.font={name:'Tajawal',size:12,bold:true};
      cell.alignment={horizontal:'center',vertical:'middle'};
      if(col<=4)cell.font={name:'Tajawal',size:12,bold:true,color:{argb:'FF0F766E'}};
      if(col===5||col===6){cell.font={name:'Tajawal',size:12,bold:true,color:{argb:'FFF59E0B'}};cell.numFmt='#,##0.00';}
    });
    ws.addRow([]);
    // Data header
    var hdrs=['#','اسم العميل','الرقم','نوع الاسترداد','تاريخ الاشتراك','نوع الباقة','المبلغ المسترد','تاريخ الإضافة','تاريخ الاسترداد','الحالة','سبب الإلغاء/الاسترداد','أضاف بواسطة'];
    var ws2=wb.addWorksheet('تفاصيل السجلات');
    ws2.columns=[{width:5},{width:22},{width:16},{width:16},{width:16},{width:22},{width:16},{width:18},{width:18},{width:14},{width:22},{width:16}];
    var h2=ws2.addRow(hdrs); h2.height=26;
    h2.eachCell(function(cell){cell.font={name:'Tajawal',size:11,bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0F766E'}};cell.alignment={horizontal:'center',vertical:'middle',wrapText:true};});
    data.forEach(function(c,idx){
      var row=ws2.addRow([idx+1,c.name,c.phone||c.mobile||'',c.refundType==='direct'?'مباشر':'اشتراك',c.subscriptionDate||'',c.packageType||'',(c.refundAmount||0),c.createdAt?fmtDtT(c.createdAt):'',c.refundDate?fmtDtT(c.refundDate):'',STL[getSt(c)]||'',c.cancelReason||'',c.addedByUsername||'']);
      row.height=22;
      row.eachCell(function(cell,col){
        cell.font={name:'Tajawal',size:11};
        cell.alignment={horizontal:'center',vertical:'middle',wrapText:true};
        cell.border={top:{style:'thin',color:{argb:'FFE2E8F0'}},bottom:{style:'thin',color:{argb:'FFE2E8F0'}},left:{style:'thin',color:{argb:'FFE2E8F0'}},right:{style:'thin',color:{argb:'FFE2E8F0'}}};
        if(idx%2===1)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF8FAFC'}};
        if(col===7){cell.font={name:'Tajawal',size:11,bold:true,color:{argb:'FF0F766E'}};cell.numFmt='#,##0.00';}
      });
    });
    wb.xlsx.writeBuffer().then(function(buf){
      var bl=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      var u=URL.createObjectURL(bl),a=document.createElement('a');
      a.href=u;a.download=fn+'.xlsx';a.click();URL.revokeObjectURL(u);
      toast('تم تصدير التقرير','s');
    }).catch(function(e){console.error(e);toast('خطأ في التصدير','e');});
  } catch(e){console.error(e);toast('خطأ','e');}
}

// ── طباعة التقرير ──
function rpPrint() {
  var data=rpGetData(), fr=document.getElementById('rpFr').value, to=document.getElementById('rpTo').value;
  var rf=data.filter(function(c){return getSt(c)==='refunded';});
  var pn=data.filter(function(c){return getSt(c)==='pending';});
  var up=data.filter(function(c){return getSt(c)==='uploaded';});
  var totalAmt=data.reduce(function(s,c){return s+(c.refundAmount||0);},0);
  var refAmt=rf.reduce(function(s,c){return s+(c.refundAmount||0);},0);
  var rows=data.map(function(c,i){
    return '<tr><td>'+(i+1)+'</td><td>'+esc(c.name)+'</td><td dir="ltr">'+esc(c.phone||c.mobile||'—')+'</td><td>'+(c.refundType==='direct'?'مباشر':'اشتراك')+'</td><td>'+(c.refundAmount||0).toFixed(2)+' ريال</td><td>'+STL[getSt(c)]+'</td><td>'+esc(c.cancelReason||'—')+'</td></tr>';
  }).join('');
  var w=window.open('','_blank');
  w.document.write('<!DOCTYPE html><html dir="rtl"><head><meta charset="utf-8"><title>التقرير المالي</title><style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Tajawal,Arial,sans-serif;padding:24px;color:#1e293b}h1{font-size:20px;font-weight:900;color:#0d9488;margin-bottom:4px}p.sub{font-size:12px;color:#64748b;margin-bottom:20px}.kpi{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:20px}.k{background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:12px 18px;min-width:140px}.k .l{font-size:10px;color:#64748b;margin-bottom:4px}.k .v{font-size:18px;font-weight:900;color:#0d9488}table{width:100%;border-collapse:collapse;font-size:11px}th{background:#0f766e;color:#fff;padding:8px;text-align:right}td{padding:6px 8px;border-bottom:1px solid #e2e8f0}tr:nth-child(even){background:#f8fafc}@media print{body{padding:10px}}</style></head><body>'+
    '<h1>التقرير المالي الشامل — V-SHAPE</h1><p class="sub">الفترة: '+(fr||'بداية')+' إلى '+(to||'نهاية')+' | تاريخ الطباعة: '+new Date().toLocaleDateString('ar-SA')+'</p>'+
    '<div class="kpi"><div class="k"><div class="l">إجمالي الطلبات</div><div class="v">'+data.length+'</div></div><div class="k"><div class="l">تم الاسترداد</div><div class="v" style="color:#16a34a">'+rf.length+'</div></div><div class="k"><div class="l">قيد المراجعة</div><div class="v" style="color:#e97a2a">'+pn.length+'</div></div><div class="k"><div class="l">تم الرفع</div><div class="v" style="color:#64748b">'+up.length+'</div></div><div class="k"><div class="l">إجمالي المبالغ</div><div class="v" style="color:#f59e0b">'+totalAmt.toFixed(2)+' ريال</div></div><div class="k"><div class="l">المستردة فعلياً</div><div class="v" style="color:#16a34a">'+refAmt.toFixed(2)+' ريال</div></div></div>'+
    '<table><thead><tr><th>#</th><th>اسم العميل</th><th>الرقم</th><th>نوع الاسترداد</th><th>المبلغ المسترد</th><th>الحالة</th><th>السبب</th></tr></thead><tbody>'+rows+'</tbody></table></body></html>');
  w.document.close();
  setTimeout(function(){ w.print(); }, 500);
}

// ============================
// قائمة التصدير
// ============================
function togRpMenu() {
  var d = document.getElementById('rpExpDrop');
  d.style.display = d.style.display === 'none' ? 'block' : 'none';
}
document.addEventListener('click', function(e) {
  if (!e.target.closest('#rpExpMenu')) {
    var d = document.getElementById('rpExpDrop');
    if (d) d.style.display = 'none';
  }
});

// ============================
// مساعد: HTML → PDF عبر html2canvas
// يحافظ على العربية والتصميم كما هو
// ============================
function htmlToPDF(element, filename, cb) {
  toast('جاري تجهيز الملف...','i');
  html2canvas(element, {
    scale: 2,
    useCORS: true,
    backgroundColor: getComputedStyle(document.documentElement).getPropertyValue('--bg') || '#ffffff',
    logging: false
  }).then(function(canvas) {
    var { jsPDF } = window.jspdf;
    var imgW = 297; // A4 landscape mm
    var imgH = (canvas.height * imgW) / canvas.width;
    var orientation = imgH > 210 ? 'portrait' : 'landscape';
    if (orientation === 'portrait') imgW = 210;

    var doc = new jsPDF({ orientation: orientation, unit:'mm', format:'a4' });
    var pageH = doc.internal.pageSize.getHeight();
    var pageW = doc.internal.pageSize.getWidth();
    var imgData = canvas.toDataURL('image/png');
    var scaledH = (canvas.height * pageW) / canvas.width;

    // قسّم على صفحات لو الطول أكبر من صفحة واحدة
    var position = 0;
    while (position < scaledH) {
      if (position > 0) doc.addPage();
      doc.addImage(imgData, 'PNG', 0, -position, pageW, scaledH);
      position += pageH;
    }

    doc.save(filename + '.pdf');
    toast('تم تصدير PDF ✓','s');
    if (cb) cb();
  }).catch(function(e) {
    console.error(e);
    toast('خطأ في التصدير','e');
  });
}

// ============================
// بناء HTML جميل للتقرير (يُطبع ويُحوّل لـ PDF)
// ============================
function buildReportHTML(title, subtitle, bodyHTML) {
  var isDk = document.documentElement.className !== 'lt';
  var bg   = isDk ? '#0b1120' : '#f1f5f9';
  var card = isDk ? '#0f172a' : '#ffffff';
  var tx   = isDk ? '#e2e8f0' : '#1e293b';
  var brd  = isDk ? '#1e293b' : '#e2e8f0';

  var div = document.createElement('div');
  div.style.cssText = 'font-family:Tajawal,sans-serif;direction:rtl;background:'+bg+';padding:24px;min-width:900px;color:'+tx;
  div.innerHTML = '<div style="background:linear-gradient(135deg,#0d9488,#0f766e);border-radius:16px;padding:20px 24px;margin-bottom:20px;display:flex;align-items:center;justify-content:space-between">'
    +'<div><div style="font-size:22px;font-weight:900;color:#fff">'+title+'</div>'
    +'<div style="font-size:12px;color:rgba(255,255,255,.7);margin-top:4px">'+subtitle+'</div></div>'
    +'<div style="font-size:28px;font-weight:900;color:rgba(255,255,255,.15)">V-SHAPE</div>'
    +'</div>'
    + bodyHTML;
  document.body.appendChild(div);
  return div;
}

function removeReportEl(el) {
  if (el && el.parentNode) el.parentNode.removeChild(el);
}

// ============================
// كل التقارير — PDF شامل
// ============================
function rpExpAllPDF() {
  var data = rpGetData();
  if (!data.length) { toast('لا توجد بيانات','w'); return; }
  var fr=document.getElementById('rpFr').value, to=document.getElementById('rpTo').value;

  var rf=data.filter(function(c){return getSt(c)==='refunded';});
  var pn=data.filter(function(c){return getSt(c)==='pending';});
  var up=data.filter(function(c){return getSt(c)==='uploaded';});
  var rj=data.filter(function(c){return getSt(c)==='rejected';});
  var totalAmt=data.reduce(function(s,c){return s+(c.refundAmount||0);},0);
  var refAmt=rf.reduce(function(s,c){return s+(c.refundAmount||0);},0);

  // اجمع صور الرسوم البيانية
  var charts = ['rpCStatus','rpCType','rpCDaily','rpCReasons','rpCUsers','rpCDiff'];
  var chartImgs = {};
  charts.forEach(function(id){
    var c=document.getElementById(id);
    if(c){ try{ chartImgs[id]=c.toDataURL('image/png'); }catch(e){} }
  });

  var kpiHTML = '<div style="display:grid;grid-template-columns:repeat(5,1fr);gap:12px;margin-bottom:20px">'
    +kpiCard('إجمالي الطلبات',data.length,'#00d4aa')
    +kpiCard('تم الاسترداد',rf.length,'#22c55e')
    +kpiCard('قيد المراجعة',pn.length,'#e97a2a')
    +kpiCard('تم الرفع',up.length,'#64748b')
    +kpiCard('مرفوض',rj.length,'#dc3545')
    +'</div>'
    +'<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px">'
    +kpiCard('إجمالي المبالغ',en(totalAmt.toFixed(2))+' ريال','#f59e0b')
    +kpiCard('المستردة فعلاً',en(refAmt.toFixed(2))+' ريال','#22c55e')
    +'</div>';

  var chartsHTML = '<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px">'
    + chartBox('توزيع الحالات', chartImgs['rpCStatus'])
    + chartBox('نوع الاسترداد (بالمبلغ)', chartImgs['rpCType'])
    +'</div>'
    + chartBox('الاسترداد اليومي', chartImgs['rpCDaily'], true)
    +'<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px">'
    + chartBox('أسباب الإلغاء', chartImgs['rpCReasons'])
    + chartBox('أداء المستخدمين', chartImgs['rpCUsers'])
    +'</div>'
    + (chartImgs['rpCDiff'] ? chartBox('فروقات التعديلات المالية', chartImgs['rpCDiff'], true) : '');

  var body = kpiHTML + chartsHTML;
  var subtitle = 'الفترة: '+(fr||'بداية')+' إلى '+(to||'نهاية')+' | التاريخ: '+new Date().toLocaleDateString('ar-SA');
  var el = buildReportHTML('التقرير المالي الشامل', subtitle, body);
  setTimeout(function(){
    htmlToPDF(el, 'V-SHAPE_Report_'+(fr||'all'), function(){ removeReportEl(el); });
  }, 300);
}

function kpiCard(label, val, color) {
  return '<div style="background:'+color+'18;border:1px solid '+color+'33;border-radius:12px;padding:14px;text-align:center">'
    +'<div style="font-size:11px;color:'+color+';font-weight:700;margin-bottom:4px">'+label+'</div>'
    +'<div style="font-size:20px;font-weight:900;color:'+color+'">'+val+'</div>'
    +'</div>';
}
function chartBox(title, imgSrc, full) {
  if (!imgSrc) return '';
  return '<div style="background:'+(document.documentElement.className==='lt'?'#fff':'#0f172a')+';border:1px solid '+(document.documentElement.className==='lt'?'#e2e8f0':'#1e293b')+';border-radius:12px;padding:14px;margin-bottom:12px'+(full?';grid-column:span 2':'')+'">'
    +'<div style="font-size:11px;font-weight:700;color:#00d4aa;margin-bottom:8px">'+title+'</div>'
    +'<img src="'+imgSrc+'" style="width:100%;border-radius:8px">'
    +'</div>';
}

// ============================
// تقرير الفروقات — PDF
// ============================
function rpExpDiffPDF() {
  var data = rpGetData();
  var fr=document.getElementById('rpFr').value, to=document.getElementById('rpTo').value;
  var recs=data.filter(function(c){return c.adminEdited&&c.originalRefundAmount!==undefined&&Math.abs((c.originalRefundAmount||0)-(c.refundAmount||0))>0.001;});
  if (!recs.length) { toast('لا توجد فروقات','w'); return; }

  var chartImg = '';
  var cv=document.getElementById('rpCDiff');
  if(cv){ try{chartImg=cv.toDataURL('image/png');}catch(e){} }

  var totalDiff=recs.reduce(function(s,c){return s+Math.abs((c.originalRefundAmount||0)-(c.refundAmount||0));},0);

  var body = '<div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:20px">'
    +kpiCard('عدد السجلات',recs.length,'#f59e0b')
    +kpiCard('إجمالي الفروقات',en(totalDiff.toFixed(2))+' ريال','#dc3545')
    +'</div>'
    +(chartImg?chartBox('مقارنة المبلغ الأصلي وبعد التعديل',chartImg,true):'')
    + tableHTML(
        ['اسم العميل','الرقم','المبلغ الأصلي','بعد التعديل','الفرق','أضاف بواسطة'],
        recs.map(function(c){
          var orig=c.originalRefundAmount||0, curr=c.refundAmount||0, diff=curr-orig;
          return [c.name||'—',c.phone||c.mobile||'—',en(orig.toFixed(2)),en(curr.toFixed(2)),
            '<span style="color:'+(diff<0?'#22c55e':'#dc3545')+';font-weight:700">'+(diff<0?'▼ ':' ▲')+en(Math.abs(diff).toFixed(2))+'</span>',
            c.addedByUsername||'—'];
        }),
        '#b45309'
      );

  var subtitle='الفترة: '+(fr||'بداية')+' إلى '+(to||'نهاية');
  var el=buildReportHTML('تقرير فروقات التعديلات المالية',subtitle,body);
  setTimeout(function(){
    htmlToPDF(el,'discrepancies_'+(fr||'all'),function(){removeReportEl(el);});
  },300);
}

// ============================
// أسباب الإلغاء — PDF
// ============================
function rpExpReasonsPDF() {
  var data=rpGetData();
  var fr=document.getElementById('rpFr').value, to=document.getElementById('rpTo').value;
  if (!data.length){ toast('لا توجد بيانات','w'); return; }

  var chartImg='';
  var cv=document.getElementById('rpCReasons');
  if(cv){try{chartImg=cv.toDataURL('image/png');}catch(e){}}

  var reasons={};
  data.forEach(function(c){var r=(c.cancelReason||'').trim()||'غير محدد';reasons[r]=(reasons[r]||0)+1;});
  var sorted=Object.entries(reasons).sort(function(a,b){return b[1]-a[1];});

  var body=(chartImg?chartBox('أكثر أسباب الإلغاء والاسترداد',chartImg,true):'')
    +tableHTML(
      ['اسم العميل','الرقم','سبب الإلغاء / الاسترداد','الحالة','المبلغ المسترد'],
      data.filter(function(c){return c.cancelReason;}).map(function(c){
        return [c.name||'—',c.phone||c.mobile||'—',c.cancelReason||'—',STL[getSt(c)]||'—',en((c.refundAmount||0).toFixed(2))+' ريال'];
      }),
      '#7c3aed'
    );

  var subtitle='الفترة: '+(fr||'بداية')+' إلى '+(to||'نهاية');
  var el=buildReportHTML('تقرير أسباب الإلغاء والاسترداد',subtitle,body);
  setTimeout(function(){
    htmlToPDF(el,'reasons_'+(fr||'all'),function(){removeReportEl(el);});
  },300);
}

// مساعد: بناء جدول HTML
function tableHTML(headers, rows, hColor) {
  var isDk=document.documentElement.className!=='lt';
  var card=isDk?'#0f172a':'#ffffff', brd=isDk?'#1e293b':'#e2e8f0';
  var html='<div style="background:'+card+';border:1px solid '+brd+';border-radius:12px;overflow:hidden;margin-bottom:16px">'
    +'<table style="width:100%;border-collapse:collapse;font-family:Tajawal,sans-serif;font-size:11px">'
    +'<thead><tr>';
  headers.forEach(function(h){
    html+='<th style="padding:10px 8px;background:'+(hColor||'#0f766e')+';color:#fff;font-weight:700;text-align:center">'+h+'</th>';
  });
  html+='</tr></thead><tbody>';
  rows.forEach(function(row,i){
    var bg=i%2===0?(isDk?'rgba(255,255,255,.02)':'#f8fafc'):(isDk?'rgba(255,255,255,.04)':'#f1f5f9');
    html+='<tr style="background:'+bg+'">';
    row.forEach(function(cell){
      html+='<td style="padding:8px;text-align:center;border-bottom:1px solid '+brd+'">'+cell+'</td>';
    });
    html+='</tr>';
  });
  html+='</tbody></table></div>';
  return html;
}

// ============================
// تقرير الفروقات — Excel
// ============================
function rpExpDiffXLS() {
  var data=rpGetData();
  var fr=document.getElementById('rpFr').value, to=document.getElementById('rpTo').value;
  var recs=data.filter(function(c){return c.adminEdited&&c.originalRefundAmount!==undefined&&Math.abs((c.originalRefundAmount||0)-(c.refundAmount||0))>0.001;});
  if(!recs.length){toast('لا توجد فروقات','w');return;}
  var wb=new ExcelJS.Workbook();
  var ws=wb.addWorksheet('تقرير الفروقات',{views:[{rightToLeft:true}]});
  ws.columns=[{width:24},{width:16},{width:16},{width:16},{width:16},{width:16}];
  ws.mergeCells('A1:F1');var tr=ws.getRow(1);tr.height=36;
  tr.getCell(1).value='تقرير فروقات التعديلات المالية — V-SHAPE';
  tr.getCell(1).font={name:'Tajawal',size:15,bold:true,color:{argb:'FFFFFFFF'}};
  tr.getCell(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF0D9488'}};
  tr.getCell(1).alignment={horizontal:'center',vertical:'middle'};
  ws.mergeCells('A2:F2');var sr=ws.getRow(2);sr.height=18;
  sr.getCell(1).value='الفترة: '+(fr||'بداية')+' إلى '+(to||'نهاية');
  sr.getCell(1).font={name:'Tajawal',size:10,color:{argb:'FF64748B'}};
  sr.getCell(1).alignment={horizontal:'center'};
  ws.addRow([]).height=6;
  var hd=['اسم العميل','الرقم','المبلغ الأصلي','بعد التعديل','الفرق','أضاف بواسطة'];
  var hr=ws.addRow(hd);hr.height=26;
  hr.eachCell(function(cell){cell.font={name:'Tajawal',size:11,bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFB45309'}};cell.alignment={horizontal:'center',vertical:'middle'};});
  recs.forEach(function(c,i){
    var orig=c.originalRefundAmount||0,curr=c.refundAmount||0,diff=curr-orig;
    var row=ws.addRow([c.name||'',c.phone||c.mobile||'',orig,curr,diff,c.addedByUsername||'']);
    row.height=22;
    row.eachCell(function(cell,col){
      cell.font={name:'Tajawal',size:11};cell.alignment={horizontal:'center',vertical:'middle'};
      cell.border={top:{style:'thin',color:{argb:'FFE2E8F0'}},bottom:{style:'thin',color:{argb:'FFE2E8F0'}},left:{style:'thin',color:{argb:'FFE2E8F0'}},right:{style:'thin',color:{argb:'FFE2E8F0'}}};
      if(i%2===1)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF8FAFC'}};
      if(col===3){cell.numFmt='#,##0.00';cell.font={name:'Tajawal',size:11,color:{argb:'FF64748B'},italic:true};}
      if(col===4){cell.numFmt='#,##0.00';cell.font={name:'Tajawal',size:11,bold:true,color:{argb:'FF0F766E'}};}
      if(col===5){cell.numFmt='#,##0.00';cell.font={name:'Tajawal',size:11,bold:true,color:{argb:diff<0?'FF16A34A':'FFDC3545'}};}
    });
  });
  wb.xlsx.writeBuffer().then(function(buf){
    var bl=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var u=URL.createObjectURL(bl),a=document.createElement('a');
    a.href=u;a.download='تقرير_الفروقات_'+(fr||'كل')+'.xlsx';a.click();URL.revokeObjectURL(u);
    toast('تم تصدير Excel الفروقات','s');
  });
}

// ============================
// أسباب الإلغاء — Excel
// ============================
function rpExpReasonsXLS() {
  var data=rpGetData();
  var fr=document.getElementById('rpFr').value,to=document.getElementById('rpTo').value;
  if(!data.length){toast('لا توجد بيانات','w');return;}
  var wb=new ExcelJS.Workbook();
  var ws=wb.addWorksheet('أسباب الإلغاء',{views:[{rightToLeft:true}]});
  ws.columns=[{width:24},{width:16},{width:14},{width:16},{width:22}];
  ws.mergeCells('A1:E1');var tr=ws.getRow(1);tr.height=36;
  tr.getCell(1).value='تقرير أسباب الإلغاء والاسترداد — V-SHAPE';
  tr.getCell(1).font={name:'Tajawal',size:15,bold:true,color:{argb:'FFFFFFFF'}};
  tr.getCell(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF7C3AED'}};
  tr.getCell(1).alignment={horizontal:'center',vertical:'middle'};
  ws.mergeCells('A2:E2');var sr=ws.getRow(2);sr.height=18;
  sr.getCell(1).value='الفترة: '+(fr||'بداية')+' إلى '+(to||'نهاية');
  sr.getCell(1).font={name:'Tajawal',size:10,color:{argb:'FF64748B'}};
  sr.getCell(1).alignment={horizontal:'center'};
  ws.addRow([]).height=6;
  var reasons={};
  data.forEach(function(c){var r=(c.cancelReason||'').trim()||'غير محدد';reasons[r]=(reasons[r]||0)+1;});
  var sorted=Object.entries(reasons).sort(function(a,b){return b[1]-a[1];});
  var sh=ws.addRow(['السبب','العدد','','','']);sh.height=24;
  sh.eachCell(function(cell,col){if(col<=2){cell.font={name:'Tajawal',size:11,bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF7C3AED'}};cell.alignment={horizontal:'center',vertical:'middle'};}});
  sorted.forEach(function(r,i){
    var row=ws.addRow([r[0],r[1],'','','']);row.height=20;
    row.getCell(1).font={name:'Tajawal',size:11};row.getCell(1).alignment={horizontal:'right'};
    row.getCell(2).font={name:'Tajawal',size:11,bold:true,color:{argb:'FF7C3AED'}};row.getCell(2).alignment={horizontal:'center'};
    if(i%2===1){row.getCell(1).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF5F3FF'}};row.getCell(2).fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF5F3FF'}};}
  });
  ws.addRow([]).height=10;
  var dh=ws.addRow(['اسم العميل','الرقم','نوع الاسترداد','المبلغ','السبب']);dh.height=24;
  dh.eachCell(function(cell){cell.font={name:'Tajawal',size:11,bold:true,color:{argb:'FFFFFFFF'}};cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FF5B21B6'}};cell.alignment={horizontal:'center',vertical:'middle'};});
  data.filter(function(c){return c.cancelReason;}).forEach(function(c,i){
    var row=ws.addRow([c.name||'',c.phone||c.mobile||'',c.refundType==='direct'?'مباشر':'اشتراك',c.refundAmount||0,c.cancelReason||'']);row.height=20;
    row.eachCell(function(cell,col){cell.font={name:'Tajawal',size:11};cell.alignment={horizontal:'center',vertical:'middle',wrapText:true};if(i%2===1)cell.fill={type:'pattern',pattern:'solid',fgColor:{argb:'FFF5F3FF'}};if(col===4)cell.numFmt='#,##0.00';});
  });
  wb.xlsx.writeBuffer().then(function(buf){
    var bl=new Blob([buf],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    var u=URL.createObjectURL(bl),a=document.createElement('a');
    a.href=u;a.download='أسباب_الإلغاء_'+(fr||'كل')+'.xlsx';a.click();URL.revokeObjectURL(u);
    toast('تم تصدير Excel الأسباب','s');
  });
}
