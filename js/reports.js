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
