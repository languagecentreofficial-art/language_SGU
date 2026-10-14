const EN2AR={sunday:'الأحد',monday:'الاثنين',tuesday:'الثلاثاء',wednesday:'الأربعاء',thursday:'الخميس',friday:'الجمعة',saturday:'السبت'};
const q=s=>encodeURIComponent(s);
const nzp=s=>{let x=String(s??'').toLowerCase().split(/\s+-\s+/)[0];x=x.replace(/class|كلاس/g,'').replace(/[^a-z0-9\u0600-\u06ff]/g,'').replace(/^c(?=\d)/,'');return x.replace(/^0+(?=.)/,'')};
function t24(s){const m=s.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);let h=+m[1];const a=(m[3]||'').toUpperCase();if(a=='PM'&&h<12)h+=12;if(a=='AM'&&h==12)h=0;if(!a&&h<7)h+=12;return String(h).padStart(2,'0')+':'+m[2]}
async function all(path){let o=[],i=0,r;do{r=await api(path+(path.includes('?')?'&':'?')+'limit=1000&offset='+i);o=o.concat(r);i+=1000}while(r.length==1000);return o}
// قراءة الشيت: يدعم أكثر من بلوك (يوم/وقت/مبنى/كلاس/جدول طلاب) في نفس الملف
function parseRows(rows){const out=[];let c=null,h=false;
 for(const r of rows){const a=String(r[0]??'').trim();
  const p=a.split('—').map(x=>x.trim());
  if(p.length>=3&&(EN2AR[p[0].toLowerCase()]||DAYS.includes(p[0]))&&/\d:\d/.test(p[2])){const [s,e]=p[2].split(/[–-]/);c={day:EN2AR[p[0].toLowerCase()]||p[0],start:t24(s),end:t24(e),building:'',cls:'',level:'',rows:[]};out.push(c);h=false;continue}
  if(!c)continue;
  let m;if(m=a.match(/^Building:\s*([^|]*)/i)){c.building=m[1].trim();continue}
  const j=r.filter(x=>x!=null).join(' ');
  if((!h||!String(r[1]??'').trim())&&/class|كلاس/i.test(j)){const mc=j.match(/(?:class(?:room)?|كلاس)\s*(?:no\.?|number|رقم)?\s*[:：#\-–—.]?\s*([A-Za-z]?\d[A-Za-z0-9]*|[A-Za-z]+[-_]?\d+)/i),ml=j.match(/level\s*[:：\-–—]?\s*([A-Za-z][A-Za-z0-9]*)/i)||j.match(/(?:class(?:room)?|كلاس)[^—–|]*[—–|-]\s*([A-Za-z]\d)\b/i);if(mc){c.cls=mc[1];if(ml)c.level=ml[1];continue}}
  if(/^(name|student name|full name|الاسم|اسم الطالب)$/i.test(a)||String(r[1]??'').trim().toLowerCase()=='code'){h=true;continue}
  if(h){if(r[1]!=null&&String(r[1]).trim())c.rows.push({name:a,code:String(r[1]).trim(),college:r[2]||null,section:r[3]||null,level:r[4]||c.level,lec:r[5]?String(r[5]).trim():null});else if(!a)h=false}}
 return out}
function downloadTemplate(){const ws=XLSX.utils.aoa_to_sheet([
 ['ENGLISH LANGUAGE COURSE'],['صيغة اليوم والوقت: اليوم بالإنجليزي — اسم الفترة — الوقت 24 ساعة أو AM/PM مثل 9:00 AM–11:00 AM'],[],
 ['Monday — First Period — 9:00 AM–11:00 AM'],['Building: Faculty of Pharmacy  |  Floor: First Floor'],['Class 0022 — English Level B2'],[],
 ['Name','Code','Faculty','Department','English Level','The Lecturer'],['اسم الطالب','261541','Computer Science','Section A1','B2','كود المحاضر (اختياري)']]);
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Sheet1');XLSX.writeFile(wb,'template.xlsx')}
async function importTab(b){
 b.innerHTML=`<div class="card"><h3>رفع جدول حضور (إكسيل)</h3><p><b>تنبيه:</b> عند الرفع تُحذف كل جروبات الأيام الموجودة في الملف وتُنشأ من جديد بالبيانات الجديدة. الطالب الجديد يُسجَّل بكلمة مرور = كوده، والموجود تُحدَّث بياناته الأكاديمية فقط دون المساس بكلمة المرور.</p><input type="file" id="f" accept=".xlsx,.xls,.csv"><div id="pv"></div><button class="btn alt" id="go" disabled>استيراد</button> <button class="btn" id="tp">تحميل القالب</button></div><div id="rp"></div>`;
 $('#tp').onclick=downloadTemplate;let blocks=[];
 $('#f').onchange=async e=>{const wb=XLSX.read(await e.target.files[0].arrayBuffer());blocks=wb.SheetNames.flatMap(n=>parseRows(XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:null})));
  const bad=blocks.filter(x=>!x.cls).length;$('#pv').innerHTML=(bad?`<div class="msg" style="background:#fde8ec;color:#8f1d33">⚠ ${bad} جدول بلا رقم كلاس — لن يتم الاستيراد. تأكد أن سطر الكلاس مثل: Class 0010 — English Level B2</div>`:'')+blocks.map(x=>`<div class="msg">${x.day} ${x.start}-${x.end} | كلاس: <b>${esc(x.cls)||'—؟'}</b> | ${esc(x.building)} | ${esc(x.level)} | ${x.rows.length} طالب</div>`).join('')||'<div class="msg">لم يتم التعرف على أي جدول، استخدم القالب</div>';$('#go').disabled=!blocks.length||bad>0};
 $('#go').onclick=async()=>{const days=[...new Set(blocks.map(x=>x.day))],old=[];for(const d of days)old.push(...await api(`slots?day=eq.${q(d)}&select=*`));
  if(!confirm(`سيتم حذف كل الجروبات الموجودة في الأيام: ${days.join('، ')} (${old.length} جروب) وإعادة إنشائها من الملف.\nيُعاد ربط المحاضرين عند تطابق اليوم والوقت والكلاس، وتبقى سجلات الغياب القديمة محفوظة. متابعة؟`))return;
  $('#go').disabled=true;const rep=[];try{
  const lm={},lm2={};old.forEach(o=>{if(!o.lecturer_code)return;lm[[o.day,fmt(o.start_time),fmt(o.end_time),nzp(o.place)].join('|')]=o.lecturer_code;const k2=o.day+'|'+nzp(o.place);lm2[k2]=(k2 in lm2&&lm2[k2]!==o.lecturer_code)?null:o.lecturer_code});
  for(const d of days)await api('slots?day=eq.'+q(d),'DELETE');
  const isn=Object.fromEntries((await all('students?select=code,is_new')).map(r=>[r.code,r.is_new])),mk={};
  for(const x of blocks){const place=x.cls+(x.building?' - '+x.building:''),key=[x.day,x.start,x.end,nzp(x.cls)].join('|');
   let s=mk[key],made=false,lcode=lm[key]||null;const lc=x.rows.find(r=>r.lec)?.lec;if(lc){const [l]=await api('staff?code=eq.'+q(lc)+'&select=code');if(l)lcode=l.code}
   if(!s){[s]=await api('slots','POST',{day:x.day,start_time:x.start,end_time:x.end,place,level:x.level,lecturer_code:lcode});mk[key]=s;made=true}
   const R=[...new Map(x.rows.map(r=>[r.code,{code:r.code,name:r.name,college:r.college,section:r.section,level:r.level,place:x.cls}])).values()],nw=R.filter(r=>!(r.code in isn)).map(r=>({...r,is_new:true})),ex=R.filter(r=>r.code in isn);
   if(nw.length)await api('students','POST',nw,'return=minimal');if(ex.length)await api('students?on_conflict=code','POST',ex,'resolution=merge-duplicates,return=minimal');nw.forEach(r=>isn[r.code]=true);
   await api('enrollments?on_conflict=student_code,slot_id','POST',R.map(r=>({student_code:r.code,slot_id:s.id})),'resolution=merge-duplicates,return=minimal');
   await notify(R.map(r=>r.code),'تحديث الجدول','تم تحديث جدولك: '+x.day+' '+x.start);
   rep.push({x,s,made,nw:nw.length,ex:ex.length,u:R.length,real:await count('enrollments','slot_id','&slot_id=eq.'+s.id)})}
  $('#pv').innerHTML='';$('#rp').innerHTML=`<div class="card wrap"><h3>نتيجة الاستيراد (تم التحقق من قاعدة البيانات)</h3><table class="table"><tr><th>اليوم والوقت</th><th>الكلاس</th><th>الجروب</th><th>جدد</th><th>موجودون</th><th>المسجلون فعلياً الآن</th></tr>${rep.map(r=>`<tr><td>${r.x.day} ${r.x.start}-${r.x.end}</td><td>${esc(r.x.cls)}</td><td>${r.made?'<b>جروب جديد</b>':'جروب قائم #'+r.s.id+(r.s.lecturer_code?' (محاضر '+esc(r.s.lecturer_code)+')':'')}</td><td>${r.nw}</td><td>${r.ex}</td><td class="${r.real<r.u?'bad':'ok'}">${r.real} / ${r.s.capacity||35}${r.real>(r.s.capacity||35)?' • تجاوز الحد':''}${r.real<r.u?' • نقص عن الملف':''}</td></tr>`).join('')}</table></div>`;
  alert('تم استيراد '+rep.reduce((a,r)=>a+r.u,0)+' طالب في '+rep.length+' جروب. راجع جدول النتيجة.');$('#go').disabled=false}catch(e){alert('خطأ: '+e.message);$('#go').disabled=false}}}

async function reportTab(b){
 const [sl,st,at]=await Promise.all([all('slots?order=day'),all('students?select=code,name,college,level,section,place&order=name'),all('attendance?select=student_code,present')]);
 const C=[['code','الكود'],['name','الاسم'],['college','الكلية'],['level','الليفل'],['section','السكشن'],['present','حضور'],['absent','غياب']];
 b.innerHTML=`<div class="card noprint"><h3>تقرير الحضور والغياب</h3><div>${C.map(([k,l])=>`<label style="display:inline-block;margin-inline-end:14px"><input type="checkbox" class="ck" value="${k}" checked> ${l}</label>`).join('')}</div><button class="btn" id="pr">تحميل PDF</button> <button class="btn alt" id="xl">تحميل إكسيل</button></div><div class="printhead"><b>كورس اللغة الإنجليزية</b><img src="assets/logo.png" height="56" alt=""></div><div class="card wrap" id="rt"></div>`;
 const m={};at.forEach(a=>{m[a.student_code]=m[a.student_code]||{p:0,a:0};a.present?m[a.student_code].p++:m[a.student_code].a++});
 const rows=st.map(s=>({...s,present:m[s.code]?.p||0,absent:m[s.code]?.a||0}));
 const draw=()=>{const k=C.filter(c=>document.querySelector(`.ck[value=${c[0]}]`).checked);$('#rt').innerHTML='<table class="table"><tr>'+k.map(c=>`<th>${c[1]}</th>`).join('')+'</tr>'+rows.map(r=>'<tr>'+k.map(c=>`<td>${esc(r[c[0]])}</td>`).join('')+'</tr>').join('')+'</table>';return k};
 document.querySelectorAll('.ck').forEach(c=>c.onchange=draw);draw();
 $('#pr').onclick=()=>print();
 $('#xl').onclick=()=>{const k=draw();const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows.map(r=>Object.fromEntries(k.map(c=>[c[1],r[c[0]]])))),'Report');XLSX.writeFile(wb,'report.xlsx')}}
