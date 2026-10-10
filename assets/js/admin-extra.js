const EN2AR={sunday:'الأحد',monday:'الاثنين',tuesday:'الثلاثاء',wednesday:'الأربعاء',thursday:'الخميس',friday:'الجمعة',saturday:'السبت'};
const q=s=>encodeURIComponent(s);
function t24(s){const m=s.match(/(\d{1,2}):(\d{2})\s*(AM|PM)?/i);let h=+m[1];const a=(m[3]||'').toUpperCase();if(a=='PM'&&h<12)h+=12;if(a=='AM'&&h==12)h=0;if(!a&&h<7)h+=12;return String(h).padStart(2,'0')+':'+m[2]}
async function all(path){let o=[],i=0,r;do{r=await api(path+(path.includes('?')?'&':'?')+'limit=1000&offset='+i);o=o.concat(r);i+=1000}while(r.length==1000);return o}
// قراءة الشيت: يدعم أكثر من بلوك (يوم/وقت/مبنى/كلاس/جدول طلاب) في نفس الملف
function parseRows(rows){const out=[];let c=null,h=false;
 for(const r of rows){const a=String(r[0]??'').trim();
  const p=a.split('—').map(x=>x.trim());
  if(p.length>=3&&(EN2AR[p[0].toLowerCase()]||DAYS.includes(p[0]))&&/\d:\d/.test(p[2])){const [s,e]=p[2].split(/[–-]/);c={day:EN2AR[p[0].toLowerCase()]||p[0],start:t24(s),end:t24(e),building:'',cls:'',level:'',rows:[]};out.push(c);h=false;continue}
  if(!c)continue;
  let m;if(m=a.match(/^Building:\s*([^|]*)/i)){c.building=m[1].trim();continue}
  if(m=a.match(/^Class\s*(\S+)\s*—\s*English Level\s*(\S+)/i)){c.cls=m[1];c.level=m[2];continue}
  if(/^(Name|الاسم)$/i.test(a)){h=true;continue}
  if(h){if(r[1]!=null&&String(r[1]).trim())c.rows.push({name:a,code:String(r[1]).trim(),college:r[2]||null,section:r[3]||null,level:r[4]||c.level,lec:r[5]?String(r[5]).trim():null});else if(!a)h=false}}
 return out}
function downloadTemplate(){const ws=XLSX.utils.aoa_to_sheet([
 ['ENGLISH LANGUAGE COURSE'],['صيغة اليوم والوقت: اليوم بالإنجليزي — اسم الفترة — الوقت 24 ساعة أو AM/PM مثل 9:00 AM–11:00 AM'],[],
 ['Monday — First Period — 9:00 AM–11:00 AM'],['Building: Faculty of Pharmacy  |  Floor: First Floor'],['Class 0022 — English Level B2'],[],
 ['Name','Code','Faculty','Department','English Level','The Lecturer'],['اسم الطالب','261541','Computer Science','Section A1','B2','كود المحاضر (اختياري)']]);
 const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Sheet1');XLSX.writeFile(wb,'template.xlsx')}
async function importTab(b){
 b.innerHTML=`<div class="card"><h3>رفع جدول حضور (إكسيل)</h3><p>أي طالب غير موجود يُسجَّل تلقائياً بكلمة مرور = كوده، والموجود تُحدَّث بياناته الأكاديمية فقط (الهاتف وكلمة المرور لا تتغيران). يُنشأ الجروب إن لم يكن موجوداً.</p><input type="file" id="f" accept=".xlsx,.xls,.csv"><div id="pv"></div><button class="btn alt" id="go" disabled>استيراد</button> <button class="btn" id="tp">تحميل القالب</button></div>`;
 $('#tp').onclick=downloadTemplate;let blocks=[];
 $('#f').onchange=async e=>{const wb=XLSX.read(await e.target.files[0].arrayBuffer());blocks=wb.SheetNames.flatMap(n=>parseRows(XLSX.utils.sheet_to_json(wb.Sheets[n],{header:1,defval:null})));
  $('#pv').innerHTML=blocks.map(x=>`<div class="msg">${x.day} ${x.start}-${x.end} | ${esc(x.cls)} | ${esc(x.level)} | ${x.rows.length} طالب</div>`).join('')||'<div class="msg">لم يتم التعرف على أي جدول، استخدم القالب</div>';$('#go').disabled=!blocks.length};
 $('#go').onclick=async()=>{$('#go').disabled=true;let n=0;try{for(const x of blocks){
   const place=x.cls+(x.building?' - '+x.building:'');
   let [s]=await api(`slots?day=eq.${q(x.day)}&start_time=eq.${x.start}&end_time=eq.${x.end}&place=eq.${q(place)}`);
   if(!s)[s]=await api('slots','POST',{day:x.day,start_time:x.start,end_time:x.end,place,level:x.level});
   const lc=x.rows.find(r=>r.lec)?.lec;if(lc){const [l]=await api('staff?code=eq.'+q(lc)+'&select=code');if(l)await api('slots?id=eq.'+s.id,'PATCH',{lecturer_code:l.code})}
   await api('students?on_conflict=code','POST',x.rows.map(r=>({code:r.code,name:r.name,college:r.college,section:r.section,level:r.level,place:x.cls})),'resolution=merge-duplicates,return=minimal');
   await api('enrollments?on_conflict=student_code,slot_id','POST',x.rows.map(r=>({student_code:r.code,slot_id:s.id})),'resolution=merge-duplicates,return=minimal');
   await notify(x.rows.map(r=>r.code),'تحديث الجدول','تم تحديث جدولك: '+x.day+' '+x.start);n+=x.rows.length}
  alert('تم استيراد '+n+' طالب');$('#pv').innerHTML=''}catch(e){alert('خطأ: '+e.message);$('#go').disabled=false}}}
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


// منع القائمة المنسدلة عند الضغط بزر الفأرة الأيمن
    document.addEventListener('contextmenu', function(e) {
        e.preventDefault();
    });

    // منع اختصارات لوحة المفاتيح الخاصة بأدوات المطور
    document.addEventListener('keydown', function(e) {
        // منع F12
        if (e.key === "F12") {
            e.preventDefault();
        }
        // منع Ctrl+Shift+I / Ctrl+Shift+J / Ctrl+Shift+C
        if (e.ctrlKey && e.shiftKey && (e.key === 'I' || e.key === 'i' || e.key === 'J' || e.key === 'j' || e.key === 'C' || e.key === 'c')) {
            e.preventDefault();
        }
        // منع Ctrl+U (عرض مصدر الصفحة)
        if (e.ctrlKey && (e.key === 'U' || e.key === 'u')) {
            e.preventDefault();
        }
    });
    setInterval(function() {
        const startTime = performance.now();
        debugger; // يوقف تنفيذ الكود إذا كانت أدوات المطور مفتوحة
        const endTime = performance.now();
        if (endTime - startTime > 100) {
            // إذا اكتشف فتح F12، يتم توجيه المستخدم أو إعادة تحميل الصفحة
            window.location.reload();
        }
    }, 1000);