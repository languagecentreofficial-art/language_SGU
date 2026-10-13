const dl=(a,n)=>{const w=XLSX.utils.book_new();XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet(a),'Sheet1');XLSX.writeFile(w,n)};
const rd=async f=>{const w=XLSX.read(await f.arrayBuffer());return XLSX.utils.sheet_to_json(w.Sheets[w.SheetNames[0]],{defval:null})};
const S=v=>v==null?null:String(v).trim()||null;
const tm=v=>typeof v=='number'?String(Math.floor(v*24)).padStart(2,'0')+':'+String(Math.round(v*1440)%60).padStart(2,'0'):t24(String(v));
const dayOf=v=>EN2AR[String(v).trim().toLowerCase()]||String(v).trim();
const upBox=(b,title,hint,tplFn,run)=>{b.innerHTML=`<div class="card"><h3>${title}</h3><p>${hint}</p><input type="file" id="f" accept=".xlsx,.xls,.csv"><button class="btn alt" id="go">رفع وتنفيذ</button> <button class="btn" id="tp">تحميل القالب</button></div>`;$('#tp').onclick=tplFn;$('#go').onclick=async()=>{const f=$('#f').files[0];if(!f)return alert('اختر ملفاً');try{alert(await run(await rd(f)))}catch(e){alert('خطأ: '+e.message)}}};
const today=()=>new Date().toLocaleDateString('en-CA',{timeZone:'Africa/Cairo'});

// ===== رفع الطلاب فقط =====
const studentsUploadTab=b=>upBox(b,'رفع / تحديث طلاب','الأعمدة: Code, Name, Phone, Faculty, Level, Section, Place. الخلية الفارغة تُبقي القيمة القديمة، وكلمة المرور = الكود للجدد.',
 ()=>dl([['Code','Name','Phone','Faculty','Level','Section','Place'],['261541','اسم الطالب','01000000000','Computer Science','B2','Section A1','Class 0022']],'students_template.xlsx'),
 async rows=>{const old=Object.fromEntries((await all('students?select=code,name,phone,college,level,section,place,is_new')).map(s=>[s.code,s]));
  const out=rows.filter(r=>S(r.Code)).map(r=>{const c=S(r.Code),o=old[c]||{};return{code:c,name:S(r.Name)||o.name||c,phone:S(r.Phone)||o.phone||null,college:S(r.Faculty)||o.college||null,level:S(r.Level)||o.level||null,section:S(r.Section)||o.section||null,place:S(r.Place)||o.place||null,is_new:o.code?o.is_new:true}});
  const nw=out.filter(r=>!old[r.code]).map(r=>({...r,is_new:true})),ex=out.filter(r=>old[r.code]).map(({is_new,...r})=>r);if(nw.length)await api('students','POST',nw,'return=minimal');if(ex.length)await api('students?on_conflict=code','POST',ex,'resolution=merge-duplicates,return=minimal');return`طلاب جدد: ${nw.length} (كلمة المرور = الكود) — موجودون: ${ex.length} (كلمة المرور لم تتغير)`});

// ===== رفع المحاضرين والتوقيتات =====
const lecTab=b=>{upBox(b,'رفع / تحديث المحاضرين وتوقيتاتهم','يحدّث هذا الرفع حقل المحاضر فقط في الجروب المطابق (نفس اليوم والوقت والكلاس) دون المساس بالطلاب أو أي بيانات أخرى. الأعمدة: Code, Name, Phone, Role, Day, From, To, Class, Level. كرّر المحاضر في أكثر من صف لتحديد أكثر من توقيت في اليوم (مثل 9-11 ليفل A1 كلاس c0012، ثم 11:30-3:30 ليفل C1 كلاس c0020).',
 ()=>dl([['Code','Name','Phone','Role','Day','From','To','Class','Level'],['L100','د. أحمد','01000000000','lecturer','Monday','9:00 AM','11:00 AM','c0012','A1'],['L100','د. أحمد','01000000000','lecturer','Monday','11:30 AM','3:30 PM','c0020','C1']],'lecturers_template.xlsx'),
 async rows=>{rows=rows.filter(r=>S(r.Code));const st={};rows.forEach(r=>st[S(r.Code)]={name:S(r.Name),phone:S(r.Phone),role:['admin','assistant','affairs'].includes(S(r.Role))?S(r.Role):'lecturer'});
  const have=new Set((await api('staff?select=code')).map(x=>x.code));let nu=0,up=0;
  for(const c in st){const x=st[c];if(have.has(c)){const d={};if(x.name)d.name=x.name;if(x.phone)d.phone=x.phone;if(Object.keys(d).length)await api('staff?code=eq.'+q(c),'PATCH',d);up++}else{await api('staff','POST',{code:c,name:x.name||c,phone:x.phone,role:x.role,roles:[x.role]},'return=minimal');nu++}}
  const nz=nzp,mk=$('#mk')?.checked;let m=0;const cr=[],sk=[];
  for(const r of rows){if(!r.Day||!r.From||!r.To)continue;const d=dayOf(r.Day),f=tm(r.From),t=tm(r.To),p=S(r.Class)||'',L=S(r.Level);
   const cand=(await api(`slots?day=eq.${q(d)}&start_time=eq.${f}&end_time=eq.${t}&select=id,place,level`)).filter(s=>nz(s.place)==nz(p));
   if(cand.length){for(const s of cand){await api('slots?id=eq.'+s.id,'PATCH',{lecturer_code:S(r.Code)})}m+=cand.length}
   else if(mk){await api('slots','POST',{day:d,start_time:f,end_time:t,place:p,level:L,lecturer_code:S(r.Code)});cr.push(d+' '+f+' '+p)}else sk.push(d+' '+f+' '+p)}
  return`محاضرون جدد: ${nu} — موجودون (لم تتغير صلاحياتهم): ${up}\nربط بجروبات قائمة بطلابها: ${m}\nجروبات جديدة أُنشئت: ${cr.length}${cr.length?'\n'+cr.join('\n'):''}\nلم تُطابق أي جروب قائم وتم تخطيها: ${sk.length}${sk.length?'\n'+sk.join('\n'):''}`});b.querySelector('#tp').insertAdjacentHTML('beforebegin','<label style="display:block;margin:10px 0"><input type="checkbox" id="mk"> إنشاء جروب جديد عند عدم وجود جروب مطابق (الافتراضي: تحديث المحاضر فقط)</label>')};

// ===== الدرجات =====
async function gradesTab(b){
 const ex=await api('exams?order=id');
 b.innerHTML=`<div class="card"><h3>إنشاء اختبار / تقييم</h3><label>الاسم</label><input id="en"><label>النوع</label><select id="ek"><option value="quiz">كويز</option><option value="midterm">اختبار فصلي</option><option value="final">اختبار نهاية العام</option><option value="weekly">تقييم أسبوعي</option><option value="monthly">تقييم شهري</option></select><label>الدرجة العظمى</label><input id="em" type="number" value="10"><button class="btn" id="ea">إضافة</button></div>
 <div class="card"><h3>الاختبارات الحالية</h3>${ex.map(e=>`<span class="pill">${esc(e.name)} (${e.max_mark}) · <a href="#" data-p="${e.id}" data-v="${e.published?0:1}">${e.published?'منشور — إخفاء':'مخفي — نشر'}</a> <a href="#" data-d="${e.id}" class="bad">✕</a></span> `).join('')||'لا يوجد'}</div><div id="up"></div>`;
 $('#ea').onclick=async()=>{await api('exams','POST',{name:$('#en').value,kind:$('#ek').value,max_mark:+$('#em').value},'return=minimal');gradesTab(b)};
 b.querySelectorAll('[data-p]').forEach(a=>a.onclick=async e=>{e.preventDefault();const v=a.dataset.v=='1';await api('exams?id=eq.'+a.dataset.p,'PATCH',{published:v});if(v){const x=ex.find(q=>q.id==a.dataset.p);notify((await all('students?select=code')).map(q=>q.code),'تم نشر درجات',x.name)}gradesTab(b)});
 b.querySelectorAll('[data-d]').forEach(a=>a.onclick=async e=>{e.preventDefault();if(confirm('حذف الاختبار ودرجاته؟')){await api('exams?id=eq.'+a.dataset.d,'DELETE',null,'return=minimal');gradesTab(b)}});
 upBox($('#up'),'رفع الدرجات','الصف الأول: Code ثم أسماء الاختبارات كأعمدة، ويمكن كتابة الدرجة العظمى بين قوسين مثل Quiz 1 (10). الاختبار الجديد يُنشأ تلقائياً.',
  ()=>dl([['Code','Name','Quiz 1 (10)','Midterm (30)'],['261541','اسم الطالب',8,25]],'grades_template.xlsx'),
  async rows=>{const cols=Object.keys(rows[0]||{}).filter(k=>!['Code','Name'].includes(k));const have=Object.fromEntries((await api('exams')).map(e=>[e.name,e]));
   for(const c of cols){const m=c.match(/^(.*?)\s*\((\d+(?:\.\d+)?)\)\s*$/),nm=m?m[1]:c;if(!have[nm])[have[nm]]=await api('exams','POST',{name:nm,max_mark:m?+m[2]:100})}
   const codes=new Set((await all('students?select=code')).map(s=>s.code));let g=[],skip=0;
   rows.forEach(r=>{const c=S(r.Code);if(!codes.has(c)){skip++;return}cols.forEach(k=>{if(r[k]==null||r[k]==='')return;const nm=(k.match(/^(.*?)\s*\(/)||[0,k])[1];g.push({exam_id:have[nm].id,student_code:c,mark:+r[k]})})});
   await api('grades?on_conflict=exam_id,student_code','POST',g,'resolution=merge-duplicates,return=minimal');
   return`تم رفع ${g.length} درجة، وتخطي ${skip} كود غير مسجل`})}

// ===== شروط الغياب والإنذارات =====
async function rulesTab(b){
 const R=Object.fromEntries((await api('rules')).map(r=>[r.key,r.value])),mo=today().slice(0,7);
 b.innerHTML=`<div class="card"><h3>شروط احتساب الغياب</h3><label>أقل عدد أيام حضور في الشهر قبل الإنذار</label><input id="r1" type="number" value="${R.min_days_per_month??4}"><label>درجة اليوم الواحد</label><input id="r2" type="number" step="0.1" value="${R.mark_per_day??0.5}"><button class="btn" id="sv">حفظ الشروط</button></div><div class="card"><h3>الإنذارات</h3><label>الشهر</label><input id="mo" type="month" value="${mo}"><div id="w" class="wrap"></div></div>`;
 $('#sv').onclick=async()=>{await api('rules?on_conflict=key','POST',[{key:'min_days_per_month',value:+$('#r1').value},{key:'mark_per_day',value:+$('#r2').value}],'resolution=merge-duplicates,return=minimal');alert('تم');rulesTab(b)};
 const show=async()=>{const m=$('#mo').value,[st,at]=await Promise.all([all('students?select=code,name,level&order=name'),all(`attendance?date=gte.${m}-01&date=lte.${m}-31&select=student_code,present`)]);
  const c={};at.forEach(a=>{c[a.student_code]=c[a.student_code]||{p:0,a:0};a.present?c[a.student_code].p++:c[a.student_code].a++});
  const min=R.min_days_per_month??4,mk=R.mark_per_day??0.5,rows=st.filter(s=>c[s.code]).map(s=>({...s,p:c[s.code].p,a:c[s.code].a,w:c[s.code].p<min}));
  $('#w').innerHTML=`<table class="table"><tr><th>الكود</th><th>الاسم</th><th>حضور</th><th>غياب</th><th>درجة الحضور</th><th></th></tr>${rows.map(r=>`<tr><td>${r.code}</td><td>${esc(r.name)}</td><td>${r.p}</td><td>${r.a}</td><td>${(r.p*mk).toFixed(1)}</td><td>${r.w?'<span class="bad">⚠ إنذار</span>':''}</td></tr>`).join('')}</table><button class="btn no" id="nw">إرسال الإنذار للمنذَرين (${rows.filter(r=>r.w).length})</button>`;
  $('#nw').onclick=async()=>{await notify(rows.filter(r=>r.w).map(r=>r.code),'إنذار حضور','حضورك أقل من '+min+' أيام في شهر '+m);alert('تم الإرسال')}};
 $('#mo').onchange=show;show()}

// ===== نقل الطلاب وتعديل الجداول =====
async function moveTab(b){
 const sl=await all('slots?select=*,lec:staff(name)&order=day,start_time'),lec=await api('staff?role=eq.lecturer&select=code,name');
 const lb=s=>`${s.day} ${fmt(s.start_time)}-${fmt(s.end_time)} | ${esc(s.place)} | ${esc(s.level)}`,opt=sel=>sl.map(s=>`<option value="${s.id}" ${s.id==sel?'selected':''}>${lb(s)}</option>`).join('');
 b.innerHTML=`<div class="card"><label>الجروب</label><select id="a">${opt()}</select><div id="ed"></div></div><div class="card"><h3>الطلاب</h3><div id="ls"></div><label>نقل المحددين إلى</label><select id="t">${opt()}</select><button class="btn" id="mv">نقل</button></div>`;
 const load=async()=>{const id=+$('#a').value,s=sl.find(x=>x.id==id);if(!s)return;
  $('#ed').innerHTML=`<label>اليوم</label><select id="d">${DAYS.map(d=>`<option ${d==s.day?'selected':''}>${d}</option>`).join('')}</select><label>من</label><input type="time" id="f" value="${fmt(s.start_time)}"><label>إلى</label><input type="time" id="e" value="${fmt(s.end_time)}"><label>المكان</label><input id="p" value="${esc(s.place)}"><label>الليفل</label><input id="l" value="${esc(s.level)}"><label>المحاضر</label><select id="lc"><option value="">-</option>${lec.map(x=>`<option value="${x.code}" ${x.code==s.lecturer_code?'selected':''}>${esc(x.name)}</option>`).join('')}</select><button class="btn alt" id="sv">حفظ التعديل</button> <button class="btn no" id="rm">حذف الجروب</button>`;
  $('#sv').onclick=async()=>{await api('slots?id=eq.'+id,'PATCH',{day:$('#d').value,start_time:$('#f').value,end_time:$('#e').value,place:$('#p').value,level:$('#l').value,lecturer_code:$('#lc').value||null});const en=await api('enrollments?slot_id=eq.'+id+'&select=student_code');notify(en.map(x=>x.student_code),'تغيير في الجدول','تم تعديل موعد أو مكان جروبك');moveTab(b)};
  $('#rm').onclick=async()=>{if(confirm('حذف الجروب؟')){await api('slots?id=eq.'+id,'DELETE',null,'return=minimal');moveTab(b)}};
  const en=await api('enrollments?slot_id=eq.'+id+'&select=student:students(code,name)');$('#ls').innerHTML=en.map(x=>`<div class="stud" onclick="this.classList.toggle('p')" data-c="${x.student.code}">${esc(x.student.name)}<small>${x.student.code}</small></div>`).join('')||'لا يوجد طلاب'};
 $('#a').onchange=load;load();
 $('#mv').onclick=async()=>{const c=[...document.querySelectorAll('.stud.p')].map(d=>d.dataset.c);if(!c.length)return alert('حدد طلاباً');const from=$('#a').value,to=$('#t').value;if(from==to)return;
  await api(`enrollments?slot_id=eq.${from}&student_code=in.(${c.join(',')})`,'DELETE',null,'return=minimal');await api('enrollments?on_conflict=student_code,slot_id','POST',c.map(x=>({student_code:x,slot_id:+to})),'resolution=merge-duplicates,return=minimal');await notify(c,'تم نقلك','تم تعديل جدولك');alert('تم النقل');load()}}

// ===== تعديل الغياب (للأدمن، أي يوم) =====
async function attTab(b){
 if(me().role!='admin'){b.innerHTML='<div class="msg">للأدمن فقط</div>';return}
 const sl=await all('slots?order=day,start_time');
 b.innerHTML=`<div class="card"><label>الجروب</label><select id="a">${sl.map(s=>`<option value="${s.id}">${s.day} ${fmt(s.start_time)} | ${esc(s.place)}</option>`).join('')}</select><label>التاريخ</label><input type="date" id="dt" value="${today()}"></div><div id="l"></div>`;
 const load=async()=>{const id=$('#a').value,d=$('#dt').value;const [en,at]=await Promise.all([api('enrollments?slot_id=eq.'+id+'&select=student:students(code,name)'),api(`attendance?slot_id=eq.${id}&date=eq.${d}`)]);const on=new Set(at.filter(x=>x.present).map(x=>x.student_code));
  $('#l').innerHTML=en.map(e=>`<div class="stud ${on.has(e.student.code)?'p':''}" data-c="${e.student.code}" onclick="this.classList.toggle('p')">${esc(e.student.name)}<small>${e.student.code}</small></div>`).join('')+'<button class="btn alt" id="sv">حفظ</button>';
  $('#sv').onclick=async()=>{await api('attendance?on_conflict=student_code,slot_id,date','POST',[...document.querySelectorAll('.stud')].map(x=>({student_code:x.dataset.c,slot_id:+id,date:d,present:x.classList.contains('p'),marked_by:me().code})),'resolution=merge-duplicates,return=minimal');alert('تم')}};
 $('#a').onchange=$('#dt').onchange=load;load()}

// ===== الصلاحيات: اختيار الشخص ثم منحه أي عدد من الصفات والأقسام =====
async function permTab(b){
 if(me().role!='admin'){b.innerHTML='<div class="msg">للأدمن فقط</div>';return}
 const as=await api('staff?select=code,name,role,roles,perms&order=name'),R=[['admin','أدمن'],['assistant','مساعد أدمن'],['affairs','شؤون الطلاب'],['lecturer','محاضر']],PRI=R.map(x=>x[0]),T=TABNAMES.filter(t=>!['الصلاحيات','حسابات'].includes(t));
 b.innerHTML=`<div class="card"><h3>اختيار الشخص</h3><label>الحساب</label><select id="ps"><option value="">اختر الشخص...</option>${as.map(a=>`<option value="${a.code}">${esc(a.name)} (${a.code})</option>`).join('')}</select></div><div id="pb"></div>`;
 $('#ps').onchange=()=>{const a=as.find(x=>x.code==$('#ps').value);if(!a){$('#pb').innerHTML='';return}const rs=a.roles||[a.role];
  $('#pb').innerHTML=`<div class="card"><h3>صلاحيات ${esc(a.name)}</h3><label>الصفات — اضغط لاختيار واحدة أو أكثر</label><div id="rc">${R.map(([k,l])=>`<button type="button" class="chip ${rs.includes(k)?'on':''}" data-k="${k}">${l}</button>`).join('')}</div><label>الأقسام المسموح له برؤيتها — اختر ما تشاء</label><div id="tc">${T.map(t=>`<button type="button" class="chip ${a.perms?.[t]?'on':''}" data-k="${t}">${t}</button>`).join('')}</div><button type="button" class="btn alt" id="sa">تحديد كل الأقسام</button><button type="button" class="btn alt" id="sn">مسح الأقسام</button><button type="button" class="btn" id="sv">حفظ الصلاحيات</button></div>`;
  $('#pb').onclick=e=>{const c=e.target.closest('.chip');if(c)c.classList.toggle('on')};
  $('#sa').onclick=()=>document.querySelectorAll('#tc .chip').forEach(c=>c.classList.add('on'));$('#sn').onclick=()=>document.querySelectorAll('#tc .chip').forEach(c=>c.classList.remove('on'));
  $('#sv').onclick=async()=>{const roles=[...document.querySelectorAll('#rc .chip.on')].map(x=>x.dataset.k),p={};document.querySelectorAll('#tc .chip.on').forEach(x=>p[x.dataset.k]=true);if(!roles.length)return alert('اختر صفة واحدة على الأقل');
   await api('staff?code=eq.'+a.code,'PATCH',{role:PRI.find(x=>roles.includes(x)),roles,perms:p});a.roles=roles;a.role=PRI.find(x=>roles.includes(x));a.perms=p;alert('تم حفظ الصلاحيات')}}}
