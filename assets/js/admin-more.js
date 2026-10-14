const dl=(a,n)=>{const w=XLSX.utils.book_new();XLSX.utils.book_append_sheet(w,XLSX.utils.aoa_to_sheet(a),'Sheet1');XLSX.writeFile(w,n)};
const rd=async f=>{const w=XLSX.read(await f.arrayBuffer());return w.SheetNames.flatMap(n=>XLSX.utils.sheet_to_json(w.Sheets[n],{defval:null}))};
const S=v=>v==null?null:String(v).trim()||null;
const tm=v=>{if(v instanceof Date)return String(v.getHours()).padStart(2,'0')+':'+String(v.getMinutes()).padStart(2,'0');if(typeof v=='number'){let m=Math.round((v-Math.floor(v))*1440)%1440;return String(Math.floor(m/60)).padStart(2,'0')+':'+String(m%60).padStart(2,'0')}return t24(String(v))};
const nA=x=>String(x).replace(/[أإآ]/g,'ا').replace(/ة/g,'ه').trim();
const dayOf=v=>{const s=String(v).trim().toLowerCase(),k=s.length>=3&&Object.keys(EN2AR).find(k=>k.startsWith(s.slice(0,3)));if(k)return EN2AR[k];return DAYS.find(d=>nA(d)==nA(s))||String(v).trim()};
const upBox=(b,title,hint,tplFn,run)=>{b.innerHTML=`<div class="card"><h3>${title}</h3><p>${hint}</p><input type="file" id="f" accept=".xlsx,.xls,.csv"><button class="btn alt" id="go">رفع وتنفيذ</button> <button class="btn" id="tp">تحميل القالب</button></div>`;$('#tp').onclick=tplFn;$('#go').onclick=async()=>{const f=$('#f').files[0];if(!f)return alert('اختر ملفاً');try{alert(await run(await rd(f)))}catch(e){alert('خطأ: '+e.message)}}};
const today=()=>new Date().toLocaleDateString('en-CA',{timeZone:'Africa/Cairo'});


// ===== تعديل بيانات أي طالب (نافذة موحّدة تعمل من أي شاشة) =====
const STUF=[['name','الاسم'],['phone','الهاتف'],['college','الكلية'],['level','الليفل'],['section','السكشن'],['place','الكلاس / المكان']];
const edBtn=c=>`<button type="button" class="chip" data-ed="${esc(c)}" style="padding:3px 12px;margin:0">✎ تعديل</button>`;
async function stuEdit(code){
 const [s]=await api('students?code=eq.'+encodeURIComponent(code)+'&select=code,name,phone,college,level,section,place');if(!s)return alert('الطالب غير موجود');
 document.getElementById('stuModal')?.remove();
 const o=document.createElement('div');o.id='stuModal';o.style.cssText='position:fixed;inset:0;background:#0a1a44b3;z-index:100;display:grid;place-items:center;padding:16px';
 o.innerHTML=`<div style="background:#fff;border-radius:16px;padding:22px;width:100%;max-width:440px;max-height:92vh;overflow:auto;border-top:5px solid #d4a63a"><h3 style="margin:0 0 4px;color:#12307a">تعديل بيانات الطالب</h3><small style="color:#5d6b8f">الكود: ${esc(s.code)} (لا يمكن تعديله)</small>${STUF.map(([k,l])=>`<label>${l}</label><input data-k="${k}" value="${esc(s[k])}">`).join('')}<div style="display:flex;gap:8px"><button class="btn" id="sm_s">حفظ</button><button class="btn alt" id="sm_c" type="button">إلغاء</button></div></div>`;
 document.body.append(o);o.querySelector('#sm_c').onclick=()=>o.remove();o.onclick=e=>{if(e.target===o)o.remove()};
 o.querySelector('#sm_s').onclick=async()=>{const d={};o.querySelectorAll('[data-k]').forEach(i=>d[i.dataset.k]=i.value.trim()||null);if(!d.name)return alert('الاسم مطلوب');
  try{await api('students?code=eq.'+encodeURIComponent(code),'PATCH',d);STUF.forEach(([k])=>document.querySelectorAll(`[data-sf="${code}|${k}"]`).forEach(el=>el.textContent=d[k]??''));o.remove()}catch(e){alert('تعذر الحفظ: '+e.message)}}}
if(!window.__stuH){window.__stuH=1;document.addEventListener('click',e=>{const k=e.target.closest('[data-ed]');if(!k)return;e.stopPropagation();e.preventDefault();stuEdit(k.dataset.ed).catch(x=>alert(x.message))},true)}

// ===== شاشة البحث عن طالب (بالاسم / الكود / الكلاس / الليفل) مع التعديل =====
async function stuSearch(b){
 const [sl,F]=await Promise.all([all('slots?select=id,day,start_time,end_time,place,level&order=id'),all('students?select=level&order=code')]);
 const SM=Object.fromEntries(sl.map(s=>[s.id,s])),cls=[...new Map(sl.map(s=>[nzp(s.place),String(s.place).split(/\s+-\s+/)[0]]))].sort((a,c)=>a[0].localeCompare(c[0],undefined,{numeric:true})),lv=[...new Set(F.map(s=>s.level).filter(Boolean))].sort();
 b.innerHTML=`<div class="card"><h3>بحث عن طالب وتعديل بياناته</h3><label>الاسم أو الكود</label><input id="ss_q" placeholder="جزء من الاسم أو الكود"><label>الكلاس</label><select id="ss_c"><option value="">كل الكلاسات</option>${cls.map(([k,l])=>`<option value="${esc(k)}">${esc(l)}</option>`).join('')}</select><label>الليفل</label><select id="ss_l"><option value="">كل الليفلات</option>${lv.map(v=>`<option>${esc(v)}</option>`).join('')}</select><button class="btn" id="ss_go">بحث</button></div><div id="ss_out"></div>`;
 const chunk=(a,n)=>{const r=[];for(let i=0;i<a.length;i+=n)r.push(a.slice(i,i+n));return r};
 $('#ss_go').onclick=async()=>{try{
  const t=$('#ss_q').value.trim().replace(/[,()*%]/g,' ').trim(),c=$('#ss_c').value,l=$('#ss_l').value,cols='code,name,phone,college,level,section,place';let f='';
  if(t){const e=encodeURIComponent(t);f+=`&or=(name.ilike.*${e}*,code.ilike.*${e}*)`}if(l)f+='&level=eq.'+encodeURIComponent(l);
  $('#ss_out').innerHTML='<div class="msg">جارٍ البحث...</div>';let R=[];
  if(c){const ids=sl.filter(s=>nzp(s.place)==c).map(s=>s.id),E=[];for(const g of chunk(ids,50))E.push(...await all(`enrollments?slot_id=in.(${g.join(',')})&select=student_code&order=student_code,slot_id`));
   for(const g of chunk([...new Set(E.map(e=>e.student_code))],100))R.push(...await api(`students?code=in.(${g.map(encodeURIComponent).join(',')})&select=${cols}${f}`));R.sort((a,d)=>String(a.name).localeCompare(d.name,'ar'))}
  else R=await api(`students?select=${cols}${f}&order=name&limit=300`);
  const G={};for(const g of chunk(R.map(r=>r.code),100))(await api(`enrollments?student_code=in.(${g.map(encodeURIComponent).join(',')})&select=student_code,slot_id`)).forEach(e=>{const s=SM[e.slot_id];if(s)(G[e.student_code]=G[e.student_code]||[]).push(`${s.day} ${fmt(s.start_time)}-${fmt(s.end_time)} | ${String(s.place).split(/\s+-\s+/)[0]}`)});
  const cell=(r,k)=>`<td data-sf="${esc(r.code)}|${k}">${esc(r[k])}</td>`;
  $('#ss_out').innerHTML=`<div class="card wrap"><h3>النتائج (${R.length})${!c&&R.length>=300?' — يظهر أول 300 فقط، ضيّق البحث':''}</h3><table class="table"><tr><th>الكود</th><th>الاسم</th><th>الهاتف</th><th>الكلية</th><th>الليفل</th><th>السكشن</th><th>الكلاس</th><th>الجروبات</th><th>تعديل</th></tr>${R.map(r=>`<tr><td>${esc(r.code)}</td>${['name','phone','college','level','section','place'].map(k=>cell(r,k)).join('')}<td>${(G[r.code]||[]).map(esc).join('<br>')||'-'}</td><td>${edBtn(r.code)}</td></tr>`).join('')||'<tr><td colspan="9">لا توجد نتائج</td></tr>'}</table></div>`}catch(e){$('#ss_out').innerHTML='<div class="msg">خطأ: '+esc(e.message)+'</div>'}};
 $('#ss_q').onkeydown=e=>{if(e.key=='Enter')$('#ss_go').click()}}

// ===== شاشة رفع الطلاب (شؤون الطلاب) + الطلاب بانتظار التوزيع =====
const cairoDay=i=>new Date(i).toLocaleDateString('en-CA',{timeZone:'Africa/Cairo'});
const cairoTime=i=>new Date(i).toLocaleTimeString('en-GB',{timeZone:'Africa/Cairo',hour:'2-digit',minute:'2-digit'});
const dayNm=d=>DAYS[new Date(d+'T12:00:00').getDay()];
// تجميع الطلاب غير الموزعين حسب يوم الرفع (بتوقيت القاهرة) + بيان فترة الرفع
function pendGroup(P){const G={};P.forEach(s=>{const d=cairoDay(s.created_at);(G[d]=G[d]||[]).push(s)});const D=Object.keys(G).sort().reverse();
 const period=!D.length?'':D.length==1?`تاريخ الرفع: ${dayNm(D[0])} ${D[0]}`:`فترة الرفع: من ${dayNm(D[D.length-1])} ${D[D.length-1]} إلى ${dayNm(D[0])} ${D[0]} (${D.length} أيام منفصلة)`;
 return{G,D,period}}
// الطلاب الذين ليس لهم أي جروب = بانتظار التوزيع
async function pendingStu(){const [st,en]=await Promise.all([all('students?select=code,name,phone,college,level,section,place,created_at&order=created_at.desc,code'),all('enrollments?select=student_code,slot_id&order=student_code,slot_id')]);
 const has=new Set(en.map(e=>e.student_code));return{P:st.filter(s=>!has.has(s.code)),en}}
async function notifyAdmins(t,b){const a=await api('staff?role=in.(admin,assistant)&select=code');await notify(a.map(x=>x.code),t,b)}
const nowAr=()=>new Date().toLocaleString('ar-EG',{timeZone:'Africa/Cairo',dateStyle:'full',timeStyle:'short'});
async function studentsUploadTab(b){
 const u=me(),canUp=['affairs','admin'].includes(u.role),canMove=['admin','assistant'].includes(u.role),isAdm=u.role=='admin';
 b.innerHTML=(canUp?'<div id="upc"></div><div id="mnc"></div>':'<div class="msg">رفع الطلاب من صلاحية شؤون الطلاب. هنا تتابع الطلاب بانتظار التوزيع.</div>')+'<div id="pn"></div><div id="ssb"></div>';
 const chunk=(a,n)=>{const r=[];for(let i=0;i<a.length;i+=n)r.push(a.slice(i,i+n));return r},lb=s=>`${s.day} ${fmt(s.start_time)}-${fmt(s.end_time)} | ${String(s.place).split(/\s+-\s+/)[0]} | ${s.level||'-'}`;
 let SL=[],NC={};
 const render=async()=>{const box=$('#pn');box.innerHTML='<div class="msg">جارٍ تحميل الطلاب بانتظار التوزيع...</div>';
  const [{P,en},sl]=await Promise.all([pendingStu(),canMove?all('slots?select=id,day,start_time,end_time,place,level,capacity&order=day,start_time'):Promise.resolve([])]);
  SL=sl;NC={};en.forEach(e=>NC[e.slot_id]=(NC[e.slot_id]||0)+1);const {G,D,period}=pendGroup(P);
  const ctl=canMove&&P.length?`<div class="card noprint"><h3>توزيع الطلاب المحددين على جروب</h3><label>ابحث عن جروب (كلاس / يوم / ليفل)</label><input id="pn_tq" placeholder="مثال: 0010 أو الثلاثاء"><label>الجروب المطلوب</label><select id="pn_t"></select><div><button class="btn alt" id="pn_mv">نقل المحددين إلى الجروب</button>${isAdm?'<button class="btn no" id="pn_rm">حذف المحددين نهائياً</button>':''}</div></div>`:'';
  box.innerHTML=`<div class="card"><h3>طلاب بانتظار التوزيع (${P.length})</h3>${period?`<div class="msg">${period}</div>`:''}<p class="note">يبقى الطالب هنا حتى يُنقل إلى جروب${canMove?'':' بواسطة الأدمن'}. عند توزيعه يختفي من هذه القائمة تلقائياً.</p>${P.length?'<label>بحث داخل القائمة</label><input id="pn_q" placeholder="اسم أو كود">':''}</div>${ctl}`+
   (D.map(d=>{const g=G[d],ts=g.map(s=>s.created_at).sort();return`<div class="dayh">${dayNm(d)} ${d}<small> — ${g.length} طالب • من ${cairoTime(ts[0])} إلى ${cairoTime(ts[ts.length-1])}</small></div><div class="card wrap"><table class="table"><tr>${canMove?`<th><input type="checkbox" class="pkd" data-d="${d}"></th>`:''}<th>الكود</th><th>الاسم</th><th>الهاتف</th><th>الكلية</th><th>الليفل</th><th>السكشن</th><th>الكلاس (من الملف)</th><th>تعديل</th></tr>${g.map(s=>`<tr data-q="${esc((s.name+' '+s.code).toLowerCase())}" data-d="${d}">${canMove?`<td><input type="checkbox" class="pk" value="${esc(s.code)}"></td>`:''}<td>${esc(s.code)}</td><td data-sf="${esc(s.code)}|name">${esc(s.name)}</td><td data-sf="${esc(s.code)}|phone">${esc(s.phone)}</td><td data-sf="${esc(s.code)}|college">${esc(s.college)}</td><td data-sf="${esc(s.code)}|level">${esc(s.level)}</td><td data-sf="${esc(s.code)}|section">${esc(s.section)}</td><td data-sf="${esc(s.code)}|place">${esc(s.place)}</td><td>${edBtn(s.code)}</td></tr>`).join('')}</table></div>`}).join('')||'<div class="msg">لا يوجد طلاب بانتظار التوزيع</div>');
  if($('#pn_q'))$('#pn_q').oninput=()=>{const t=$('#pn_q').value.trim().toLowerCase();document.querySelectorAll('#pn tr[data-q]').forEach(r=>r.hidden=!!t&&!r.dataset.q.includes(t))};
  document.querySelectorAll('.pkd').forEach(c=>c.onchange=()=>document.querySelectorAll(`#pn tr[data-d="${c.dataset.d}"]:not([hidden]) .pk`).forEach(k=>k.checked=c.checked));
  if(!$('#pn_t'))return;
  const fill=()=>{const t=$('#pn_tq').value.trim().toLowerCase();$('#pn_t').innerHTML='<option value="">اختر الجروب...</option>'+SL.filter(s=>!t||lb(s).toLowerCase().includes(t)).map(s=>`<option value="${s.id}">${esc(lb(s))} (${NC[s.id]||0}/${s.capacity||35})</option>`).join('')};$('#pn_tq').oninput=fill;fill();
  const picked=()=>[...document.querySelectorAll('.pk:checked')].map(x=>x.value);
  $('#pn_mv').onclick=async()=>{try{const c=picked(),id=+$('#pn_t').value;if(!c.length)return alert('حدد طلاباً أولاً');if(!id)return alert('اختر الجروب');const s=SL.find(x=>x.id==id),cap=s.capacity||35,cur=await count('enrollments','slot_id','&slot_id=eq.'+id);
    if(cur+c.length>cap){if(!isAdm)return alert(`الجروب لا يتسع (${cur}/${cap})، يحتاج الأدمن لرفع الحد`);if(!confirm(`الجروب سيتجاوز الحد (${cur+c.length}/${cap}). المتابعة؟`))return}
    for(const g of chunk(c,200))await api('enrollments?on_conflict=student_code,slot_id','POST',g.map(x=>({student_code:x,slot_id:id})),'resolution=merge-duplicates,return=minimal');
    await notify(c,'تم توزيعك على جروب',lb(s));alert('تم نقل '+c.length+' طالب إلى: '+lb(s));render()}catch(e){alert(e.message)}};
  if($('#pn_rm'))$('#pn_rm').onclick=async()=>{const c=picked();if(!c.length)return alert('حدد طلاباً أولاً');if(!confirm(`حذف ${c.length} طالب نهائياً من النظام؟ لا يمكن التراجع.`))return;try{for(const g of chunk(c,100))await api('students?code=in.('+g.map(encodeURIComponent).join(',')+')','DELETE',null,'return=minimal');render()}catch(e){alert('تعذر الحذف: '+e.message)}}};
 if(canUp){
  upBox($('#upc'),'رفع طلاب بملف إكسيل (شؤون الطلاب)','الأعمدة: Code, Name, Phone, Faculty, Level, Section, Place. يُضاف الطلاب الجدد إلى قائمة "بانتظار التوزيع" أسفل الصفحة ويصل إشعار للأدمن. الخلية الفارغة للطالب الموجود تُبقي القيمة القديمة، وكلمة المرور = الكود للجدد.',
   ()=>dl([['Code','Name','Phone','Faculty','Level','Section','Place'],['261541','اسم الطالب','01000000000','Computer Science','B2','Section A1','Class 0022']],'students_template.xlsx'),
   async rows=>{const old=Object.fromEntries((await all('students?select=code,name,phone,college,level,section,place&order=code')).map(s=>[s.code,s])),seen=new Set(),out=[];
    rows.forEach(r=>{const c=S(r.Code);if(!c||seen.has(c))return;seen.add(c);const o=old[c]||{};out.push({code:c,name:S(r.Name)||o.name||c,phone:S(r.Phone)||o.phone||null,college:S(r.Faculty)||o.college||null,level:S(r.Level)||o.level||null,section:S(r.Section)||o.section||null,place:S(r.Place)||o.place||null})});
    if(!out.length)throw new Error('لا توجد صفوف بها عمود Code');
    const nw=out.filter(r=>!old[r.code]),ex=out.filter(r=>old[r.code]);
    for(let i=0;i<nw.length;i+=500)await api('students','POST',nw.slice(i,i+500),'return=minimal');
    for(let i=0;i<ex.length;i+=500)await api('students?on_conflict=code','POST',ex.slice(i,i+500),'resolution=merge-duplicates,return=minimal');
    if(nw.length)await notifyAdmins('طلاب جدد من شؤون الطلاب',`تم رفع ${nw.length} طالب جديد بواسطة ${u.name} (${nowAr()}). بانتظار توزيعهم على الجروبات من تبويب "رفع طلاب".`);
    render();return`طلاب جدد: ${nw.length} (أُضيفوا لقائمة بانتظار التوزيع، وكلمة المرور = الكود) — موجودون: ${ex.length} (حُدّثت بياناتهم ولم تتغير كلمة المرور)${nw.length?'\nتم إشعار الأدمن':''}`});
  $('#mnc').innerHTML=`<div class="card"><h3>إضافة طالب يدوياً</h3><label>الكود</label><input id="mu_code"><label>الاسم</label><input id="mu_name"><label>الهاتف</label><input id="mu_phone"><label>الكلية</label><input id="mu_college"><label>الليفل</label><input id="mu_level"><label>السكشن</label><input id="mu_section"><label>الكلاس / المكان</label><input id="mu_place"><button class="btn" id="mu_s">حفظ وإضافة للقائمة</button></div>`;
  $('#mu_s').onclick=async()=>{try{const d=Object.fromEntries(['code','name','phone','college','level','section','place'].map(k=>[k,$('#mu_'+k).value.trim()||null]));if(!d.code||!d.name)return alert('الكود والاسم مطلوبان');
   if((await api('students?code=eq.'+encodeURIComponent(d.code)+'&select=code')).length)return alert('هذا الكود مسجل مسبقاً');
   await api('students','POST',d,'return=minimal');await notifyAdmins('طالب جديد من شؤون الطلاب',`تمت إضافة الطالب ${d.name} (${d.code}) بواسطة ${u.name} (${nowAr()}). بانتظار التوزيع.`);
   ['code','name','phone','college','level','section','place'].forEach(k=>$('#mu_'+k).value='');alert('تمت الإضافة، كلمة المرور = الكود');render()}catch(e){alert('تعذر الحفظ: '+e.message)}}}
 await Promise.all([render(),stuSearch($('#ssb'))])}

// ===== رفع المحاضرين والتوقيتات =====
const LA={code:['code','كود','الكود','lecturercode'],name:['name','الاسم','اسم'],phone:['phone','mobile','tel','الهاتف','رقمالهاتف'],role:['role','الصلاحية'],day:['day','اليوم'],from:['from','start','starttime','من'],to:['to','end','endtime','الى','إلى'],cls:['class','classroom','room','place','كلاس','الكلاس'],level:['level','englishlevel','الليفل']};
const lecRow=r=>{const K={};Object.keys(r).forEach(k=>K[k.toLowerCase().replace(/[\s_.\-]/g,'')]=r[k]);const o={};for(const f in LA){const k=LA[f].find(n=>K[n]!=null&&K[n]!=='');o[f]=k?K[k]:null}return o};
const mins=t=>{const [h,m]=String(t).split(':');return +h*60+ +m};
// مطابقة الجروب: نفس اليوم + نفس الكلاس، ثم التوقيت (مطابق تماماً أو متداخل)
const lecMatch=(sl,d,f,t,p)=>{const same=sl.filter(s=>s.day==d&&nzp(s.place)==nzp(p)),A=mins(f),B=mins(t),df=s=>[Math.abs(mins(fmt(s.start_time))-A),Math.abs(mins(fmt(s.end_time))-B)];
 let hit=same.filter(s=>{const [x,y]=df(s);return x==0&&y==0});if(hit.length)return{hit,how:'exact'};
 hit=same.filter(s=>{const [x,y]=df(s);return x<=15&&y<=15});if(hit.length)return{hit:[hit.sort((u,v)=>df(u)[0]+df(u)[1]-df(v)[0]-df(v)[1])[0]],how:'near'};
 return{hit:[],why:same.length?'الكلاس له مواعيد أخرى في هذا اليوم ('+same.map(s=>fmt(s.start_time)+'-'+fmt(s.end_time)).join('، ')+')':'لا يوجد هذا الكلاس في هذا اليوم'}};
const lecTab=b=>{upBox(b,'رفع / تحديث المحاضرين وتوقيتاتهم','ارفع هذا الملف بعد رفع جدول الطلاب. يربط كل محاضر بالجروبات المطابقة (نفس اليوم والكلاس والوقت) دون المساس بالطلاب. الأعمدة: Code, Name, Phone, Role, Day, From, To, Class, Level. كرّر المحاضر في أكثر من صف لتحديد أكثر من توقيت في اليوم.',
 ()=>dl([['Code','Name','Phone','Role','Day','From','To','Class','Level'],['L100','د. أحمد','01000000000','lecturer','Monday','9:00 AM','11:00 AM','c0012','A1'],['L100','د. أحمد','01000000000','lecturer','Monday','11:30 AM','3:30 PM','c0020','C1']],'lecturers_template.xlsx'),
 async rows0=>{const rows=[...new Map(rows0.map(lecRow).filter(r=>S(r.code)).map(r=>[[S(r.code),r.day,r.from,r.to,r.cls].join('|'),r])).values()];
  if(!rows.length)throw new Error('لم يتم التعرف على الأعمدة. المطلوب: Code, Name, Phone, Role, Day, From, To, Class, Level');
  const st={};rows.forEach(r=>{const c=S(r.code),ro=(S(r.role)||'').toLowerCase();st[c]=st[c]||{name:S(r.name),phone:S(r.phone),role:['admin','assistant','affairs'].includes(ro)?ro:'lecturer'}});
  const have=new Set((await all('staff?select=code')).map(x=>x.code));let nu=0,up=0;
  for(const c in st){const x=st[c];if(have.has(c)){const d={};if(x.name)d.name=x.name;if(x.phone)d.phone=x.phone;if(Object.keys(d).length)await api('staff?code=eq.'+q(c),'PATCH',d);up++}else{await api('staff','POST',{code:c,name:x.name||c,phone:x.phone,role:x.role,roles:[x.role]},'return=minimal');nu++}}
  const sl=await all('slots?select=id,day,start_time,end_time,place,level,lecturer_code&order=id'),mk=$('#mk')?.checked,L={};let ex=0,ov=0,noT=0;const cr=[],sk=[];
  for(const r of rows){if(!r.day||!r.from||!r.to){noT++;continue}
   let d,f,t;try{d=dayOf(r.day);f=tm(r.from);t=tm(r.to)}catch(e){sk.push(S(r.code)+': وقت غير مفهوم ('+r.from+' / '+r.to+')');continue}
   const p=S(r.cls)||'',c=S(r.code),m=lecMatch(sl,d,f,t,p);
   if(m.hit.length){for(const s of m.hit){await api('slots?id=eq.'+s.id,'PATCH',{lecturer_code:c});s.lecturer_code=c}if(m.how=='exact')ex+=m.hit.length;else ov+=m.hit.length;L[c]=(L[c]||0)+m.hit.length}
   else if(mk){const sib=sl.find(s=>nzp(s.place)==nzp(p)),[n]=await api('slots','POST',{day:d,start_time:f,end_time:t,place:sib?sib.place:p,level:S(r.level)||sib?.level||null,lecturer_code:c});sl.push(n);cr.push(d+' '+f+'-'+t+' '+p);L[c]=(L[c]||0)+1}
   else sk.push(`${c} | ${d} ${f}-${t} | كلاس ${p} ← ${m.why}`)}
  for(const c in L)await notify([c],'تحديث جدولك','تم ربطك بـ '+L[c]+' جروب. راجع جدول محاضراتك');
  return`محاضرون جدد: ${nu} — موجودون: ${up}\nربط مطابق للوقت: ${ex} جروب\nربط بوقت مقارب (±15 دقيقة): ${ov} جروب\nجروبات جديدة أُنشئت: ${cr.length}${cr.length?'\n'+cr.join('\n'):''}\nصفوف بلا يوم/وقت (حسابات فقط): ${noT}\nلم تُربط (${sk.length}):${sk.length?'\n'+sk.join('\n'):''}`});
 b.querySelector('#tp').insertAdjacentHTML('beforebegin','<label style="display:block;margin:10px 0"><input type="checkbox" id="mk" checked> إنشاء الموعد تلقائياً إذا لم يوجد جروب بنفس اليوم والكلاس والوقت (مثل كلاس له موعدان بمحاضرين مختلفين)</label>')};

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
  const en=await api('enrollments?slot_id=eq.'+id+'&select=student:students(code,name,level)');$('#ls').innerHTML=en.map(x=>`<div class="stud" onclick="this.classList.toggle('p')" data-c="${x.student.code}"><span><span data-sf="${x.student.code}|name">${esc(x.student.name)}</span><br><small>${x.student.code} • <span data-sf="${x.student.code}|level">${esc(x.student.level)||'-'}</span></small></span>${edBtn(x.student.code)}</div>`).join('')||'لا يوجد طلاب'};
 $('#a').onchange=load;load();
 $('#mv').onclick=async()=>{const c=[...document.querySelectorAll('.stud.p')].map(d=>d.dataset.c);if(!c.length)return alert('حدد طلاباً');const from=$('#a').value,to=$('#t').value;if(from==to)return;
  await api(`enrollments?slot_id=eq.${from}&student_code=in.(${c.join(',')})`,'DELETE',null,'return=minimal');await api('enrollments?on_conflict=student_code,slot_id','POST',c.map(x=>({student_code:x,slot_id:+to})),'resolution=merge-duplicates,return=minimal');await notify(c,'تم نقلك','تم تعديل جدولك');alert('تم النقل');load()}}

// ===== تعديل الغياب (للأدمن، أي يوم) =====
async function attTab(b){
 if(me().role!='admin'){b.innerHTML='<div class="msg">للأدمن فقط</div>';return}
 const sl=await all('slots?order=day,start_time');
 b.innerHTML=`<div class="card"><label>الجروب</label><select id="a">${sl.map(s=>`<option value="${s.id}">${s.day} ${fmt(s.start_time)} | ${esc(s.place)}</option>`).join('')}</select><label>التاريخ</label><input type="date" id="dt" value="${today()}"></div><div id="l"></div>`;
 const load=async()=>{const id=$('#a').value,d=$('#dt').value;const [en,at]=await Promise.all([api('enrollments?slot_id=eq.'+id+'&select=student:students(code,name,level)'),api(`attendance?slot_id=eq.${id}&date=eq.${d}`)]);const on=new Set(at.filter(x=>x.present).map(x=>x.student_code));
  $('#l').innerHTML=en.map(e=>`<div class="stud ${on.has(e.student.code)?'p':''}" data-c="${e.student.code}" onclick="this.classList.toggle('p')"><span><span data-sf="${e.student.code}|name">${esc(e.student.name)}</span><br><small>${e.student.code} • <span data-sf="${e.student.code}|level">${esc(e.student.level)||'-'}</span></small></span>${edBtn(e.student.code)}</div>`).join('')+'<button class="btn alt" id="sv">حفظ</button>';
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
