-- شغّله مرة واحدة في Supabase SQL Editor لإنشاء حساب الأدمن
-- الدخول: الكود admin / كلمة المرور admin123  (غيّرها فور الدخول)
insert into staff(code,name,phone,role) values('admin','مدير النظام','','admin')
on conflict(code) do update set role='admin';
update staff set password_hash=crypt('admin123',gen_salt('bf')) where code='admin';
