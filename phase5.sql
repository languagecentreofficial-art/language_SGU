-- 1) رفع الإكسيل لا يتوقف عند تجاوز 35 (التنبيه يظهر في الواجهة)
drop trigger if exists enr_cap on enrollments;
-- 2) أكثر من صفة للشخص الواحد
alter table staff add column if not exists roles text[];
update staff set roles=array[role] where roles is null;
grant select(roles) on staff to anon;
create or replace function login(p_code text, p_pass text) returns json language sql security definer as $$
 select coalesce(
  (select json_build_object('code',code,'name',name,'role',role,'roles',coalesce(roles,array[role]),'phone',phone) from staff where code=p_code and password_hash=crypt(p_pass,password_hash)),
  (select json_build_object('code',code,'name',name,'role','student','roles',array['student'],'phone',phone) from students where code=p_code and password_hash=crypt(p_pass,password_hash)))
$$;
