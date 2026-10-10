-- شغّل هذا الملف كاملاً في Supabase > SQL Editor
create extension if not exists pgcrypto;

create table staff(
  code text primary key, name text not null, phone text,
  role text not null check (role in ('lecturer','admin','assistant')),
  perms jsonb default '{}', password_hash text, created_at timestamptz default now());

create table students(
  code text primary key, name text not null, phone text, college text,
  level text, section text, place text,
  password_hash text, created_at timestamptz default now());

-- الجروب = يوم + وقت من/إلى + مكان + ليفل + محاضر (يمكن أكثر من فترة في نفس اليوم)
create table slots(
  id bigserial primary key, day text not null, start_time time not null, end_time time not null,
  place text not null, level text, lecturer_code text references staff(code) on update cascade);

create table enrollments(
  student_code text references students(code) on update cascade on delete cascade,
  slot_id bigint references slots(id) on delete cascade,
  primary key(student_code, slot_id));

create table attendance(
  id bigserial primary key,
  student_code text references students(code) on update cascade on delete cascade,
  slot_id bigint references slots(id) on delete cascade,
  date date not null, present boolean not null default false, marked_by text,
  unique(student_code, slot_id, date));

create table exams(id bigserial primary key, name text not null, kind text, max_mark numeric not null);
create table grades(
  exam_id bigint references exams(id) on delete cascade,
  student_code text references students(code) on update cascade on delete cascade,
  mark numeric, primary key(exam_id, student_code));

create table complaints(id bigserial primary key, student_code text references students(code),
  body text not null, reply text, status text default 'open', created_at timestamptz default now());
create table change_requests(id bigserial primary key, student_code text references students(code),
  from_slot bigint references slots(id), to_slot bigint references slots(id),
  reason text not null, status text default 'pending', reject_reason text, created_at timestamptz default now());
create table announcements(id bigserial primary key, title text not null, body text,
  created_by text, created_at timestamptz default now());
create table notifications(id bigserial primary key, target_code text not null,
  title text, body text, read boolean default false, created_at timestamptz default now());
create table rules(key text primary key, value numeric);  -- مثال: min_days_per_month=4 , mark_per_day=0.5
insert into rules values('min_days_per_month',4),('mark_per_day',0.5);

-- كلمة المرور الافتراضية = الكود (مشفّرة)
create function set_default_pass() returns trigger language plpgsql as $$
begin if new.password_hash is null then new.password_hash := crypt(new.code, gen_salt('bf')); end if; return new; end $$;
create trigger t1 before insert on staff for each row execute function set_default_pass();
create trigger t2 before insert on students for each row execute function set_default_pass();

create function login(p_code text, p_pass text) returns json language sql security definer as $$
 select coalesce(
  (select json_build_object('code',code,'name',name,'role',role,'phone',phone) from staff where code=p_code and password_hash=crypt(p_pass,password_hash)),
  (select json_build_object('code',code,'name',name,'role','student','phone',phone) from students where code=p_code and password_hash=crypt(p_pass,password_hash)))
$$;

create function change_password(p_code text, p_old text, p_new text) returns boolean language plpgsql security definer as $$
declare n int; begin
 update students set password_hash=crypt(p_new,gen_salt('bf')) where code=p_code and password_hash=crypt(p_old,password_hash);
 get diagnostics n = row_count;
 if n=0 then update staff set password_hash=crypt(p_new,gen_salt('bf')) where code=p_code and password_hash=crypt(p_old,password_hash); get diagnostics n = row_count; end if;
 return n>0; end $$;

-- صلاحيات anon (المرحلة 1). كلمات المرور لا تُقرأ أبداً.
alter table staff enable row level security; alter table students enable row level security;
do $$ declare t text; begin
 foreach t in array array['staff','students','slots','enrollments','attendance','exams','grades','complaints','change_requests','announcements','notifications','rules'] loop
  execute format('alter table %I enable row level security',t);
  execute format('create policy "p_%s" on %I for all to anon using (true) with check (true)',t,t);
  execute format('grant select,insert,update,delete on %I to anon',t);
 end loop; end $$;
revoke select on staff, students from anon;
grant select(code,name,phone,role,perms,created_at) on staff to anon;
grant select(code,name,phone,college,level,section,place,created_at) on students to anon;
grant usage on all sequences in schema public to anon;
grant execute on function login(text,text), change_password(text,text,text) to anon;
