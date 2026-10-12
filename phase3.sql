-- شغّله مرة واحدة في Supabase SQL Editor
alter table staff drop constraint if exists staff_role_check;
alter table staff add constraint staff_role_check check (role in ('lecturer','admin','assistant','affairs'));
alter table students add column if not exists is_new boolean default false;
alter table slots add column if not exists capacity int;           -- null = الحد الافتراضي 35
insert into rules values('default_cap',35) on conflict do nothing;
grant select(is_new) on students to anon;

create table if not exists blocks(
  id bigserial primary key, student_code text references students(code) on update cascade on delete cascade,
  reason text not null, active boolean default true, created_by text, created_at timestamptz default now(),
  lifted_by text, lifted_at timestamptz, lift_note text);
alter table blocks enable row level security;
drop policy if exists p_blocks on blocks;
create policy p_blocks on blocks for all to anon using (true) with check (true);
grant select,insert,update,delete on blocks to anon; grant usage on all sequences in schema public to anon;

-- حد أقصى لعدد طلاب الكلاس (35 افتراضياً، والأدمن يغيّره لكل كلاس)
create or replace function check_cap() returns trigger language plpgsql as $$
declare cap int; n int; begin
 if exists(select 1 from enrollments where slot_id=new.slot_id and student_code=new.student_code) then return new; end if;
 select coalesce(s.capacity,(select value::int from rules where key='default_cap'),35) into cap from slots s where s.id=new.slot_id;
 select count(*) into n from enrollments where slot_id=new.slot_id;
 if n>=cap then raise exception 'الكلاس ممتلئ، الحد الأقصى % طالب',cap; end if;
 return new; end $$;
drop trigger if exists enr_cap on enrollments;
create trigger enr_cap before insert on enrollments for each row execute function check_cap();

-- المحجوب لا يستطيع تقديم طلب تغيير
create or replace function chk_block() returns trigger language plpgsql as $$
begin if exists(select 1 from blocks where student_code=new.student_code and active) then raise exception 'أنت محجوب عن تقديم طلبات التغيير'; end if; return new; end $$;
drop trigger if exists cr_block on change_requests;
create trigger cr_block before insert on change_requests for each row execute function chk_block();
