-- توحيد بيانات الكلاس: حقل "الكلاس" في بيانات الطالب يُشتق دائماً من تسكينه الفعلي في الجروبات
-- (لا يبقى كلاس قديم لطالب غير مسكَّن، ولا يختلف بين التبويبات)

create or replace function sync_student_place() returns trigger
language plpgsql security definer set search_path = public as $$
declare c text; p text;
begin
  c := case when TG_OP = 'DELETE' then old.student_code else new.student_code end;
  select s.place into p from enrollments e join slots s on s.id = e.slot_id
   where e.student_code = c order by e.slot_id desc limit 1;
  update students set place = p where code = c and place is distinct from p;
  return null;
end $$;
drop trigger if exists t_enr_place on enrollments;
create trigger t_enr_place after insert or delete on enrollments for each row execute function sync_student_place();

-- عند تغيير اسم الكلاس في الجروب
create or replace function sync_place_slot() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if new.place is distinct from old.place then
    update students set place = new.place
     where code in (select student_code from enrollments where slot_id = new.id);
  end if;
  return null;
end $$;
drop trigger if exists t_slot_place on slots;
create trigger t_slot_place after update of place on slots for each row execute function sync_place_slot();

-- مزامنة لمرة واحدة لكل الطلاب الحاليين (غير المسكَّن يصبح بلا كلاس)
update students s set place = (
  select sl.place from enrollments e join slots sl on sl.id = e.slot_id
   where e.student_code = s.code order by e.slot_id desc limit 1)
where s.place is distinct from (
  select sl.place from enrollments e join slots sl on sl.id = e.slot_id
   where e.student_code = s.code order by e.slot_id desc limit 1);
