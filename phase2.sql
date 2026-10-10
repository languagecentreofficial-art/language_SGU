-- شغّله مرة واحدة: يمنع تعديل غياب يوم سابق إلا للأدمن
create or replace function lock_att() returns trigger language plpgsql as $$
begin
 if new.date < (now() at time zone 'Africa/Cairo')::date
    and coalesce((select role from staff where code=new.marked_by),'') <> 'admin'
 then raise exception 'لا يمكن تعديل غياب يوم سابق إلا بواسطة الأدمن'; end if;
 return new; end $$;
drop trigger if exists att_lock on attendance;
create trigger att_lock before insert or update on attendance for each row execute function lock_att();
