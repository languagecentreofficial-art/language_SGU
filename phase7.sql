-- يسمح بحذف/إعادة إنشاء الجروبات دون فقدان سجلات الغياب وطلبات التغيير القديمة
do $$ declare r record; begin
 for r in select conrelid::regclass t, conname from pg_constraint
   where contype='f' and confrelid='slots'::regclass and conrelid in ('attendance'::regclass,'change_requests'::regclass)
 loop execute format('alter table %s drop constraint %I', r.t, r.conname); end loop; end $$;
alter table attendance      add foreign key (slot_id)   references slots(id) on delete set null;
alter table change_requests add foreign key (from_slot) references slots(id) on delete set null;
alter table change_requests add foreign key (to_slot)   references slots(id) on delete set null;
