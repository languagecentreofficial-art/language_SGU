-- حماية كلمات المرور: لا يمكن لأي رفع أو تحديث عبر API أن يغيّر password_hash
-- (التغيير الوحيد المسموح هو من خلال دالة change_password)
revoke update on students from anon;
grant update(code,name,phone,college,level,section,place,is_new) on students to anon;
revoke update on staff from anon;
grant update(code,name,phone,role,roles,perms) on staff to anon;

-- (اختياري) عرض الجروبات الفارغة التي أنشأها رفع المحاضرين سابقاً قبل الإصلاح، للمراجعة قبل حذفها يدوياً:
-- select s.id,s.day,s.start_time,s.place,s.level,s.lecturer_code from slots s
--  where not exists(select 1 from enrollments e where e.slot_id=s.id) order by s.day,s.start_time;
