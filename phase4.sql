-- درجات الاختبارات مخفية عن الطلاب حتى ينشرها الأدمن
alter table exams add column if not exists published boolean default false;
