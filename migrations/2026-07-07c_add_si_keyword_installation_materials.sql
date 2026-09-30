-- "ซื้อวัสดุงาน SI" (id 45) matched no category — add "si" as a keyword for
-- ค่าวัสดุสิ้นเปลืองงานติดตั้ง (SI = installation job site name) and backfill the row.

UPDATE categories SET keywords = keywords || ARRAY['si'] WHERE name = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง';

UPDATE slip_transactions SET category = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง' WHERE id = 45 AND category IS NULL;
