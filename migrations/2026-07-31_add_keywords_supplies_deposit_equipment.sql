-- ids 91, 92, 96 matched no category — add keywords and backfill.
-- "ซื้อของ" (generic purchase reimbursement) -> installation materials, matches existing manual pattern (ids 24,26,32,38,41,47,66,86).
-- "poe" -> equipment/goods (POE extender is network equipment, cost-of-goods).
-- "วางประกัน" (performance/contract deposit) -> deposit category, matches existing manual pattern (ids 35, 89).

UPDATE categories SET keywords = keywords || ARRAY['ซื้อของ'] WHERE name = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง';
UPDATE categories SET keywords = keywords || ARRAY['poe'] WHERE name = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)';
UPDATE categories SET keywords = keywords || ARRAY['วางประกัน'] WHERE name = 'เงินประกันผลงาน/สัญญา';

UPDATE slip_transactions SET category = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง' WHERE id = 91 AND category IS NULL;
UPDATE slip_transactions SET category = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)' WHERE id = 92 AND category IS NULL;
UPDATE slip_transactions SET category = 'เงินประกันผลงาน/สัญญา' WHERE id = 96 AND category IS NULL;
