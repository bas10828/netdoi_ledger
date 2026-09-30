-- บันได is reusable equipment technicians keep, not a per-job consumable —
-- move it from ค่าวัสดุสิ้นเปลืองงานติดตั้ง to ค่าเครื่องมือ/อุปกรณ์ช่าง.
UPDATE categories SET keywords = array_remove(keywords, 'บันได')
WHERE name = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง';

UPDATE categories SET keywords = array_append(keywords, 'บันได')
WHERE name = 'ค่าเครื่องมือ/อุปกรณ์ช่าง';

UPDATE slip_transactions SET category = 'ค่าเครื่องมือ/อุปกรณ์ช่าง'
WHERE id IN (42, 71);
