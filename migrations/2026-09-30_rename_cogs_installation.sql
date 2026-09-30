-- With ต้นทุนสินค้าขายช่าง split out, the old catch-all name reads too broad.
-- Rename it to pair with the technician-sales category, and give the group a short name.
-- audit_log keeps the old name on purpose (history).
UPDATE categories SET name = 'ต้นทุนสินค้างานติดตั้ง'
WHERE name = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)';

UPDATE categories SET group_name = 'ต้นทุนขาย'
WHERE group_name = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)';

UPDATE slip_transactions SET category = 'ต้นทุนสินค้างานติดตั้ง'
WHERE category = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)';
