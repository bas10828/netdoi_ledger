UPDATE slip_transactions SET category = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)' WHERE id = 72;
UPDATE slip_transactions SET category = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง' WHERE id IN (69, 74);
UPDATE slip_transactions SET category = 'ค่าเครื่องมือ/อุปกรณ์ช่าง' WHERE id = 70;
UPDATE slip_transactions SET category = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง' WHERE id = 75; -- arrived mid-migration

