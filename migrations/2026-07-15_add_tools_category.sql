-- New category for reusable technician tools, distinct from
-- ค่าวัสดุสิ้นเปลืองงานติดตั้ง (per-job consumables) and
-- ค่าอุปกรณ์/สินค้า (ต้นทุนขาย) (equipment installed for the client).
UPDATE categories SET sort_order = sort_order + 1 WHERE sort_order >= 6;

INSERT INTO categories (name, keywords, group_name, sort_order)
VALUES ('ค่าเครื่องมือ/อุปกรณ์ช่าง', ARRAY['เครื่องมือช่าง', 'ชุดเครื่องมือ'], NULL, 6);
