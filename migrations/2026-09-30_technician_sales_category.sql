-- Cameras/equipment bought to resell to technicians (separate technician price list,
-- not the storefront price) get their own category, grouped under ต้นทุนขาย so the
-- reports page can split technician sales from other cost of goods.
UPDATE categories SET group_name = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)'
WHERE name = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)';

UPDATE categories SET sort_order = sort_order + 1
WHERE sort_order > (SELECT sort_order FROM categories WHERE name = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)');

INSERT INTO categories (name, keywords, group_name, sort_order)
SELECT 'ต้นทุนสินค้าขายช่าง', ARRAY['สั่งกล้อง'], 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)', sort_order + 1
FROM categories WHERE name = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)'
ON CONFLICT (name) DO NOTHING;

-- 'สั่งกล้อง' was added to the generic ต้นทุนขาย category earlier today; move it.
UPDATE categories SET keywords = array_remove(keywords, 'สั่งกล้อง')
WHERE name = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)';

UPDATE slip_transactions SET category = 'ต้นทุนสินค้าขายช่าง' WHERE id = 141;

-- Other spellings of the hardware/lighting stores seen in memos (id 156 typo) or likely.
UPDATE categories SET keywords = keywords || ARRAY(
    SELECT k FROM unnest(ARRAY['โกลลอลเฮ้าส์', 'global house', 'hub lighting']) AS k
    WHERE NOT (k = ANY(keywords))
)
WHERE name = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง';
