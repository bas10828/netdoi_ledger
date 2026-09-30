-- Uncategorized review 2026-09-30 (ids 141, 154-156) + one unknown-direction row (121).

-- Keyword gaps: "สั่งกล้อง" (ordering a camera for a client) matched nothing — only
-- "กล้องวงจรปิด" existed. Bare "กล้อง" would steal "ซื้อของงานติดกล้อง" (consumables)
-- since ต้นทุนขาย sorts first, so keep it specific.
UPDATE categories SET keywords = array_append(keywords, 'สั่งกล้อง')
WHERE name = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)' AND NOT ('สั่งกล้อง' = ANY(keywords));

-- Hardware/lighting store names that show up in reimbursement memos.
UPDATE categories SET keywords = array_append(keywords, 'ฮับไลท์ติ้ง')
WHERE name = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง' AND NOT ('ฮับไลท์ติ้ง' = ANY(keywords));
UPDATE categories SET keywords = array_append(keywords, 'โกลบอลเฮ้าส์')
WHERE name = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง' AND NOT ('โกลบอลเฮ้าส์' = ANY(keywords));

UPDATE slip_transactions SET category = 'ค่าอุปกรณ์/สินค้า (ต้นทุนขาย)' WHERE id = 141;
UPDATE slip_transactions SET category = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง' WHERE id IN (155, 156);
-- Lump-sum reimbursement with no itemization — same treatment as id 49 (user decision).
UPDATE slip_transactions SET category = 'อื่นๆ' WHERE id = 154;

-- Company paid HUB LIGHTING for job materials — expense, like id 26.
UPDATE slip_transactions SET direction = 'expense' WHERE id = 121 AND direction = 'unknown';
