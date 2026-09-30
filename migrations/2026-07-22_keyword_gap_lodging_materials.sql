-- Keyword gaps found via uncategorized transaction review (id 77, 79, 84):
-- "ค่าห้องพัก" didn't match ที่พัก keywords; "ซื้อท่อ..." and "ซื้อลูกตะปู"
-- didn't match any วัสดุสิ้นเปลืองงานติดตั้ง keyword.
UPDATE categories SET keywords = array_append(keywords, 'ห้องพัก')
WHERE name = 'ค่าเดินทาง/ที่พักงานต่างพื้นที่';

UPDATE categories SET keywords = array_append(keywords, 'ซื้อท่อ')
WHERE name = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง';

UPDATE categories SET keywords = array_append(keywords, 'ตะปู')
WHERE name = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง';

UPDATE slip_transactions SET category = 'ค่าเดินทาง/ที่พักงานต่างพื้นที่' WHERE id = 77;
UPDATE slip_transactions SET category = 'ค่าวัสดุสิ้นเปลืองงานติดตั้ง' WHERE id IN (79, 84);

-- id 81, 82: payments to NT (โทรคมนาคมแห่งชาติ), empty memo, unclear purpose — dumped to อื่นๆ.
UPDATE slip_transactions SET category = 'อื่นๆ' WHERE id IN (81, 82);
