-- Follow-up to 2026-07-07_vehicle_category_keywords.sql.
-- 1) The repair sub-categories were missing a "เปลี่ยนแบต <vehicle>" keyword pair
--    (the generic parent category has "เปลี่ยนแบต" alone, but the per-vehicle ones only
--    had "ซ่อม"/"อะไหล่" context words), so memos like "เปลี่ยนแบตรถ Vigo" still fell
--    through to the generic "ค่าซ่อมแซมและบำรุงรักษา" category.
-- 2) Re-run categorization on the 3 existing rows that were miscategorized before the fix.

UPDATE categories SET keywords = keywords || ARRAY['เปลี่ยนแบต dmax','dmax เปลี่ยนแบต'] WHERE name = 'ค่าซ่อมบำรุง - Dmax';
UPDATE categories SET keywords = keywords || ARRAY['เปลี่ยนแบต revo','revo เปลี่ยนแบต'] WHERE name = 'ค่าซ่อมบำรุง - Revo';
UPDATE categories SET keywords = keywords || ARRAY['เปลี่ยนแบต vigo','vigo เปลี่ยนแบต','เปลี่ยนแบตรถ vigo'] WHERE name = 'ค่าซ่อมบำรุง - Vigo';
UPDATE categories SET keywords = keywords || ARRAY['เปลี่ยนแบตรถตู้','รถตู้เปลี่ยนแบต'] WHERE name = 'ค่าซ่อมบำรุง - รถตู้';

-- Retroactive fix for rows categorized before this migration existed.
UPDATE slip_transactions SET category = 'ค่าซ่อมบำรุง - Vigo' WHERE id = 9 AND category = 'ค่าซ่อมแซมและบำรุงรักษา';
UPDATE slip_transactions SET category = 'ค่าน้ำมัน - Revo' WHERE id = 44 AND category = 'ค่าน้ำมันเชื้อเพลิง';
UPDATE slip_transactions SET category = 'ค่าน้ำมัน - Vigo' WHERE id = 46 AND category = 'ค่าน้ำมันเชื้อเพลิง';
