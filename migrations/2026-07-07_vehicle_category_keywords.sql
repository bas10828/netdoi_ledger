-- Fix auto-categorization for per-vehicle fuel/repair categories added in 2026-07-02_vehicle_categories.sql
-- Problem 1: keywords were left empty, so nothing ever matched them.
-- Problem 2: the generic "ค่าน้ำมันเชื้อเพลิง" (sort_order 6) / "ค่าซ่อมแซมและบำรุงรักษา" (sort_order 7)
--   categories sit ahead of the per-vehicle ones in sort_order, so guess_category()'s
--   first-match loop always hits the generic category first even with keywords filled in.
-- Fix: renumber so per-vehicle categories are checked before their generic parent, and
--   fill in keywords that pair the vehicle name with fuel/repair context words (guess_category
--   only does substring OR-matching, not AND, so a bare vehicle name would collide between the
--   fuel and repair sub-categories for the same vehicle).

UPDATE categories SET sort_order = 6 WHERE name = 'ค่าน้ำมัน - Dmax';
UPDATE categories SET sort_order = 7 WHERE name = 'ค่าน้ำมัน - Revo';
UPDATE categories SET sort_order = 8 WHERE name = 'ค่าน้ำมัน - Vigo';
UPDATE categories SET sort_order = 9 WHERE name = 'ค่าน้ำมัน - รถตู้';
UPDATE categories SET sort_order = 10 WHERE name = 'ค่าน้ำมันเชื้อเพลิง';
UPDATE categories SET sort_order = 11 WHERE name = 'ค่าซ่อมบำรุง - Dmax';
UPDATE categories SET sort_order = 12 WHERE name = 'ค่าซ่อมบำรุง - Revo';
UPDATE categories SET sort_order = 13 WHERE name = 'ค่าซ่อมบำรุง - Vigo';
UPDATE categories SET sort_order = 14 WHERE name = 'ค่าซ่อมบำรุง - รถตู้';
UPDATE categories SET sort_order = 15 WHERE name = 'ค่าซ่อมแซมและบำรุงรักษา';

UPDATE categories SET keywords = ARRAY[
    'น้ำมัน dmax','dmax น้ำมัน','น้ำมันดีแม็กซ์','ดีแม็กซ์น้ำมัน'
] WHERE name = 'ค่าน้ำมัน - Dmax';

UPDATE categories SET keywords = ARRAY[
    'น้ำมัน revo','revo น้ำมัน','น้ำมันรีโว่','รีโว่น้ำมัน','น้ำมันรีโว','รีโวน้ำมัน'
] WHERE name = 'ค่าน้ำมัน - Revo';

UPDATE categories SET keywords = ARRAY[
    'น้ำมัน vigo','vigo น้ำมัน','น้ำมันวีโก้','วีโก้น้ำมัน','น้ำมันวีโก','วีโกน้ำมัน'
] WHERE name = 'ค่าน้ำมัน - Vigo';

UPDATE categories SET keywords = ARRAY[
    'น้ำมันรถตู้','รถตู้น้ำมัน'
] WHERE name = 'ค่าน้ำมัน - รถตู้';

UPDATE categories SET keywords = ARRAY[
    'ซ่อม dmax','dmax ซ่อม','ซ่อมดีแม็กซ์','ดีแม็กซ์ซ่อม','อะไหล่ dmax','dmax อะไหล่'
] WHERE name = 'ค่าซ่อมบำรุง - Dmax';

UPDATE categories SET keywords = ARRAY[
    'ซ่อม revo','revo ซ่อม','ซ่อมรีโว่','รีโว่ซ่อม','ซ่อมรีโว','รีโวซ่อม','อะไหล่ revo'
] WHERE name = 'ค่าซ่อมบำรุง - Revo';

UPDATE categories SET keywords = ARRAY[
    'ซ่อม vigo','vigo ซ่อม','ซ่อมวีโก้','วีโก้ซ่อม','ซ่อมวีโก','วีโกซ่อม','อะไหล่ vigo'
] WHERE name = 'ค่าซ่อมบำรุง - Vigo';

UPDATE categories SET keywords = ARRAY[
    'ซ่อมรถตู้','รถตู้ซ่อม'
] WHERE name = 'ค่าซ่อมบำรุง - รถตู้';
