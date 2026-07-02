ALTER TABLE categories ADD COLUMN IF NOT EXISTS group_name TEXT;

UPDATE categories SET group_name = 'ค่าน้ำมันเชื้อเพลิง'
    WHERE name IN ('ค่าน้ำมันเชื้อเพลิง', 'ค่าน้ำมัน - Dmax', 'ค่าน้ำมัน - Revo', 'ค่าน้ำมัน - Vigo', 'ค่าน้ำมัน - รถตู้');

UPDATE categories SET group_name = 'ค่าซ่อมแซมและบำรุงรักษา'
    WHERE name IN ('ค่าซ่อมแซมและบำรุงรักษา', 'ค่าซ่อมบำรุง - Dmax', 'ค่าซ่อมบำรุง - Revo', 'ค่าซ่อมบำรุง - Vigo', 'ค่าซ่อมบำรุง - รถตู้');
