UPDATE categories SET sort_order = sort_order + 8 WHERE sort_order >= 8;

INSERT INTO categories (name, keywords, sort_order) VALUES
    ('ค่าน้ำมัน - Dmax', '{}', 8),
    ('ค่าน้ำมัน - Revo', '{}', 9),
    ('ค่าน้ำมัน - Vigo', '{}', 10),
    ('ค่าน้ำมัน - รถตู้', '{}', 11),
    ('ค่าซ่อมบำรุง - Dmax', '{}', 12),
    ('ค่าซ่อมบำรุง - Revo', '{}', 13),
    ('ค่าซ่อมบำรุง - Vigo', '{}', 14),
    ('ค่าซ่อมบำรุง - รถตู้', '{}', 15)
ON CONFLICT (name) DO NOTHING;
