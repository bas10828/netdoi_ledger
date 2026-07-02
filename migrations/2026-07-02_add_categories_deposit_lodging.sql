INSERT INTO categories (name, keywords, sort_order) VALUES
    ('เงินประกันผลงาน/สัญญา', '{}', 23),
    ('ค่าเดินทาง/ที่พักงานต่างพื้นที่', ARRAY['ค่าที่พัก', 'ที่พักงาน'], 24)
ON CONFLICT (name) DO NOTHING;
