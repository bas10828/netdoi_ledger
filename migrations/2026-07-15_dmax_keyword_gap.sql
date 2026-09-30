-- Memo "เติมน้ำมัน D-Max" (hyphenated, "เติม" prefix) doesn't match any existing
-- Dmax keyword and falls through to the generic ค่าน้ำมันเชื้อเพลิง category.
-- Bare "d-max" catches it without colliding with Revo/Vigo/รถตู้ keywords.
UPDATE categories SET keywords = array_append(keywords, 'd-max')
WHERE name = 'ค่าน้ำมัน - Dmax';
