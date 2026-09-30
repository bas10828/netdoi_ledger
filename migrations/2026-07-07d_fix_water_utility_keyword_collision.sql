-- "ค่าน้ำ" keyword on ค่าน้ำค่าไฟ (utility bills) is a substring of "ค่าน้ำมัน" (fuel),
-- so any fuel memo phrased "ค่าน้ำมัน ..." instead of "เติมน้ำมัน ..." would be
-- miscategorized as a water/electric utility bill before ever reaching the fuel
-- categories (id3 sort_order 3, ahead of fuel at sort_order 6-10). Not yet triggered
-- in existing data, but a real landmine.
--
-- Replace with specific phrases, and add "น้ำดื่ม" explicitly so id 48
-- ("ค่าน้ำดื่ม" — drinking water, kept under this monthly-recurring utility category
-- per user decision) keeps matching now that the broad "ค่าน้ำ" keyword is gone.

UPDATE categories
SET keywords = ARRAY['ค่าน้ำประปา','ค่าน้ำไฟฟ้า','ค่าไฟ','การไฟฟ้า','การประปา','น้ำดื่ม']
WHERE name = 'ค่าน้ำค่าไฟ';
