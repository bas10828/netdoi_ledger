-- id 53 "สลักหลัง PO พะเยา" matched no category — endorsing a PO for billing is
-- PO-related admin work, same bucket as the deposit-per-PO rows (35, 89, 96).
-- Add "สลักหลัง" not "po" — "po" as a bare keyword risks false-matching unrelated memos.

UPDATE categories SET keywords = keywords || ARRAY['สลักหลัง'] WHERE name = 'เงินประกันผลงาน/สัญญา';

UPDATE slip_transactions SET category = 'เงินประกันผลงาน/สัญญา' WHERE id = 53 AND category = 'อื่นๆ';
