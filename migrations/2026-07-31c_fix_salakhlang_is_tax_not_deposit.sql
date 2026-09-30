-- Correction: id 53 "สลักหลัง PO พะเยา" was miscategorized as เงินประกันผลงาน/สัญญา.
-- Checked the slip image: it's a bill-payment to biller "MAE SAI" with a taxpayer ID
-- and "ประเภทรายได้" field — "สลักหลังตราสาร" is Revenue Department terminology for
-- paying stamp duty in lieu of affixing a paper stamp on a contract/PO. This is a
-- tax expense (ภาษี), not a performance/contract deposit.

UPDATE categories SET keywords = array_remove(keywords, 'สลักหลัง') WHERE name = 'เงินประกันผลงาน/สัญญา';
UPDATE categories SET keywords = keywords || ARRAY['สลักหลัง'] WHERE name = 'ภาษี';

UPDATE slip_transactions SET category = 'ภาษี' WHERE id = 53;
