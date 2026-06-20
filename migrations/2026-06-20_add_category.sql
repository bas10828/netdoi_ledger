ALTER TABLE slip_transactions ADD COLUMN IF NOT EXISTS category text;
CREATE INDEX IF NOT EXISTS idx_slip_transactions_category ON slip_transactions (category);
