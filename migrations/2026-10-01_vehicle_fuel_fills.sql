-- Fuel fills read from tax-invoice and odometer photos, one row per fill. Feeds the
-- /vehicles cost page. Deliberately separate from slip_transactions: a fill may be paid
-- in cash, by card or by a staff member's own QR, so it need not have a bank slip.
CREATE TABLE IF NOT EXISTS vehicle_fuel_fills (
    id            SERIAL PRIMARY KEY,
    fill_date     DATE NOT NULL,
    vehicle       TEXT NOT NULL,
    plate         TEXT NOT NULL DEFAULT '',
    fuel_type     TEXT NOT NULL,              -- 'diesel' | 'benzine'
    fuel_product  TEXT NOT NULL DEFAULT '',   -- as printed on the receipt
    price_per_l   NUMERIC(8, 2),              -- NULL when only a bank slip exists
    liters        NUMERIC(10, 3),
    amount        NUMERIC(12, 2) NOT NULL,
    odometer_km   INTEGER,
    note          TEXT NOT NULL DEFAULT '',
    station       TEXT NOT NULL DEFAULT '',
    created_at    TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE INDEX IF NOT EXISTS vehicle_fuel_fills_vehicle_date ON vehicle_fuel_fills (vehicle, fill_date);
