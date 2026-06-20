CREATE TABLE IF NOT EXISTS audit_log (
    id           serial PRIMARY KEY,
    txn_id       integer,
    action       text NOT NULL CHECK (action IN ('create', 'update', 'delete')),
    changed_by   text NOT NULL,
    changed_at   timestamptz NOT NULL DEFAULT now(),
    before_data  jsonb,
    after_data   jsonb
);

CREATE INDEX IF NOT EXISTS idx_audit_log_txn_id ON audit_log (txn_id);
CREATE INDEX IF NOT EXISTS idx_audit_log_changed_at ON audit_log (changed_at);
