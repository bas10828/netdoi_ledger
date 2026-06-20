-- slip_ledger: backup of every file sent to the LINE group + parsed slip transactions.
-- Kept separate from raw_files since not every file sent to the group is a parseable slip.

CREATE TABLE owner_accounts (
    id            serial PRIMARY KEY,
    name_hint     text NOT NULL,
    account_hint  text NOT NULL,
    created_at    timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE raw_files (
    id             serial PRIMARY KEY,
    line_message_id text,
    line_group_id   text,
    line_user_id    text,
    file_type       text NOT NULL,
    storage_path    text NOT NULL,
    received_at     timestamptz NOT NULL DEFAULT now(),
    processed       boolean NOT NULL DEFAULT false,
    is_slip         boolean NOT NULL DEFAULT false
);

CREATE TABLE slip_transactions (
    id                serial PRIMARY KEY,
    raw_file_id       integer REFERENCES raw_files(id),
    bank              text,
    txn_date          date,
    txn_time          time,
    amount            numeric(12, 2) NOT NULL,
    fee               numeric(12, 2) NOT NULL DEFAULT 0,
    sender_name       text,
    sender_account    text,
    receiver_name     text,
    receiver_account  text,
    memo              text,
    printed_ref       text,
    qr_trans_ref      text,
    direction         text CHECK (direction IN ('expense', 'income', 'unknown')) NOT NULL DEFAULT 'unknown',
    category          text,
    ai_model          text NOT NULL,
    verified_bank     boolean NOT NULL DEFAULT false,
    raw_ai_response   jsonb,
    status            text NOT NULL DEFAULT 'pending_review',
    created_at        timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX idx_slip_transactions_raw_file_id ON slip_transactions (raw_file_id);
CREATE INDEX idx_slip_transactions_txn_date ON slip_transactions (txn_date);
CREATE INDEX idx_slip_transactions_qr_trans_ref ON slip_transactions (qr_trans_ref);
CREATE INDEX idx_slip_transactions_category ON slip_transactions (category);

CREATE TABLE dashboard_users (
    id            serial PRIMARY KEY,
    username      text NOT NULL UNIQUE,
    password_hash text NOT NULL,
    role          text NOT NULL DEFAULT 'admin',
    created_at    timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE line_messages (
    id              serial PRIMARY KEY,
    line_message_id text,
    line_group_id   text,
    line_user_id    text,
    message_type    text NOT NULL DEFAULT 'text',
    text            text,
    received_at     timestamptz NOT NULL DEFAULT now()
);

CREATE INDEX idx_line_messages_group_id ON line_messages (line_group_id);
CREATE INDEX idx_line_messages_received_at ON line_messages (received_at);

CREATE TABLE ai_usage_log (
    id             serial PRIMARY KEY,
    model          text NOT NULL,
    input_tokens   integer NOT NULL,
    output_tokens  integer NOT NULL,
    cost_usd       numeric(10, 6) NOT NULL,
    created_at     timestamptz NOT NULL DEFAULT now()
);

CREATE TABLE app_settings (
    key   text PRIMARY KEY,
    value text NOT NULL
);

INSERT INTO app_settings (key, value) VALUES ('ai_starting_balance_usd', '4.78');

CREATE TABLE audit_log (
    id           serial PRIMARY KEY,
    txn_id       integer,
    action       text NOT NULL CHECK (action IN ('create', 'update', 'delete')),
    changed_by   text NOT NULL,
    changed_at   timestamptz NOT NULL DEFAULT now(),
    before_data  jsonb,
    after_data   jsonb
);

CREATE INDEX idx_audit_log_txn_id ON audit_log (txn_id);
CREATE INDEX idx_audit_log_changed_at ON audit_log (changed_at);
