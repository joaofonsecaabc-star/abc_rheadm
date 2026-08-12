PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS stores (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL COLLATE NOCASE UNIQUE,
  active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1)),
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS positions (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL COLLATE NOCASE UNIQUE,
  active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1)),
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS unregistered_reasons (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  description TEXT NOT NULL COLLATE NOCASE UNIQUE,
  active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1)),
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS users (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  username TEXT NOT NULL COLLATE NOCASE UNIQUE,
  full_name TEXT NOT NULL,
  password_hash TEXT NOT NULL,
  password_salt TEXT NOT NULL,
  role TEXT NOT NULL DEFAULT 'operator' CHECK(role IN ('admin','operator')),
  active INTEGER NOT NULL DEFAULT 1 CHECK(active IN (0,1)),
  last_login_at TEXT,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS sessions (
  id TEXT PRIMARY KEY,
  user_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
  expires_at TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS employees (
  id INTEGER PRIMARY KEY,
  employee_name TEXT NOT NULL,
  cpf TEXT,
  role_name TEXT NOT NULL,
  store_name TEXT NOT NULL,
  phone TEXT,
  hired_at TEXT,
  birth_date TEXT,
  active INTEGER NOT NULL DEFAULT 1,
  employment_status TEXT NOT NULL DEFAULT 'Ativo',
  formal_employment INTEGER NOT NULL DEFAULT 1,
  unregistered_start_date TEXT,
  unregistered_reason TEXT,
  experience_days INTEGER NOT NULL DEFAULT 90,
  experience_critical INTEGER NOT NULL DEFAULT 0,
  notice_start TEXT,
  notice_end TEXT,
  termination_date TEXT,
  receives_transit INTEGER NOT NULL DEFAULT 1,
  receives_cost_assistance INTEGER NOT NULL DEFAULT 0,
  cost_assistance_amount REAL NOT NULL DEFAULT 0,
  card_type TEXT,
  card_daily_fare REAL NOT NULL DEFAULT 0,
  second_card_type TEXT,
  second_card_daily_fare REAL NOT NULL DEFAULT 0,
  recharge_day INTEGER,
  advance_days INTEGER NOT NULL DEFAULT 3,
  schedule_type TEXT NOT NULL DEFAULT 'Personalizada',
  schedule_start_date TEXT,
  work_days_json TEXT NOT NULL DEFAULT '[1,2,3,4,5]',
  data_json TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE INDEX IF NOT EXISTS idx_employees_cpf ON employees(cpf) WHERE cpf IS NOT NULL AND cpf <> '';
CREATE INDEX IF NOT EXISTS idx_employees_store ON employees(store_name);
CREATE INDEX IF NOT EXISTS idx_employees_termination ON employees(termination_date);
CREATE INDEX IF NOT EXISTS idx_employees_active ON employees(active);

CREATE TABLE IF NOT EXISTS recharge_events (
  id INTEGER PRIMARY KEY,
  employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
  period TEXT NOT NULL,
  scheduled_date TEXT NOT NULL,
  completed_date TEXT NOT NULL,
  charged_days INTEGER,
  card_amount REAL,
  second_card_amount REAL,
  total_amount REAL,
  responsible_user_id INTEGER REFERENCES users(id),
  data_json TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_recharge_employee_period ON recharge_events(employee_id, period);
CREATE INDEX IF NOT EXISTS idx_recharge_completed ON recharge_events(completed_date);

CREATE TABLE IF NOT EXISTS hr_occurrences (
  id INTEGER PRIMARY KEY,
  employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
  occurrence_date TEXT NOT NULL,
  end_date TEXT,
  type TEXT NOT NULL CHECK(type IN ('Falta','Atestado','Atraso','Aviso')),
  hours INTEGER NOT NULL DEFAULT 0,
  minutes INTEGER NOT NULL DEFAULT 0,
  days INTEGER NOT NULL DEFAULT 0,
  note TEXT,
  data_json TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_occurrences_date ON hr_occurrences(occurrence_date);
CREATE INDEX IF NOT EXISTS idx_occurrences_employee ON hr_occurrences(employee_id);

CREATE TABLE IF NOT EXISTS app_settings (
  key TEXT PRIMARY KEY,
  value_json TEXT NOT NULL,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS audit_logs (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id INTEGER REFERENCES users(id),
  action TEXT NOT NULL,
  entity_type TEXT NOT NULL,
  entity_id TEXT,
  before_json TEXT,
  after_json TEXT,
  ip_address TEXT,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_audit_created ON audit_logs(created_at);
CREATE INDEX IF NOT EXISTS idx_audit_entity ON audit_logs(entity_type, entity_id);

INSERT OR IGNORE INTO app_settings(key,value_json) VALUES
 ('advance_days','3'),
 ('schema_version','1');
