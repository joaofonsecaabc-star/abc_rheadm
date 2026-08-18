PRAGMA foreign_keys = OFF;

CREATE TABLE hr_occurrences_new (
  id INTEGER PRIMARY KEY,
  employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
  occurrence_date TEXT NOT NULL,
  end_date TEXT,
  type TEXT NOT NULL CHECK(type IN ('Falta','Atestado','Atraso','Aviso','Férias')),
  hours INTEGER NOT NULL DEFAULT 0,
  minutes INTEGER NOT NULL DEFAULT 0,
  days INTEGER NOT NULL DEFAULT 0,
  note TEXT,
  data_json TEXT NOT NULL,
  created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

INSERT INTO hr_occurrences_new SELECT * FROM hr_occurrences;
DROP TABLE hr_occurrences;
ALTER TABLE hr_occurrences_new RENAME TO hr_occurrences;
CREATE INDEX idx_occurrences_date ON hr_occurrences(occurrence_date);
CREATE INDEX idx_occurrences_employee ON hr_occurrences(employee_id);
INSERT OR IGNORE INTO app_settings(key,value_json) VALUES ('state_revision','0');

CREATE TABLE IF NOT EXISTS login_attempts (
  attempt_key TEXT PRIMARY KEY,
  attempts INTEGER NOT NULL DEFAULT 0,
  blocked_until TEXT,
  updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

PRAGMA foreign_keys = ON;
