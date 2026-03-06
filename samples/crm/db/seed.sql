PRAGMA foreign_keys = ON;

CREATE TABLE IF NOT EXISTS groups (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE,
  sort_order INTEGER NOT NULL DEFAULT 9999
);

CREATE TABLE IF NOT EXISTS contacts (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  first_name TEXT NOT NULL,
  last_name TEXT NOT NULL,
  email TEXT,
  phone TEXT,
  company TEXT,
  notes TEXT,
  group_id INTEGER,
  created_at TEXT NOT NULL DEFAULT (datetime('now')),
  updated_at TEXT,
  FOREIGN KEY(group_id) REFERENCES groups(id) ON DELETE SET NULL
);

CREATE INDEX IF NOT EXISTS idx_contacts_name ON contacts(last_name, first_name);
CREATE INDEX IF NOT EXISTS idx_contacts_group ON contacts(group_id);

INSERT OR IGNORE INTO groups(id,name,sort_order) VALUES (1,'Customers',1);
INSERT OR IGNORE INTO groups(id,name,sort_order) VALUES (2,'Leads',2);
INSERT OR IGNORE INTO groups(id,name,sort_order) VALUES (3,'Partners',3);

INSERT OR IGNORE INTO contacts(id,first_name,last_name,email,phone,company,notes,group_id,created_at)
VALUES (1,'Anna','Berg','anna@example.com','+31 20 123 4567','Northwind BV','Interested in yearly contract',2,datetime('now'));
