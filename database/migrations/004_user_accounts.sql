-- Migration 004: Dashboard 本地帳號（A/B/C 級別）
-- 為「帳號權限」頁的帳號管理 UI 建立資料表。
-- 級別：A = 讀全部（含後台管理、資料管理）
--       B = 不能讀後台管理（其餘可讀）
--       C = 不能讀後台管理與資料管理（一般身分）
-- 密碼：PBKDF2-SHA256（10000 iterations、32 bytes），每帳號獨立隨機 salt。

BEGIN;

CREATE TABLE IF NOT EXISTS kpl_dashboard.user_accounts (
  id            SERIAL      PRIMARY KEY,
  user_id       TEXT        NOT NULL,
  password_hash TEXT        NOT NULL,
  password_salt TEXT        NOT NULL,
  level         CHAR(1)     NOT NULL DEFAULT 'C' CHECK (level IN ('A','B','C')),
  is_active     BOOLEAN     NOT NULL DEFAULT TRUE,
  created_by    TEXT        NOT NULL DEFAULT '',
  created_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  updated_at    TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  CONSTRAINT user_accounts_unique_uid UNIQUE (user_id)
);

COMMENT ON TABLE  kpl_dashboard.user_accounts                IS 'Dashboard 本地帳號（A/B/C 級別）；密碼為 PBKDF2-SHA256 + 每帳號隨機 salt';
COMMENT ON COLUMN kpl_dashboard.user_accounts.user_id        IS '登入帳號（唯一，程式以小寫比對）';
COMMENT ON COLUMN kpl_dashboard.user_accounts.password_hash  IS 'PBKDF2-SHA256(password, salt, 10000, 32) 的 hex';
COMMENT ON COLUMN kpl_dashboard.user_accounts.password_salt  IS '每帳號隨機 salt（hex）';
COMMENT ON COLUMN kpl_dashboard.user_accounts.level          IS '權限級別 A/B/C';
COMMENT ON COLUMN kpl_dashboard.user_accounts.is_active      IS '停用不刪除，方便保留稽核記錄；停用後無法登入';
COMMENT ON COLUMN kpl_dashboard.user_accounts.created_by     IS '建立者帳號（user_id）';

COMMIT;
