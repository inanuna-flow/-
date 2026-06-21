-- Migration 003: IP 黑名單管理
-- 為「後臺限制 IP 管理」頁面建立資料表
-- 邏輯：黑名單（封鎖清單中的 IP 無法存取 API）

CREATE TABLE IF NOT EXISTS kpl_dashboard.ip_blocklist (
  id          SERIAL      PRIMARY KEY,
  ip_cidr     TEXT        NOT NULL,
  label       TEXT        NOT NULL DEFAULT '',
  created_by  TEXT        NOT NULL DEFAULT '',
  created_at  TIMESTAMPTZ NOT NULL DEFAULT NOW(),
  is_active   BOOLEAN     NOT NULL DEFAULT TRUE,
  CONSTRAINT ip_blocklist_unique_ip UNIQUE (ip_cidr)
);

COMMENT ON TABLE  kpl_dashboard.ip_blocklist            IS 'IP 黑名單：列於此表的 IP 將被 server.js 阻擋存取 API';
COMMENT ON COLUMN kpl_dashboard.ip_blocklist.ip_cidr    IS 'IPv4 單一位址或 CIDR，例：192.168.1.1 或 10.0.0.0/8';
COMMENT ON COLUMN kpl_dashboard.ip_blocklist.label      IS '備註說明（原因、來源等）';
COMMENT ON COLUMN kpl_dashboard.ip_blocklist.created_by IS '建立者帳號（user_id）';
COMMENT ON COLUMN kpl_dashboard.ip_blocklist.is_active  IS '停用不刪除，方便保留稽核記錄';
