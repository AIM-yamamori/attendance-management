-- 勤怠管理システム アプリDB スキーマ定義
-- 基本設計書 4章「データベース設計」に基づく

-- ============================================
-- users（ユーザーマスタ）
-- 基本設計書 4.3.1節
-- ============================================
CREATE TABLE IF NOT EXISTS users (
    employee_id   TEXT PRIMARY KEY,
    last_name     TEXT NOT NULL,
    first_name    TEXT NOT NULL,
    department    TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    role          TEXT NOT NULL CHECK(role IN ('admin','general')),
    created_at    TEXT NOT NULL,
    updated_at    TEXT NOT NULL
);

-- ============================================
-- locks（ロック状態管理）
-- 基本設計書 4.3.2節
-- ============================================
CREATE TABLE IF NOT EXISTS locks (
    id            INTEGER PRIMARY KEY AUTOINCREMENT,
    target_month  TEXT NOT NULL,
    employee_id   TEXT NOT NULL REFERENCES users(employee_id),
    is_locked     INTEGER NOT NULL DEFAULT 0,
    locked_at     TEXT,
    unlocked_at   TEXT,
    updated_at    TEXT NOT NULL,
    UNIQUE(target_month, employee_id)
);

CREATE INDEX IF NOT EXISTS idx_locks_target_month ON locks(target_month);

-- ============================================
-- settings（システム設定）
-- 基本設計書 4.3.3節
-- ============================================
CREATE TABLE IF NOT EXISTS settings (
    key         TEXT PRIMARY KEY,
    value       TEXT NOT NULL,
    updated_at  TEXT NOT NULL
);