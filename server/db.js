'use strict';
const path = require('path');
const fs = require('fs');
const Database = require('better-sqlite3');
const config = require('./config');

fs.mkdirSync(config.dataDir, { recursive: true });
const db = new Database(path.join(config.dataDir, 'jobs.db'));
db.pragma('journal_mode = WAL');

db.exec(`
CREATE TABLE IF NOT EXISTS jobs (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  original_name TEXT NOT NULL,        -- 该任务的展示名（取首个文件名，便于列表查看）
  stored_name   TEXT NOT NULL,        -- 主输入文件名（selected_product 对应文件）
  input_path    TEXT NOT NULL,        -- 主输入（selected_product）所在路径
  step_path     TEXT,                 -- 转换产物（兼容旧单文件语义）
  step_paths    TEXT,                 -- JSON 数组：本次任务产出的全部 .stp 路径（多根时多个）
  roots         TEXT,                 -- JSON 数组：被判定为“根产品/顶层装配”的文件名列表
  status        TEXT NOT NULL DEFAULT 'pending',
  attempts      INTEGER NOT NULL DEFAULT 0,
  error         TEXT,
  meta          TEXT,                 -- JSON：{ files:[], products:[], parts:[], selected_product }
  created_at    TEXT NOT NULL,
  started_at    TEXT,
  finished_at   TEXT,
  expires_at    TEXT
);
CREATE INDEX IF NOT EXISTS idx_jobs_status ON jobs(status, id);
`);

// 老库迁移：补缺失列（先建索引会导致 no such column）
const cols = db.prepare(`PRAGMA table_info(jobs)`).all().map((c) => c.name);
for (const c of ['step_path', 'meta', 'expires_at', 'step_paths', 'roots']) {
  if (!cols.includes(c)) db.exec(`ALTER TABLE jobs ADD COLUMN ${c} TEXT`);
}
db.exec(`CREATE INDEX IF NOT EXISTS idx_jobs_expires ON jobs(expires_at)`);
{
  const ms = config.retentionDays * 86400000;
  const now = Date.now();
  const upd = db.prepare(`UPDATE jobs SET expires_at=? WHERE id=?`);
  for (const row of db.prepare(`SELECT id, created_at FROM jobs WHERE expires_at IS NULL`).all()) {
    const base = row.created_at ? new Date(row.created_at).getTime() : now;
    upd.run(new Date(base + ms).toISOString(), row.id);
  }
}

// 进程重启后，把上次遗留的 processing 任务退回队列（由看门狗/重试接管）
db.prepare(`UPDATE jobs SET status='pending', started_at=NULL WHERE status='processing'`).run();

// meta 存 JSON；建任务时只填 files 骨架，产品/零件分类与 selected_product 由调用方决定
function createJob({ original_name, stored_name, input_path, meta }) {
  const now = new Date().toISOString();
  const expires = new Date(Date.now() + config.retentionDays * 86400000).toISOString();
  const info = db.prepare(
    `INSERT INTO jobs (original_name, stored_name, input_path, status, attempts, meta, created_at, expires_at)
     VALUES (?, ?, ?, 'pending', 0, ?, ?, ?)`
  ).run(original_name, stored_name, input_path, meta == null ? null : JSON.stringify(meta), now, expires);
  return info.lastInsertRowid;
}

// 原子地认领下一个 pending 任务（单 Worker，串行）
function claimNext() {
  const tx = db.transaction(() => {
    const row = db.prepare(
      `SELECT * FROM jobs WHERE status='pending' ORDER BY id ASC LIMIT 1`
    ).get();
    if (!row) return null;
    db.prepare(
      `UPDATE jobs SET status='processing', started_at=?, attempts=attempts+1 WHERE id=?`
    ).run(new Date().toISOString(), row.id);
    return db.prepare(`SELECT * FROM jobs WHERE id=?`).get(row.id);
  });
  return tx();
}

function completeJob(id, step_paths, roots) {
  db.prepare(`UPDATE jobs SET status='done', step_path=?, step_paths=?, roots=?, finished_at=? WHERE id=?`)
    .run(
      JSON.stringify(step_paths && step_paths.length ? step_paths[0] : null),
      JSON.stringify(step_paths || []),
      JSON.stringify(roots || []),
      new Date().toISOString(), id,
    );
}

function failJob(id, error, willRetry) {
  db.prepare(
    `UPDATE jobs SET status=?, error=?, finished_at=? WHERE id=?`
  ).run(
    willRetry ? 'pending' : 'failed',
    willRetry ? null : String(error),
    willRetry ? null : new Date().toISOString(),
    id
  );
}

function updateMeta(id, meta) {
  db.prepare(`UPDATE jobs SET meta=? WHERE id=?`).run(JSON.stringify(meta), id);
}

function getJob(id) {
  return db.prepare(`SELECT * FROM jobs WHERE id=?`).get(id);
}

function listJobs(limit = 100, offset = 0) {
  return db.prepare(`SELECT * FROM jobs ORDER BY id DESC LIMIT ? OFFSET ?`).all(limit, offset);
}

function countJobs() {
  return db.prepare(`SELECT COUNT(*) AS n FROM jobs`).get().n;
}

function queuePosition(id) {
  const row = db.prepare(
    `SELECT COUNT(*) AS n FROM jobs WHERE status='pending' AND id < ?`
  ).get(id);
  return row ? row.n : 0;
}

module.exports = { db, createJob, claimNext, completeJob, failJob, updateMeta, getJob, listJobs, countJobs, queuePosition };
