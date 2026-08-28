'use strict';
const fs = require('fs');
const path = require('path');
const config = require('./config');
const { db } = require('./db');

// 清理过期任务：删除上传文件与转换结果，并把记录标记为失效（expired）。
// 跳过 processing，避免与正在进行的转换产生竞态。
function cleanupExpired() {
  const now = new Date().toISOString();
  const rows = db.prepare(
    `SELECT id FROM jobs
     WHERE expires_at IS NOT NULL AND expires_at < ?
       AND status NOT IN ('expired', 'processing')`
  ).all(now);

  for (const r of rows) {
    // 部分沙箱/环境会拦截 fs.rm（安全删除钩子），失败时忽略，避免清理任务本身崩溃
    try { fs.rmSync(path.join(config.uploadsDir, String(r.id)), { recursive: true, force: true }); } catch (_) {}
    try { fs.rmSync(path.join(config.resultsDir, String(r.id)), { recursive: true, force: true }); } catch (_) {}
    db.prepare(`UPDATE jobs SET status='expired', step_path=NULL, error=NULL, finished_at=? WHERE id=?`)
      .run(now, r.id);
  }
  if (rows.length) {
    console.log(`[cleanup] expired ${rows.length} job(s): ${rows.map((r) => '#' + r.id).join(', ')}`);
  }
  return rows.length;
}

// 启动定时清理（默认每小时）。unref 使定时器不阻止进程退出。
function startCleanup() {
  cleanupExpired();
  const timer = setInterval(cleanupExpired, config.cleanupIntervalMs);
  timer.unref();
  return timer;
}

module.exports = { cleanupExpired, startCleanup };
