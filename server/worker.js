'use strict';
const fs = require('fs');
const path = require('path');
const config = require('./config');
const { claimNext, completeJob, failJob, updateMeta } = require('./db');
const { convert, killCatia } = require('./catia');

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
let running = true;
process.on('SIGINT', () => { running = false; });
process.on('SIGTERM', () => { running = false; });

async function processOne() {
  const job = claimNext();
  if (!job) { await sleep(config.pollIntervalMs); return; }
  console.log(`[worker] start job #${job.id} (${job.original_name}) attempt ${job.attempts}`);

  let meta = null;
  try { meta = JSON.parse(job.meta || '{}'); } catch (_) { meta = {}; }
  const files = Array.isArray(meta.files) ? meta.files : [];
  const batchFiles = files.map((f) => f.path).filter(Boolean);

  if (!batchFiles.length) {
    failJob(job.id, '任务缺少上传文件列表', false);
    return;
  }

  // 手动改选的根产品（可选）：若设置则只导出这些，否则由驱动自动探测根
  const forcedRoots = Array.isArray(meta.selected_roots) ? meta.selected_roots : null;

  const outDir = path.join(config.resultsDir, String(job.id));
  fs.mkdirSync(outDir, { recursive: true });
  // 清除上一次（可能改选根后）遗留的旧 .stp，避免结果目录里有多个过期文件
  for (const f of fs.readdirSync(outDir)) {
    if (/\.stp$/i.test(f)) { try { fs.rmSync(path.join(outDir, f), { force: true }); } catch (_) {} }
  }

  try {
    await killCatia(); // 确保起始无残留 CATIA 占用 License
    const { roots, steps } = await convert(batchFiles, outDir, config.catiaTimeoutMs, forcedRoots);
    // 把产物与根信息写回 meta，便于前端展示
    meta.roots = roots;
    meta.steps = steps.map((s) => ({ name: s.name, path: s.path }));
    updateMeta(job.id, meta);
    completeJob(job.id, steps.map((s) => s.path), roots);
    console.log(`[worker] job #${job.id} done -> ${steps.length} step(s), roots=${roots.join(', ')}`);
  } catch (  e) {
    const willRetry = job.attempts < config.maxAttempts;
    failJob(job.id, (e && e.message) || String(e), willRetry);
    console.error(`[worker] job #${job.id} failed: ${(e && e.message) || e} ${willRetry ? '(will retry)' : '(no retry)'}`);
  } finally {
    await killCatia(); // job 之间强制清理，释放 License，防队列卡死
  }
}

async function loop() {
  console.log(`[worker] serial worker started (mock=${config.mock}, timeout=${config.catiaTimeoutMs}ms, maxAttempts=${config.maxAttempts})`);
  while (running) {
    try {
      await processOne();
    } catch (e) {
      console.error('[worker] unexpected error in loop:', e);
      await sleep(config.pollIntervalMs);
    }
  }
  console.log('[worker] stopped');
}

if (require.main === module) {
  loop();
}

module.exports = { loop };
