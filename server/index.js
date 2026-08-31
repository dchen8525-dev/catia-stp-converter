'use strict';
const express = require('express');
const multer = require('multer');
const fs = require('fs');
const path = require('path');
const config = require('./config');
const yazl = require('yazl');
const { createJob, getJob, listJobs, countJobs, queuePosition, updateMeta, db } = require('./db');
const { loop: startWorker } = require('./worker');
const { startCleanup } = require('./cleanup');

fs.mkdirSync(config.uploadsDir, { recursive: true });
fs.mkdirSync(config.resultsDir, { recursive: true });

const app = express();
app.use(express.json());
app.use(express.static(path.join(config.rootDir, 'public')));

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, config.uploadsDir),
  filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname),
});
const upload = multer({
  storage,
  limits: { fileSize: config.maxFileMB * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (/\.(catpart|catproduct)$/i.test(file.originalname)) return cb(null, true);
    cb(new Error('仅支持 .CATPart / .CATProduct 文件'));
  },
});

function safeJson(v, fallback) {
  if (v == null) return fallback;
  try { return JSON.parse(v); } catch { return fallback; }
}
function publicView(j) {
  const meta = safeMeta(j.meta);
  const roots = safeJson(j.roots, []);
  const stepPaths = safeJson(j.step_paths, []);
  return {
    id: j.id,
    original_name: j.original_name,
    status: j.status,
    attempts: j.attempts,
    error: j.error,
    meta,
    roots,
    selected_roots: Array.isArray(meta.selected_roots) ? meta.selected_roots : null,
    steps: stepPaths,
    step_count: stepPaths.length,
    queue_position: j.status === 'pending' ? queuePosition(j.id) + 1 : null,
    created_at: j.created_at,
    started_at: j.started_at,
    finished_at: j.finished_at,
    download_url: j.status === 'done' && stepPaths.length ? `/api/download/${j.id}` : null,
  };
}
function safeMeta(m) {
  try { return JSON.parse(m || '{}'); } catch { return {}; }
}

// 把上传的文件落盘到 uploads/<jobId>/，返回文件清单（含落盘后的真实路径）
// 把 multer 落盘的文件封装为清单（此时文件仍在 multer 默认目录，尚未归置）
function stageFiles(req) {
  const files = req.files || [];
  return files.map((f) => ({ filename: f.originalname, path: f.path }));
}

// 分类：.CATProduct = 产品（装配）；.CATPart = 零件
// heuristicRoots：上传时无法启动 CATIA 扫描引用，先用“所有产品即根（无产品则所有零件即根）”
// 作为前端展示/默认勾选；真正的根由驱动在转换时通过引用扫描判定。
function classifyAndPick(staged) {
  const products = staged.filter((f) => /\.catproduct$/i.test(f.filename));
  const parts = staged.filter((f) => /\.catpart$/i.test(f.filename));
  const productNames = products.map((p) => p.filename);
  const partNames = parts.map((p) => p.filename);
  const heuristicRoots = productNames.length ? productNames : partNames;
  return {
    products: productNames,
    parts: partNames,
    roots: heuristicRoots,
    selected_product: heuristicRoots[0] || null,
  };
}

// 批量上传：一次提交多个文件（零件 + 产品），组成一个“模型”任务
// 注意：直接归置到 uploads/<jobId>/，不再经过临时目录，避免触发环境的 rm 拦截。
app.post('/api/upload-batch', upload.array('files', 100), (req, res) => {
  const files = req.files || [];
  if (!files.length) return res.status(400).json({ error: '未收到文件' });
  const staged = stageFiles(req);
  const { products, parts, roots, selected_product } = classifyAndPick(staged);
  if (!selected_product) {
    // 无法判定目标：把 multer 临时文件删掉（若环境拦截 rm，则忽略，留待清理）
    staged.forEach((s) => { try { fs.rmSync(s.path, { force: true }); } catch (_) {} });
    return res.status(400).json({ error: '未识别到可转换的文件（需至少一个 .CATPart 或 .CATProduct）' });
  }

  // 先建任务拿 jobId，再用 jobId 目录作为最终归置位置
  const jobId = createJob({
    original_name: staged[0].filename,
    stored_name: selected_product,
    input_path: '',
    meta: { files: [], products, parts, roots, selected_roots: null, selected_product },
  });
  const finalDir = path.join(config.uploadsDir, String(jobId));
  fs.mkdirSync(finalDir, { recursive: true });
  for (const s of staged) {
    const nd = path.join(finalDir, s.filename);
    try { fs.renameSync(s.path, nd); } catch (e) {
      return res.status(500).json({ error: '归置上传文件失败: ' + e.message });
    }
    s.path = nd;
  }
  const selRec = staged.find((s) => s.filename === selected_product);
  if (selRec) db.prepare(`UPDATE jobs SET input_path=? WHERE id=?`).run(selRec.path, jobId);
  updateMeta(jobId, { files: staged, products, parts, roots, selected_roots: null, selected_product });
  // 上传即写入启发式根列表（产品优先）；驱动完成后会用实际根覆盖
  db.prepare(`UPDATE jobs SET roots=? WHERE id=?`).run(JSON.stringify(roots), jobId);

  res.json({ jobId, status: 'pending', products, parts, roots, selected_product });
});

// 改选“要导出的根产品”（可多选）并重新入队。不传 roots 或传空数组 = 恢复自动探测。
app.post('/api/jobs/:id/select-roots', (req, res) => {
  const j = getJob(req.params.id);
  if (!j) return res.status(404).json({ error: '任务不存在' });
  const meta = safeMeta(j.meta);
  const allRoots = (meta.products && meta.products.length) ? meta.products : (meta.parts || []);
  let roots = Array.isArray(req.body && req.body.roots) ? req.body.roots : null;
  if (roots) {
    roots = roots.filter((r) => allRoots.includes(r));
    if (!roots.length) return res.status(400).json({ error: '未提供有效的根产品' });
  }
  meta.selected_roots = roots; // null => 自动探测
  updateMeta(j.id, meta);
  db.prepare(`UPDATE jobs SET status='pending', started_at=NULL, error=NULL, step_path=NULL, step_paths=NULL, finished_at=NULL WHERE id=?`).run(j.id);
  res.json({ jobId: j.id, status: 'pending', roots: roots || 'auto', selected_roots: roots });
});

// 分页任务列表（按 id 倒序，最新在前）。默认每页 50（上限 100），返回 { items, total, page, pageSize }
app.get('/api/jobs', (req, res) => {
  const page = Math.max(parseInt(req.query.page, 10) || 1, 1);
  const pageSize = Math.min(Math.max(parseInt(req.query.pageSize, 10) || 50, 1), 100);
  const total = countJobs();
  const items = listJobs(pageSize, (page - 1) * pageSize).map(publicView);
  res.json({ items, total, page, pageSize });
});

app.get('/api/jobs/:id', (req, res) => {
  const j = getJob(req.params.id);
  if (!j) return res.status(404).json({ error: '任务不存在' });
  res.json(publicView(j));
});

function jobSteps(j) {
  const arr = safeJson(j.step_paths, []);
  return arr.filter((p) => typeof p === 'string' && fs.existsSync(p))
    .map((p) => ({ path: p, name: path.basename(p) }));
}

// 下载单个 STEP：/api/download/:id/:idx （idx 为该任务内第几个 STEP，从 0 开始）
// 不带 idx：若只有一个 STEP 直接下载；若有多个则打包成 zip。
app.get('/api/download/:id/:idx?', (req, res) => {
  const j = getJob(req.params.id);
  if (!j) return res.status(404).json({ error: '任务不存在' });
  if (j.status === 'expired') return res.status(404).json({ error: '文件已过期，无法下载' });
  if (j.status !== 'done') return res.status(404).json({ error: 'STEP 尚未就绪' });
  const steps = jobSteps(j);
  if (!steps.length) return res.status(404).json({ error: 'STEP 文件丢失' });

  const idxParam = req.params.idx;
  if (idxParam == null) {
    if (steps.length === 1) {
      return res.download(steps[0].path, steps[0].name);
    }
    // 多个 → 打包 zip
    res.setHeader('Content-Type', 'application/zip');
    res.setHeader('Content-Disposition', `attachment; filename="job-${j.id}-steps.zip"`);
    const zip = new yazl.ZipFile();
    zip.outputStream.on('error', (err) => { res.destroy(err); });
    zip.outputStream.pipe(res);
    for (const s of steps) zip.addFile(s.path, s.name);
    zip.end();
    return;
  }
  const idx = parseInt(idxParam, 10);
  if (!Number.isInteger(idx) || idx < 0 || idx >= steps.length) {
    return res.status(404).json({ error: 'STEP 索引越界' });
  }
  const s = steps[idx];
  res.download(s.path, s.name);
});

// 跨任务批量下载：收集所有任务的全部 STEP 打包
app.get('/api/download-batch', (req, res) => {
  const ids = String(req.query.ids || '').split(',')
    .map((s) => parseInt(s, 10)).filter((n) => Number.isInteger(n));
  const steps = [];
  for (const id of ids) {
    const j = getJob(id);
    if (!j) continue;
    for (const s of jobSteps(j)) steps.push({ ...s, jobId: j.id });
  }
  if (!steps.length) return res.status(404).json({ error: '没有可下载的 STEP' });

  res.setHeader('Content-Type', 'application/zip');
  res.setHeader('Content-Disposition', 'attachment; filename="catia-steps.zip"');
  const zip = new yazl.ZipFile();
  zip.outputStream.on('error', (err) => { res.destroy(err); });
  zip.outputStream.pipe(res);
  for (const s of steps) {
    zip.addFile(s.path, `${s.jobId}_${s.name}`);
  }
  zip.end();
});

// 错误处理（multer 过滤 / 大小限制等）
app.use((err, req, res, next) => {
  if (res.headersSent) return next(err);
  const msg = err && err.message ? err.message : '服务器错误';
  const status = /文件|仅支持/.test(msg) ? 400 : 500;
  res.status(status).json({ error: msg });
});

startCleanup();
if (config.startWorker) startWorker();
app.listen(config.port, () => {
  console.log(`catia-stp-converter listening on http://localhost:${config.port}  (mock=${config.mock})`);
});
