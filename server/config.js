'use strict';
const path = require('path');

const rootDir = path.join(__dirname, '..');

module.exports = {
  port: parseInt(process.env.PORT || '3000', 10),
  rootDir,
  uploadsDir: path.join(rootDir, 'uploads'),
  resultsDir: path.join(rootDir, 'results'),
  dataDir: path.join(rootDir, 'data'),

  // VBScript 驱动（cscript 运行）。见 catia.js：PowerShell 的 COM 绑定在本机对 CATIA 失效。
  driverScript: path.join(rootDir, 'server', 'catia-driver', 'convert.vbs'),

  // 单 Worker 串行；CATIA 启动 + 转换可能很慢，给足超时（秒 → 毫秒）
  catiaTimeoutMs: (parseInt(process.env.CATIA_TIMEOUT_SEC || '300', 10)) * 1000,

  // 失败后重试次数（不含首次）
  maxAttempts: parseInt(process.env.MAX_ATTEMPTS || '2', 10),

  // Worker 空闲轮询间隔（毫秒）
  pollIntervalMs: parseInt(process.env.POLL_INTERVAL_MS || '2000', 10),

  // 单文件大小上限（MB）
  maxFileMB: parseInt(process.env.MAX_FILE_MB || '200', 10),

  // 是否随 Web 进程一起启动 Worker（同一台机器时设为 true 最方便）
  startWorker: process.env.START_WORKER !== 'false',

  // MOCK 模式：无 CATIA 时直接生成占位 .step，用于验证整条队列链路
  mock: process.env.MOCK === '1',

  // 文件保留天数（从提交时间起算）：到期后清理上传文件/转换结果并标记任务失效
  retentionDays: parseInt(process.env.RETENTION_DAYS || '15', 10),

  // 过期清理周期（毫秒），默认每小时
  cleanupIntervalMs: parseInt(process.env.CLEANUP_INTERVAL_MS || '3600000', 10),
};
