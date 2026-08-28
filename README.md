# CATIA 模型转 STEP (catia-stp-converter)

Web 服务：上传 CATIA 三维模型（`.CATPart` 零件 / `.CATProduct` 产品），通过驱动本机 CATIA V5 自动转换为 STEP（`.stp`）。

- **批量上传**：一次提交一个模型的全部文件（产品 + 被引用的零件），后台串行逐个转换
- **多根自动探测**：批次含多个 `.CATProduct` 时，驱动会扫描各产品的引用关系，找出「不被其它产品引用」的根装配，每个根各自导出一个 STEP；前端可对要导出的根做勾选覆盖
- **批量下载**：勾选多个已完成任务，或同一任务内的多个 STEP，打包成 zip 下载
- **任务队列**：SQLite 持久化 + 单 Worker 串行，避免多个 CATIA 实例争抢 License

## 工作原理

```
浏览器 ──POST /api/upload-batch──▶ Express ──▶ SQLite 队列(jobs.db)
                                                 │
                              单 Worker 串行认领 ─┘
                                                 ▼
                  复制整批文件到 %TEMP% 干净目录（扁平，便于解析引用）
                                                 ▼
                    cscript.exe → convert.vbs（VBScript）
                                                 ▼
                    CATIA V5 扫描各产品引用 → 判定「根产品」→ 逐个导出 STEP
                      （每个根产品把整棵装配树：零件+子产品 导出为一个 .stp）
                                                 ▼
                        结果写回 results/<jobId>/，前端可对每个 STEP 下载
```

关键设计说明：

- **用 VBScript 而不是 PowerShell**：沿用 catia-pdf-converter 的工程结论——部分机器上 .NET 的 COM 互操作层对 `CATIA.Application` 返回死对象，而 VBScript 的 IDispatch 晚绑定正常。
- **干净临时目录**：CATIA V5 无法打开含非常规 Unicode 字符（如 U+2011 不间断连字符）的路径，Node 侧先把整批文件复制到 `%TEMP%\catia-stp-*` 再转换；被产品引用的零件/子产品一并复制，CATIA 才能在同目录就近解析引用。
- **多产品怎么转（核心问题）**：

  CATIA 三维模型分两类：`CATPart`（零件）和 `CATProduct`（产品/装配）。一个 `CATProduct` 通过引用把若干 `CATPart`（甚至子 `CATProduct`）组装起来。**转换成 STEP 时，应当转换「顶层的根产品」**——它会把整棵装配树（零件 + 子产品）一并导出为一个 STEP。

  上传多个产品时，处理策略：

  1. 按扩展名分类：`CATProduct` → 产品，`CATPart` → 零件。
  2. **驱动自动探测根**：转换时在 VBScript 里用 CATIA 打开每个产品，遍历 Product 树收集「被引用的文档名」，构建「谁引用谁」的图；**不被任何其他产品引用的即为顶层根装配**。
  3. **单根 → 导出一个 STEP**（囊括整棵装配树）；**多个独立根 → 各自导出一个 STEP**，并按 `<jobId>_<根名>.stp` 一并打包。
  4. **手动覆盖**：前端列出检测到的根，可勾选/取消要导出的根（或勾选部分），「重转勾选的根」触发重新转换；不勾选任何 / 不改则恢复自动探测。
  5. 没有产品、只有零件：每个零件各导出一个 STEP。

  > 说明：MOCK 模式（无 CATIA）无法扫描引用，会保守地把「所有产品都当作根」各自导出，以保证不漏转；真实模式才做精确的根探测。

## 环境要求

- Windows（CATIA V5 仅支持 Windows）
- 已安装 CATIA V5
- Node.js ≥ 18

## 安装与启动

```bash
npm install
npm start          # 启动 Web 服务 + Worker（默认 http://localhost:3000）
```

浏览器打开 http://localhost:3000 即可上传转换。

如果 Web 服务与 Worker 分开部署：

```bash
node server/index.js   # 只启动 Web
node server/worker.js  # 只启动 Worker
```

## 配置（环境变量）

| 变量 | 默认 | 说明 |
|---|---|---|
| `PORT` | `3000` | Web 服务端口 |
| `CATIA_TIMEOUT_SEC` | `300` | 单个任务转换超时（秒），超时强杀 CATIA |
| `MAX_ATTEMPTS` | `2` | 失败重试次数（不含首次） |
| `POLL_INTERVAL_MS` | `2000` | Worker 空闲轮询间隔（毫秒） |
| `MAX_FILE_MB` | `200` | 单文件上传大小上限（MB） |
| `START_WORKER` | `true` | 设为 `false` 时不随 Web 进程启动 Worker |
| `MOCK` | — | 设为 `1` 时不调 CATIA，生成占位 .step（验证队列链路用） |
| `RETENTION_DAYS` | `15` | 文件保留天数（从提交时间起算），到期自动清理 |
| `CLEANUP_INTERVAL_MS` | `3600000` | 过期清理周期（毫秒），默认每小时 |

## API

| 方法 | 路径 | 说明 |
|---|---|---|
| POST | `/api/upload-batch` | 多文件上传（字段名 `files`，支持 .CATPart/.CATProduct），返回 `{ jobId, products, parts, roots }` |
| POST | `/api/jobs/:id/select-roots` | 改选要导出的根（body `{roots:[...]}`），任务重新入队；传空数组/null 恢复自动探测 |
| GET | `/api/jobs` | 任务列表（含状态、根列表、STEP 列表、队列位、提交/完成时间） |
| GET | `/api/jobs/:id` | 单个任务状态 |
| GET | `/api/download/:id` | 该任务 STEP：仅 1 个直接下载；多个则打包 zip |
| GET | `/api/download/:id/:idx` | 下载该任务内第 `idx` 个 STEP（从 0 起） |
| GET | `/api/download-batch?ids=1,2,3` | 批量下载选中任务的全部 STEP（zip，`<jobId>_<根名>.stp`） |

示例批量上传：

```bash
curl -F "files=@assy.CATProduct" -F "files=@part1.CATPart" -F "files=@part2.CATPart" http://localhost:3000/api/upload-batch
```

## 目录结构

```
server/
  index.js          # Express + 上传/下载/选产品 API + 启动定时清理
  worker.js         # 串行任务 Worker
  db.js             # SQLite 队列（含 meta：文件清单/产品/零件/所选产品）
  config.js         # 配置（环境变量）
  cleanup.js        # 过期文件清理
  catia.js          # CATIA 转换编排：临时目录、spawn cscript、超时看门狗、多根解析、killCatia
  catia-driver/
    convert.vbs     # VBScript 驱动：CreateObject → 扫描各产品引用 → 判定根 → 逐个导出 STEP
public/
  index.html        # Web 前端（批量上传 / 根产品勾选覆盖 / 多 STEP 下载）
data/               # 运行期 SQLite 数据库（不入库）
uploads/            # 上传的模型文件（按 jobId 归档，不入库）
results/            # 转换出的 STEP（不入库）
```

## 已知说明

- Worker 会执行 `taskkill /IM CNEXT.exe /F` 强制清理 CATIA，请确保该服务运行在**专用**登录会话，避免误杀同一台机器上其他人打开的 CATIA。
- 数据库文件（`data/jobs.db`）、上传模型、转换结果均通过 `.gitignore` 排除，不入 Git。
