# CATIA 模型转 STEP (catia-stp-converter)

Web 服务：上传 CATIA 三维模型（`.CATPart` 零件 / `.CATProduct` 产品），通过驱动本机 CATIA V5 自动转换为 STEP（`.stp`）。

- **批量上传**：一次提交一个模型的全部文件（产品 + 被引用的零件），后台串行逐个转换
- **多根自动探测**：批次含多个 `.CATProduct` 时，驱动会扫描各产品的引用关系，找出「不被其它产品引用」的根装配，每个根各自导出一个 STEP；前端可对要导出的根做勾选覆盖
- **批量下载**：勾选多个已完成任务，或同一任务内的多个 STEP，打包成 zip 下载
- **任务队列**：SQLite 持久化 + 单 Worker 串行，避免多个 CATIA 实例争抢 License

## 技术栈

| 层 | 选型 | 说明 |
|---|---|---|
| 语言 | Node.js ≥ 18（CommonJS） | 服务端全部 JavaScript，无 TS / 构建步骤 |
| Web 框架 | Express 4 | 上传 / 下载 / 任务列表 API |
| 数据库 | SQLite（better-sqlite3，WAL 模式） | 任务队列持久化，同步 API，文件在 `data/jobs.db` |
| 文件上传 | multer | 多文件上传，落盘到 `uploads/<jobId>/` |
| zip 打包 | yazl | 批量下载时把多个 STEP 打成 zip |
| 前端 | 原生 HTML / CSS / JS（`public/index.html`） | 单文件页面，无框架、无打包 |
| 转换驱动 | VBScript（cscript 运行） | COM IDispatch 晚绑定驱动 CATIA；.NET COM 互操作在本机不可用，故不用 PowerShell |
| 转换引擎 | CATIA V5（本机安装） | 交互会话扫描引用 + 官方无头批处理 CATBatchStarter 导出 STEP |

## 工作原理

```
浏览器 ──POST /api/upload-batch──▶ Express ──▶ SQLite 队列(jobs.db)
                                                 │
                              单 Worker 串行认领 ─┘
                                                 ▼
                  复制整批文件到 %TEMP% 干净目录（扁平，便于解析引用）
                                                 ▼
                    cscript.exe → convert.vbs（VBScript 驱动）
                                                 ▼
      （需要时）启动 CATIA 会话扫描引用 → 判定「根产品」
                                                 ▼
           逐个根：写单文件 BatchDataExchange XML → 调
           CATSTART -run "CATBatchStarter -input ..." 无头导出 STEP
             （每个根产品把整棵装配树：零件+子产品 导出为一个 .stp）
                                                 ▼
                        结果写回 results/<jobId>/，前端可对每个 STEP 下载
```

关键设计说明：

- **导出走 CATIA 官方无头批处理（Batch DXF-IGES-STEP），而非 `Document.ExportData(..., "stp")`**：本机（该服务实际部署机）上 `ExportData` 对 STEP 持续 E_FAIL（连空零件也失败），而 IGES 导出与 GUI 手动导出均正常；排除证书/环境差异后确认是 API 路径问题，官方批处理走独立的 C++ 转换路径可用。驱动逐根生成一份**单文件** batch XML，经 `CATSTART -run "CATBatchStarter -input <xml>"` 无头转换，产出 AP242 文本 STEP（`ISO-10303-21;` HEADER/DATA 结构）。该批处理 XML 只接受一个输入文件（多文件返回 RC=208），故每个根单独一次调用。
- **完成判定看文件、不看进程**：轮询目标 `.stp` 是否「存在且两次采样（间隔 5s）大小稳定」判定写完；不依赖批处理进程退出——其包装进程 CATSTART 可能在文件写完很久后仍挂着，且占住 stdout 管道会导致阻塞读取死锁（曾踩过坑）。
- **手动根严格覆盖自动探测**：只要手动指定了根（`selected_roots` 非 null），驱动不再启动 CATIA 会话扫描，直接导出勾选项。
- **用 VBScript 而不是 PowerShell**：沿用 catia-pdf-converter 的工程结论——部分机器上 .NET 的 COM 互操作层对 `CATIA.Application` 返回死对象，而 VBScript 的 IDispatch 晚绑定正常。（另本机实测 VBScript 的 `Like` 运算符报「未定义 Sub 或 Function」，脚本统一用 `InStr`/`Left` 做字符串匹配。）
- **干净临时目录**：CATIA V5 无法打开含非常规 Unicode 字符（如 U+2011 不间断连字符）的路径，Node 侧先把整批文件复制到 `%TEMP%\catia-stp-*` 再转换；被产品引用的零件/子产品一并复制，CATIA 才能在同目录就近解析引用。
- **多产品怎么转（核心问题）**：

  CATIA 三维模型分两类：`CATPart`（零件）和 `CATProduct`（产品/装配）。一个 `CATProduct` 通过引用把若干 `CATPart`（甚至子 `CATProduct`）组装起来。**转换成 STEP 时，应当转换「顶层的根产品」**——它会把整棵装配树（零件 + 子产品）一并导出为一个 STEP。

  上传多个产品时，处理策略：

  1. 按扩展名分类：`CATProduct` → 产品，`CATPart` → 零件。
  2. **驱动自动探测根**（仅在**未手动指定根**且存在 CATProduct 时执行；纯零件批次不启动 CATIA 会话）：转换时在 VBScript 里用 CATIA 打开每个产品，遍历 Product 树收集「被引用的文档名」，构建「谁引用谁」的图；**不被任何其他产品引用的即为顶层根装配**。
  3. **单根 → 导出一个 STEP**（囊括整棵装配树）；**多个独立根 → 各自导出一个 STEP**，并按 `<jobId>_<根名>.stp` 一并打包。
  4. **手动覆盖（严格优先）**：前端列出检测到的根，可勾选/取消要导出的根，点「重转勾选的根」触发重新转换；只要手动指定了根，驱动就不做自动探测、直接导出勾选项。**恢复自动探测 = 把该任务 `selected_roots` 置回 null**（接口不传 `roots` 或不再调用该接口）；前端「全取消勾选」提交的是空数组，会返回 400，并不会恢复自动探测。
  5. 没有产品、只有零件：每个零件各导出一个 STEP。

  > 说明：MOCK 模式（无 CATIA）无法扫描引用，会保守地把「所有产品都当作根」各自导出，以保证不漏转；真实模式才做精确的根探测。

## 环境要求

- Windows（CATIA V5 仅支持 Windows）
- 已安装 CATIA V5，且证书（License）可用（需 **HD2 + ST1**，见下节）
- Node.js ≥ 18

## CATIA 证书（License）

- **所需许可项：HD2 + ST1**（部署机实测，两个都要）：
  - **HD2**（CATIA V5 Design Product 配置）：打开/装配 `CATPart`、`CATProduct` 所需。自动探测根时驱动启动的 CATIA 扫描会话（`convert.vbs` 里 `Documents.Open` 遍历引用）依赖它。
  - **ST1**（STEP 3D 数据交换接口）：STEP 导出所需。导出走 Batch DXF-IGES-STEP（`CATBatchStarter`）产出 AP242 STEP，靠的就是它。
  - 纯零件批次虽不启动扫描会话，但导出同样要经批处理打开并转换文档，两个许可仍都用到；缺任一项，任务会转换失败。MOCK 模式不调 CATIA，不占证书。
- **前提**：部署机上的 CATIA V5 必须已正确授权，无头批处理（CATBatchStarter）与交互会话一样需要证书可用。建议部署时先手动打开一次 CATIA，确认能正常拿到 License，再启动本服务。
- **占用方式**：本服务按「同一时刻至多一个 CATIA 实例」设计——SQLite 队列单 Worker 串行认领，避免多个 CATIA 实例同时争抢证书导致取不到。
- **证书释放**：Worker 在启动时与每个任务结束后都会 `taskkill /IM CNEXT.exe /F` 强杀残留 CATIA 释放证书；自动探测根的扫描会话用完即 `catia.Quit`。因此服务所在登录会话不要同时被他人手动使用 CATIA（会被误杀，见「已知说明」）。
- **排查提示**：转换失败若怀疑是证书问题，可先在同一账号下用 GUI 手动执行一次同样的导出验证。本部署机曾出现 STEP `ExportData` 持续 E_FAIL，已排除证书原因（是 API 路径问题，故改走官方批处理，见「工作原理」）。

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
| `CATIA_TIMEOUT_SEC` | `900` | 单个任务转换超时（秒），超时强杀 CATIA；逐根批处理启动 + 大零件导出耗时高（200MB 级约 2~6 分钟/根） |
| `MAX_ATTEMPTS` | `2` | 失败重试次数（不含首次） |
| `POLL_INTERVAL_MS` | `2000` | Worker 空闲轮询间隔（毫秒） |
| `MAX_FILE_MB` | `500` | 单文件上传大小上限（MB）；实测大 CATPart 可达 273MB |
| `START_WORKER` | `true` | 设为 `false` 时不随 Web 进程启动 Worker |
| `MOCK` | — | 设为 `1` 时不调 CATIA，生成占位 .step（验证队列链路用） |
| `RETENTION_DAYS` | `15` | 文件保留天数（从提交时间起算），到期自动清理 |
| `CLEANUP_INTERVAL_MS` | `3600000` | 过期清理周期（毫秒），默认每小时 |

## API

| 方法 | 路径 | 说明 |
|---|---|---|
| POST | `/api/upload-batch` | 多文件上传（字段名 `files`，支持 .CATPart/.CATProduct），返回 `{ jobId, products, parts, roots }` |
| POST | `/api/jobs/:id/select-roots` | 改选要导出的根（body `{roots:[...]}`），任务重新入队；**不传 `roots`**（`selected_roots` 置 null）恢复自动探测，传空数组返回 400 |
| GET | `/api/jobs?page=1&pageSize=50` | 分页任务列表（按 id 倒序，最新在前；`page` 从 1 起，`pageSize` 默认 50、上限 100），返回 `{ items, total, page, pageSize }`，每条含状态、根列表、STEP 列表、队列位、提交/完成时间 |
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
    convert.vbs     # VBScript 驱动：需要时启会话扫描引用判根 → 逐根写单文件 batch XML → CATBatchStarter 导出
public/
  index.html        # Web 前端（批量上传 / 根产品勾选覆盖 / 多 STEP 下载）
data/               # 运行期 SQLite 数据库（不入库）
uploads/            # 上传的模型文件（按 jobId 归档，不入库）
results/            # 转换出的 STEP（不入库）
```

## 已知说明

- Worker 会执行 `taskkill /IM CNEXT.exe /F` 强制清理 CATIA，请确保该服务运行在**专用**登录会话，避免误杀同一台机器上其他人打开的 CATIA。
- 部署机的 VBScript 环境有两处坑：`.vbs` 须保持**纯 ASCII**（cscript 按 ANSI 解析，UTF-8 中文注释会吞掉换行符）；`Like` 运算符不可用（报「未定义 Sub 或 Function」）。改动 `convert.vbs` 时遵守这两条。
- 数据库文件（`data/jobs.db`）、上传模型、转换结果均通过 `.gitignore` 排除，不入 Git。
