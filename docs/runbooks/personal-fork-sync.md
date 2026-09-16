# 个人 fork 同步与 Excel 适配

2026-09-17：将个人版本 `8aba26b` 与上游 `e2ebabe81e066e0498bdb58000942ef5199d338c`
合并，保留双方 Git 历史。正式运行采用上游 Vue、Thread/Run/Event、鉴权、Provider、原生
Milvus BM25、分页和独立持久 worker；移除旧平铺 backend、旧静态页面、旧 Excel store 与
旧入口。未执行旧数据库、旧索引或历史运行数据转换，未启动基础设施或修改个人 Secret。

## Excel 的正式路径

上传仍只创建 durable Index Job。Worker 调用 `backend.indexing.document_loader.DocumentLoader`，
在 candidate Document Version scope 中解析、持久写入、核验 manifest 并原子发布。

- `.xlsx` 使用 openpyxl 只读解析，保留 Sheet、字符串单元格、多行内容、Markdown 样文本、
  公式原文、原始行号、空值与有序表头；不计算公式，也不执行单元格内容。
- 第一条非空行作为表头；空表头生成 `column_N`，重复表头生成唯一后缀，并保留原有带后缀的
  列名，避免 `[价格,价格,价格_2]` 导致字典值覆盖。全空数据行跳过；只有表头的表仍可检索。
- 上游文本净化规则继续适用（如移除零宽字符）；值与旧实现一样规范为去除首尾空白的字符串。
- L3 行叶子的 `text` 是紧凑 JSON：`sheet`、`row`、`values`；L2 父级表格分页包含
  `sheet`、`sheet_index`、`headers`、`rows`。JSON 与引用 identity 一起进入现有 Milvus / ParentChunk
  存储。没有另一个全局 Excel 数据库，也没有上传后按 filename 再补取表格的检索路径。
- 每页至多 25 行，JSON 至多 `min(MAX_PAGE_CHARACTERS, 12000)` 字符及 60000 UTF-8 字节；
  Sheet 数与总表格分页数均受 `MAX_DOCUMENT_PAGES` 限制。单行或表头超出页容量明确解析失败，
  不截断单元格、不生成部分 JSON。空白 Sheet 不产生索引。引用 page 是工作簿内的表格分页，
  原 Sheet 序号、Sheet 名与 Excel 原始行号保留在内容中。
- RAG 仍按正式 L3 target 检索并按现有阈值展开 L2 父块，遵守冻结的 Tenant / Version scope、
  rerank 和 Evidence context budget。默认正常小表可完整打包；大页仍可能被 Evidence 预算
  省略或标记截断，不保证一次回答召回整张工作簿。
- `.xls` 保留上游 Unstructured 格式解析；结构化增强仅针对 `.xlsx`。这两者按扩展名分派，
  不在 XLSX 失败后静默切换 parser。本次轻量环境未验证真实 `.xls` 依赖栈。

默认 parser 为 `document-loader-v3-excel-json`，chunker 为 `three-level-800-100-excel-pages-v3`。
未来新部署若沿用旧配置，必须检查显式 `DOCUMENT_PARSER_VERSION` / `DOCUMENT_CHUNKER_VERSION`
override 并更新或移除；本次只更新 `.env.example`，不读写真实 `.env`。API 与 worker 必须
使用同一 build profile。现有旧数据不会自动迁移或触发 collection drop；需要当前索引时，
在部署环境完成上游迁移前置检查后通过正式上传流程重新构建。

## 学习资料与历史测试

`langchain-study/` 的 13 个 Python/notebook、`experiments/loader_probe/` 的 README、脚本和
两个样例文件共 17 个资产按个人提交逐字保留。它们是历史学习材料，不是当前生产入口；
probe README 中的旧说明和学习脚本依赖不构成当前架构说明。本次未执行会生成或覆盖样例的
probe，也未执行可能调用在线模型的 notebook。学习代码的环境、API 和依赖需要另行核对。

为保持原文，这两个明确的学习目录不参加 Ruff 格式/lint；生产模块和正式测试仍执行完整规则。
pytest 的正式发现目录为 `tests/`，不会执行学习脚本。

- 旧 Excel 测试的原文保存在[历史快照](../history/excel-loader-test-8aba26b.md)，其行为已迁移为
  `tests/test_excel_structured_loader.py`，并新增版本化 pipeline 回归。
- 旧根目录 LangSmith 脚本改为[历史源码文档](../history/langsmith-eval-8aba26b.md)，不再导入已删除
  的旧 Agent 或在测试发现时联网。当前评测使用 `backend/evaluation` 与 `scripts/evaluate_rag.py`。
- 工具运行日志 `.omx/logs/` 不保留在当前树；原始提交仍可查看。独立的 MemoryAgentBench 数据集
  仓库不属于此合并，不复制或提交其内容。

## 离线验证方式与边界

从仓库根目录使用 Python 3.12 的独立 `.venv`。本次为避免安装大型本地模型运行库，使用：

```powershell
$env:APP_ENV = "test"
$env:DATABASE_URL = "sqlite+pysqlite:///:memory:"
$env:JWT_SECRET_KEY = "offline-test-only-synthetic-signing-key"
$env:PYTHON_DOTENV_DISABLED = "1"
$env:HF_HUB_OFFLINE = "1"
$env:TRANSFORMERS_OFFLINE = "1"
uv sync --dev --locked --python 3.12 --no-install-package torch --no-install-package sentence-transformers --no-install-package transformers --no-install-package spacy --no-install-package unstructured
uv run --no-sync pytest -q tests/test_excel_structured_loader.py tests/test_excel_versioned_pipeline.py
uv run --no-sync ruff format --check .
uv run --no-sync ruff check .
uv run --no-sync mypy
```

这是锁定依赖的轻量测试环境，不是完整运行环境；完整安装仍使用上游 `uv sync --dev --locked`。
新增 Excel 回归使用真实生成的 XLSX、真实 SQLite Catalog/ParentChunk、正式 Publication、
MilvusWriter 与 RAG retrieval/Evidence；Embedding 和 Milvus transport 使用 fake Adapter。
测试涵盖分页、碰撞表头、原始公式/换行、空表、限额失败、metadata 隔离、持久化后父块重载、
引用以及失败 candidate 不替换已发布版本。没有真实模型、付费 LLM 或生产数据库调用。

Windows 原生缺少 `O_DIRECTORY` / `O_NOFOLLOW`，上游 SkillRegistry 明确 fail-closed。
依赖其初始化的完整应用/Registry 门禁无法在本环境通过；不通过补写标志或绕过路径校验解决。
应在具备安全文件打开能力的 Linux 环境执行完整上游门禁。真实 PostgreSQL / Redis / Milvus、
Docker sandbox、模型 smoke、在线评测和部署均需要另行验证。

RAG baseline 按正式 score 命令重建。解析与 publication 源码 fingerprint 改变，因此生成的
baseline metadata 更新；Dataset、静态 observations、Gate 和指标不修改。这仍是
`contract_smoke`，不是真实 Excel 检索质量或在线回答质量结论。

### 本次实际结果

以下按各次执行分别记录，部分集合重叠，不累加为独立用例数：

| 验证 | 结果 |
| --- | --- |
| 新增 Excel loader + versioned pipeline | 8 passed |
| Excel、artifact、Catalog、Publication、Worker、ParentChunk、MilvusWriter、RAG parent/targets | 110 passed |
| SQLite 迁移、RAG eval/trace/evidence/short-circuit/fault、Provider bridge/retry、Rerank | 109 passed，11 subtests passed |
| Ruff check / format | 全仓生产范围通过，303 files already formatted |
| mypy | 上游配置的 61 文件通过；另单独检查修改的 loader 通过 |
| Contract / RAG schema 生成检查 | 通过 |
| RAG score / baseline 比较 | 通过；只有生成的 source fingerprint 变化 |
| 非模型 Runtime benchmark | 全部预算通过 |

额外尝试全仓 `pytest --continue-on-collection-errors --cov=backend --cov=scripts`：
897 passed、5 skipped、51 failed、15 collection errors、108 subtests passed；coverage 71.43%，
未达到上游 80% 门禁。完整门禁未通过，不能把通过的离线子集代替完整发布验证。

全仓失败暴露了一个新增测试的 import-shape 问题，现已改为绝对模块导入，并定向复测通过。
测试 JWT 未配置造成的一项失败，在显式注入合成测试 key 后通过。该轮定向重测
（Excel、import-shape、auth、endpoint I/O、upload security、embedding runtime）为
41 passed、2 failed、5 subtests passed；剩余两个失败是未修改的上游 endpoint I/O 测试
要求 40ms 内 `ticks > 3`，本 Windows 环境观测为 2。全仓运行中的另外两个时间敏感失败
在定向重测时通过；没有放宽断言或改写上游逻辑。

其余完整门禁阻塞包括前述 SkillRegistry POSIX flags、Unix `resource` 模块、Windows
symlink 权限、POSIX shell 启动；Registry validate 同样受安全 flags 限制，生产 Compose
检查脚本使用 `/dev/null`，在 Windows 解析为无效 env-file 路径。本次没有修改这些平台
边界，也没有启动服务来绕过检查。全仓 coverage 结果因这些缺口不构成 Linux 发布证据。
