# 季度简报辅助系统（Briefing Support System）

从公募规模 **底稿 Excel（终表）** 自动生成：

1. **表图 Excel**（`*_tables.xlsx`）——带数据条 / 色阶 / 示例高亮，**主产物**，人工微调后贴进 PPT  
2. **PPT 文字初稿**（`*_briefing.pptx`）——数字从终表读出自动成文；表图默认不自动贴

原则：**数字以底稿终表为准**，避免数据日手抄与双路径计算导致口径漂移。

---

## 文档入口（接手先看这里）

| 文档 | 给谁看 | 内容 |
|------|--------|------|
| **[出报指引.md](出报指引.md)** | 出报操作员 | 日常流程、底稿规范、粘贴清单、验收勾选 |
| **[docs/需求与维护说明.md](docs/需求与维护说明.md)** | 开发 / 维护 | 需求背景、模块职责、改哪里、故障排查 |
| 本文 README | 所有人 | 安装、快速开始、目录结构、命令一览 |

网页操作台侧栏也可打开「出报指引 / README」。

---

## 快速开始

### 1. 安装

```bash
cd /path/to/BriefingSupportSystem
python -m pip install -r requirements.txt
```

建议 Python 3.10+。依赖见 `requirements.txt`（pandas / openpyxl / python-pptx / FastAPI 等）。

### 2. 推荐用法：网页操作台

```bash
python main.py serve --open
```

浏览器打开后：上传底稿 → 确认期别与日期 → 生成 → 分别下载 Excel / PPT。  
有校验 **error** 时禁止生成；请先改底稿再重新上传。

### 3. 命令行等价流程

```bash
# 1) 校验底稿（不改文件）
python main.py validate-draft --xlsx "/path/to/底稿.xlsx" --strict

# 2) 导出表图 Excel（主产物）
python main.py export-tables \
  --xlsx "/path/to/底稿.xlsx" \
  --output output/26H1_tables.xlsx \
  --period 26H1 \
  --focus 示例

# 3) 生成 PPT 文字（从终表读数；默认不贴表）
python main.py full-deck \
  --config config/report_26h1.yaml \
  --data data/26h1_from_draft \
  --draft-xlsx "/path/to/底稿.xlsx" \
  --output output/26H1_briefing.pptx \
  --strict
```

成文后按 [出报指引.md](出报指引.md) 的「终表 → PPT 粘贴清单」人工贴表，并审改观点句。

新一期配置：复制 `config/report_template.yaml`（或上期 `report_*.yaml`），改 `period_label` / `dates` / `focus_company`。

---

## 端到端数据流

```text
底稿.xlsx
  → validate_draft      # 结构 / 增速刻度检查（error 拦生成）
  → normalize_draft     # sheet 名、Top30、表头别名（不重算数字）
  → load_final_tables   # 按列名抽终表
  → deck_workbook       # → *_tables.xlsx   【主产物】
  → table_facts + sheet_narratives
  → pptx_filler         # → *_briefing.pptx 【文字初稿，默认清旧图】
```

期别叙事自动适配：**Q1**（单季）/ **Q3**（单季主导 + YTD）/ **H1·Q4**（期别主导）。详见维护文档。

默认模版与映射：`templates/25Q4_template.pptx` + `config/slide_map_q4.yaml`。

---

## 目录结构

```text
BriefingSupportSystem/
├── main.py                 # 入口：python main.py <command>
├── requirements.txt
├── 出报指引.md              # 操作员文档（权威）
├── README.md               # 本文件
├── config/
│   ├── report_template.yaml    # 新一期配置模板
│   ├── report_*.yaml           # 各期示例配置
│   ├── slide_map_q4.yaml       # 【主】PPT 槽位映射
│   └── slide_map.yaml          # 旧 H1 映射（遗留）
├── templates/
│   ├── 25Q4_template.pptx      # 【主】统一版式母版
│   └── 25H1_template.pptx      # 旧母版（遗留）
├── data/                   # 输入与网页临时任务（见 .gitignore）
│   └── sample_25h1/        # CSV 示例（遗留路径调试用）
├── output/                 # 生成结果（gitignore）
├── docs/
│   ├── 需求与维护说明.md    # 需求 + 模块维护（权威）
│   └── 新一期出报指引.md    # 指向根目录出报指引
└── briefing/               # 业务代码
    ├── cli.py              # 子命令实现
    ├── full_deck.py        # 【主】全册编排
    ├── period_profile.py   # 期别模式（Q1/Q3/H1/Q4）
    ├── models.py           # ReportConfig 等
    ├── importers/          # 底稿校验 / 规范化 / xlsx→CSV
    ├── deck/               # 终表抽取、事实解析、10 页叙事
    ├── excel/              # tables.xlsx、条件格式、双向数据条
    ├── render/             # PPT / HTML 灌版
    ├── web/                # FastAPI 操作台
    ├── analytics/          # 【遗留】CSV 排名计算
    ├── narrative/          # 【遗留】CSV 叙事规则
    ├── pipeline.py         # 【遗留】HTML 生成
    ├── phase1.py           # 【遗留】早期 1–2 页方案
    └── excel_ppt_loop.py   # 【遗留】最小闭环调试
```

**维护时优先改主路径**（`importers` / `deck` / `excel` / `full_deck` / `web`）。标「遗留」的模块仅在无底稿、走 CSV 时回退使用，新需求一般不必动。

改需求速查表见 [docs/需求与维护说明.md](docs/需求与维护说明.md) 第 B.4 节。

---

## 命令一览

| 命令 | 用途 | 日常 |
|------|------|------|
| `serve` | 网页操作台（上传→校验→生成→下载） | ✅ 推荐 |
| `validate-draft` | 检查底稿结构与增速刻度 | ✅ |
| `export-tables` | 底稿 → 表图 Excel | ✅ |
| `full-deck` | 终表 → PPT 文字（可选 `--paste-tables`） | ✅ |
| `import-xlsx` | 底稿 → `fund_metrics.csv`（辅路径） | 按需 |
| `generate` / `phase1` / `excel-loop` | HTML / 早期 PPT / 调试闭环 | 遗留 |

```bash
python main.py --help
python main.py <command> --help
```

---

## 产物与 10 页对应

| 产物 | 说明 |
|------|------|
| `output/{期}_tables.xlsx` | 主产物：各页终表 + 样式 +「使用说明」粘贴清单 |
| `output/{期}_briefing.pptx` | 文字初稿；默认不贴表图 |

| 页 | 内容 |
|----|------|
| 1–4 | 行业整体 / 总规模 / 非货 / 非货增量 |
| 5–6 | 主动权益 / 权益 ETF（左右表） |
| 7–10 | 货币 / 固收 / 固收+ / FOF |

---

## 注意事项

1. **表图以终表为准**；观点句仍建议人工审改，数字句勿手抄。  
2. 同一增速列禁止混用「百分数 `2`」与「小数 `0.02`」，否则会出现 1%→100%。  
3. 默认不自动贴表（跨平台更稳）；需要系统贴图时加 `--paste-tables`（Mac 可选 Excel 导出）。  
4. 关注公司默认为「示例基金 / 示例」，可在 YAML 的 `focus_company` 或 CLI `--focus` 修改。  
5. 本地临时目录（`data/_web_jobs/`、`output/`、`.tmp_*`）不要提交。

---

## CSV 辅路径（非推荐主流程）

若暂无底稿、只有标准化长表，可使用 `data/*/fund_metrics.csv` + `full-deck`（不传 `--draft-xlsx`）。  
字段约定见历史示例：`data/sample_25h1/fund_metrics.csv`。  
**正式出报请走底稿终表路径**，以保证与排名表同口径。

---

## 交接打包（导出整个项目）

在仓库**上一级目录**执行（排除 `.git`、IDE、缓存、临时产物与网页任务目录）：

```bash
cd /path/to/project

zip -r BriefingSupportSystem-handoff.zip BriefingSupportSystem \
  -x 'BriefingSupportSystem/.git/*' \
  -x 'BriefingSupportSystem/.idea/*' \
  -x 'BriefingSupportSystem/**/__pycache__/*' \
  -x 'BriefingSupportSystem/.tmp_*/*' \
  -x 'BriefingSupportSystem/output/*' \
  -x 'BriefingSupportSystem/data/_web_jobs/*' \
  -x 'BriefingSupportSystem/data/_web_uploads/*' \
  -x 'BriefingSupportSystem/tmp_pptx_imgs/*' \
  -x 'BriefingSupportSystem/tmp_pptx_media/*' \
  -x 'BriefingSupportSystem/**/.DS_Store' \
  -x 'BriefingSupportSystem/**/*.pyc'
```

若已把当前改动全部 commit，也可用仅含已跟踪文件的干净归档：

```bash
cd /path/to/project
git archive --format=zip -o ../BriefingSupportSystem-handoff.zip HEAD
```

接手方解压后按本文「快速开始」安装即可。
