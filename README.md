# 季度简报辅助系统 (Briefing Support System)

根据公募规模数据，自动生成与现有风格一致的 **PPT 简报初稿**（叙事文字 + 带数据条/色阶的排名表），避免数据日当天在 PPT 里手抄数字出错。

## 当前能力

| 能力 | 状态 |
|------|------|
| 标准化 CSV → 排名/增量计算 | ✅ |
| 规则化叙事（银华点评、竞品位次变化） | ✅ 全册 10 页 |
| Excel 数据表（可打开核对） | ✅ 总规模排名 |
| PPT 表图（数据条负左正右；排名色阶；银华高亮） | ✅ 全册 |
| Notebook / Oracle 直连 | ⏳ 通过导出 CSV 对接 |

**推荐主命令：** `full-deck`（全册 10 页文字 + 表图）。  
调试第 2 页可用：`excel-loop`。

> **新一期怎么用：** 见根目录 [出报指引.md](出报指引.md)。  
> **推荐：** `python main.py serve --open`（上传底稿 → 生成 **PPT 文字稿 + Excel 表图**）。




## 环境安装

```bash
cd /Users/thomasliu/PycharmProjects/BriefingSupportSystem

# 建议在业务 conda 环境中安装
conda activate enhancedindexfundanalysis   # 或你的环境名
python -m pip install -r requirements.txt
```

依赖主要包括：`pandas`、`openpyxl`、`python-pptx`、`Pillow`、`PyYAML`、`Jinja2`。

## 新一期数字出来后怎么用

### 1. 准备数据目录

例如 `data/25Q3/`，其中必须有：

```
data/25Q3/fund_metrics.csv
```

### 2. 准备配置

复制 `config/report_25h1.yaml` 为新文件（如 `config/report_25q3.yaml`），修改周期与日期：

```yaml
report:
  title: "{period_label}公募行业数据简报"
  period_label: "25Q3"
  period_type: "quarter"          # quarter | half_year | year
  data_source_note: "如无特别注明，以下数据口径均未剔除联接，数据来源wind"

dates:
  current: "20250930"             # 报告期末
  previous_quarter: "20250630"    # 上季度末
  year_start: "20241231"          # 年初/上年末

focus_company: "银华基金"
focus_company_short: "银华"

ranking:
  top_n: 30
```

### 3. 生成 PPT（全册 10 页）

```bash
python main.py full-deck \
  --config config/report_25q3.yaml \
  --data data/25Q3 \
  --template templates/25H1_template.pptx \
  --output output/25Q3_briefing.pptx
```

用示例数据试跑：

```bash
python main.py full-deck \
  --config config/report_25h1.yaml \
  --data data/sample_25h1 \
  --template templates/25H1_template.pptx \
  --output output/25H1_full_deck.pptx
```

仅灌第 2 页（调试用）：

```bash
python main.py excel-loop --output output/25H1_excel_loop.pptx
```

### 4. 查看产物

| 文件 | 说明 |
|------|------|
| `output/xxx.pptx` | 全册简报 PPT（10 页文字 + 表图） |
| `output/_full_deck/*.png` | 各页表图（含双向数据条 / 色阶） |
| `output/_full_deck/total_ranking.xlsx` | 总规模排名表，可用 Excel 核对 |

打开方式：

```bash
open output/25H1_full_deck.pptx
```

**页面对应：**

| 页 | 内容 |
|----|------|
| 1 | 公募规模整体情况 |
| 2 | 总规模（含货币）排名 |
| 3 | 非货规模排名 |
| 4 | 非货增量概览 |
| 5 | 主动权益排名 |
| 6 | 权益 ETF（含/不含联接） |
| 7 | 货币排名 |
| 8 | 固收排名 |
| 9 | 固收+排名 |
| 10 | FOF 排名 |
## 数据格式：`fund_metrics.csv`

长表，每行一条「公司 × 品类 × 时点」：

| 字段 | 必填 | 说明 |
|------|------|------|
| `company` | ✅ | 基金公司全称，如 `银华基金`；行业合计用 `__industry__` |
| `category` | ✅ | 见下方品类取值 |
| `date` | ✅ | `YYYYMMDD`，需覆盖期末 / 上季末 / 年初 |
| `aum` | ✅ | 规模（亿元） |
| `new_issue` | 可选 | **本期/YTD 新发贡献**（记在报告期末行；年初行一般为 0） |
| `nav_change` | 可选 | **本期/YTD 净值变化贡献**（同上） |
| `holding_sales` | 可选 | **本期/YTD 持营贡献**（同上） |

**`category` 取值：**

| key | 含义 |
|-----|------|
| `total` | 总规模（含货币） |
| `non_money` | 非货 |
| `money` | 货币 |
| `active_equity` | 主动权益 |
| `passive_equity` | 被动权益 |
| `fixed_income` | 固收 |
| `fixed_income_plus` | 固收+ |
| `fof` | FOF |

示例（节选）：

```csv
company,category,date,aum,new_issue,nav_change,holding_sales
__industry__,total,20241231,322447,0,0,0
__industry__,total,20250630,344167,0,0,0
银华基金,total,20241231,5344,0,0,0
银华基金,total,20250630,5806,0,0,0
银华基金,non_money,20250630,2423,0,0,0
银华基金,money,20250630,3384,0,0,0
```

完整示例见：`data/sample_25h1/fund_metrics.csv`。

### 从现有 Pandas Notebook 导出

在分析脚本末尾增加：

```python
# df 需已整理为上述列
out_dir = Path("data/25Q3")
out_dir.mkdir(parents=True, exist_ok=True)
df.to_csv(out_dir / "fund_metrics.csv", index=False, encoding="utf-8")
```

然后执行 `full-deck` 即可（详见 [出报指引.md](出报指引.md)）。

## 其他命令

```bash
# 生成 HTML 简报草稿（五章结构，偏调试用）
python main.py generate \
  --config config/report_25h1.yaml \
  --data data/sample_25h1 \
  --output output/25H1_briefing.html

# Phase1：第 1–2 页文字 + 简单表图（早期方案，表图无双向数据条）
python main.py phase1 \
  --config config/report_25h1.yaml \
  --data data/sample_25h1 \
  --output output/25H1_phase1.pptx
```

日常出简报请优先用 **`full-deck`**（见 [出报指引.md](出报指引.md)）。

## 设计说明（为什么是 Excel → PPT）

样例 25H1 PPT 中的「可视化表」本质是 **Excel 条件格式**（数据条 / 色阶）再贴进 PPT，不是 PPT 原生 Chart。

当前闭环：

```
fund_metrics.csv
    → 分析引擎（排名、增速、排名变化、增量拆解）
    → 各页表图渲染（双向数据条 / 色阶 / 银华高亮）
    → 灌入 PPT 模版全册 10 页（文字 + 等比贴图）
```

说明：

- **PPT 表图**使用程序渲染的双向数据条，观感接近样例。
- 同时会写出 `total_ranking.xlsx` 便于核对；Excel 内数据条可能为单向，**以 PPT 图为准**。
- 示例 CSV 中部分细分品类为演示补全；正式使用请导出真实全量数据。
## 目录结构

```
BriefingSupportSystem/
├── main.py                      # 入口
├── 出报指引.md                   # 新一期出报操作指引
├── requirements.txt
├── config/
│   ├── report_25h1.yaml         # 报告周期配置示例
│   ├── report_template.yaml     # 配置模板
│   └── slide_map.yaml           # PPT 页/形状映射
├── templates/
│   └── 25H1_template.pptx       # 简报 PPT 母版
├── data/
│   └── sample_25h1/             # 示例 CSV
├── output/                      # 生成结果（可 gitignore）
└── briefing/
    ├── analytics/               # 排名与增量计算
    ├── narrative/               # 叙事规则
    ├── excel/                   # Excel 构建与表图渲染
    ├── render/                  # PPT / HTML 灌版
    ├── excel_ppt_loop.py        # 主闭环
    ├── phase1.py
    ├── pipeline.py
    └── cli.py
```

## 注意事项

1. 模版路径：`--template templates/25H1_template.pptx`（由 25H1 定稿 PPT 复制而来）。
2. 关注公司默认银华，可在 YAML 的 `focus_company` 修改。
3. 叙事中带判断的句子（如「值得关注的是…」）仍建议人工审改；**数字句由系统写入，勿手抄**。
4. 若打开到旧文件，请确认文件名与修改时间，或使用带日期的 `--output` 路径。

## 后续计划

- Notebook 一键导出模板  
- 可选：本机 Excel 自动化导出更高保真表图（需授权控制 Excel）
