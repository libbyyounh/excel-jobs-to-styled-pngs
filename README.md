---
name: excel-jobs-to-styled-pngs
description: Use when converting a recruitment / job-listing Excel into designed PNG images (one full image + per-company slices) that match a teal-badge card style reference. Triggers on phrases like "把这表格转成图片", "按这个样式做新的招聘 PNG", "把岗位 Excel 做成招聘海报", or whenever the user shows prior output like `b1.png…bN.png` and asks to reproduce the look for a new spreadsheet. Same-shape input: rows of 序号 / 公司 / 岗位 / 薪酬 (with optional 招聘人数 / 工作地点 columns). Output: one full long image plus N per-company slices for slide insertion.
---

# excel-jobs-to-styled-pngs

把招聘岗位 Excel 转成 teal-badge 风格 PNG 的三阶段 skill：`clean → render → crop`。

## When to invoke

- 用户给了一份招聘 / 岗位 Excel（列含 `序号 / 公司 / 岗位 / 薪酬`，可选 `招聘人数 / 工作地点`）
- 用户要求：做成图片 / 转成 PNG / 按 b1-bN 样式出图 / 做招聘海报
- 输出：一张完整长图 + N 张按公司切片的 PNG

不要用：

- 图表 / dashboard → 用 `dataviz`
- 完整 PPT → 用 `slides` 或 `ppt-master`
- 一行一公司、岗位挤在一个单元格里（这种 schema 不在本 skill 假设范围）
- 非招聘内容（噪声过滤关键字是按 "政策 / 贷款 / 详见" 调的）

## Pipeline

```
Excel  ─► clean.py  ─► *_cleaned.xlsx + *_clean_report.txt
                       │
                       ▼
                 render.py  ─► {prefix}.png
                       │
                       ▼
                  crop.py  ─► {prefix}_1.png … {prefix}_N.png
```

**Hard gate**: 必须先跑 `clean.py`，再 `render.py`，再 `crop.py`。没跑 `clean` 就 render 会出现噪声行（"政策 / 贷款"行、大标题行等）。

详细字段含义、过滤规则、CJK 字体兜底链见 `SKILL.md`。本文件只讲 agent 视角的安装 / 入口。

## Install (one-time, per machine)

脚本依赖极少，没有特殊二进制；但 `crop.py` 故意用纯 Pillow（PEP 668 在受管 Python 上会拦截 `pip install numpy`，所以不要尝试装 numpy）。

```bash
# 用 venv（推荐，PEP 668 友好）
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# 或者直接装核心依赖
pip install Pillow openpyxl
```

依赖列表：`requirements.txt`（Pillow ≥ 10，openpyxl ≥ 3.1）。

CJK 字体在 `render.py` 顶部有 `FONT_CANDIDATES` 兜底链，覆盖 macOS / Linux 常见路径。**不要硬编码一个字体路径**——多平台用户会踩坑。

## Inputs (per run)

执行前向用户确认两件事：

1. **Excel** — `.xlsx`，包含 `序号 / 公司 / 岗位 / 薪酬`（可选 `招聘人数 / 工作地点`）。
2. **样式参考图** — 形如 `b1.png … bN.png` 的青绿徽章卡片样张。如果用户没提供，使用 `render.py` 默认的 teal-badge 配色。

任一缺失就先问，**不要**直接进入 render。

## Configure (per Excel)

每份新 Excel 都要改三份脚本顶部的常量：

| 脚本 | 常量 | 含义 |
|---|---|---|
| `clean.py` | `CONFIGS` | `(sheet_name, min_data_row, company_col, job_col, salary_col)` |
| `render.py` | `TARGETS` | 要渲染的 `(sheet_name, 输出前缀)` |
| `render.py` | 视觉常量 | `WIDTH` / `TEAL` / `BADGE_W` 等，按样式参考图调 |
| `crop.py` | `TARGETS` | 要切片的 `(输出前缀, 长图文件名)` |

`openpyxl` 的列号从 1 开始数。`min_data_row` 是「第一条数据行」的行号（含表头的话通常表头+1）。

## Run

```bash
# Stage 0 — clean（必跑）
python3 clean.py path/to/recruitment.xlsx
# → *_cleaned.xlsx
# → *_clean_report.txt   ← render 之前必读

# Stage 1 — render
python3 render.py path/to/recruitment_cleaned.xlsx
# → {prefix}.png

# Stage 2 — crop
python3 crop.py
# → {prefix}_1.png … {prefix}_N.png
```

跑完每一阶段都 `print` / `ls` 确认产物真的生成了，再进入下一阶段。

## Output naming

| 文件 | 内容 |
|---|---|
| `{stem}_cleaned.xlsx` | Stage 0：四列清洗后的数据 |
| `{stem}_clean_report.txt` | Stage 0：被丢弃的行 + 原因 |
| `{prefix}.png` | Stage 1：完整长图（所有卡片堆叠） |
| `{prefix}_1.png … {prefix}_N.png` | Stage 2：按公司切片 |

每个 sheet 用自解释前缀（`c_gzh.png` 公众号、`c_zpxq.png` 招聘详情），输出本身就是文档。

## Common mistakes

| 症状 | 原因 / 修复 |
|---|---|
| 输出里公司是 `1` / `2` / `3` | `CONFIGS` 列号错了 — `openpyxl` 从 1 开始，先 dump 前 3 行核对 |
| `clean_report` 里 `kept=0` | 所有行都被丢 — 通常 `min_data_row` 或薪酬列号错了 |
| 出现 `政策` 行 | `clean.py` 的 `NOISE_KEYWORDS` 没覆盖到 — 补关键字再重跑 |
| 第一家公司不见了 | `render.py` 里硬编码了 `rows[2:]` — 清洗后已经去掉表头，不要再手动切片 |
| 中文变方块（□） | CJK 字体没找到 — 检查 `FONT_CANDIDATES` 任一项是否可用 |
| `ModuleNotFoundError: numpy` | **不要** `pip install numpy` — `crop.py` 是纯 Pillow 实现 |
| 一张超高长图，没切片 | 忘了跑 Stage 2 — 切片才是给幻灯片用的成品 |
| `面议` 变成 `面议元/月` | 不要臆造单位 — 原样渲染 cell 内容 |
| 切片数量不对（多了） | 灰色分隔线检测阈值过低 — 阈值提到 `80%` |
| 薪酬挤到岗位文字上 | 薪酬锚到固定列了 — 应该是 `job_bbox.right + 40` |

完整排错表见 `SKILL.md`。

## Files

```
excel-jobs-to-styled-pngs/
├── README.md              # 本文件：agent 安装 / 入口
├── SKILL.md               # 完整说明：字段、视觉 spec、troubleshooting
├── requirements.txt       # Pillow + openpyxl
├── clean.py               # Stage 0
├── render.py              # Stage 1
└── crop.py                # Stage 2
```