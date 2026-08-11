---
name: excel-jobs-to-styled-pngs
description: Use when converting a recruitment / job-listing Excel into designed PNG images (one full image + per-company slices) that match a teal-badge card style reference. Triggers on phrases like "把这表格转成图片", "按这个样式做新的招聘 PNG", "把岗位 Excel 做成招聘海报", or whenever the user shows prior output like `b1.png…bN.png` and asks to reproduce the look for a new spreadsheet. Same-shape input: rows of 序号 / 公司 / 岗位 / 薪酬 (with optional 招聘人数 / 工作地点 columns). Output: one full long image plus N per-company slices for slide insertion.
---

# Excel Jobs → Styled PNGs

## Overview

A three-stage pipeline — **clean → render → crop**. Each stage is a separate script with a clear contract. The skill is gated: cleaning must finish before render starts; the Excel must be cleaned before either render or crop runs.

```
Excel  ─► clean.py  ─► *_cleaned.xlsx + clean_report.txt
                           │
                           ▼
                     render.py  ─► {prefix}.png
                           │
                           ▼
                      crop.py  ─► {prefix}_1.png ... {prefix}_N.png
```

Built on Pillow + openpyxl, no numpy, no Playwright.

## Inputs (required)

Before running, confirm both are present:

1. **Excel** — a recruitment / job-listing xlsx with rows of 序号 / 公司 / 岗位 / 薪酬 (and optional 招聘人数 / 工作地点). If the sheet has section headers (e.g. "一、招聘岗位"), policy noise rows, or duplicate `招聘人数` columns, **clean.py will handle it** — that's its job.
2. **Style reference image** — a prior output like `b1.png…bN.png` (teal-badge card style). The agent samples colors and proportions from it. If the user does **not** supply one, fall back to the **default teal-badge spec** documented in the "Visual spec" section below.

If either is missing, ask. Do not start render without confirming both.

## When to Use

- Recruitment / job-listing Excel with rows of 序号 / 公司 / 岗位 / 薪酬
- Output: one long designed image (for 公众号 / 海报) + per-company slices (for slide insertion)
- User has a style reference image OR says "做成和 b1-b15 一样的"

## When NOT to Use

- Charts / dashboards (use dataviz)
- Slide decks (use slides / ppt-master)
- One-row-per-company Excel where jobs are crammed into one cell — this skill assumes one row per (company, job)
- Non-recruitment content (the noise filter is tuned for "政策 / 贷款 / 详见")

## Stage 0 — Clean (must run first)

The cleaning stage produces a **tidy xlsx** with exactly four columns in this order: `序号, 公司, 岗位, 薪酬`. All subsequent stages assume this schema. Cleaning also writes a `*_clean_report.txt` listing every dropped row and why — read this before render to catch silent surprises.

**Default (zero-config)**: leave `CONFIGS = []` at the top of `clean.py`. The script scans every sheet in the workbook, finds the row whose first cell is `"序号"` (= header row), then locates the company / job / salary columns by matching common header strings (`公司名称 / 所属公司 / 公司`, `岗位名称 / 岗位`, `薪资范围 / 薪酬范围 / 薪酬 / 薪资`). Sheets that don't match (e.g. a 字典项 sheet) are skipped with a warning.

**Manual override**: fill `CONFIGS` with explicit tuples if auto-detection fails on a sheet:

```python
# (sheet_name, min_data_row, company_col, job_col, salary_col)
CONFIGS = [
    ('公众号最新版', 2, 2, 3, 6),   # has duplicate 招聘人数 col → salary at 6
    ('招聘详情',    3, 2, 3, 5),   # extra "一、招聘岗位" title row → min_row 3
]
```

**Filters applied in order** (each `drop` is logged):

| Filter | Why | Action |
|---|---|---|
| `序号` starts with `"二、"`, `"三、"`, … | New section (政策宣讲 etc.) — not recruitment | **break** the whole sheet |
| `序号` starts with `"一、"`, … | Section title row (大标题) | skip |
| `序号` cell is the literal string `"序号"` | Header row | skip |
| All three data cells empty | Spacer / blank line | skip |
| Job or salary empty | Row can't render | skip |
| Company cell contains `政策` / `贷款` / `详见` | Policy / reference noise | skip |
| No company has been seen yet (stray data before any group) | Stray | skip |

**Forward-fill**: first row of a company group has `公司` filled; subsequent rows have it empty. clean.py propagates the value downward.

**Newline collapse**: `中国人寿…\n第四营销服务部` becomes `中国人寿… 第四营销服务部` (the `\n` → space happens in clean.py, so render.py doesn't have to think about it).

```bash
python3 clean.py path/to/recruitment.xlsx
# → path/to/recruitment_cleaned.xlsx
# → path/to/recruitment_clean_report.txt
```

**Read the report before render.** If "kept=N" looks wrong (e.g. 0 jobs kept), the column indices in CONFIGS are wrong — fix and re-run.

## Stage 1 — Render

Reads `*_cleaned.xlsx` (schema fixed by clean.py), draws a designed long image, writes one PNG per sheet.

**Default (zero-config)**: leave `TARGETS = []` at the top of `render.py`. The script processes every sheet in the cleaned xlsx and derives the output filename from the sheet name (`公众号最新版` → `c_公众号最新版.png`). Special characters in sheet names (`/\:*?"<>|.` and whitespace) are sanitized to underscores.

**Manual override**: set `TARGETS` to map specific sheets to specific output filenames:

```python
TARGETS = [
    ('公众号最新版', 'c_gzh.png'),
    ('招聘详情',    'c_zpxq.png'),
]
```

### Visual spec (default teal-badge card)

These constants live at the top of `render.py`. Override them if the user supplies a different style reference — sample colors from the reference image (5 px around the badge center) and adjust.

| Constant | Value | Notes |
|---|---|---|
| `WIDTH` | 1200 | Canvas width |
| `MARGIN_LEFT` / `MARGIN_RIGHT` | 60 | Side padding |
| `BG_COLOR` | `(255,255,255)` | White background |
| `TEAL` | `(100,149,137)` | Badge fill — change to match brand |
| `DARK_TEXT` | `(51,51,51)` | Job / salary text |
| `BADGE_W` × `BADGE_H` | 180 × 56 | Rounded rect, radius 12 |
| `COMPANY_LINE_H` | 42 | Per wrap-line of company name |
| `JOB_ROW_H` | 58 | Per job row |
| `SECTION_GAP` | 30 | Space between cards |
| `SEPARATOR_COL` | `(200,200,200)` | 2 px line between cards (skip after last) |

**Per-card draw order**: (1) teal rounded badge with `"招聘岗位"` white text top-left, (2) company name to the right of the badge, wrapped to fit `WIDTH - margins - BADGE_W - 24`, (3) each job row: `<job>` + spacing + `<salary>` on the same line, (4) gray separator line below the card (except last).

**Card header height** (where badge + company name live) is `max(BADGE_H, COMPANY_LINE_H * n_lines)`, **not** a sum. Badge and company are laid out **side by side**, not stacked — adding another line of company name makes the header wider, not taller.

**Salary x-position** is `job_bbox.right + 40` (right after the job name, with a ~40 px gap), **not** a fixed column anchored to the badge. A fixed column overlaps long job names like `"非标机械设计工程师"`.

### CJK font fallback (critical)

Non-CJK fonts render Chinese as **boxes** (`□`). Always probe a fallback chain — never hardcode one path:

```python
FONT_CANDIDATES = [
    '/System/Library/Fonts/STHeiti Medium.ttc',         # macOS 华文黑体
    '/System/Library/Fonts/Hiragino Sans GB.ttc',       # macOS 冬青黑
    '/Library/Fonts/PingFang.ttc',                      # macOS 苹方 (some versions)
    '/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc',     # Linux 文泉驿
    '/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc',  # Linux Noto CJK
]
# Sizes: badge 38, company 40, job 32, salary 30
# Pick the first that ImageFont.truetype(path, 12) loads without OSError.
```

### Don't fabricate salary suffixes

If a salary is `"面议"`, `"3500+提成"`, `"20-60w年薪"`, or `"/"`, render it **as-is**. Do **not** append `"元/月"` — the inconsistency looks worse than no unit. Only render what the cell already says.

```bash
python3 render.py path/to/recruitment_cleaned.xlsx
# → c_gzh.png, c_zpxq.png, ...
```

## Stage 2 — Crop (pure Pillow, no numpy)

`pip install numpy` is **blocked by PEP 668** on managed Python installs. The pure-Pillow row scan is fast enough at 1200 px width.

```python
from PIL import Image

img = Image.open(src).convert('RGB')
w, h = img.size
pixels = img.load()

sep_rows = []
for y in range(h):
    gray = 0
    for x in range(w):
        r, g, b = pixels[x, y]
        if 190 <= r <= 210 and 190 <= g <= 210 and 190 <= b <= 210:
            gray += 1
            if gray * 10 > w * 7:    # early-exit on first 70%
                sep_rows.append(y)
                break
```

Then:
1. **Merge consecutive rows** within 3 px (anti-aliasing of the same line)
2. **Boundaries** = `[0, *sep_rows, h]`
3. **Crop** each `[boundaries[i], boundaries[i+1]]` strip → `{prefix}_{i+1}.png`

**Expected yield**: N−1 separator lines for N cards. If you get many more (false positives from text bleeding into the gray range), raise the threshold to 80% or render the separator more saturated (e.g. `(180,180,180)`).

**Default (zero-config)**: leave `TARGETS = []` at the top of `crop.py`. The script globs every `c_*.png` in the current directory (skipping files that already look like slices, e.g. `c_foo_3.png`) and processes them automatically.

**Manual override**: set `TARGETS` to process specific files:

```python
TARGETS = [
    ('c_gzh',  'c_gzh.png'),
    ('c_zpxq', 'c_zpxq.png'),
]
```

```bash
python3 crop.py
# → c_gzh_1.png ... c_gzh_N.png, c_zpxq_1.png ... c_zpxq_N.png
```

## Bundled scripts

Three scripts are bundled with this skill, in the same directory as `SKILL.md`:

- `clean.py` — Stage 0 (Excel → cleaned xlsx + report)
- `render.py` — Stage 1 (cleaned xlsx → long PNG)
- `crop.py` — Stage 2 (long PNG → per-company slices)

**Default usage is zero-config** — `CONFIGS = []` in `clean.py`, `TARGETS = []` in both `render.py` and `crop.py`. The scripts auto-detect the Excel schema and auto-discover rendered PNGs.

**Typical workflow for a new Excel** (zero edits required):

```bash
cp /path/to/skill/{clean.py,render.py,crop.py} working_dir/
cp your_recruitment.xlsx working_dir/
cd working_dir
python3 clean.py your_recruitment.xlsx
python3 render.py your_recruitment_cleaned.xlsx
python3 crop.py
```

**When you need to override** (auto-detection fails, or you want a different style):

- Edit `CONFIGS` in `clean.py` (column indices per sheet) — only when auto-detection picks the wrong column
- Edit the "Visual spec" constants at the top of `render.py` (colors, sizes, layout) — only when matching a non-default style reference
- Edit `TARGETS` in `render.py` / `crop.py` (sheet → filename mapping) — only when you want explicit control over output names

## Common mistakes

| Symptom | Cause / Fix |
|---|---|
| `clean.py` skips a sheet with "could not auto-detect header" | Sheet doesn't have a `"序号"` cell, or its 公司 / 岗位 / 薪酬 columns use unrecognized header strings. Add them to `COMPANY_HEADERS` / `JOB_HEADERS` / `SALARY_HEADERS` in `clean.py`, or add a manual `CONFIGS` entry |
| Output shows `1`, `2`, `3` as companies | Auto-detection picked the wrong salary column (some files have multiple 招聘人数 columns). Add a manual `CONFIGS` entry with the correct salary index |
| Clean report shows `kept=0` | All rows dropped — usually wrong `min_row` or wrong salary column. Add a manual `CONFIGS` entry |
| 政策 rows appear in output | `NOISE_KEYWORDS` in `clean.py` doesn't cover this dataset. Add the keyword and re-run |
| Sheet name with `.` becomes `c_7_14岗位.png` | The `.` is sanitized to `_` in render.py's `safe_filename`. Functional but ugly — add a manual `TARGETS` entry to control the output name |
| crop.py finds nothing | No `c_*.png` in the current directory. Either run `render.py` first, or set `TARGETS` manually with absolute paths |
| Crop re-processes already-sliced files | The slice pattern (`_*N*.png`) only filters numbers at the END. If your sheet name ends in a digit (e.g. `c_2.png`), name slices differently or use `TARGETS` |
| Company name overflows the badge | Wrap not triggered — long company. Verify `WIDTH - margins - BADGE_W - 24` is positive |
| Chinese text → boxes | Font is not CJK. Probe the fallback chain in `render.py` |
| `ModuleNotFoundError: numpy` | `crop.py` should be pure Pillow — see Stage 2 above |
| One tall unreadable image | Skipped Stage 2 — always crop, slices are what get inserted into slides |
| `面议` rendered as `面议元/月` | Don't fabricate the suffix — render the cell as-is |
| Wrong number of slices (e.g. 32 instead of 18) | Separator detection threshold too low. Raise to 80% in `crop.py` |
| Huge empty space at bottom of canvas | Card header height summed (`BADGE_H + COMPANY_LINE_H * lines`) instead of `max(...)` — see Visual spec |
| Salary overlaps long job names | Salary anchored to a fixed column. Use `job_bbox.right + 40` instead — see Visual spec |
| Skipped clean.py → noisy output | Cleaning is a hard prerequisite. Run it first, read the report, THEN render |

## When the user provides a different style reference

1. Open the reference image, sample 5 px around the badge center to get the accent color
2. Visually estimate font sizes by proportion (badge text usually 35–45 px at 1200 wide)
3. Note layout: badge left vs centered, with/without divider, with/without company name wrap
4. Override only the constants in the "Visual spec" table — keep the algorithm
5. Always re-run Stage 1 against the new constants before Stage 2; the old `c_*.png` will be overwritten

## Output naming

| File | Content |
|---|---|
| `{stem}_cleaned.xlsx` | Stage 0 output — tidy data for render |
| `{stem}_clean_report.txt` | Stage 0 output — what was dropped and why |
| `{prefix}.png` | Stage 1 output — full long image (all cards stacked) |
| `{prefix}_1.png` … `{prefix}_N.png` | Stage 2 output — per-card slices (N = separators + 1) |

Use a descriptive prefix per sheet (`c_gzh.png` for 公众号, `c_zpxq.png` for 招聘详情) so the outputs are self-documenting.
