# excel-jobs-to-styled-pngs

把招聘岗位 Excel 转成 teal-badge 风格 PNG（一张长图 + 按公司切片）的 skill。  
触发场景：「把这表格转成图片」「按 b1-bN 样式出图」「把岗位 Excel 做成招聘海报」等。详细字段、视觉 spec、troubleshooting 见 [`excel-jobs-to-styled-pngs/SKILL.md`](excel-jobs-to-styled-pngs/SKILL.md)。

## Pipeline

```
Excel ─► clean.py ─► render.py ─► crop.py
         ↓            ↓            ↓
      cleaned      长图 .png     _1.._N .png
       xlsx
```

## 安装

### 全局安装（agent 自动加载）

skill 源文件已放在仓库的 `excel-jobs-to-styled-pngs/` 子目录。全局安装命令：

```bash
npx skills add -g -y https://github.com/libbyyounh/excel-jobs-to-styled-pngs/tree/main/excel-jobs-to-styled-pngs
```

安装后全局位置：`~/.claude/skills/excel-jobs-to-styled-pngs/`。

### 本地运行（手动调用脚本）

依赖：Python 3.9+、Pillow、openpyxl。`crop.py` 故意不依赖 numpy（PEP 668 友好）。

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r excel-jobs-to-styled-pngs/requirements.txt
```

然后按顺序跑：

```bash
python3 excel-jobs-to-styled-pngs/clean.py  path/to/recruitment.xlsx
python3 excel-jobs-to-styled-pngs/render.py path/to/recruitment_cleaned.xlsx
python3 excel-jobs-to-styled-pngs/crop.py
```

每份新 Excel 需要先改 `clean.py` / `render.py` / `crop.py` 顶部的 `CONFIGS` / `TARGETS`，详见 SKILL.md。