# excel-jobs-to-styled-pngs

把招聘岗位 Excel 转成 teal-badge 风格 PNG（一张长图 + 按公司切片）的 skill。  
触发场景：「把这表格转成图片」「按 b1-bN 样式出图」「把岗位 Excel 做成招聘海报」等。  
详细字段、视觉 spec、troubleshooting 见 [`excel-jobs-to-styled-pngs/SKILL.md`](excel-jobs-to-styled-pngs/SKILL.md)。

## Pipeline

```
Excel ─► clean.py ─► render.py ─► crop.py
         ↓            ↓            ↓
      cleaned      长图 .png     _1.._N .png
       xlsx
```

## 安装

### 支持的 agent

Claude Code / Codex / OpenCode / Cursor / Cline / GitHub Copilot / Windsurf / Aider Desk 等（72+ 个 CLI coding agent，`npx skills` 都识别）。

### 全局安装（推荐）

`npx skills` 默认会探测本机已安装的 agent，symlink 到各 agent 的全局 skills 目录，一次性装到全部 agent：

```bash
npx skills add -g -y https://github.com/libbyyounh/excel-jobs-to-styled-pngs/tree/main/excel-jobs-to-styled-pngs
```

如果只装部分 agent，用 `-a`：

```bash
npx skills add -g -y -a claude-code -a codex -a opencode https://github.com/libbyyounh/excel-jobs-to-styled-pngs/tree/main/excel-jobs-to-styled-pngs
```

各 agent 的全局路径（symlink 源）：

| Agent | 全局路径 |
|---|---|
| Claude Code | `~/.claude/skills/excel-jobs-to-styled-pngs` |
| Codex | `~/.codex/skills/excel-jobs-to-styled-pngs` |
| OpenCode | `~/.config/opencode/skills/excel-jobs-to-styled-pngs` |
| Cursor | `~/.cursor/skills/excel-jobs-to-styled-pngs` |
| Cline / Copilot / Gemini CLI / Firebender 等 | `~/.agents/skills/excel-jobs-to-styled-pngs` |
| Windsurf | `~/.codeium/windsurf/skills/excel-jobs-to-styled-pngs` |

### 项目级安装

去掉 `-g` 即装到当前仓库的对应 agent 子目录，跟团队共享：

```bash
npx skills add -y https://github.com/libbyyounh/excel-jobs-to-styled-pngs/tree/main/excel-jobs-to-styled-pngs
```

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