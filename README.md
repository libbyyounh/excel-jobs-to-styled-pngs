# excel-jobs-to-styled-pngs

把招聘岗位 Excel 表格转换为统一风格的 PNG 图片，方便直接插入演示文稿、公众号文章或海报。

## 功能

- 读取一份招聘岗位 Excel（包含 `序号 / 公司 / 岗位 / 薪酬` 等列，可选 `招聘人数 / 工作地点`）。
- 生成 **一张完整长图**（所有岗位一览），以及 **N 张按公司分组的切片图**，便于插入 PPT / Keynote / 飞书幻灯片 / 微信图文。

## 卡片样式

输出采用青绿徽章风格的招聘卡片（teal-badge card），与历史样张 `b1.png … bN.png` 保持一致的视觉规范。  
视觉常量（背景色、徽章色、字号、间距等）集中在 `render.py` 顶部的「Visual spec」区域，可按需调整。

## 三阶段处理流程

```
Excel  ─► clean.py  ─► *_cleaned.xlsx + *_clean_report.txt
                       │
                       ▼
                 render.py  ─► {prefix}.png        (一张完整长图)
                       │
                       ▼
                  crop.py  ─► {prefix}_1.png … {prefix}_N.png  (按公司切片)
```

- **Stage 0 — clean.py**：清理噪声行（大标题、政策行、空行等），统一为 `序号 / 公司 / 岗位 / 薪酬` 四列，生成 `*_cleaned.xlsx` 与 `*_clean_report.txt`。
- **Stage 1 — render.py**：读取清洗后的 Excel，绘制带青绿徽章 + 公司名 + 岗位 + 薪酬的招聘卡片长图。
- **Stage 2 — crop.py**：纯 Pillow（不依赖 numpy）扫描灰色分隔线，将长图切成 N 张公司卡片。

详细的字段含义、过滤规则、CJK 字体兜底链见 [`excel-jobs-to-styled-pngs/SKILL.md`](excel-jobs-to-styled-pngs/SKILL.md)。

## 安装

### 1. 准备 Python 环境

需要 **Python 3.9+**。建议用 `venv` 隔离依赖，避免和系统 Python 冲突：

```bash
# 进入项目目录
cd excel-jobs-to-styled-pngs

# 创建并激活虚拟环境
python3 -m venv .venv
source .venv/bin/activate          # macOS / Linux
# .venv\Scripts\activate           # Windows
```

### 2. 安装依赖

skill 依赖的库很少，全部是纯 Python 包：

```bash
pip install -r requirements.txt
```

如果项目里还没有 `requirements.txt`，可直接安装核心依赖：

```bash
pip install Pillow openpyxl
```

> **注意**：`crop.py` 故意使用纯 Pillow 实现，**不依赖 numpy**。在 PEP 668 管理的 Python 安装上 `pip install numpy` 会被拦截。

### 3. 准备输入文件

把以下两份文件放到工作目录（例如 `./input/`）：

1. **招聘 Excel**（`.xlsx`）— 列包含 `序号 / 公司 / 岗位 / 薪酬`，可选 `招聘人数 / 工作地点`。
2. **样式参考图**（可选）— 形如 `b1.png … bN.png` 的青绿徽章卡片样张；不提供时会使用 `render.py` 里默认的 teal-badge 配色。

### 4. 配置脚本参数

使用前需根据实际 Excel 修改三份脚本顶部的常量：

| 脚本 | 配置项 | 含义 |
|---|---|---|
| `clean.py` | `CONFIGS` | 每个 sheet 的 `(sheet_name, min_data_row, company_col, job_col, salary_col)` |
| `render.py` | `TARGETS` | 要渲染的 `(sheet_name, 输出前缀)` |
| `render.py` | 视觉常量 | `WIDTH` / `TEAL` / `BADGE_W` 等，可按样式参考图调整 |
| `crop.py`  | `TARGETS` | 要切片的 `(输出前缀, 长图文件名)` |

CJK 字体路径已经在 `render.py` 顶部以 `FONT_CANDIDATES` 兜底链的形式列出，覆盖 macOS / Linux 常见路径，无需手动配置。

## 使用方法

按顺序执行三阶段：

```bash
# Stage 0 — 清理 Excel（必须先跑）
python3 excel-jobs-to-styled-pngs/clean.py input/recruitment.xlsx

# Stage 1 — 渲染长图
python3 excel-jobs-to-styled-pngs/render.py input/recruitment_cleaned.xlsx

# Stage 2 — 切片
python3 excel-jobs-to-styled-pngs/crop.py
```

输出物：

| 文件 | 说明 |
|---|---|
| `recruitment_cleaned.xlsx` | 清洗后的四列数据 |
| `recruitment_clean_report.txt` | 被丢弃的行及其原因（运行 Stage 1 前必读） |
| `{prefix}.png` | 一张完整长图 |
| `{prefix}_1.png … {prefix}_N.png` | 按公司切片的 N 张卡片图 |

## 常见问题

- **输出 PNG 中中文变成方块（□）**：CJK 字体没找到。检查系统是否装了任意一个 `FONT_CANDIDATES` 中的字体，或在 `render.py` 顶部补一条路径。
- **`kept=0`（clean_report 里所有行都被丢弃）**：一般是 `CONFIGS` 里的 `min_data_row` 或薪酬列号写错了。`openpyxl` 列号从 1 开始计数。
- **`ModuleNotFoundError: numpy`**：不要用 `pip install numpy`，`crop.py` 已经做了纯 Pillow 实现。
- **切片数量对不上**：灰色分隔线检测阈值在 `crop.py` 里，阈值偏低时会出现误检，按需上调到 `80%` 左右。

更多 troubleshooting 见 [SKILL.md](excel-jobs-to-styled-pngs/SKILL.md) 的「Common mistakes」表。

## 目录结构

```
excel-jobs-to-styled-pngs/
├── README.md
├── input/                                  # 你的 Excel / 参考图
└── excel-jobs-to-styled-pngs/              # skill 本体
    ├── SKILL.md
    ├── clean.py
    ├── render.py
    └── crop.py
```

## 许可

脚本与文档随仓库分发，按需修改使用。