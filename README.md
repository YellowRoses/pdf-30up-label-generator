# pdf-30up-label-generator

批量把一文件夹的 PDF 转成可打印的 **30-up（Avery 5160 风格）Word 标签文档**。

每个源 PDF 生成 1 个 Word 文档，文档里整版 30 个格子全部放同一张图（即一页 30 张相同标签，方便整版打印）。

## 核心特性

- **不改模板格子尺寸**：脚本直接在你提供的 Word 模板的**原有格子**里塞图，格子宽/行高/边框/位置原封不动。
- **适配任意模板结构**：只要模板里是 30 个 `<w:tc>` 格子（单个 10×3 表，或 3 个 10×1 表——包括藏在浮动文本框里的），脚本都能把图注入每一个。
- **图片拉伸铺满格子**：图片显示尺寸可配（默认 `2.602″ × 0.992″`，对应 2.625″×1.0″ 的格子）。
- **每个文档只内嵌一次图**：PNG 在 `word/media/` 只存一份，30 个格子共用同一张图，文件小巧。
- **零依赖直出 PDF**：没有 Word/LibreOffice 的环境里，也能从渲染图直接生成几何一致的 30-up PDF（`scripts/png_to_pdf.py`，纯 Python 标准库）。

## 安装

需要 Python 3.8+。渲染后端按优先级自动选择：

```bash
pip install lxml
# 任选一个渲染后端：
pip install pymupdf        # PyMuPDF（推荐，纯 Python、跨平台）
# 或
pip install pdf2image      # 需要额外安装 poppler
```

> 本仓库作者环境用的是 `workbuddy-pdf.exe`（`--pdf-cli` 指定路径），效果等同于上面的开源后端，二选一即可。

## 使用

**运行前先确认三件事**（AI 执行本技能时应向使用者询问，不要假设默认值）：

1. **源 PDF 目录** —— 标签 PDF 放在哪；
2. **Word 模版** —— 使用者自己的 `.docx`，或直接用内置的 `examples/amazon_fnsku_label_template_30up.docx`（真实 30-up、Avery 5160 规格，开箱即用）。**排版严格遵循模版**：脚本只往模版既有格子里注入图片，绝不自建网格；
3. **标签显示尺寸** —— 宽 × 高，**厘米或英寸都可以**，脚本只认英寸（cm ÷ 2.54）。参考值：亚马逊 FNSKU 标签显示尺寸为 **6.61 cm × 2.52 cm**（= `2.602″ × 0.992″`）。

```bash
# 1) 把每个 PDF 首页渲染成 PNG（按 PDF 原名命名）
python scripts/render_pdfs.py --src "<源 PDF 目录>" --out pdfpng --dpi 600 \
    --pdf-cli "C:/.../workbuddy-pdf.exe"

# 2) 把 PNG 注入模版的 30 个格子，每个 PDF 生成一个 .docx
python scripts/build_labels.py \
    --template examples/amazon_fnsku_label_template_30up.docx \
    --src "<源 PDF 目录>" --png-dir pdfpng \
    --out STYX_label_word_docs --imgw-in 2.602 --imgh-in 0.992
```

输出目录 `STYX_label_word_docs/` 下每个 `.docx` 即一张 30-up 标签 sheet。

### 没有自己的模版？

直接用内置的 `examples/amazon_fnsku_label_template_30up.docx`（真实 30-up 模版，非玩具demo）。也可以生成一个合成演示模版：

```bash
python examples/make_sample_template.py demo_template.docx
```

### 转成 PDF（两条路径，二选一）

**路径 A — 经 Word sheet 转换**（适合同时需要 `.docx` 的场景）：

```bash
python scripts/docx_to_pdf.py --src STYX_label_word_docs          # 整文件夹
python scripts/docx_to_pdf.py --file STYX_label_word_docs/AP15500\ v2.docx --out pdf_out
```

转换后端（自动选择，可用 `--backend` 强制）：
- **LibreOffice**（默认优先）：跨平台、可无头运行，适合 Linux 服务器；从 https://libreoffice.org 安装。
- **Microsoft Word COM**（Windows）：需 `pip install pywin32` 且本机装有 Word。

**路径 B — 零依赖直接出 PDF**（无需 Word/LibreOffice/PyMuPDF，纯标准库）：

```bash
python scripts/png_to_pdf.py --src "<源 PDF 目录>" --png-dir pdfpng --out pdf_out \
    --img-w-in 2.602 --img-h-in 0.992
```

直接从第 1 步渲染出的 PNG 重建与 Word sheet 几何一致的 30-up PDF（Letter、3×10 网格、虚线格子边框、图片每份 PDF 只内嵌一次）。PNG 目录三种布局都认：`<base>.png`（本仓库 `render_pdfs.py` 的输出）、`<编号>/page-0001.png`（pdf-cli 批量输出）、`<base>/page-0001.png`。可用 `--limit 2` 做冒烟测试；格子数、图片尺寸可用 `--cols/--rows/--img-w-in/--img-h-in` 调整。

## ⚠️ 两个踩坑点（重要）

1. **不要新建表格去"假装"30-up 布局**（例如 `python-docx add_table(10,3)`）。那样和原模板几何对不上，格子会看起来被"扩大"。正确做法是**往模板已有的格子里注入图片**。
2. **注入图片时不要加 `noChangeAspect`**，否则图片无法拉伸铺满格子，会被锁成原始比例留白边。

## 校验清单

每个产物 `.docx` 应满足：
- `document.xml` 结构良构；
- 绘图数 = 格子数（30）；
- 所有 `r:embed` 指向同一个 `media/label.png`；
- 图片 extent 等于配置尺寸；
- 格子的 `tcW` / `trH` 与模板逐字节一致（证明没被改动）。

## 目录结构

```
pdf-30up-label-skill/
├── SKILL.md                  # WorkBuddy skill 定义
├── README.md
├── LICENSE
├── scripts/
│   ├── render_pdfs.py        # PDF -> PNG（多后端）
│   ├── build_labels.py       # PNG -> 注入模板 30 格
│   ├── docx_to_pdf.py        # .docx -> PDF（LibreOffice / Word COM）
│   └── png_to_pdf.py         # PNG -> 30-up PDF（零依赖，路径 B）
└── examples/
    ├── amazon_fnsku_label_template_30up.docx   # 内置 30-up Word 模版（开箱即用）
    └── make_sample_template.py                 # 合成演示模版生成器
```
