# pdf-30up-label-generator

批量把一文件夹的 PDF 转成可打印的 **30-up（Avery 5160 风格）Word 标签文档**。

每个源 PDF 生成 1 个 Word 文档，文档里整版 30 个格子全部放同一张图（即一页 30 张相同标签，方便整版打印）。

## 核心特性

- **不改模板格子尺寸**：脚本直接在你提供的 Word 模板的**原有格子**里塞图，格子宽/行高/边框/位置原封不动。
- **适配任意模板结构**：只要模板里是 30 个 `<w:tc>` 格子（单个 10×3 表，或 3 个 10×1 表——包括藏在浮动文本框里的），脚本都能把图注入每一个。
- **图片拉伸铺满格子**：图片显示尺寸可配（默认 `2.602″ × 0.992″`，对应 2.625″×1.0″ 的格子）。
- **每个文档只内嵌一次图**：PNG 在 `word/media/` 只存一份，30 个格子共用同一张图，文件小巧。

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

```bash
# 1) 把每个 PDF 首页渲染成 PNG（按 PDF 原名命名）
python scripts/render_pdfs.py --src "STYX (new SKU code) Label" --out pdfpng --dpi 600 \
    --pdf-cli "C:/.../workbuddy-pdf.exe"

# 2) 把 PNG 注入模板的 30 个格子，每个 PDF 生成一个 .docx
python scripts/build_labels.py \
    --template amazon_fnsku_label_template_30up.docx \
    --src "STYX (new SKU code) Label" --png-dir pdfpng \
    --out STYX_label_word_docs --imgw-in 2.602 --imgh-in 0.992
```

输出目录 `STYX_label_word_docs/` 下每个 `.docx` 即一张 30-up 标签 sheet。

### 没有模板？生成演示模板

```bash
python examples/make_sample_template.py demo_template.docx
```

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
│   └── build_labels.py       # PNG -> 注入模板 30 格
└── examples/
    └── make_sample_template.py
```
