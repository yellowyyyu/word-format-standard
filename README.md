# word-format-standard —— 个人 Word 格式标准与核对流程

一套"把 Word 文档格式管住"的完整方案，由三部分组成：

1. **一份格式标准** —— 把个人/论文的格式要求（字体、字号、缩进、图表、公式……）汇成一张表；
2. **一套四步核对流程** —— 备份 → 自动检查 → 公式符号专项测量 → 人工分遍核对，每一步都有明确的"过关标准"；
3. **三个命令行工具** —— 格式自动检查器、公式符号对齐测量、公式符号大小统一。

## 为什么需要它

人工（或 AI）改 Word 文档时，最容易出问题的是"格式漂移"：字体混了、缩进乱了、
公式符号忽大忽小、图表题注不成套……靠肉眼一条条查，工作量大、又容易漏。

这套方案的思路：

- **把要求写成标准**（一张表，谁都能对着看）；
- **把标准变成程序能查的规则**（跑一遍就能列出所有问题，迭代到全绿）；
- **把"看"这一环拆成几遍做**（每遍只看几项，看完一遍再看下一遍——一次看太多容易漏）。

## 快速开始

1. 打开 [`docs/Word文档格式标准与核对流程.md`](docs/Word文档格式标准与核对流程.md)
   （同目录还有排版好的 `.docx` 版，本身就是"按标准排版"的活样板），按你的需要修改格式要求；
2. 打开 `tools/word_format_check.py`，把顶部"允许的字体、字号"等换成你的标准；
3. 以后每次改完文档：
   - **自动检查**：`python tools/word_format_check.py 文件.docx` —— 输出 `[全部通过]` 才算过，
     有任何问题修好再跑，直到全绿；
   - **公式符号测量**（改动涉及公式时）：把文档导出 PDF，运行
     `python tools/word_symbol_align_check.py 文件.pdf` —— 全部符号与文字的上下偏差
     ≤ 0.5 磅才算过；
   - **人工核对**：按标准里"分四遍"的清单走：字体字号 → 段落版式 → 公式专项 → 整体通读。

## 仓库结构

```text
word-format-standard/
├── docs/
│   ├── Word文档格式标准与核对流程.md      # 标准正文（Markdown 版，方便改）
│   └── Word文档格式标准与核对流程.docx    # 排版好的样板（按标准排版，可直接看效果）
└── tools/
    ├── word_format_check.py              # 自动检查：字体/字号/行距/段距/缩进/标题/题注/注行/表格/公式对象尺寸
    ├── word_symbol_align_check.py        # 公式专项：导出 PDF 后逐字符测量符号与文字的位置差
    └── word_inline_eq_resize.py          # 公式修复：把行内 MathType 符号缩到与段落文字同字号
```

## 工具说明

- 三个工具都是 Python 脚本，命令行运行；依赖：`lxml`（检查器）、`PyMuPDF`（测量器）。
- `word_format_check.py` 的规则直接对应标准文本，改标准时同步改脚本顶部的允许值即可。
- `word_symbol_align_check.py` 针对的是 **MathType 公式对象**在 PDF 里的字形：
  主字形与所在行中文基线的垂直偏差应 ≤ 0.5 磅（≈0.2 毫米，肉眼不可见）。
- `word_inline_eq_resize.py` 修复的典型病症：MathType 对象内部字形默认 12pt，
  在五号（10.5pt）/小五（9pt）正文里会"顶天立地"（偏上偏下）——
  工具按 `段落字号 ÷ 12` 缩放对象显示尺寸并校准垂直位置。

## 相关仓库

- [word-formula-to-mathtype](https://github.com/yellowyyyu/word-formula-to-mathtype) ——
  Word 公式批量转 MathType 的完整流程（公式转换的上游工具）
- [claude-user-profile](https://github.com/yellowyyyu/claude-user-profile) ——
  让 AI 长期记住个人习惯的档案系统
- [project-goal-loop](https://github.com/yellowyyyu/project-goal-loop) ——
  多步项目的汇报格式技能

## License

MIT —— 见 `LICENSE`。
