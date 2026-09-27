# 线性代数笔记：源文件

每页手写笔记识别成 `pages/pNN.txt`（一行一段，公式用 LaTeX），再由脚本生成 Word（公式为 Word 原生公式）。

- `lib.py`：解析源文件、LaTeX 转 Word 公式、生成 .docx
- `build.py`：`python3 build.py pages/p01.txt pages/p03.txt … -o 输出.docx`
- `preview.py`：生成单页预览图（需要 LibreOffice）

依赖：Python 3、lxml、pypandoc（含 pandoc）、PyMuPDF、Pillow。
