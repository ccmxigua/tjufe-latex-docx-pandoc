# 天津财经大学博士毕业论文模版：tjufe-latex-docx-pandoc

一个面向“LaTeX 论文源文件 → Word DOCX”的 Pandoc 转换工具链示例。仓库包含转换脚本、Lua filter、DOCX 后处理脚本、公开占位样例与参考样式生成脚本。

> 注意：本仓库是公开模板和工具链，不包含任何真实论文正文、真实数据、真实参考文献库、个人日志或私有信息。

## 功能概览

- 使用 Pandoc 将 LaTeX 转为 DOCX。
- 通过 Lua filter 处理部分论文结构、标题、图表和交叉引用。
- 通过 Python 后处理 DOCX 样式、页眉页脚、段落、图表题注和部分公式显示。
- 可生成/使用 `reference.docx` 作为 Word 样式参考。
- 提供最小公开样例：`samples/sample-thesis.tex` 与 `samples/sample-metadata.yaml`。

## 目录结构

```text
.
├── convert_tjufe.sh
├── filters/
│   └── tjufe-thesis.lua
├── scripts/
│   ├── build_reference_docx.py
│   ├── preprocess_tjufe_tex.py
│   └── postprocess_tjufe_docx.py
├── samples/
│   ├── sample-thesis.tex
│   ├── sample-metadata.yaml
│   └── figures/example.png
├── reference.docx
├── .gitignore
└── LICENSE
```

## 依赖

建议环境：macOS / Linux shell。

必需工具：

- `pandoc`
- `python3`
- Python 包：`python-docx`、`lxml`（如脚本运行提示缺包，请先安装）
- LaTeX 环境可选；本工具链主要依赖 Pandoc 解析 LaTeX 输入

示例安装：

```bash
brew install pandoc
python3 -m pip install python-docx lxml
```

## 快速开始

```bash
chmod +x convert_tjufe.sh
./convert_tjufe.sh samples/sample-thesis.tex samples/sample-metadata.yaml out/sample.docx
```

输出文件默认写入 `out/`。该目录已被 `.gitignore` 排除。

## 生成 reference.docx

如果需要重新生成参考样式文件：

```bash
python3 scripts/build_reference_docx.py reference.docx
```

## 隐私与发布建议

如果你基于本仓库处理真实论文，请不要提交以下内容到公开仓库：

- 真实论文正文、章节草稿、评审意见
- 真实参考文献库或未公开数据
- 生成的 DOCX/PDF/日志文件
- 个人姓名、学号、导师信息等敏感信息
- 本地绝对路径、API key、token、账号凭据

建议把真实项目放在私有目录或 private 仓库，只把通用脚本和公开样例保留在 public 仓库。

## License

MIT License. See `LICENSE`.
