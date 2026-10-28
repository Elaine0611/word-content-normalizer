# Word Content Normalizer

一个用于整理现有 Word `.docx` 文档的 Codex Skill。它在原文档的副本上清理封面、目录等指定内容，并按标题与字数规则在正文中标记内容段，同时尽量保留原有文字、图片、图表、绘图和嵌入对象。

## 适用场景

- 需要从现有 Word 文档生成一份简化、规范化的副本。
- 需要按章节和内容长度划分正文，供后续阅读或处理。
- 希望保留文档中的视觉对象，而不是提取文字后重新制作 DOCX。

默认流程只生成一份规范化 DOCX，不覆盖原文件。对于无法可靠识别的删除边界或可能影响正文视觉对象的情况，应先检查并确认，不应直接交付结果。

## 主要处理规则

- 清理各分节的页眉页脚、可识别的封面装饰内容、版本修订记录和目录；保留普通业务表格及非目标内容。
- 根据 Word 标题样式和正文编号识别标题层级，整理多余空段落与段落间距。
- 按标题和字数合并或拆分内容段：不足 150 字的初始块向后合并，合并后不超过 1000 字；超过 800 字时优先在标题或元素边界拆分。
- 在每个最终内容段前插入独立的一行 `##`，包括首段；不使用分页作为分段标记。
- 不在表格或视觉对象内部切分，不重建或重新编码已有图片。
- 对输出文件进行结构验证，包括正文视觉对象、图片媒体、清理结果和分段覆盖情况。

字数规则参照 Word 状态栏的“字数”：连续英文、数字和下划线计为一个词，中文和标点逐个计数，空白及图片本身不计数。更完整的边界判断见 [分类规则](references/classification-rules.md) 和 [转换规则](references/transformation-rules.md)。

## 环境准备

需要 Python 3，并安装脚本使用的依赖：

```bash
python -m pip install python-docx lxml
```

在仓库根目录运行以下命令。路径中有空格时，请给文件路径加引号。

## 使用方法

### 1. 检查原文档

```bash
python scripts/inspect_docx.py "input.docx"
```

脚本会输出 JSON 检查结果，包含文档结构、分节和正文块等信息。如需保存报告，可加 `--json report.json`。

### 2. 生成规范化副本

```bash
python scripts/clean_docx.py "input.docx"
```

脚本在原文件所在目录生成一份新 DOCX，并在清理后完成文内分段。输出文件采用 `<原文件名>_规范化_<北京时间日期>-V<版本号>.docx` 格式；同一天再次运行会递增版本号，不覆盖已有文件。脚本会打印实际输出路径。

默认段落间距为 12 磅，可以通过 `--spacing-pt` 调整：

```bash
python scripts/clean_docx.py "input.docx" --spacing-pt 10
```

### 3. 验证输出

将第二步打印的实际路径传给验证脚本：

```bash
python scripts/verify_docx.py "input.docx" "生成的规范化文件.docx"
```

脚本输出 JSON 验证报告；验证失败时返回非零退出码。如需保存报告，可加 `--json verify-report.json`。验证失败或发现正文视觉对象丢失时，不应将输出文件当作完成品。

### 可选：导出独立的分段 DOCX

仓库还提供单独的分段导出脚本，用于将文档按安全边界导出到指定目录。这是独立于上述“一份规范化副本”流程的功能：

```bash
python scripts/export_segments.py "input.docx" "segments"
```

可用 `--preamble first|separate|omit` 指定首个标题之前内容的处理方式，默认是 `first`。脚本会输出包含覆盖与验证结果的 JSON 报告；目标目录已有文件时，只有明确指定 `--overwrite` 才允许覆盖。

当前独立导出功能不支持多分节文档；多分节文档仍可使用上面的规范化副本流程。

## 作为 Skill 使用

本仓库根目录的 [SKILL.md](SKILL.md) 定义了完整的工作流程与安全边界；`agents/openai.yaml` 提供 Skill 的显示名称、简介和默认提示词。在支持 Skills 的 Codex 环境中安装本目录后，可以使用 `$word-content-normalizer` 调用。

## 项目结构

```text
.
├── SKILL.md                  # Skill 工作流程与交付规则
├── agents/openai.yaml        # Skill 界面信息
├── references/               # 分类与转换细则
├── scripts/                  # 检查、清理、分段、导出和验证脚本
└── tests/                    # 分段与标题处理测试
```

运行测试：

```bash
python -m unittest discover -s tests -v
```
