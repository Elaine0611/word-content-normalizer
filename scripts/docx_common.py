##读取 Word 原始结构+识别视觉对象+识别标题+识别分段边界+识别目录/修订历史+验证图片
## 公共识别与读取工具：负责识别标题、目录域、修订表、视觉对象、图片哈希等，但它本身不会修改或输出 Word 文件。

from __future__ import annotations

import hashlib
import re
from datetime import date, datetime, timedelta, timezone
from pathlib import Path

from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.ns import qn


VISUAL_TAGS = (
    "<w:drawing",
    "<w:pict",
    "<v:shape",
    "<v:imagedata",
    "<wp:anchor",
    "<wp:inline",
    "<o:OLEObject",
    "<w:object",
)


def iter_blocks(doc):
    for child in doc.element.body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, doc)
        elif child.tag.endswith("}tbl"):
            yield Table(child, doc)


def has_visual(obj):
    xml = obj._element.xml if hasattr(obj, "_element") else obj.xml
    return any(tag in xml for tag in VISUAL_TAGS)


def norm(text):
    return re.sub(
        r"[\s【】\[\]<>《》]",
        "",
        text or "",
    ).lower()


def paragraph_style_name(p):
    """Return a paragraph style name without assuming the style resolves."""

    style = getattr(p, "style", None)
    return getattr(style, "name", None) or ""


def image_hashes(doc):
    out = []

    for rel in doc.part.rels.values():
        part = rel.target_part

        if getattr(
            part,
            "content_type",
            "",
        ).startswith("image/"):
            out.append(
                hashlib.sha256(
                    part.blob
                ).hexdigest()
            )

    return sorted(out)


def body_visual_count(doc):
    xml = doc.element.body.xml

    return sum(
        xml.count(tag)
        for tag in (
            "<w:drawing",
            "<w:pict",
            "<w:object",
        )
    )


# ------------------------------------------------------
# 标题识别
# ------------------------------------------------------

CHINESE_NUM = "一二三四五六七八九十百"


def decimal_heading_level(number):
    """
    根据数字编号计算标题层级。

    1       -> 1
    1.0     -> 1
    2.0     -> 1

    1.1     -> 2
    2.1     -> 2

    1.1.1   -> 3
    2.1.1.1 -> 4

    2.1.0   -> 2

    末尾 .0 不增加层级。
    """

    parts = number.split(".")

    while (
        len(parts) > 1
        and parts[-1] == "0"
    ):
        parts.pop()

    return len(parts)


def detect_heading(text):
    """
    根据段落文字识别标题。

    如果不是标题，返回 None。
    """

    text = (text or "").strip()

    if not text:
        return None

    # 防止长正文因为以数字开头被误识别
    if len(text) > 80:
        return None

    # --------------------------------------------------
    # 1. 中文一级标题
    #
    # 一、目的
    # 二、适用范围
    # --------------------------------------------------

    m = re.match(
        rf"^([{CHINESE_NUM}]+)、\s*(\S.*)$",
        text,
    )

    if m:
        return {
            "level": 1,
            "kind": "chinese_major",
            "number": m.group(1),
            "text": text,
        }

    # --------------------------------------------------
    # 2. 第一章 / 第一节 / 第一部分
    # --------------------------------------------------

    m = re.match(
        rf"^第([{CHINESE_NUM}]+)[章节部分]\s*(\S.*)$",
        text,
    )

    if m:
        return {
            "level": 1,
            "kind": "chinese_major",
            "number": m.group(1),
            "text": text,
        }

    # --------------------------------------------------
    # 3. 中文括号标题
    #
    # （一）范围
    # （二）说明
    # --------------------------------------------------

    m = re.match(
        rf"^[（(]([{CHINESE_NUM}]+)[）)]\s*(\S.*)$",
        text,
    )

    if m:
        return {
            "level": 2,
            "kind": "chinese_parenthesized",
            "number": m.group(1),
            "text": text,
        }

    # --------------------------------------------------
    # 4. 数字标题
    #
    # 1.目的
    # 1.0目的
    # 1.1 范围
    # 1.1.1 使用方法
    # 2.1.1.1 TCR的作用
    #
    # 注意：
    # 空格可有可无。
    # --------------------------------------------------

    # 多级编号允许紧跟标题，例如 1.1范围；
    # 单个整数必须带“.”或“、”，避免把“2024年报告”误判为标题。
    m = re.match(
        r"^(\d+(?:\.\d+)+)(?:[.、])?\s*([^\d\s].*)$",
        text,
    )

    if not m:
        m = re.match(
            r"^(\d+)[.、]\s*([^\d\s].*)$",
            text,
        )

    if m:
        number = m.group(1)

        return {
            "level": decimal_heading_level(
                number
            ),
            "kind": "decimal",
            "number": number,
            "text": text,
        }

    return None


# ------------------------------------------------------
# Word 标题样式识别
# ------------------------------------------------------

def _heading_name_level(name):
    """
    Heading 1 -> 1
    Heading 5 -> 5
    标题 2 -> 2
    """

    name = (name or "").strip()

    m = re.match(
        r"^(?:Heading|标题)\s*([1-9]\d*)$",
        name,
        re.I,
    )

    if not m:
        return None

    return int(m.group(1))


def _style_chain(style):
    """
    当前样式 -> 父样式 -> 更上级父样式
    """

    visited = set()

    while style is not None:

        style_id = getattr(
            style,
            "style_id",
            None,
        )

        if style_id in visited:
            break

        if style_id:
            visited.add(style_id)

        yield style

        style = getattr(
            style,
            "base_style",
            None,
        )


def _get_style_ilvl(style):
    """
    获取样式里的 Word 自动编号层级 ilvl。
    """

    pPr = style._element.find(
        qn("w:pPr")
    )

    if pPr is None:
        return None

    numPr = pPr.find(
        qn("w:numPr")
    )

    if numPr is None:
        return None

    ilvl = numPr.find(
        qn("w:ilvl")
    )

    if ilvl is None:
        return None

    value = ilvl.get(
        qn("w:val")
    )

    if value is None:
        return None

    return int(value)


def _get_paragraph_ilvl(p):
    """
    获取段落自身配置的自动编号层级。
    """

    pPr = p._p.pPr

    if pPr is None:
        return None

    numPr = pPr.find(
        qn("w:numPr")
    )

    if numPr is None:
        return None

    ilvl = numPr.find(
        qn("w:ilvl")
    )

    if ilvl is None:
        return None

    value = ilvl.get(
        qn("w:val")
    )

    if value is None:
        return None

    return int(value)


def _get_style_outline_level(style):
    """
    获取 Word 样式 outline level。
    """

    pPr = style._element.find(
        qn("w:pPr")
    )

    if pPr is None:
        return None

    outline = pPr.find(
        qn("w:outlineLvl")
    )

    if outline is None:
        return None

    value = outline.get(
        qn("w:val")
    )

    if value is None:
        return None

    return int(value)


def detect_heading_style(p):
    """
    根据 Word 样式、继承关系以及自动编号判断标题。

    普通 Normal 列表即使存在编号，
    也不会仅仅因为编号而被判断为标题。
    """

    chain = list(
        _style_chain(p.style)
    )

    # --------------------------------------------------
    # 1. 先确认它属于标题样式体系
    # --------------------------------------------------

    is_heading_style = False

    for style in chain:

        if _heading_name_level(
            style.name
        ) is not None:
            is_heading_style = True
            break

        if _get_style_outline_level(
            style
        ) is not None:
            is_heading_style = True
            break

    if not is_heading_style:
        return None

    # --------------------------------------------------
    # 2. 优先读取段落自己的编号层级
    # --------------------------------------------------

    ilvl = _get_paragraph_ilvl(p)

    # --------------------------------------------------
    # 3. 再读取样式的编号层级
    # --------------------------------------------------

    if ilvl is None:

        for style in chain:

            ilvl = _get_style_ilvl(
                style
            )

            if ilvl is not None:
                break

    if ilvl is not None:

        return {
            "level": ilvl + 1,
            "kind": "word_heading_style",
            "number": None,
            "text": p.text.strip(),
            "style": paragraph_style_name(p),
            "source": "word_numbering",
        }

    # --------------------------------------------------
    # 4. 再尝试 outline level
    # --------------------------------------------------

    for style in chain:

        outline = _get_style_outline_level(
            style
        )

        if outline is not None:

            return {
                "level": outline + 1,
                "kind": "word_heading_style",
                "number": None,
                "text": p.text.strip(),
                "style": paragraph_style_name(p),
                "source": "outline_level",
            }

    # --------------------------------------------------
    # 5. 最后使用 Heading N 样式名称
    # --------------------------------------------------

    for style in chain:

        level = _heading_name_level(
            style.name
        )

        if level is not None:

            return {
                "level": level,
                "kind": "word_heading_style",
                "number": None,
                "text": p.text.strip(),
                "style": paragraph_style_name(p),
                "source": "style_name",
            }

    return None


def detect_heading_paragraph(p):
    """
    完整标题识别入口。

    优先级：

    1. 段落实际文字中的明确编号
    2. Word 标题样式 / 自动编号
    """

    # 先看文字本身
    info = detect_heading(
        p.text
    )

    if info:
        info["source"] = "text"
        return info

    # 再看 Word 样式
    return detect_heading_style(p)

# ------------------------------------------------------
# 初始内容块边界识别
# ------------------------------------------------------

def detect_numbered_item_heading(text):
    """
    识别括号数字编号。

    示例：
    (1) 在TCR主页面点击+ADD New按钮创建TCR。
    (2) TCR创建页面介绍
    （1）说明
    （2）配置

    这种编号不一定属于正式文档标题，
    但在内容分段阶段作为新的初始内容块边界。
    """

    text = (text or "").strip()

    if not text:
        return None

    m = re.match(
        r"^[（(]\s*(\d+)\s*[）)]\s*(\S.*)$",
        text,
    )

    if not m:
        return None

    return {
        "level": None,
        "kind": "numeric_parenthesized",
        "number": m.group(1),
        "text": text,
        "source": "text",
    }


def detect_block_heading_paragraph(p):
    """
    内容分段阶段统一使用的边界识别函数。

    优先级：
    1. 正式标题
       - Word Heading
       - Word 自动编号
       - 一、目的
       - 1.0目的
       - 1.1标题
       等

    2. 括号数字编号
       - (1)
       - (2)
       - （1）
       - （2）

    返回 None 表示当前段落不是新的初始内容块边界。
    """

    # 先识别正式标题
    info = detect_heading_paragraph(p)

    if info:
        return info

    # 再识别分段用的括号数字编号
    return detect_numbered_item_heading(
        p.text
    )


# ------------------------------------------------------
# 兼容现有 clean_docx.py
# ------------------------------------------------------

def major_heading(text):
    info = detect_heading(text)

    return bool(
        info
        and info["level"] == 1
    )


def subheading(text):
    info = detect_heading(text)

    return bool(
        info
        and info["level"] == 2
    )


# ------------------------------------------------------
# TOC
# ------------------------------------------------------

def toc_field(p):
    return (
        "w:instrText" in p._p.xml
        and re.search(
            r"\bTOC\b",
            p._p.xml,
            re.I,
        )
        is not None
    )


# ------------------------------------------------------
# Revision History
# ------------------------------------------------------

REVISION_HEADINGS = {
    norm(text)
    for text in (
        "版本更新内容", "版本记录", "版本变更", "修订记录", "修改记录",
        "版本修订记录", "Revision History", "Revision Record",
        "Change History", "Document History",
    )
}


def is_revision_heading(text):
    return norm(text).rstrip(":：") in REVISION_HEADINGS


def revision_score(table):
    if not table.rows:
        return 0

    header = "|".join(
        c.text.strip()
        for c in table.rows[0].cells
    ).casefold()

    groups = (
        ("版本", "version", "revision"),
        ("日期", "时间", "date"),
        ("更新页码", "页码", "page"),
        ("更新内容", "变更内容", "内容", "comment", "description of changes", "change description"),
        (
            "编辑",
            "作者",
            "修改人",
            "修订人",
            "编制",
            "author",
            "editor",
            "approved",
            "reviewed by",
        ),
    )

    return sum(
        any(term in header for term in group)
        for group in groups
    )


def revision_table(table, preceding_text=""):
    """Require several header signals, or an explicit adjacent revision title."""
    return revision_score(table) >= 3 or (
        is_revision_heading(preceding_text) and revision_score(table) >= 1
    )

def default_output(path: Path, on_date: date | None = None):
    """Choose the next dated version beside the source DOCX."""
    path = Path(path)
    day = on_date or datetime.now(timezone(timedelta(hours=8))).date()
    prefix = f"{path.stem}_规范化_{day:%Y%m%d}-V"
    pattern = re.compile(re.escape(prefix) + r"([1-9]\d*)" + re.escape(path.suffix))
    versions = [
        int(match.group(1))
        for entry in path.parent.iterdir()
        if entry.is_file() and (match := pattern.fullmatch(entry.name))
    ]
    return path.with_name(f"{prefix}{max(versions, default=0) + 1}{path.suffix}")
