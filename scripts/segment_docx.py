from __future__ import annotations

import argparse
from copy import deepcopy
import json
import re
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from docx_common import (
    iter_blocks,
    has_visual,
    detect_block_heading_paragraph,
    paragraph_style_name,
)

MIN_CHARS = 150
MERGE_MAX_CHARS = 1000

SPLIT_TRIGGER_CHARS = 800
SPLIT_TARGET_CHARS = 800
SPLIT_BACKWARD_MIN = 500
PRIMARY_HEADING_MAX_STYLE_LEVEL = 3


WORD_COUNT_PATTERN = re.compile(r"[A-Za-z0-9_]+|[^\s]")


def count_chars(text: str) -> int:
    """按 Word 状态栏的“字数”口径统计分块长度。

    连续英文、数字和下划线算一个词；中文和标点逐个计数；
    空白与图片本身不计数。保留原函数名以兼容现有调用。
    """
    return sum(1 for _ in WORD_COUNT_PATTERN.finditer(text or ""))


def table_text(table: Table) -> str:
    """
    提取表格中的可见文字，仅用于字符统计和日志检查。

    不修改表格。
    """
    parts = []

    for row in table.rows:
        for cell in row.cells:
            text = cell.text.strip()

            if text:
                parts.append(text)

    return "\n".join(parts)


def element_info(index, block):
    """
    将一个 Word block 转成便于检查的日志信息。

    注意：
    这里只生成日志数据，
    不修改 Word，不复制图片。
    """

    if isinstance(block, Table):

        text = table_text(block)

        return {
            "index": index,
            "type": "table",
            "text": text,
            "char_count": count_chars(text),
            "has_visual": has_visual(block),
        }

    # Paragraph
    text = block.text.strip()

    return {
        "index": index,
        "type": "paragraph",
        "text": text,
        "char_count": count_chars(text),
        "has_visual": has_visual(block),
        "style": paragraph_style_name(block),
    }


def heading_tier(paragraph, heading):
    """Classify a detected heading for segmentation without changing its style.

    Use the actual Heading style before Word numbering metadata: a Heading 5
    list item can report numbering level 1 while remaining a small heading.
    """
    if heading.get("kind") == "numeric_parenthesized":
        return "minor"
    if heading.get("source") == "text":
        level = heading.get("level")
        return "major" if level is not None and level <= 3 else "minor"
    style = paragraph.style
    while style is not None:
        match = re.fullmatch(
            r"(?:Heading|标题)\s*([1-9]\d*)", style.name, re.IGNORECASE
        )
        if match:
            return (
                "major" if int(match.group(1)) <= PRIMARY_HEADING_MAX_STYLE_LEVEL
                else "minor"
            )
        style = style.base_style
    level = heading.get("level")
    return "major" if level is not None and level <= 3 else "minor"


def summarize_elements(elements):
    """
    汇总一个 Initial Block 内所有元素。
    """

    return {
        "char_count": sum(
            item["char_count"]
            for item in elements
        ),
        "visual_count": sum(
            1
            for item in elements
            if item["has_visual"]
        ),
        "start_index": (
            elements[0]["index"]
            if elements
            else None
        ),
        "end_index": (
            elements[-1]["index"]
            if elements
            else None
        ),
    }

def merge_short_blocks(blocks):
    """
    按长度规则向后合并连续 Initial Blocks。

    规则：
    - 当前块 >= 150 字：直接保留
    - 当前块 < 150 字：尝试向后合并完整的后续初始块
    - 后续块若只有标题，连同紧随其后的有内容块一起合并
    - 合并后 <= 1000 字：允许合并
    - 合并后 > 1000 字：停止
    - 即使停止后仍 < 150 字，也允许保留

    不限制标题 level。
    只要求原文顺序连续。
    """

    merged = []
    i = 0

    while i < len(blocks):

        # 复制当前 block，避免修改原始 blocks
        current = {
            **blocks[i],
            "elements": list(
                blocks[i]["elements"]
            ),
            "source_block_indices": [
                blocks[i]["block_index"]
            ],
        }

        # 当前块不足150字时，不断尝试向后合并
        while (
            current["char_count"] < MIN_CHARS
            and i + 1 < len(blocks)
        ):
            end = i + 1
            while end + 1 < len(blocks) and not any(
                item["type"] == "table" or item["has_visual"] or item["text"].strip()
                for item in blocks[end]["elements"][1:]
            ):
                end += 1

            following = blocks[i + 1:end + 1]
            new_count = current["char_count"] + sum(
                block["char_count"] for block in following
            )

            # 加入下一块以后超过1000，则停止
            if new_count > MERGE_MAX_CHARS:
                break

            # 整组加入，避免只把一个标题留在当前分段末尾。
            for next_block in following:
                current["elements"].extend(next_block["elements"])
                current["source_block_indices"].append(next_block["block_index"])

            # 重新计算合并后的统计信息
            current.update(
                summarize_elements(
                    current["elements"]
                )
            )

            i = end

        current["segment_index"] = len(
            merged
        )

        merged.append(current)

        i += 1

    return merged

def compact_elements_text(elements):
    """
    将 elements 中的文字拼起来，
    去掉空格、换行、Tab。

    只用于计算拆分位置。
    """
    return "".join(
        re.sub(
            r"\s+",
            "",
            item.get("text", ""),
        )
        for item in elements
    )


def find_sentence_split_count(elements):
    """
    以 800 字为目标寻找句号切点。

    规则：
    1. 先找 800 前面的最后一个“。”
    2. 如果该位置 >= 500，就在那里切
    3. 如果前面的句号导致不足 500，
       就从 800 往后找第一个“。”
    4. 前后都没有句号，最后才在 800 字硬切

    返回值是：
    从当前内容开头数，需要保留多少个可见字符。
    """

    text = compact_elements_text(
        elements
    )

    if len(text) <= SPLIT_TRIGGER_CHARS:
        return None

    # 800 前面的最后一个句号
    previous_period = text.rfind(
        "。",
        0,
        SPLIT_TARGET_CHARS,
    )

    if previous_period != -1:

        previous_cut = (
            previous_period + 1
        )

        if (
            previous_cut
            >= SPLIT_BACKWARD_MIN
        ):
            return previous_cut

    # 前面的切点太短，
    # 从800以后找第一个句号
    next_period = text.find(
        "。",
        SPLIT_TARGET_CHARS,
    )

    if next_period != -1:
        return next_period + 1

    # 实在没有句号
    return min(
        SPLIT_TARGET_CHARS,
        len(text),
    )


def split_text_by_visible_count(
    text,
    cut_count,
):
    """
    按“去掉空白后的字符数”切一个文本。

    返回：
    left_text, right_text
    """

    visible_count = 0

    for pos, char in enumerate(text):

        if not char.isspace():
            visible_count += 1

        if visible_count >= cut_count:

            return (
                text[:pos + 1],
                text[pos + 1:],
            )

    return text, ""


def split_elements_at_count(
    elements,
    cut_count,
):
    """
    根据字符位置拆 elements。

    如果切点正好落在一个 paragraph 内，
    就把这个 paragraph 的文字拆成左右两部分。

    当前版本主要用于纯文字测试。
    """

    left = []
    right = []

    remaining = cut_count

    for position, item in enumerate(
        elements
    ):

        item_count = item[
            "char_count"
        ]

        # 已经到切点
        if remaining <= 0:

            right.extend(
                dict(x)
                for x in elements[position:]
            )

            break

        # 0字对象，例如图片：
        # 当前阶段先保持原顺序
        if item_count == 0:

            left.append(
                dict(item)
            )

            continue

        # 整个 element 都属于左边
        if item_count <= remaining:

            left.append(
                dict(item)
            )

            remaining -= item_count
            continue

        # ----------------------------------------------
        # 切点落在当前 paragraph 内部
        # ----------------------------------------------

        # 表格或视觉对象暂时不从内部拆
        if (
            item.get("type") == "table"
            or item.get(
                "has_visual",
                False,
            )
        ):

            right.extend(
                dict(x)
                for x in elements[position:]
            )

            break

        left_text, right_text = (
            split_text_by_visible_count(
                item.get(
                    "text",
                    "",
                ),
                remaining,
            )
        )

        if left_text:

            left_item = dict(item)

            left_item[
                "text"
            ] = left_text

            left_item[
                "char_count"
            ] = count_chars(
                left_text
            )

            left_item[
                "split_fragment"
            ] = True

            left.append(
                left_item
            )

        if right_text:

            right_item = dict(item)

            right_item[
                "text"
            ] = right_text

            right_item[
                "char_count"
            ] = count_chars(
                right_text
            )

            right_item[
                "split_fragment"
            ] = True

            right.append(
                right_item
            )

        right.extend(
            dict(x)
            for x in elements[
                position + 1:
            ]
        )

        break

    return left, right


def make_split_part(
    original_block,
    elements,
):
    """
    根据一部分 elements 生成拆分结果。

    title 仍然继承原来的标题。
    """

    part = {
        key: value
        for key, value
        in original_block.items()
        if key != "elements"
    }

    part["elements"] = elements

    part.update(
        summarize_elements(
            elements
        )
    )

    return part


def split_long_block(block):
    """
    拆分单个 >800 字的 block。
    """

    if (
        block["char_count"]
        <= SPLIT_TRIGGER_CHARS
    ):
        return [
            {
                **block,
                "elements": list(
                    block["elements"]
                ),
                "part_index": 1,
                "part_total": 1,
                "display_title": (
                    block["title"]
                ),
                "split_from_long": False,
            }
        ]

    parts = []

    remaining_elements = list(
        block["elements"]
    )

    while True:

        remaining_count = sum(
            item["char_count"]
            for item
            in remaining_elements
        )

        if (
            remaining_count
            <= SPLIT_TRIGGER_CHARS
        ):
            break

        cut_count = (
            find_sentence_split_count(
                remaining_elements
            )
        )

        if cut_count is None:
            break

        left, right = (
            split_elements_at_count(
                remaining_elements,
                cut_count,
            )
        )

        # 防止异常情况下死循环
        if not left or not right:
            break

        parts.append(
            make_split_part(
                block,
                left,
            )
        )

        remaining_elements = right

    if remaining_elements:

        parts.append(
            make_split_part(
                block,
                remaining_elements,
            )
        )

    total = len(parts)

    for index, part in enumerate(
        parts,
        start=1,
    ):

        part[
            "part_index"
        ] = index

        part[
            "part_total"
        ] = total

        part[
            "split_from_long"
        ] = total > 1

        if total > 1:

            part[
                "display_title"
            ] = (
                f"{block['title']}"
                f"（{index}/{total}）"
            )

        else:

            part[
                "display_title"
            ] = block["title"]

    return parts


def split_long_blocks(blocks):
    """
    对所有 merged blocks 执行长块拆分。
    """

    result = []

    for block in blocks:

        parts = split_long_block(
            block
        )

        for part in parts:

            part[
                "segment_index"
            ] = len(result)

            result.append(
                part
            )

    return result


def _safe_boundary_position(elements):
    """Prefer major headings, then minor headings, then other element edges."""

    if len(elements) < 2:
        return None

    cumulative = 0
    candidates = []
    total = sum(item["char_count"] for item in elements)

    for position, item in enumerate(elements[:-1], start=1):
        cumulative += item["char_count"]

        # Do not leave a short heading paragraph by itself when more content
        # is available. A continuation may exceed 800 characters instead.
        if (
            position == 1
            and item.get("type") == "paragraph"
            and item["char_count"] < MIN_CHARS
        ):
            continue

        priority = {"major": 0, "minor": 1}.get(
            elements[position].get("heading_tier"), 2
        )
        candidates.append((position, cumulative, total - cumulative, priority))

    if not candidates:
        return None

    balanced = [
        candidate for candidate in candidates
        if candidate[1] >= MIN_CHARS and candidate[2] >= MIN_CHARS
    ]
    pool = balanced or candidates
    best_priority = min(candidate[3] for candidate in pool)
    pool = [candidate for candidate in pool if candidate[3] == best_priority]
    preferred = [candidate for candidate in pool if candidate[1] >= SPLIT_BACKWARD_MIN]
    pool = preferred or pool

    return min(
        pool,
        key=lambda candidate: (
            abs(candidate[1] - SPLIT_TARGET_CHARS),
            candidate[1] > SPLIT_TARGET_CHARS,
        ),
    )[0]


def split_long_block_at_boundaries(block):
    """Split a long block only between paragraphs or tables."""

    remaining = [dict(item) for item in block["elements"]]
    parts = []

    while (
        sum(item["char_count"] for item in remaining)
        > SPLIT_TRIGGER_CHARS
    ):
        position = _safe_boundary_position(remaining)

        if position is None:
            break

        left = remaining[:position]
        right = remaining[position:]

        if not left or not right:
            break

        parts.append(make_split_part(block, left))
        remaining = right

    if remaining:
        parts.append(make_split_part(block, remaining))

    total = len(parts)

    for index, part in enumerate(parts, start=1):
        part["part_index"] = index
        part["part_total"] = total
        part["split_from_long"] = total > 1
        part["split_mode"] = "element_boundary"
        part["display_title"] = (
            f"{block['title']}（{index}/{total}）"
            if total > 1
            else block["title"]
        )

    return parts


def split_long_blocks_at_boundaries(blocks):
    """Create export-safe segments without splitting paragraph XML."""

    result = []

    for block in blocks:
        for part in split_long_block_at_boundaries(block):
            part["segment_index"] = len(result)
            result.append(part)

    return result



def build_initial_blocks(doc, include_blank=False):
    """
    根据较大标题边界生成 Initial Blocks。

    一个 Initial Block：

        一个标题
        +
        到下一个较大标题之前的所有内容（包括小标题及其说明）

    第一个标题之前的内容单独放入 preamble。
    """

    preamble_elements = []
    blocks = []

    current_block = None

    for index, block in enumerate(
        iter_blocks(doc)
    ):

        # --------------------------------------------------
        # 表格不会直接作为标题
        # --------------------------------------------------

        if isinstance(block, Table):

            info = element_info(
                index,
                block,
            )

            if current_block is None:
                preamble_elements.append(info)
            else:
                current_block[
                    "elements"
                ].append(info)

            continue

        # --------------------------------------------------
        # Paragraph
        # --------------------------------------------------

        heading = (
            detect_block_heading_paragraph(
                block
            )
        )

        # 普通日志可以忽略纯空白段落；导出模式必须保留它们，
        # 因为其中可能包含分页、书签或其他布局信息。
        if (
            not block.text.strip()
            and not has_visual(block)
            and not include_blank
        ):
            continue

        info = element_info(
            index,
            block,
        )
        if heading:
            info["heading_tier"] = heading_tier(block, heading)

        # --------------------------------------------------
        # 遇到新的标题
        # --------------------------------------------------

        if heading and info["heading_tier"] == "major":

            # 先结束前一个 block
            if current_block is not None:

                summary = summarize_elements(
                    current_block["elements"]
                )

                current_block.update(
                    summary
                )

                blocks.append(
                    current_block
                )

            # 开始新的 block
            current_block = {
                "block_index": len(blocks),
                "title": block.text.strip(),
                "heading_level": heading.get(
                    "level"
                ),
                "heading_kind": heading.get(
                    "kind"
                ),
                "heading_number": heading.get(
                    "number"
                ),
                "heading_source": heading.get(
                    "source"
                ),
                "elements": [
                    info
                ],
            }

            continue

        # --------------------------------------------------
        # 普通正文 / 图片 / 表格
        # --------------------------------------------------

        if current_block is None:

            # 第一个标题之前
            preamble_elements.append(
                info
            )

        else:

            current_block[
                "elements"
            ].append(info)

    # ------------------------------------------------------
    # 文档结束时收尾
    # ------------------------------------------------------

    if current_block is not None:

        summary = summarize_elements(
            current_block["elements"]
        )

        current_block.update(
            summary
        )

        blocks.append(
            current_block
        )

    # ------------------------------------------------------
    # Preamble
    # ------------------------------------------------------

    preamble_summary = summarize_elements(
        preamble_elements
    )

    preamble = {
        **preamble_summary,
        "elements": preamble_elements,
    }

    return {
        "preamble": preamble,
        "blocks": blocks,
    }


def segment_for_document(doc, title="document", preamble_mode="first"):
    """Plan contiguous segments in one DOCX without changing its body elements."""
    if preamble_mode not in {"first", "separate", "omit"}:
        raise ValueError(
            "preamble_mode must be one of: first, separate, omit"
        )

    result = build_initial_blocks(doc, include_blank=True)
    initial_blocks = result["blocks"]
    preamble = result["preamble"]
    merge_inputs = list(initial_blocks)
    if preamble_mode == "first" and preamble["elements"] and merge_inputs:
        first = {
            **merge_inputs[0],
            "elements": list(preamble["elements"]) + list(merge_inputs[0]["elements"]),
        }
        first.update(summarize_elements(first["elements"]))
        merge_inputs[0] = first
    merged_blocks = merge_short_blocks(merge_inputs)

    if preamble["elements"] and preamble_mode != "omit":
        preamble_block = {
            "block_index": -1,
            "title": title,
            "heading_level": None,
            "heading_kind": "preamble",
            "heading_number": None,
            "heading_source": None,
            "elements": list(preamble["elements"]),
            "source_block_indices": [-1],
        }
        preamble_block.update(
            summarize_elements(preamble_block["elements"])
        )

        if preamble_mode == "first" and merged_blocks:
            merged_blocks[0]["source_block_indices"].insert(0, -1)
        else:
            merged_blocks.insert(0, preamble_block)

    if not merged_blocks:
        raise ValueError("No exportable paragraphs or tables were found")

    final_blocks = split_long_blocks_at_boundaries(merged_blocks)

    return {
        "boundary_mode": "paragraph_or_table",
        "preamble_mode": preamble_mode,
        "initial_block_count": len(initial_blocks),
        "merged_block_count": len(merged_blocks),
        "final_block_count": len(final_blocks),
        "preamble": preamble,
        "initial_blocks": initial_blocks,
        "merged_blocks": merged_blocks,
        "final_blocks": final_blocks,
    }


def segment_for_export(path: Path, preamble_mode="first"):
    """Build DOCX-exportable segments using only whole body elements."""
    path = Path(path)
    doc = Document(path)
    if len(doc.sections) != 1:
        raise ValueError("Segment export supports single-section documents only")
    plan = segment_for_document(doc, path.stem, preamble_mode)
    plan["file"] = str(path.resolve())
    return plan


def segment_starts(plan):
    """Return the first body index of every segment, including the first."""
    return [block["start_index"] for block in plan["final_blocks"]]


def segment_coverage(plan, block_count):
    indices = [
        item["index"]
        for block in plan["final_blocks"]
        for item in block["elements"]
    ]
    return indices == list(range(block_count))


def _sentence_cut(text, target_count):
    """Find a sentence end near the target using the segment word metric."""
    candidates = []
    for match in re.finditer(r"[。！？!?]|(?<!\d)\.(?=\s|$)", text):
        position = match.end()
        left = count_chars(text[:position])
        right = count_chars(text[position:])
        if left >= MIN_CHARS and right >= MIN_CHARS:
            candidates.append((abs(left - target_count), position))
    return min(candidates)[1] if candidates else None


def _simple_text_runs(paragraph):
    """Return text runs only when a paragraph has no objects or field XML."""
    if paragraph.style.name.lower().startswith("list"):
        return None
    node = paragraph._p
    if node.pPr is not None and (
        node.pPr.find(qn("w:numPr")) is not None
        or node.pPr.find(qn("w:sectPr")) is not None
    ):
        return None
    runs = []
    for child in node:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag != qn("w:r"):
            return None
        texts = [part for part in child if part.tag == qn("w:t")]
        if len(texts) != 1 or any(
            part.tag not in {qn("w:rPr"), qn("w:t")} for part in child
        ):
            return None
        runs.append(child)
    return runs if runs else None


def _split_plain_paragraph(paragraph, cut):
    """Split text-only run XML at a sentence end, preserving run formatting."""
    runs = _simple_text_runs(paragraph)
    if runs is None:
        return False
    original = paragraph.text
    if not 0 < cut < len(original):
        return False
    right_node = deepcopy(paragraph._p)
    right_runs = list(right_node.findall(qn("w:r")))
    cursor = 0
    for left_run, right_run in zip(runs, right_runs):
        left_text_node = left_run.find(qn("w:t"))
        right_text_node = right_run.find(qn("w:t"))
        run_text = left_text_node.text or ""
        left_length = max(0, min(len(run_text), cut - cursor))
        left_text = run_text[:left_length]
        right_text = run_text[left_length:]
        if left_text:
            left_text_node.text = left_text
        else:
            paragraph._p.remove(left_run)
        if right_text:
            right_text_node.text = right_text
            if right_text[0].isspace() or right_text[-1].isspace():
                right_text_node.set(
                    "{http://www.w3.org/XML/1998/namespace}space", "preserve"
                )
        else:
            right_node.remove(right_run)
        cursor += len(run_text)
    paragraph._p.addnext(right_node)
    if paragraph.text + Paragraph(right_node, paragraph._parent).text != original:
        raise RuntimeError("Sentence split changed paragraph text")
    return True


def _split_oversized_sentences(doc):
    """Add sentence boundaries only where heading and element edges cannot split."""
    for _ in range(1000):
        plan = segment_for_document(doc)
        blocks = list(iter_blocks(doc))
        changed = False
        for segment in plan["final_blocks"]:
            if segment["char_count"] <= SPLIT_TRIGGER_CHARS:
                continue
            prefix_count = 0
            for item in segment["elements"]:
                if item["type"] == "paragraph" and not item.get("heading_tier"):
                    paragraph = blocks[item["index"]]
                    if _simple_text_runs(paragraph) is not None:
                        cut = _sentence_cut(
                            paragraph.text, SPLIT_TARGET_CHARS - prefix_count
                        )
                        if cut is not None and _split_plain_paragraph(paragraph, cut):
                            changed = True
                            break
                prefix_count += item["char_count"]
            if changed:
                break
        if not changed:
            return
    raise RuntimeError("Sentence splitting did not converge")


def apply_segments_in_place(doc):
    """Place one independent `##` line immediately before every segment."""
    _split_oversized_sentences(doc)
    plan = segment_for_document(doc)
    blocks = list(iter_blocks(doc))
    if not segment_coverage(plan, len(blocks)):
        raise RuntimeError("Segment plan must cover every body block exactly once")
    try:
        marker_style = doc.styles["WCN Segment Marker"]
    except KeyError:
        marker_style = doc.styles.add_style("WCN Segment Marker", WD_STYLE_TYPE.PARAGRAPH)
    for index in segment_starts(plan):
        marker = OxmlElement("w:p")
        blocks[index]._element.addprevious(marker)
        paragraph = Paragraph(marker, doc)
        paragraph.style = marker_style
        paragraph.add_run("##")
    return plan


def verify_in_document_segments(doc):
    blocks = list(iter_blocks(doc))
    marked_indices = []
    original_index = 0
    for block in blocks:
        if (isinstance(block, Paragraph) and block.style.name == "WCN Segment Marker"
                and block.text == "##"):
            marked_indices.append(original_index)
            block._element.getparent().remove(block._element)
        else:
            original_index += 1
    plan = segment_for_document(doc)
    starts = segment_starts(plan)
    coverage = segment_coverage(plan, original_index)
    marked = coverage and marked_indices == starts
    return {
        "segment_count": len(plan["final_blocks"]),
        "segment_char_counts": [block["char_count"] for block in plan["final_blocks"]],
        "segment_coverage_passed": coverage,
        "segment_boundaries_marked": bool(marked),
        "segment_boundary_indices": starts,
        "marked_boundary_indices": marked_indices,
    }


def segment(path: Path):
    """Inspect the same safe boundaries used by in-document normalization."""
    return segment_for_export(path)


def main():
    ap = argparse.ArgumentParser(
        description=(
            "Generate initial content blocks "
            "from DOCX headings."
        )
    )

    ap.add_argument(
        "input",
        type=Path,
    )

    ap.add_argument(
        "--json",
        type=Path,
    )

    args = ap.parse_args()

    result = segment(
        args.input
    )

    payload = json.dumps(
        result,
        ensure_ascii=False,
        indent=2,
    )

    if args.json:

        args.json.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        args.json.write_text(
            payload,
            encoding="utf-8",
        )

    print(payload)


if __name__ == "__main__":
    main()
