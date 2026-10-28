## 可直接输出多份 DOCX

from __future__ import annotations

import argparse
import json
import os
import re
import tempfile
import zipfile
from pathlib import Path

from lxml import etree

from segment_docx import segment_for_export
from verify_docx import verify_segment


DOCUMENT_XML = "word/document.xml"
BODY_BLOCK_TAGS = {"p", "tbl"}
SUPPORTED_BODY_TAGS = BODY_BLOCK_TAGS | {"sectPr"}


def safe_filename(value, fallback="segment"):
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value or "")
    value = re.sub(r"\s+", "_", value).strip(" ._")
    return (value[:48].rstrip(" ._") or fallback)


def _selected_document_xml(source, keep_indices):
    with zipfile.ZipFile(source) as archive:
        xml = archive.read(DOCUMENT_XML)

    parser = etree.XMLParser(resolve_entities=False, remove_blank_text=False)
    root = etree.fromstring(xml, parser)
    bodies = root.xpath(
        "//w:body",
        namespaces={
            "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        },
    )

    if len(bodies) != 1:
        raise ValueError("DOCX must contain exactly one document body")

    body = bodies[0]
    unsupported = [
        etree.QName(child).localname
        for child in body
        if etree.QName(child).localname not in SUPPORTED_BODY_TAGS
    ]

    if unsupported:
        raise ValueError(
            "Unsupported top-level body elements: "
            + ", ".join(sorted(set(unsupported)))
        )

    keep_indices = set(keep_indices)
    block_index = 0

    for child in list(body):
        local_name = etree.QName(child).localname

        if local_name not in BODY_BLOCK_TAGS:
            continue

        if block_index not in keep_indices:
            body.remove(child)

        block_index += 1

    missing = keep_indices - set(range(block_index))

    if missing:
        raise IndexError(
            "Segment refers to missing body indices: "
            + ", ".join(str(index) for index in sorted(missing))
        )

    return etree.tostring(
        root,
        xml_declaration=True,
        encoding="UTF-8",
        standalone=True,
    )


def write_selected_blocks(source, output, keep_indices):
    """Copy a DOCX while retaining only selected body paragraphs/tables."""

    replacement_xml = _selected_document_xml(source, keep_indices)
    output.parent.mkdir(parents=True, exist_ok=True)
    handle = tempfile.NamedTemporaryFile(
        prefix=".segment_",
        suffix=".docx",
        dir=output.parent,
        delete=False,
    )
    temporary = Path(handle.name)
    handle.close()

    try:
        with zipfile.ZipFile(source) as source_zip:
            with zipfile.ZipFile(temporary, "w") as output_zip:
                for info in source_zip.infolist():
                    payload = (
                        replacement_xml
                        if info.filename == DOCUMENT_XML
                        else source_zip.read(info.filename)
                    )
                    output_zip.writestr(info, payload)

        os.replace(temporary, output)
    finally:
        if temporary.exists():
            temporary.unlink()


def _segment_indices(segment):
    indices = [item["index"] for item in segment["elements"]]

    if indices != sorted(set(indices)):
        raise ValueError("Export segment contains duplicate or unordered indices")

    if any(right != left + 1 for left, right in zip(indices, indices[1:])):
        raise ValueError("Export segment body indices must be contiguous")

    return indices


def _expected_export_indices(plan, preamble_mode):
    if preamble_mode != "omit":
        all_indices = [
            item["index"]
            for segment in plan["final_blocks"]
            for item in segment["elements"]
        ]
        return list(range(max(all_indices) + 1)) if all_indices else []

    return sorted(
        item["index"]
        for block in plan["initial_blocks"]
        for item in block["elements"]
    )


def export_segments(
    source,
    output_dir,
    *,
    preamble_mode="first",
    overwrite=False,
):
    source = Path(source)
    output_dir = Path(output_dir)

    if source.suffix.lower() != ".docx":
        raise ValueError("Input must be a .docx file")

    if not source.is_file():
        raise FileNotFoundError(source)

    plan = segment_for_export(source, preamble_mode=preamble_mode)
    planned = []

    for position, segment in enumerate(plan["final_blocks"], start=1):
        title = segment.get("display_title") or segment.get("title") or "segment"
        filename = f"segment_{position:03d}_{safe_filename(title)}.docx"
        final_path = output_dir / filename
        planned.append((segment, _segment_indices(segment), final_path))

    exported_indices = [
        index
        for _, indices, _ in planned
        for index in indices
    ]
    expected_indices = _expected_export_indices(plan, preamble_mode)
    coverage_passed = (
        len(exported_indices) == len(set(exported_indices))
        and sorted(exported_indices) == expected_indices
    )

    if not coverage_passed:
        raise RuntimeError(
            "Segment plan does not cover the expected source body blocks "
            "exactly once"
        )

    output_dir.mkdir(parents=True, exist_ok=True)

    conflicts = [path for _, _, path in planned if path.exists()]

    if conflicts and not overwrite:
        raise FileExistsError(
            "Refusing to overwrite existing segment files: "
            + ", ".join(path.name for path in conflicts)
        )

    exported = []

    with tempfile.TemporaryDirectory(
        prefix=".segment_export_",
        dir=output_dir,
    ) as staging_dir:
        staging_dir = Path(staging_dir)

        for segment, indices, final_path in planned:
            staged_path = staging_dir / final_path.name
            write_selected_blocks(source, staged_path, indices)
            verification = verify_segment(source, staged_path, indices)

            if not verification["passed"]:
                raise RuntimeError(
                    f"Segment verification failed for {final_path.name}: "
                    + json.dumps(verification, ensure_ascii=False)
                )

            exported.append(
                {
                    "segment_index": segment["segment_index"],
                    "title": segment.get("title"),
                    "display_title": segment.get("display_title"),
                    "output": str(final_path.resolve()),
                    "source_indices": indices,
                    "char_count": segment["char_count"],
                    "visual_count": segment["visual_count"],
                    "verification_passed": True,
                }
            )

        for _, _, final_path in planned:
            os.replace(staging_dir / final_path.name, final_path)

    return {
        "source": str(source.resolve()),
        "output_dir": str(output_dir.resolve()),
        "boundary_mode": plan["boundary_mode"],
        "preamble_mode": preamble_mode,
        "segment_count": len(exported),
        "coverage_passed": coverage_passed,
        "covered_source_indices": sorted(exported_indices),
        "segments": exported,
        "passed": (
            coverage_passed
            and all(item["verification_passed"] for item in exported)
        ),
    }


def main():
    parser = argparse.ArgumentParser(
        description=(
            "Export DOCX segments at safe paragraph/table boundaries "
            "without rebuilding visual objects."
        )
    )
    parser.add_argument("input", type=Path)
    parser.add_argument("output_dir", type=Path)
    parser.add_argument(
        "--preamble",
        choices=("first", "separate", "omit"),
        default="first",
    )
    parser.add_argument("--overwrite", action="store_true")
    parser.add_argument("--json", type=Path)
    args = parser.parse_args()

    report = export_segments(
        args.input,
        args.output_dir,
        preamble_mode=args.preamble,
        overwrite=args.overwrite,
    )
    payload = json.dumps(report, ensure_ascii=False, indent=2)

    if args.json:
        args.json.parent.mkdir(parents=True, exist_ok=True)
        args.json.write_text(payload, encoding="utf-8")

    print(payload)


if __name__ == "__main__":
    main()
