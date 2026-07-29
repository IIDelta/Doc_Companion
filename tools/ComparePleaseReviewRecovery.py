"""
ComparePleaseReviewRecovery.py

Compares:
1. A generated PleaseReview DOCX/DOCM output.
2. The same file after Microsoft Word recovered and re-saved it.

The report intentionally does not export document text. It reports only:
- package part names;
- XML parse status;
- element/tag counts;
- tracked-change marker counts;
- run-property sequence counts;
- styles and relationship summaries;
- package part hashes.

Usage:
    python tools/ComparePleaseReviewRecovery.py ^
        "C:\\Path\\Generated_PleaseReview.docx" ^
        "C:\\Path\\Generated_PleaseReview_recovered.docx"

Optional output path:
    python tools/ComparePleaseReviewRecovery.py ^
        "C:\\Path\\Generated_PleaseReview.docx" ^
        "C:\\Path\\Generated_PleaseReview_recovered.docx" ^
        "C:\\Path\\PleaseReviewRecoveryComparison.json"
"""

from __future__ import annotations

import hashlib
import json
import sys
import zipfile
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

NSMAP = {
    "w": W_NS,
    "rel": REL_NS,
}


def qn(local_name: str) -> str:
    """Return a WordprocessingML qualified name."""
    return f"{{{W_NS}}}{local_name}"


def local_name(tag_name: str) -> str:
    """Return the local element name without its XML namespace."""
    if "}" in tag_name:
        return tag_name.split("}", 1)[1]

    return tag_name


W_INS = qn("ins")
W_DEL = qn("del")
W_MOVE_FROM = qn("moveFrom")
W_MOVE_TO = qn("moveTo")
W_R = qn("r")
W_RPR = qn("rPr")
W_RPR_CHANGE = qn("rPrChange")
W_PPR_CHANGE = qn("pPrChange")
W_TBL_PR_CHANGE = qn("tblPrChange")
W_TR_PR_CHANGE = qn("trPrChange")
W_TC_PR_CHANGE = qn("tcPrChange")
W_SECT_PR_CHANGE = qn("sectPrChange")
W_RSTYLE = qn("rStyle")
W_COLOR = qn("color")
W_STRIKE = qn("strike")
W_HIGHLIGHT = qn("highlight")

STYLE_PART_NAME = "word/styles.xml"
MAIN_DOCUMENT_PART_NAME = "word/document.xml"
MAIN_RELATIONSHIPS_PART_NAME = "word/_rels/document.xml.rels"


TRACKED_CHANGE_TAGS = {
    W_INS: "insertions",
    W_DEL: "deletions",
    W_MOVE_FROM: "moved_from",
    W_MOVE_TO: "moved_to",
    W_RPR_CHANGE: "run_property_changes",
    W_PPR_CHANGE: "paragraph_property_changes",
    W_TBL_PR_CHANGE: "table_property_changes",
    W_TR_PR_CHANGE: "table_row_property_changes",
    W_TC_PR_CHANGE: "table_cell_property_changes",
    W_SECT_PR_CHANGE: "section_property_changes",
}


def parse_xml(xml_bytes: bytes) -> etree._Element:
    """Parse XML securely without resolving entities."""
    parser = etree.XMLParser(
        resolve_entities=False,
        no_network=True,
        remove_blank_text=False,
        huge_tree=True,
    )

    return etree.fromstring(xml_bytes, parser=parser)


def sha256_bytes(value: bytes) -> str:
    """Return a SHA-256 digest without exposing source bytes."""
    return hashlib.sha256(value).hexdigest()


def is_xml_part(part_name: str) -> bool:
    """Return True for XML and relationship package parts."""
    normalized = part_name.lower()

    return normalized.endswith(".xml") or normalized.endswith(".rels")


def is_content_part(part_name: str) -> bool:
    """Return True for visible Word content parts."""
    normalized = part_name.lower()

    if normalized == "word/document.xml":
        return True

    if normalized.startswith("word/header") and normalized.endswith(".xml"):
        return True

    if normalized.startswith("word/footer") and normalized.endswith(".xml"):
        return True

    return normalized in {
        "word/footnotes.xml",
        "word/endnotes.xml",
        "word/comments.xml",
        "word/commentsextended.xml",
    }


def read_package(document_path: Path) -> Dict[str, bytes]:
    """Read all ZIP package parts into a dictionary."""
    with zipfile.ZipFile(document_path, "r") as package:
        return {
            package_part.filename: package.read(package_part.filename)
            for package_part in package.infolist()
        }


def summarize_tracked_changes(root: etree._Element) -> Dict[str, int]:
    """Count tracked-change-related XML elements."""
    counts = Counter()

    for element in root.iter():
        category = TRACKED_CHANGE_TAGS.get(element.tag)

        if category is not None:
            counts[category] += 1

    return dict(sorted(counts.items()))


def summarize_redaction_markers(root: etree._Element) -> Dict[str, int]:
    """
    Count redaction-related run formatting markers.

    This does not export any text or run content.
    """
    summary = {
        "runs": 0,
        "runs_with_character_style": 0,
        "runs_with_red_color": 0,
        "runs_with_strike": 0,
        "runs_with_yellow_highlight": 0,
    }

    for run in root.iter(W_R):
        summary["runs"] += 1

        run_properties = run.find(W_RPR)

        if run_properties is None:
            continue

        if run_properties.find(W_RSTYLE) is not None:
            summary["runs_with_character_style"] += 1

        color = run_properties.find(W_COLOR)

        if color is not None:
            color_value = (
                color.get(qn("val"))
                or color.get("val")
                or ""
            ).upper()

            if color_value in {"FF0000", "RED"}:
                summary["runs_with_red_color"] += 1

        if run_properties.find(W_STRIKE) is not None:
            summary["runs_with_strike"] += 1

        highlight = run_properties.find(W_HIGHLIGHT)

        if highlight is not None:
            highlight_value = (
                highlight.get(qn("val"))
                or highlight.get("val")
                or ""
            ).lower()

            if highlight_value == "yellow":
                summary["runs_with_yellow_highlight"] += 1

    return summary


def summarize_run_property_sequences(
    root: etree._Element,
) -> Dict[str, int]:
    """
    Count w:rPr child-element sequences.

    Only XML tag names are included. No document text, author data, comments,
    field values, or protocol content is exported.
    """
    sequences = Counter()

    for run_properties in root.iter(W_RPR):
        child_names = [
            local_name(child.tag)
            for child in run_properties
        ]

        sequence = " > ".join(child_names)

        if not sequence:
            sequence = "(empty)"

        sequences[sequence] += 1

    return dict(
        sequences.most_common(50)
    )


def summarize_xml_part(
    part_name: str,
    xml_bytes: bytes,
) -> Dict[str, Any]:
    """Summarize a single XML package part without exposing document text."""
    summary: Dict[str, Any] = {
        "part": part_name,
        "sha256": sha256_bytes(xml_bytes),
        "parsed": False,
    }

    try:
        root = parse_xml(xml_bytes)
        summary["parsed"] = True
        summary["root_element"] = local_name(root.tag)
        summary["element_count"] = sum(1 for _ in root.iter())

        if is_content_part(part_name):
            summary["tracked_changes"] = summarize_tracked_changes(root)
            summary["redaction_markers"] = summarize_redaction_markers(root)
            summary["run_property_sequences"] = (
                summarize_run_property_sequences(root)
            )

        if part_name == STYLE_PART_NAME:
            style_count = len(
                root.xpath(
                    "./w:style",
                    namespaces={"w": W_NS},
                )
            )

            matching_redaction_styles = root.xpath(
                "./w:style[@w:styleId='DCPleaseReviewDeletedText']",
                namespaces={"w": W_NS},
            )

            summary["style_count"] = style_count
            summary["please_review_style_found"] = bool(
                matching_redaction_styles
            )

        if part_name == MAIN_RELATIONSHIPS_PART_NAME:
            relationships = root.xpath(
                "./rel:Relationship",
                namespaces={"rel": REL_NS},
            )

            relationship_types = Counter(
                relationship.get("Type", "")
                for relationship in relationships
            )

            summary["relationship_type_counts"] = dict(
                sorted(relationship_types.items())
            )

    except Exception as exc:
        summary["parse_error"] = str(exc)

    return summary


def summarize_package(
    document_path: Path,
) -> Dict[str, Any]:
    """Create a content-free structural summary of a DOCX/DOCM package."""
    package_parts = read_package(document_path)

    summary: Dict[str, Any] = {
        "file_name": document_path.name,
        "file_size_bytes": document_path.stat().st_size,
        "part_count": len(package_parts),
        "parts": {},
    }

    for part_name in sorted(package_parts):
        if is_xml_part(part_name):
            summary["parts"][part_name] = summarize_xml_part(
                part_name,
                package_parts[part_name],
            )
        else:
            summary["parts"][part_name] = {
                "part": part_name,
                "sha256": sha256_bytes(package_parts[part_name]),
                "binary_part": True,
                "size_bytes": len(package_parts[part_name]),
            }

    return summary


def compare_part_summaries(
    generated_summary: Dict[str, Any],
    recovered_summary: Dict[str, Any],
) -> Dict[str, Any]:
    """Compare package summaries without exposing XML or protocol text."""
    generated_parts = generated_summary["parts"]
    recovered_parts = recovered_summary["parts"]

    generated_part_names = set(generated_parts)
    recovered_part_names = set(recovered_parts)

    only_in_generated = sorted(
        generated_part_names - recovered_part_names
    )

    only_in_recovered = sorted(
        recovered_part_names - generated_part_names
    )

    changed_parts: List[Dict[str, Any]] = []

    for part_name in sorted(
        generated_part_names & recovered_part_names
    ):
        generated_part = generated_parts[part_name]
        recovered_part = recovered_parts[part_name]

        if generated_part.get("sha256") == recovered_part.get("sha256"):
            continue

        part_change: Dict[str, Any] = {
            "part": part_name,
            "generated_sha256": generated_part.get("sha256"),
            "recovered_sha256": recovered_part.get("sha256"),
        }

        if (
            "tracked_changes" in generated_part
            or "tracked_changes" in recovered_part
        ):
            part_change["generated_tracked_changes"] = (
                generated_part.get("tracked_changes", {})
            )
            part_change["recovered_tracked_changes"] = (
                recovered_part.get("tracked_changes", {})
            )

        if (
            "redaction_markers" in generated_part
            or "redaction_markers" in recovered_part
        ):
            part_change["generated_redaction_markers"] = (
                generated_part.get("redaction_markers", {})
            )
            part_change["recovered_redaction_markers"] = (
                recovered_part.get("redaction_markers", {})
            )

        if (
            "run_property_sequences" in generated_part
            or "run_property_sequences" in recovered_part
        ):
            part_change["generated_run_property_sequences"] = (
                generated_part.get("run_property_sequences", {})
            )
            part_change["recovered_run_property_sequences"] = (
                recovered_part.get("run_property_sequences", {})
            )

        changed_parts.append(part_change)

    return {
        "parts_only_in_generated": only_in_generated,
        "parts_only_in_recovered": only_in_recovered,
        "changed_parts": changed_parts,
    }


def build_default_report_path(
    generated_path: Path,
) -> Path:
    """Create a default content-free comparison report path."""
    return generated_path.with_name(
        f"{generated_path.stem}_RecoveryComparison.json"
    )


def main() -> int:
    """Run comparison and write the content-free JSON report."""
    if len(sys.argv) not in {3, 4}:
        print(
            "Usage:\n"
            "  python tools/ComparePleaseReviewRecovery.py "
            "\"Generated_PleaseReview.docx\" "
            "\"Recovered_PleaseReview.docx\" "
            "[OptionalOutputReport.json]"
        )
        return 2

    generated_path = Path(sys.argv[1]).expanduser().resolve()
    recovered_path = Path(sys.argv[2]).expanduser().resolve()

    report_path = (
        Path(sys.argv[3]).expanduser().resolve()
        if len(sys.argv) == 4
        else build_default_report_path(generated_path)
    )

    if not generated_path.exists():
        print(f"Generated file not found: {generated_path}")
        return 2

    if not recovered_path.exists():
        print(f"Recovered file not found: {recovered_path}")
        return 2

    try:
        generated_summary = summarize_package(generated_path)
        recovered_summary = summarize_package(recovered_path)

        report = {
            "report_type": "PleaseReview OOXML Recovery Comparison",
            "generated_utc": datetime.now(timezone.utc).isoformat(),
            "generated_document": generated_summary,
            "recovered_document": recovered_summary,
            "comparison": compare_part_summaries(
                generated_summary,
                recovered_summary,
            ),
        }

        report_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        with report_path.open("w", encoding="utf-8") as report_file:
            json.dump(
                report,
                report_file,
                indent=2,
                ensure_ascii=False,
            )

        print(f"Comparison report written to: {report_path}")
        print(
            "The report contains no extracted document text. "
            "It contains structural XML/package metadata only."
        )

        return 0

    except zipfile.BadZipFile:
        print(
            "One of the selected files is not a valid DOCX/DOCM ZIP package."
        )
        return 1

    except Exception as exc:
        print(f"Comparison failed: {exc}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
