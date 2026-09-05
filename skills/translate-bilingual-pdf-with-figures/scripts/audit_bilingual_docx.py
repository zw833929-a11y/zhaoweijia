#!/usr/bin/env python3
"""Audit a bilingual DOCX against its canonical JSON model."""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import re
import unicodedata
import zipfile
from pathlib import Path
from typing import Any

from lxml import etree


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
PLACEHOLDER_RE = re.compile(
    r"\b(?:TODO|TBD|LOREM IPSUM)\b|\[PLACEHOLDER\]|待翻译|图片缺失|MISSING IMAGE",
    re.IGNORECASE,
)
SUPERSCRIPT_MAP = str.maketrans("⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻", "0123456789+-")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Audit a bilingual DOCX against its JSON model.")
    parser.add_argument("docx", type=Path, help="DOCX to inspect")
    parser.add_argument("--model", type=Path, required=True, help="Canonical bilingual JSON model")
    return parser.parse_args()


def bookmark_name(item_id: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_]", "_", item_id)[:24]
    digest = hashlib.sha1(item_id.encode("utf-8")).hexdigest()[:8]
    return f"item_{cleaned}_{digest}"


def normalize_text(value: str) -> str:
    value = unicodedata.normalize("NFKC", value).translate(SUPERSCRIPT_MAP)
    value = value.replace("−", "-").replace("–", "-").replace("—", "-")
    value = value.replace("ﬁ", "fi").replace("ﬂ", "fl")
    value = re.sub(r"\^\{([^}]+)\}", r"\1", value)
    value = re.sub(r"\^([+\-]?\d+|[A-Za-z]+)", r"\1", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def xml_root(archive: zipfile.ZipFile, name: str) -> etree._Element:
    return etree.fromstring(archive.read(name))


def paragraph_text(paragraph: etree._Element) -> str:
    return "".join(paragraph.xpath(".//w:t/text()", namespaces=NS))


def styles_map(archive: zipfile.ZipFile) -> dict[str, str]:
    root = xml_root(archive, "word/styles.xml")
    result: dict[str, str] = {}
    for style in root.xpath("//w:style", namespaces=NS):
        style_id = style.get(f"{{{NS['w']}}}styleId", "")
        names = style.xpath("./w:name/@w:val", namespaces=NS)
        result[style_id] = names[0] if names else style_id
    return result


def reference_numbers(references: list[Any]) -> list[int] | None:
    numbers: list[int] = []
    for reference in references:
        if isinstance(reference, dict):
            value = str(reference.get("text", ""))
        else:
            value = str(reference)
        match = re.match(r"^\s*(?:\[(\d+)\]|(\d+)[.)])", value)
        if not match:
            return None
        numbers.append(int(match.group(1) or match.group(2)))
    return numbers


def add_check(checks: list[dict[str, Any]], name: str, passed: bool, details: Any) -> None:
    checks.append({"name": name, "passed": bool(passed), "details": details})


def main() -> int:
    args = parse_args()
    docx_path = args.docx.expanduser().resolve()
    model_path = args.model.expanduser().resolve()
    if not docx_path.is_file():
        raise SystemExit(f"DOCX not found: {docx_path}")
    if not model_path.is_file():
        raise SystemExit(f"Model not found: {model_path}")

    model = json.loads(model_path.read_text(encoding="utf-8"))
    items = model.get("items", [])
    references = model.get("references", [])
    if not isinstance(items, list) or not isinstance(references, list):
        raise SystemExit("Model must contain 'items' and 'references' arrays")

    with zipfile.ZipFile(docx_path) as archive:
        names = archive.namelist()
        document = xml_root(archive, "word/document.xml")
        style_names = styles_map(archive)

        paragraph_records: list[tuple[str, str]] = []
        for paragraph in document.xpath("//w:body/w:p", namespaces=NS):
            style_ids = paragraph.xpath("./w:pPr/w:pStyle/@w:val", namespaces=NS)
            style_id = style_ids[0] if style_ids else "Normal"
            paragraph_records.append((style_names.get(style_id, style_id), paragraph_text(paragraph)))

        style_counts = collections.Counter(style for style, _ in paragraph_records)
        body_text = normalize_text("\n".join(text_value for _, text_value in paragraph_records))
        bookmark_values = document.xpath("//w:bookmarkStart/@w:name", namespaces=NS)
        bookmark_counts = collections.Counter(bookmark_values)
        inline_count = len(document.xpath("//wp:inline", namespaces=NS))
        anchor_count = len(document.xpath("//wp:anchor", namespaces=NS))
        blip_count = len(document.xpath("//a:blip", namespaces=NS))
        media_files = [name for name in names if name.startswith("word/media/") and not name.endswith("/")]

        field_texts: list[str] = []
        all_visible_text: list[str] = [body_text]
        for name in names:
            if not (name.startswith("word/") and name.endswith(".xml")):
                continue
            try:
                root = xml_root(archive, name)
            except etree.XMLSyntaxError:
                continue
            field_texts.extend(root.xpath("//w:instrText/text()", namespaces=NS))
            all_visible_text.extend(root.xpath("//w:t/text()", namespaces=NS))

    checks: list[dict[str, Any]] = []
    expected_bookmarks = [bookmark_name(str(item.get("id", ""))) for item in items]
    missing_bookmarks = [name for name in expected_bookmarks if bookmark_counts[name] == 0]
    duplicate_bookmarks = [name for name in expected_bookmarks if bookmark_counts[name] != 1]
    add_check(
        checks,
        "modeled item bookmarks",
        not missing_bookmarks and not duplicate_bookmarks,
        {"expected": len(expected_bookmarks), "missing": missing_bookmarks, "not_exactly_once": duplicate_bookmarks},
    )

    empty_translation_ids: list[str] = []
    text_coverage_missing: list[dict[str, str]] = []
    for item in items:
        item_id = str(item.get("id", ""))
        if item.get("type") == "figure":
            fields = ("caption_en", "caption_zh")
        else:
            fields = ("en", "zh")
        for field in fields:
            value = str(item.get(field, "")).strip()
            if field.endswith("zh") and not value:
                empty_translation_ids.append(item_id)
            normalized = normalize_text(value)
            if normalized and normalized not in body_text:
                text_coverage_missing.append({"id": item_id, "field": field})
    add_check(checks, "nonempty translations", not empty_translation_ids, empty_translation_ids)
    add_check(checks, "modeled text coverage", not text_coverage_missing, text_coverage_missing)

    source_pairs = style_counts["Source Paragraph"] + sum(
        style_counts[f"Source Heading {level}"] for level in range(1, 4)
    )
    translation_pairs = style_counts["Translation Paragraph"] + sum(
        style_counts[f"Translation Heading {level}"] for level in range(1, 4)
    )
    add_check(
        checks,
        "source/translation paragraph pairing",
        source_pairs == translation_pairs,
        {"source": source_pairs, "translation": translation_pairs},
    )

    expected_figures = sum(1 for item in items if item.get("type") == "figure")
    add_check(
        checks,
        "figure count",
        inline_count == expected_figures and blip_count == expected_figures,
        {
            "expected": expected_figures,
            "inline": inline_count,
            "drawing_references": blip_count,
            "media_files": len(media_files),
        },
    )
    add_check(checks, "no floating images", anchor_count == 0, {"anchors": anchor_count})
    add_check(
        checks,
        "bilingual figure legends",
        style_counts["Source Caption"] == expected_figures
        and style_counts["Translation Caption"] == expected_figures,
        {
            "expected": expected_figures,
            "source_captions": style_counts["Source Caption"],
            "translation_captions": style_counts["Translation Caption"],
        },
    )

    add_check(
        checks,
        "reference count",
        style_counts["Reference"] == len(references),
        {"expected": len(references), "actual": style_counts["Reference"]},
    )
    numbers = reference_numbers(references)
    add_check(
        checks,
        "reference sequence",
        numbers is None or numbers == list(range(1, len(numbers) + 1)),
        {"numbered": numbers is not None, "numbers": numbers},
    )

    visible_text = "\n".join(all_visible_text)
    placeholders = sorted(set(match.group(0) for match in PLACEHOLDER_RE.finditer(visible_text)))
    add_check(checks, "no placeholders", not placeholders, placeholders)

    field_codes = " ".join(field_texts).upper()
    add_check(
        checks,
        "page fields",
        "PAGE" in field_codes and "NUMPAGES" in field_codes,
        {"PAGE": "PAGE" in field_codes, "NUMPAGES": "NUMPAGES" in field_codes},
    )

    passed = all(check["passed"] for check in checks)
    report = {
        "docx": str(docx_path),
        "model": str(model_path),
        "passed": passed,
        "summary": {
            "items": len(items),
            "figures": expected_figures,
            "references": len(references),
            "inline_images": inline_count,
            "floating_images": anchor_count,
        },
        "checks": checks,
    }
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
