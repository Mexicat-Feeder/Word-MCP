"""V2 live Word tools: grouped, agent-friendly facade over COM helpers.

The v2 surface is inspired by SuperDoc's MCP shape: a few lifecycle tools plus
grouped intent tools. It keeps the existing live COM implementations underneath
so v2 can improve tool ergonomics without rewriting Word automation logic.
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
import uuid
from pathlib import Path
from typing import Any

from word_document_server.defaults import DEFAULT_AUTHOR
from word_document_server.tools import live_read_tools, live_tools


_sessions: dict[str, dict[str, Any]] = {}
_handles: dict[str, dict[str, dict[str, Any]]] = {}

DEFAULT_TEMPLATE_ID = "default_plain"
SUPPORTED_BLUEPRINT_BLOCKS = {
    "title_page",
    "toc",
    "paragraph",
    "heading",
    "list",
    "table",
    "image",
    "image_placeholder",
    "page_break",
    "section_break",
}

WD_ACTIVE_END_PAGE_NUMBER = 3
EMU_PER_POINT = 12700
IMAGE_EXTENSIONS = {".bmp", ".dib", ".emf", ".gif", ".jpg", ".jpeg", ".png", ".tif", ".tiff", ".wmf"}

PAGE_SIZES = {
    "letter": (612, 792),
    "a4": (595, 842),
}

DEFAULT_PAGE_SETUP = {
    "size": "letter",
    "orientation": "portrait",
    "margins": {
        "top": 72,
        "bottom": 72,
        "left": 72,
        "right": 72,
    },
}

DEFAULT_STYLE_PROFILE = {
    "Normal": {
        "font_name": "Aptos",
        "font_size": 11,
        "space_after": 6,
        "line_spacing_rule": 0,
    },
    "Heading 1": {
        "font_name": "Aptos Display",
        "font_size": 18,
        "bold": True,
        "space_before": 18,
        "space_after": 6,
        "keep_with_next": True,
    },
    "Heading 2": {
        "font_name": "Aptos Display",
        "font_size": 14,
        "bold": True,
        "space_before": 12,
        "space_after": 4,
        "keep_with_next": True,
    },
    "Heading 3": {
        "font_name": "Aptos Display",
        "font_size": 12,
        "bold": True,
        "space_before": 8,
        "space_after": 4,
        "keep_with_next": True,
    },
    "List Paragraph": {
        "font_name": "Aptos",
        "font_size": 11,
        "left_indent": 36,
        "space_after": 4,
    },
    "Caption": {
        "font_name": "Aptos",
        "font_size": 9,
        "italic": True,
        "space_after": 8,
    },
    "Timer": {
        "font_name": "Aptos",
        "font_size": 11,
        "bold": True,
        "font_color": "#C00000",
        "space_before": 6,
        "space_after": 6,
    },
    "Key Finding": {
        "font_name": "Aptos",
        "font_size": 11,
        "bold": True,
        "font_color": "#1F4E79",
        "space_before": 8,
        "space_after": 8,
    },
}


def _load_result(raw: str | dict) -> dict[str, Any]:
    if isinstance(raw, dict):
        return raw
    try:
        return json.loads(raw)
    except Exception:
        return {"success": False, "raw": raw}


def _dump(result: dict[str, Any]) -> str:
    return json.dumps(result, ensure_ascii=False)


def _color_to_word(hex_color: str) -> int | None:
    if not hex_color:
        return None
    c = str(hex_color).lstrip("#")
    if len(c) != 6:
        return None
    try:
        r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    except ValueError:
        return None
    return r + (g << 8) + (b << 16)


def _natural_sort_key(value: str | os.PathLike) -> list[Any]:
    text = str(value)
    return [int(part) if part.isdigit() else part.lower() for part in re.split(r"(\d+)", text)]


def _asset_paths(asset_dir: str | None) -> list[str]:
    if not asset_dir:
        return []
    root = Path(asset_dir)
    if not root.exists():
        return []
    files = [
        path
        for path in root.rglob("*")
        if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS
    ]
    return [str(path) for path in sorted(files, key=lambda path: _natural_sort_key(path.name))]


def _asset_sequence_key(value: str | os.PathLike) -> list[Any]:
    name = Path(value).name
    numbers = [int(part) for part in re.findall(r"\d+", name)]
    return [numbers[0] if numbers else 999999, *_natural_sort_key(name)]


def _split_asset_paths(asset_paths: list[str]) -> tuple[list[str], list[str]]:
    title_assets = [
        path for path in asset_paths
        if "title" in Path(path).stem.lower()
    ]
    title_set = set(title_assets)
    body_assets = [path for path in asset_paths if path not in title_set]
    return sorted(title_assets, key=_asset_sequence_key), body_assets


def _shape_points(value: Any) -> float | None:
    try:
        return round(float(value), 1)
    except Exception:
        return None


def _shape_int(value: Any) -> int | None:
    try:
        return int(value)
    except Exception:
        return None


def _safe_attr(obj: Any, name: str, default: Any = None) -> Any:
    try:
        return getattr(obj, name)
    except Exception:
        return default


def _word_alignment_name(value: int | None) -> str | None:
    return {0: "left", 1: "center", 2: "right", 3: "justify"}.get(value)


def _word_wrap_name(value: int | None) -> str:
    return {
        0: "square",
        1: "tight",
        2: "topbottom",
        3: "behind",
        4: "infront",
    }.get(value, "inline")


def _blueprint_list_format(numbering: dict | None) -> dict[str, Any] | None:
    if not isinstance(numbering, dict) or numbering.get("kind") != "word_list":
        return None
    marker = str(numbering.get("list_string") or "").strip()
    level = numbering.get("level")
    try:
        zero_based_level = max(0, int(level) - 1)
    except Exception:
        zero_based_level = 0
    list_type = "bullet"
    if any(char.isdigit() for char in marker) or marker[:1].isalpha():
        list_type = "number"
    try:
        list_value = int(numbering.get("list_value") or 1)
    except Exception:
        list_value = 1
    return {
        "list_type": list_type,
        "level": zero_based_level,
        "continue_previous": list_value > 1,
    }


def _range_page(rng) -> int | None:
    try:
        return int(rng.Information(WD_ACTIVE_END_PAGE_NUMBER))
    except Exception:
        return None


def _range_start(rng) -> int:
    try:
        return int(rng.Start)
    except Exception:
        return 0


def _safe_style_name(style: Any) -> str:
    try:
        return str(style) if style else "Normal"
    except Exception:
        return "Normal"


def _page_role(page: int, first_body_page: int | None, toc_pages: set[int]) -> str:
    if first_body_page and page < first_body_page:
        if page == 1:
            return "title_page"
        if page in toc_pages or page == first_body_page - 1:
            return "toc"
        return "front_matter"
    if page in toc_pages:
        return "toc"
    return "body"


def _assign_page_roles(pages: list[dict[str, Any]], first_body_page: int | None, toc_pages: set[int]) -> None:
    for page in pages:
        page["role"] = _page_role(page["page"], first_body_page, toc_pages)


def _infer_numbering(text: str, style: str, para=None) -> dict[str, Any]:
    clean = (text or "").strip()
    if para is not None:
        try:
            list_format = para.Range.ListFormat
            list_type = int(list_format.ListType)
            if list_type:
                return {
                    "kind": "word_list",
                    "list_type": list_type,
                    "level": int(list_format.ListLevelNumber),
                    "list_string": str(list_format.ListString),
                    "list_value": int(list_format.ListValue),
                }
        except Exception:
            pass

    literal = re.match(r"^(\d+(?:\.\d+)*\.?)\s+\S", clean)
    if literal and style.lower().startswith("heading"):
        return {"kind": "literal", "value": literal.group(1).rstrip(".")}
    if literal and style.lower() in {"normal", "normalordirect", "list paragraph", "listparagraph"}:
        return {"kind": "literal", "value": literal.group(1).rstrip(".")}
    return {"kind": "none"}


def _looks_like_toc(style: str, text: str) -> bool:
    lowered_style = (style or "").replace(" ", "").lower()
    lowered_text = (text or "").strip().lower()
    return lowered_style.startswith("toc") or lowered_text == "table of contents"


def _normalize_margins(margins: dict | None) -> dict[str, float]:
    if margins is not None and not isinstance(margins, dict):
        raise ValueError("margins must be an object with top, bottom, left, and right values")
    merged = dict(DEFAULT_PAGE_SETUP["margins"])
    for key, value in (margins or {}).items():
        if key in merged and value is not None:
            merged[key] = float(value)
    return merged


def _normalize_page_setup(page_setup: dict | None) -> dict[str, Any]:
    incoming = page_setup or {}
    normalized = {
        "size": (incoming.get("size") or DEFAULT_PAGE_SETUP["size"]).lower(),
        "orientation": (incoming.get("orientation") or DEFAULT_PAGE_SETUP["orientation"]).lower(),
        "margins": _normalize_margins(incoming.get("margins")),
    }
    if normalized["size"] == "custom":
        normalized["width"] = float(incoming.get("width") or PAGE_SIZES[DEFAULT_PAGE_SETUP["size"]][0])
        normalized["height"] = float(incoming.get("height") or PAGE_SIZES[DEFAULT_PAGE_SETUP["size"]][1])
    elif normalized["size"] not in PAGE_SIZES:
        raise ValueError(f"Unsupported page size: {normalized['size']}. Use one of {sorted([*PAGE_SIZES, 'custom'])}")
    if normalized["orientation"] not in {"portrait", "landscape"}:
        raise ValueError("orientation must be 'portrait' or 'landscape'")
    return normalized


def _normalize_blueprint(blueprint: dict | None) -> dict[str, Any]:
    if blueprint is not None and not isinstance(blueprint, dict):
        return {
            "template": DEFAULT_TEMPLATE_ID,
            "page_setup": {},
            "properties": {},
            "blocks": blueprint,
        }
    source = blueprint or {}
    document = source.get("document") if isinstance(source.get("document"), dict) else {}
    sections = source.get("sections")
    if sections is None:
        sections = document.get("sections")
    pages = source.get("pages")
    if pages is None:
        pages = document.get("pages")

    raw_blocks = source.get("blocks")
    if raw_blocks is None:
        raw_blocks = document.get("blocks")
    has_explicit_blocks = raw_blocks is not None
    if raw_blocks is None:
        raw_blocks = []
    blocks = list(raw_blocks) if isinstance(raw_blocks, list) else raw_blocks
    if isinstance(blocks, list) and not has_explicit_blocks and sections:
        for section in sections:
            if isinstance(section, dict):
                blocks.extend(section.get("blocks") or [])
    if isinstance(blocks, list) and not has_explicit_blocks and pages:
        for page in pages:
            if isinstance(page, dict):
                blocks.extend(page.get("blocks") or [])

    properties = {}
    for container in (document, source):
        for key in ("title", "subject", "author", "keywords", "comments", "category", "manager", "company", "last_author"):
            if container.get(key) is not None:
                properties[key] = container.get(key)
        if isinstance(container.get("properties"), dict):
            properties.update({k: v for k, v in container["properties"].items() if v is not None})

    return {
        "template": source.get("template") or document.get("template") or DEFAULT_TEMPLATE_ID,
        "page_setup": source.get("page_setup") or document.get("page_setup") or {},
        "properties": properties,
        "blocks": blocks,
    }


def _validate_blueprint(blueprint: dict | None) -> list[str]:
    normalized = _normalize_blueprint(blueprint)
    errors: list[str] = []
    try:
        _normalize_page_setup(normalized["page_setup"])
    except ValueError as exc:
        errors.append(str(exc))

    blocks = normalized["blocks"]
    if not isinstance(blocks, list):
        return ["blueprint blocks must be a list"]

    for index, block in enumerate(blocks):
        if not isinstance(block, dict):
            errors.append(f"block {index} must be an object")
            continue
        block_type = block.get("type")
        if block_type not in SUPPORTED_BLUEPRINT_BLOCKS:
            errors.append(
                f"block {index} has unsupported type {block_type!r}; "
                f"use one of {sorted(SUPPORTED_BLUEPRINT_BLOCKS)}"
            )
            continue
        if block_type in {"paragraph", "heading"} and block.get("text") is None:
            errors.append(f"block {index} ({block_type}) requires text")
        if block_type == "heading":
            level = int(block.get("level") or 1)
            if level < 1 or level > 3:
                errors.append(f"block {index} heading level must be 1, 2, or 3")
        if block_type == "title_page":
            title = block.get("title") or block.get("text")
            images = block.get("images")
            if title is None and images is not None and not isinstance(images, list):
                errors.append(f"block {index} title_page images must be an array")
        if block_type == "toc":
            levels = int(block.get("levels") or 3)
            if levels < 1 or levels > 9:
                errors.append(f"block {index} toc levels must be between 1 and 9")
        if block_type == "list":
            items = block.get("items")
            if not isinstance(items, list) or not items:
                errors.append(f"block {index} list requires a non-empty items array")
        if block_type == "table":
            rows = block.get("rows") or block.get("data")
            if not isinstance(rows, list) or not rows:
                errors.append(f"block {index} table requires non-empty rows")
            elif not all(isinstance(row, list) for row in rows):
                errors.append(f"block {index} table rows must be arrays")
            else:
                width = max(len(row) for row in rows)
                if width == 0 or any(len(row) != width for row in rows):
                    errors.append(f"block {index} table rows must all have the same non-zero width")
        if block_type == "image":
            if not block.get("path"):
                errors.append(f"block {index} image requires path")
    return errors


def _ensure_word_style(doc, name: str, profile: dict[str, Any]) -> str | None:
    try:
        style = doc.Styles(name)
    except Exception:
        try:
            style = doc.Styles.Add(Name=name, Type=1)  # wdStyleTypeParagraph
        except Exception as exc:
            return str(exc)

    try:
        font = style.Font
        if profile.get("font_name"):
            font.Name = profile["font_name"]
        if profile.get("font_size") is not None:
            font.Size = profile["font_size"]
        if profile.get("bold") is not None:
            font.Bold = bool(profile["bold"])
        if profile.get("italic") is not None:
            font.Italic = bool(profile["italic"])
        color = _color_to_word(profile.get("font_color"))
        if color is not None:
            font.Color = color

        paragraph = style.ParagraphFormat
        if profile.get("space_before") is not None:
            paragraph.SpaceBefore = profile["space_before"]
        if profile.get("space_after") is not None:
            paragraph.SpaceAfter = profile["space_after"]
        if profile.get("line_spacing_rule") is not None:
            paragraph.LineSpacingRule = profile["line_spacing_rule"]
        if profile.get("left_indent") is not None:
            paragraph.LeftIndent = profile["left_indent"]
        if profile.get("keep_with_next") is not None:
            paragraph.KeepWithNext = bool(profile["keep_with_next"])
    except Exception as exc:
        return str(exc)
    return None


def _apply_word_page_setup(doc, page_setup: dict | None) -> dict[str, Any]:
    normalized = _normalize_page_setup(page_setup)
    if normalized["size"] == "custom":
        width, height = normalized["width"], normalized["height"]
    else:
        width, height = PAGE_SIZES[normalized["size"]]
    if normalized["orientation"] == "landscape":
        width, height = height, width

    setup = doc.PageSetup
    setup.Orientation = 1 if normalized["orientation"] == "landscape" else 0
    setup.PageWidth = width
    setup.PageHeight = height
    setup.TopMargin = normalized["margins"]["top"]
    setup.BottomMargin = normalized["margins"]["bottom"]
    setup.LeftMargin = normalized["margins"]["left"]
    setup.RightMargin = normalized["margins"]["right"]
    return normalized


def _apply_default_template_live(filename: str) -> dict[str, Any]:
    if sys.platform != "win32":
        return {"error": "Default template initialization is only available on Windows"}

    from word_document_server.core.word_com import find_document, get_word_app, undo_record

    app = get_word_app()
    doc = find_document(app, filename)
    style_errors: dict[str, str] = {}

    with undo_record(app, "MCP: Apply Default Template"):
        page_setup = _apply_word_page_setup(doc, DEFAULT_PAGE_SETUP)
        for name, profile in DEFAULT_STYLE_PROFILE.items():
            error = _ensure_word_style(doc, name, profile)
            if error:
                style_errors[name] = error

        try:
            props = doc.BuiltInDocumentProperties
            if DEFAULT_AUTHOR:
                props("Author").Value = DEFAULT_AUTHOR
        except Exception as exc:
            style_errors["properties"] = str(exc)

    return {
        "success": True,
        "template": DEFAULT_TEMPLATE_ID,
        "page_setup": page_setup,
        "styles": list(DEFAULT_STYLE_PROFILE.keys()),
        "style_errors": style_errors or None,
    }


def _resolve_com_range(doc, position: str = "end", paragraph_index: int = None):
    if paragraph_index is not None:
        if paragraph_index < 1 or paragraph_index > doc.Paragraphs.Count:
            raise ValueError(f"paragraph_index {paragraph_index} out of range (1-{doc.Paragraphs.Count})")
        start = doc.Paragraphs(paragraph_index).Range.Start
        return doc.Range(start, start)
    if position == "start":
        return doc.Range(0, 0)
    if position == "end":
        end_pos = doc.Content.End - 1
        return doc.Range(end_pos, end_pos)
    try:
        offset = int(position)
    except Exception as exc:
        raise ValueError("position must be 'start', 'end', or a character offset") from exc
    return doc.Range(offset, offset)


def _layout_page_setup_live(filename: str, page_setup: dict | None) -> dict[str, Any]:
    if sys.platform != "win32":
        return {"error": "Layout page setup is only available on Windows"}
    from word_document_server.core.word_com import find_document, get_word_app, undo_record

    app = get_word_app()
    doc = find_document(app, filename)
    with undo_record(app, "MCP: Page Setup"):
        normalized = _apply_word_page_setup(doc, page_setup)
    return {"success": True, "document": doc.Name, "page_setup": normalized}


def _layout_insert_break_live(
    filename: str,
    break_kind: str,
    position: str = "end",
    paragraph_index: int = None,
    break_type: str = "next_page",
) -> dict[str, Any]:
    if sys.platform != "win32":
        return {"error": "Break insertion is only available on Windows"}
    from word_document_server.core.word_com import find_document, get_word_app, undo_record

    app = get_word_app()
    doc = find_document(app, filename)
    section_breaks = {
        "next_page": 2,
        "continuous": 3,
        "even_page": 4,
        "odd_page": 5,
    }
    if break_kind == "page":
        word_break = 7
    else:
        word_break = section_breaks.get((break_type or "next_page").lower())
        if word_break is None:
            return {"error": f"Invalid section break_type: {break_type}", "valid_break_types": sorted(section_breaks)}

    with undo_record(app, f"MCP: Insert {break_kind.title()} Break"):
        rng = _resolve_com_range(doc, position=position, paragraph_index=paragraph_index)
        rng.InsertBreak(Type=word_break)

    return {
        "success": True,
        "document": doc.Name,
        "break": break_kind,
        "break_type": break_type if break_kind == "section" else "page",
        "position": position,
        "paragraph_index": paragraph_index,
    }


def _insert_toc_live(
    filename: str,
    title: str = "Table of Contents",
    levels: int = 3,
    page_break_after: bool = True,
) -> dict[str, Any]:
    if sys.platform != "win32":
        return {"error": "TOC insertion is only available on Windows"}
    from word_document_server.core.word_com import find_document, get_word_app, undo_record

    app = get_word_app()
    doc = find_document(app, filename)
    with undo_record(app, "MCP: Insert TOC"):
        rng = doc.Range()
        rng.Collapse(0)
        if title:
            rng.InsertAfter(f"{title}\r")
            title_start = max(0, rng.End - len(title) - 1)
            try:
                doc.Range(title_start, title_start + len(title)).Style = "TOC Heading"
            except Exception:
                pass
            rng = doc.Range(rng.End, rng.End)
        toc = doc.TablesOfContents.Add(
            Range=rng,
            UseHeadingStyles=True,
            UpperHeadingLevel=1,
            LowerHeadingLevel=int(levels),
            IncludePageNumbers=True,
            RightAlignPageNumbers=True,
            UseHyperlinks=True,
        )
        try:
            toc.Update()
        except Exception:
            pass
        if page_break_after:
            after = doc.Range(toc.Range.End, toc.Range.End)
            after.InsertBreak(Type=7)

    return {
        "success": True,
        "document": doc.Name,
        "title": title,
        "levels": int(levels),
        "page_break_after": page_break_after,
    }


def _update_tocs_live(filename: str) -> dict[str, Any]:
    if sys.platform != "win32":
        return {"error": "TOC update is only available on Windows"}
    from word_document_server.core.word_com import find_document, get_word_app

    app = get_word_app()
    doc = find_document(app, filename)
    updated = 0
    for index in range(1, doc.TablesOfContents.Count + 1):
        try:
            doc.TablesOfContents(index).Update()
            updated += 1
        except Exception:
            continue
    try:
        doc.Fields.Update()
    except Exception:
        pass
    return {"success": True, "document": doc.Name, "updated_tocs": updated}


def _paragraph_index_for_range(rng, paragraph_ranges: list[tuple[int, int, int]]) -> int | None:
    start = _range_start(rng)
    for index, para_start, para_end in paragraph_ranges:
        if para_start <= start <= para_end:
            return index
    return None


def _inspect_tables_live(doc, paragraph_ranges: list[tuple[int, int, int]]) -> list[dict[str, Any]]:
    tables: list[dict[str, Any]] = []
    for table_index in range(1, doc.Tables.Count + 1):
        try:
            table = doc.Tables(table_index)
            rng = table.Range
            row_count = int(table.Rows.Count)
            col_count = int(table.Columns.Count)
            rows = []
            for row in range(1, row_count + 1):
                values = []
                for col in range(1, col_count + 1):
                    try:
                        values.append(str(table.Cell(row, col).Range.Text).rstrip("\r\x07"))
                    except Exception:
                        values.append("")
                rows.append(values)
            tables.append({
                "type": "table",
                "table_index": table_index,
                "page": _range_page(rng),
                "paragraph_index": _paragraph_index_for_range(rng, paragraph_ranges),
                "range_start": _range_start(rng),
                "rows": rows,
                "row_count": row_count,
                "col_count": col_count,
                "style": _safe_style_name(getattr(table, "Style", None)),
            })
        except Exception as exc:
            tables.append({"type": "table", "table_index": table_index, "error": str(exc)})
    return tables


def _inspect_inline_images_live(doc, paragraph_ranges: list[tuple[int, int, int]]) -> list[dict[str, Any]]:
    images: list[dict[str, Any]] = []
    for image_index in range(1, doc.InlineShapes.Count + 1):
        try:
            shape = doc.InlineShapes(image_index)
            rng = shape.Range
            try:
                alignment = int(rng.ParagraphFormat.Alignment)
            except Exception:
                alignment = None
            title = _safe_attr(shape, "Title", "") or f"InlineShape {image_index}"
            images.append({
                "type": "image",
                "image_index": image_index,
                "source": "inline",
                "name": title,
                "page": _range_page(rng),
                "paragraph_index": _paragraph_index_for_range(rng, paragraph_ranges),
                "range_start": _range_start(rng),
                "width_pt": _shape_points(_safe_attr(shape, "Width")),
                "height_pt": _shape_points(_safe_attr(shape, "Height")),
                "wrapping": "inline",
                "alignment": _word_alignment_name(alignment),
                "paragraph_after": True,
            })
        except Exception as exc:
            images.append({"type": "image", "image_index": image_index, "source": "inline", "error": str(exc)})
    return images


def _inspect_shape_text(shape) -> dict[str, Any]:
    result: dict[str, Any] = {}
    text_frame = _safe_attr(shape, "TextFrame")
    if text_frame is None:
        return result
    try:
        has_text = bool(int(_safe_attr(text_frame, "HasText", 0)))
    except Exception:
        has_text = False
    result["has_text"] = has_text
    if not has_text:
        return result
    text_range = _safe_attr(text_frame, "TextRange")
    if text_range is None:
        return result
    text = str(_safe_attr(text_range, "Text", "") or "").rstrip("\r\x07")
    result["text"] = text
    font = _safe_attr(text_range, "Font")
    if font is not None:
        result["font_name"] = str(_safe_attr(font, "Name", "") or "") or None
        result["font_size"] = _shape_points(_safe_attr(font, "Size"))
        result["bold"] = _shape_int(_safe_attr(font, "Bold"))
        result["font_color"] = _shape_int(_safe_attr(font, "Color"))
    return result


def _inspect_shape_visuals(shape) -> dict[str, Any]:
    result: dict[str, Any] = {}
    fill = _safe_attr(shape, "Fill")
    if fill is not None:
        result["fill_visible"] = _shape_int(_safe_attr(fill, "Visible"))
        fore_color = _safe_attr(fill, "ForeColor")
        result["fill_color"] = _shape_int(_safe_attr(fore_color, "RGB"))
    line = _safe_attr(shape, "Line")
    if line is not None:
        result["line_visible"] = _shape_int(_safe_attr(line, "Visible"))
        fore_color = _safe_attr(line, "ForeColor")
        result["line_color"] = _shape_int(_safe_attr(fore_color, "RGB"))
        result["line_weight"] = _shape_points(_safe_attr(line, "Weight"))
    return result


def _inspect_shapes_live(doc, paragraph_ranges: list[tuple[int, int, int]]) -> list[dict[str, Any]]:
    shapes: list[dict[str, Any]] = []
    for shape_index in range(1, doc.Shapes.Count + 1):
        try:
            shape = doc.Shapes(shape_index)
            anchor = _safe_attr(shape, "Anchor")
            wrap_format = _safe_attr(shape, "WrapFormat")
            name = str(_safe_attr(shape, "Name", "") or f"Shape {shape_index}")
            record = {
                "type": "shape",
                "shape_index": shape_index,
                "name": name,
                "page": _range_page(anchor) if anchor is not None else None,
                "paragraph_index": _paragraph_index_for_range(anchor, paragraph_ranges) if anchor is not None else None,
                "range_start": _range_start(anchor) if anchor is not None else None,
                "width_pt": _shape_points(_safe_attr(shape, "Width")),
                "height_pt": _shape_points(_safe_attr(shape, "Height")),
                "left_pt": _shape_points(_safe_attr(shape, "Left")),
                "top_pt": _shape_points(_safe_attr(shape, "Top")),
                "shape_type": _shape_int(_safe_attr(shape, "Type")),
                "auto_shape_type": _shape_int(_safe_attr(shape, "AutoShapeType")),
                "wrap_type": _shape_int(_safe_attr(wrap_format, "Type")),
                "relative_horizontal_position": _shape_int(_safe_attr(shape, "RelativeHorizontalPosition")),
                "relative_vertical_position": _shape_int(_safe_attr(shape, "RelativeVerticalPosition")),
            }
            record.update(_inspect_shape_text(shape))
            record.update(_inspect_shape_visuals(shape))
            if anchor is None:
                record["anchor_missing"] = True
            shapes.append(record)
        except Exception as exc:
            shapes.append({"type": "shape", "shape_index": shape_index, "error": str(exc)})
    return shapes


def _picture_shapes_as_images(shapes: list[dict[str, Any]]) -> list[dict[str, Any]]:
    images: list[dict[str, Any]] = []
    for shape in shapes:
        name = str(shape.get("name") or "")
        is_picture = shape.get("shape_type") == 13 or name.lower().startswith("picture")
        if shape.get("error") or not is_picture:
            continue
        wrap_type = shape.get("wrap_type")
        images.append({
            "type": "image",
            "image_index": shape.get("shape_index"),
            "source": "floating_shape",
            "name": shape.get("name"),
            "page": shape.get("page"),
            "paragraph_index": shape.get("paragraph_index"),
            "range_start": shape.get("range_start") or 0,
            "width_pt": shape.get("width_pt"),
            "height_pt": shape.get("height_pt"),
            "wrapping": _word_wrap_name(wrap_type),
            "wrap_type": wrap_type,
            "left_pt": shape.get("left_pt"),
            "top_pt": shape.get("top_pt"),
            "relative_horizontal_position": shape.get("relative_horizontal_position"),
            "relative_vertical_position": shape.get("relative_vertical_position"),
            "paragraph_after": True,
        })
    return images


def _is_replayable_title_shape(shape: dict[str, Any]) -> bool:
    name = str(shape.get("name") or "").lower()
    return bool(shape.get("text")) or name.startswith("text box") or name.startswith("rectangle") or shape.get("auto_shape_type") == 1


def _is_picture_shape(shape: dict[str, Any]) -> bool:
    name = str(shape.get("name") or "").lower()
    return shape.get("shape_type") == 13 or name.startswith("picture")


def _map_assets_to_body_images(
    images: list[dict[str, Any]],
    asset_paths: list[str],
    first_body_page: int | None,
    toc_pages: set[int],
) -> list[str]:
    warnings: list[str] = []
    body_images = [
        image for image in images
        if not image.get("error")
        and _page_role(int(image.get("page") or 0), first_body_page, toc_pages) == "body"
    ]
    body_images.sort(key=lambda item: (item.get("page") or 0, item.get("range_start") or 0, item.get("image_index") or 0))
    for image, asset_path in zip(body_images, asset_paths):
        image["asset_path"] = asset_path
        image["path"] = asset_path
    if asset_paths and len(asset_paths) != len(body_images):
        warnings.append(
            f"asset_dir contains {len(asset_paths)} images but reference inspection found {len(body_images)} body images"
        )
    return warnings


def _map_assets_to_title_images(
    images: list[dict[str, Any]],
    asset_paths: list[str],
    first_body_page: int | None,
    toc_pages: set[int],
) -> list[str]:
    warnings: list[str] = []
    title_images = [
        image for image in images
        if not image.get("error")
        and _page_role(int(image.get("page") or 0), first_body_page, toc_pages) == "title_page"
    ]
    title_images.sort(key=lambda item: (item.get("range_start") or 0, item.get("image_index") or 0))
    for image, asset_path in zip(title_images, asset_paths):
        image["asset_path"] = asset_path
        image["path"] = asset_path
    if asset_paths and len(asset_paths) != len(title_images):
        warnings.append(
            f"asset_dir contains {len(asset_paths)} title images but reference inspection found {len(title_images)} title-page picture shapes"
        )
    return warnings


def _build_page_blueprint(
    paragraph_records: list[dict[str, Any]],
    image_records: list[dict[str, Any]],
    table_records: list[dict[str, Any]],
    shape_records: list[dict[str, Any]],
    page_count: int,
    first_body_page: int | None,
    toc_pages: set[int],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[str]]:
    pages = [{"page": page, "role": "body", "blocks": []} for page in range(1, page_count + 1)]
    _assign_page_roles(pages, first_body_page, toc_pages)
    page_by_number = {page["page"]: page for page in pages}
    warnings: list[str] = []
    flow_items: list[tuple[int, int, int, dict[str, Any]]] = []

    title_shapes = [shape for shape in shape_records if shape.get("page") == 1 and not shape.get("error")]
    title_images = [image for image in image_records if image.get("page") == 1 and not image.get("error")]
    title_text = " ".join(
        record["text"] for record in paragraph_records
        if record.get("page") == 1 and record.get("text") and record.get("text") != "\f"
    ).strip()
    if pages and pages[0]["role"] == "title_page":
        flow_items.append((
            1,
            0,
            0,
            {
                "type": "title_page",
                "page": 1,
                "role": "title_page",
                "text": title_text,
                "shapes": title_shapes,
                "images": title_images,
                "page_break_after": True,
                "fidelity": "inspected_only",
            },
        ))
        unsupported_title_shapes = [
            shape for shape in title_shapes
            if not _is_replayable_title_shape(shape) and not _is_picture_shape(shape)
        ]
        unmapped_title_images = [image for image in title_images if not image.get("path")]
        if unsupported_title_shapes or unmapped_title_images:
            warnings.append("title_page contains pictures or unsupported decorative shapes without mapped assets; creation remains approximate")

    toc_added_pages: set[int] = set()
    for record in paragraph_records:
        page = int(record.get("page") or 1)
        role = _page_role(page, first_body_page, toc_pages)
        if role == "title_page":
            continue
        text = record.get("text") or ""
        style = record.get("style") or "Normal"
        if text == "\f":
            flow_items.append((page, record.get("range_start") or 0, 0, {
                "type": "page_break",
                "page": page,
                "source": "reference",
            }))
            continue
        if role == "toc" and _looks_like_toc(style, text):
            if page not in toc_added_pages:
                flow_items.append((page, record.get("range_start") or 0, 0, {
                    "type": "toc",
                    "page": page,
                    "role": "toc",
                    "title": "Table of Contents",
                    "levels": 3,
                    "page_break_after": True,
                }))
                toc_added_pages.add(page)
            continue

        block = {k: v for k, v in record.items() if k not in {"range_start", "in_table"}}
        flow_items.append((page, record.get("range_start") or 0, 0, block))

    for table in table_records:
        if table.get("error"):
            warnings.append(f"table {table.get('table_index')} could not be inspected: {table.get('error')}")
            continue
        page = int(table.get("page") or 1)
        block = {k: v for k, v in table.items() if k not in {"range_start"}}
        flow_items.append((page, table.get("range_start") or 0, 1, block))

    for image in image_records:
        if image.get("error"):
            warnings.append(f"image {image.get('image_index')} could not be inspected: {image.get('error')}")
            continue
        page = int(image.get("page") or 1)
        if _page_role(page, first_body_page, toc_pages) == "title_page":
            continue
        block = {k: v for k, v in image.items() if k not in {"range_start", "source"}}
        if block.get("paragraph_index") is not None:
            block["anchor_paragraph_index"] = block.pop("paragraph_index")
        flow_items.append((page, image.get("range_start") or 0, 2, block))

    flow_items.sort(key=lambda item: (item[0], item[1], item[2]))
    blocks = [item[3] for item in flow_items]
    for page, _, _, block in flow_items:
        if page not in page_by_number:
            page_by_number[page] = {"page": page, "role": _page_role(page, first_body_page, toc_pages), "blocks": []}
            pages.append(page_by_number[page])
        page_by_number[page]["blocks"].append(block)

    return sorted(pages, key=lambda item: item["page"]), blocks, warnings


def _inspect_blueprint_live(filename: str, asset_dir: str = None) -> dict[str, Any]:
    if sys.platform != "win32":
        return {"error": "Blueprint inspection is only available on Windows"}

    from word_document_server.core.word_com import find_document, get_word_app

    app = get_word_app()
    doc = find_document(app, filename)
    setup = doc.PageSetup
    orientation = "landscape" if int(setup.Orientation) == 1 else "portrait"
    page_setup = {
        "size": "custom",
        "orientation": orientation,
        "width": float(setup.PageWidth),
        "height": float(setup.PageHeight),
        "margins": {
            "top": float(setup.TopMargin),
            "bottom": float(setup.BottomMargin),
            "left": float(setup.LeftMargin),
            "right": float(setup.RightMargin),
        },
    }

    paragraph_records: list[dict[str, Any]] = []
    paragraph_ranges: list[tuple[int, int, int]] = []
    first_body_page = None
    toc_pages: set[int] = set()

    for index in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(index)
        text = para.Range.Text.rstrip("\r\x07")
        style = _safe_style_name(para.Style)
        page = _range_page(para.Range)
        range_start = _range_start(para.Range)
        try:
            range_end = int(para.Range.End)
        except Exception:
            range_end = range_start
        paragraph_ranges.append((index, range_start, range_end))
        in_table = False
        try:
            in_table = bool(para.Range.Tables.Count)
        except Exception:
            in_table = False

        if page and _looks_like_toc(style, text):
            toc_pages.add(page)
        if page and first_body_page is None and style.replace(" ", "").lower().startswith("heading1"):
            if re.match(r"^\s*1(?:\.|\s)", text or ""):
                first_body_page = page

        block_type = "paragraph"
        level = None
        lowered = style.lower()
        if lowered.startswith("heading"):
            block_type = "heading"
            try:
                level = int(style.split()[-1])
            except Exception:
                level = 1
        if text == "\f":
            block_type = "page_break"
        block = {
            "type": block_type,
            "text": text,
            "style": style,
            "page": page,
            "paragraph_index": index,
            "range_start": range_start,
            "in_table": in_table,
        }
        if level is not None:
            block["level"] = level
        if block_type in {"heading", "paragraph"}:
            block["numbering"] = _infer_numbering(text, style, para)
        if not in_table or block_type == "page_break":
            paragraph_records.append(block)

    page_count = int(doc.ComputeStatistics(2))
    if first_body_page is None:
        for record in paragraph_records:
            if (
                record.get("page")
                and record.get("type") == "heading"
                and not _looks_like_toc(record.get("style"), record.get("text"))
            ):
                first_body_page = int(record["page"])
                break

    shapes = _inspect_shapes_live(doc, paragraph_ranges)
    tables = _inspect_tables_live(doc, paragraph_ranges)
    images = _inspect_inline_images_live(doc, paragraph_ranges) + _picture_shapes_as_images(shapes)
    asset_paths = _asset_paths(asset_dir)
    title_asset_paths, body_asset_paths = _split_asset_paths(asset_paths)
    warnings = _map_assets_to_title_images(images, title_asset_paths, first_body_page, toc_pages)
    warnings.extend(_map_assets_to_body_images(images, body_asset_paths, first_body_page, toc_pages))
    images.sort(key=lambda item: (item.get("page") or 0, item.get("range_start") or 0, item.get("source") or "", item.get("image_index") or 0))
    pages, blocks, page_warnings = _build_page_blueprint(
        paragraph_records,
        images,
        tables,
        shapes,
        page_count,
        first_body_page,
        toc_pages,
    )
    warnings.extend(page_warnings)

    return {
        "success": True,
        "session_blueprint": {
            "template": DEFAULT_TEMPLATE_ID,
            "document": {
                "title": _safe_doc_property(doc, "Title"),
                "subject": _safe_doc_property(doc, "Subject"),
                "author": _safe_doc_property(doc, "Author"),
                "page_setup": page_setup,
                "stats": {
                    "pages": page_count,
                    "words": doc.ComputeStatistics(0),
                    "paragraphs": doc.Paragraphs.Count,
                    "tables": doc.Tables.Count,
                    "inline_shapes": doc.InlineShapes.Count,
                    "shapes": doc.Shapes.Count,
                },
                "reference": {
                    "asset_dir": asset_dir,
                    "asset_count": len(asset_paths),
                    "title_asset_count": len(title_asset_paths),
                    "body_asset_count": len(body_asset_paths),
                    "first_body_page": first_body_page,
                    "toc_pages": sorted(toc_pages),
                },
                "pages": pages,
                "blocks": blocks,
                "images": images,
                "tables": tables,
                "shapes": shapes,
                "warnings": warnings,
                "sections": [{"blocks": blocks}],
            },
        },
        "warnings": warnings,
    }


def _blueprint_expected_counts(blueprint: dict | None) -> dict[str, int]:
    normalized = _normalize_blueprint(blueprint)
    paragraphs = 0
    tables = 0
    images = 0
    image_placeholders = 0
    if not isinstance(normalized["blocks"], list):
        return {"paragraphs": 0, "tables": 0, "images": 0, "image_placeholders": 0, "blocks": 0}
    for block in normalized["blocks"]:
        block_type = block.get("type")
        if block_type in {"paragraph", "heading", "image_placeholder"}:
            paragraphs += 1
            if block_type == "image_placeholder":
                image_placeholders += 1
        elif block_type == "title_page":
            if block.get("title") or block.get("text"):
                paragraphs += 1
        elif block_type == "toc":
            paragraphs += 1
        elif block_type == "image":
            images += 1
        elif block_type == "list":
            paragraphs += len(block.get("items") or [])
        elif block_type == "table":
            tables += 1
        elif block_type in {"page_break", "section_break"}:
            pass
    return {
        "paragraphs": paragraphs,
        "tables": tables,
        "images": images,
        "image_placeholders": image_placeholders,
        "blocks": len(normalized["blocks"]),
    }


def _safe_doc_property(doc, prop_name: str) -> str:
    try:
        value = doc.BuiltInDocumentProperties(prop_name).Value
        return str(value) if value is not None else ""
    except Exception:
        return ""


def _new_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:12]}"


def _register_session(open_result: dict[str, Any]) -> str:
    session_id = _new_id("word")
    filename = open_result.get("document") or open_result.get("name")
    _sessions[session_id] = {
        "session_id": session_id,
        "filename": filename,
        "full_path": open_result.get("full_path"),
        "template": open_result.get("template"),
        "created_at": time.time(),
        "updated_at": time.time(),
    }
    _handles[session_id] = {}
    return session_id


def _resolve_filename(session_id: str = None, filename: str = None) -> str | None:
    if filename:
        return filename
    if not session_id:
        return None
    session = _sessions.get(session_id)
    if not session:
        raise ValueError(f"Unknown session_id: {session_id}")
    return session.get("full_path") or session.get("filename")


def _touch_session(session_id: str = None, filename: str = None, full_path: str = None) -> None:
    if not session_id or session_id not in _sessions:
        return
    if filename:
        _sessions[session_id]["filename"] = filename
    if full_path:
        _sessions[session_id]["full_path"] = full_path
    _sessions[session_id]["updated_at"] = time.time()


def _session_records() -> list[dict[str, Any]]:
    """Return agent-facing session metadata without exposing large cached objects."""
    records = []
    now = time.time()
    for session_id, session in sorted(_sessions.items()):
        records.append({
            "session_id": session_id,
            "filename": session.get("filename"),
            "full_path": session.get("full_path"),
            "template": session.get("template"),
            "created_at": session.get("created_at"),
            "updated_at": session.get("updated_at"),
            "age_seconds": round(now - session.get("created_at", now), 1),
            "has_blueprint": bool(session.get("blueprint")),
            "handle_count": len(_handles.get(session_id, {})),
        })
    return records


def _long_find_usage(length: int, field: str = "find_text") -> dict[str, Any]:
    return {
        "error": f"{field} is {length} chars (Word Find limit: 255).",
        "usage": (
            "Do not split blindly if this is a long paragraph edit. Use a shorter unique "
            "search string with word_v2_search(context_chars=...), then edit with the returned "
            "handle; or replace a full paragraph by passing paragraph_index to word_v2_edit."
        ),
        "alternatives": [
            "word_v2_search(session_id, find_text='<short unique anchor>', context_chars=120)",
            "word_v2_edit(session_id, action='replace', paragraph_index=3, text='<new paragraph>', track_changes=True)",
            "word_v2_get_content(session_id, action='page_text', page=1, end_page=1) to get char offsets, then use start/end",
        ],
    }


def _with_session(result: dict[str, Any]) -> dict[str, Any]:
    session_id = _register_session(result)
    result["session_id"] = session_id
    result["mode"] = "live_com"
    result["next"] = [
        "word_v2_get_content(session_id, action='info')",
        "word_v2_search(session_id, find_text='...')",
    ]
    return result


def _matches_open_document(entry: dict[str, Any], path: str) -> bool:
    target = (path or "").strip().lower()
    if not target:
        return False

    candidates = [
        str(entry.get("index") or ""),
        str(entry.get("name") or ""),
        str(entry.get("full_path") or ""),
    ]
    full_path = str(entry.get("full_path") or "")
    if full_path:
        candidates.append(full_path.replace("\\", "/"))

    return any(candidate.lower() == target for candidate in candidates if candidate)


async def _attach_open_document(path: str = None) -> dict[str, Any]:
    listed = _load_result(await live_read_tools.word_live_list_open())
    if listed.get("error"):
        return listed

    documents = listed.get("documents", [])
    if not documents:
        return {
            "error": "No documents are open in Word. Provide path to open a file or use action='new'.",
            "documents": [],
        }

    selected = None
    if path:
        selected = next((doc for doc in documents if _matches_open_document(doc, path)), None)
        if selected is None:
            return {
                "error": f"Open document not found: {path}",
                "documents": documents,
                "usage": "Call word_v2_open(action='list') to inspect open documents, then attach by name, full_path, or index.",
            }
    else:
        selected = next((doc for doc in documents if doc.get("active")), None) or documents[0]

    return {
        "success": True,
        "document": selected.get("name"),
        "full_path": selected.get("full_path"),
        "pages": selected.get("pages"),
        "saved": selected.get("saved"),
        "track_revisions": selected.get("track_revisions"),
        "message": "Attached to an already-open Word document.",
        "open_documents": documents,
    }


def _store_handle(session_id: str, match: dict[str, Any], index: int) -> dict[str, Any]:
    handle = f"match_{index + 1}"
    target = {
        "kind": "selection",
        "start": match.get("start"),
        "end": match.get("end"),
    }
    enriched = {
        **match,
        "handle": handle,
        "target": target,
    }
    _handles.setdefault(session_id, {})[handle] = enriched
    return enriched


def _resolve_target(
    session_id: str = None,
    handle: str = None,
    target: dict[str, Any] = None,
    start: int = None,
    end: int = None,
) -> dict[str, Any]:
    if target:
        return target
    if handle:
        if not session_id:
            raise ValueError("session_id is required when using handle")
        stored = _handles.get(session_id, {}).get(handle)
        if not stored:
            raise ValueError(f"Unknown handle for session: {handle}")
        return stored["target"]
    if start is not None and end is not None:
        return {"kind": "selection", "start": start, "end": end}
    raise ValueError("Provide handle, target, or start/end")


def _target_range(target: dict[str, Any]) -> tuple[int, int]:
    if target.get("kind") != "selection":
        raise ValueError("target.kind must be 'selection'")
    start = target.get("start")
    end = target.get("end")
    if start is None or end is None:
        raise ValueError("selection target requires start and end")
    return int(start), int(end)


def _resolve_paragraph_target(filename: str, paragraph_index: int) -> dict[str, Any]:
    if paragraph_index is None:
        raise ValueError("paragraph_index is required")
    if paragraph_index < 1:
        raise ValueError("paragraph_index is 1-based; use 1 for the first paragraph")

    if sys.platform != "win32":
        raise ValueError("paragraph_index targeting is only available on Windows live sessions")

    from word_document_server.core.word_com import get_word_app, find_document

    app = get_word_app()
    doc = find_document(app, filename)
    total = doc.Paragraphs.Count
    if paragraph_index > total:
        raise ValueError(f"paragraph_index {paragraph_index} out of range (1-{total})")

    rng = doc.Paragraphs(paragraph_index).Range
    start = int(rng.Start)
    end = int(rng.End)
    try:
        text = str(rng.Text)
        while end > start and text.endswith(("\r", "\x07")):
            end -= 1
            text = text[:-1]
    except Exception:
        if end > start:
            end -= 1
    return {"kind": "selection", "start": start, "end": end, "paragraph_index": paragraph_index}


def _lower_insert_paragraph_index(paragraph_index: int | None) -> int | None:
    if paragraph_index is None:
        return None
    if paragraph_index < 1:
        raise ValueError("paragraph_index is 1-based; use 1 for the first paragraph")
    return paragraph_index - 1


def _normalize_mutation_tool_name(tool: str) -> str:
    normalized = (tool or "").lower().strip()
    if normalized.startswith("mcp_word_"):
        normalized = normalized[len("mcp_word_"):]
    if normalized.startswith("word_v2_"):
        normalized = normalized[len("word_v2_"):]
    return {
        "comments": "comment",
        "revision": "track_changes",
        "revisions": "track_changes",
        "track": "track_changes",
        "track_change": "track_changes",
        "tables": "table",
        "images": "media",
        "image": "media",
        "picture": "media",
        "pictures": "media",
        "properties": "layout",
    }.get(normalized, normalized)


async def word_v2_open(
    path: str = None,
    directory: str = ".",
    visible: bool = True,
    read_only: bool = False,
    password: str | None = None,
    action: str = "open",
) -> str:
    """Open, attach to, list, or create Word documents and return a v2 session when applicable."""
    action = (action or "open").lower().strip()
    path_alias = (path or "").lower().strip()

    if action == "list":
        result = _load_result(await live_read_tools.word_live_list_open())
        if not result.get("error"):
            result["session_count"] = len(_sessions)
            result["sessions"] = _session_records()
            result["usage"] = (
                "Call word_v2_open() to attach to the active document, "
                "word_v2_open(action='attach', path='<name|full_path|index>') to attach a listed document, "
                "word_v2_open(action='sessions') to list MCP session IDs, "
                "or word_v2_open(path='<file.docx>') to open a file."
            )
        return _dump(result)
    if action in {"sessions", "list_sessions"}:
        sessions = _session_records()
        return _dump({
            "success": True,
            "count": len(sessions),
            "sessions": sessions,
            "usage": "Use an existing session_id with word_v2_get_content, word_v2_search, word_v2_edit, word_v2_comment, word_v2_save, and word_v2_close.",
        })

    if action in {"active", "attach"} or (action == "open" and path_alias in {"", "active", "current"}):
        result = await _attach_open_document(None if path_alias in {"", "active", "current"} else path)
    elif action == "new" or path_alias in {"new", "blank", "create", "create_new"}:
        result = _load_result(await live_tools.word_live_create_document(visible=visible))
        if not result.get("error"):
            template_result = _apply_default_template_live(result.get("document"))
            result["template"] = DEFAULT_TEMPLATE_ID
            result["template_result"] = template_result
    elif action == "open":
        result = _load_result(await live_tools.word_live_open_document(
            filename=path,
            directory=directory,
            visible=visible,
            read_only=read_only,
            password=password,
        ))
    else:
        return _dump({
            "error": "Invalid action",
            "valid_actions": ["open", "active", "attach", "list", "sessions", "new"],
        })

    if result.get("error"):
        return _dump(result)

    return _dump(_with_session(result))


async def word_v2_save(session_id: str, out: str = None) -> str:
    """Save a v2 live session in place, to a new path, or export as PDF."""
    try:
        session = _sessions.get(session_id)
        if not session:
            raise ValueError(f"Unknown session_id: {session_id}")
        if not out and not session.get("full_path"):
            return _dump({
                "error": "This session is an unsaved document. Provide out='path.docx' for the first save."
            })
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    if out and out.lower().endswith(".pdf"):
        result = _load_result(await live_tools.word_live_convert_to_pdf(filename=filename, pdf_path=out))
        result["session_id"] = session_id
        return _dump(result)

    result = _load_result(await live_tools.word_live_save(filename=filename, save_as=out))
    if not result.get("error"):
        _touch_session(
            session_id,
            filename=result.get("document"),
            full_path=result.get("saved_as") or result.get("path"),
        )
        result["session_id"] = session_id
    return _dump(result)


async def word_v2_close(session_id: str, save_changes: str = "save") -> str:
    """Close a v2 live session. Unsaved changes follow save_changes."""
    try:
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    result = _load_result(await live_tools.word_live_close_document(
        filename=filename,
        save_changes=save_changes,
    ))
    if not result.get("error"):
        _sessions.pop(session_id, None)
        _handles.pop(session_id, None)
        result["session_id"] = session_id
    return _dump(result)


async def word_v2_get_content(
    session_id: str,
    action: str = "text",
    page: int = 1,
    end_page: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    include_runs: bool = False,
) -> str:
    """Read document content from a live session.

    Actions: text, page_text, info, comments, revisions, paragraph_format, snapshot, diff.
    """
    try:
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    action = (action or "text").lower()
    if action == "text":
        result = _load_result(await live_read_tools.word_live_get_text(filename))
    elif action == "page_text":
        result = _load_result(await live_read_tools.word_live_get_page_text(filename, page, end_page))
    elif action == "info":
        result = _load_result(await live_read_tools.word_live_get_info(filename))
    elif action == "comments":
        result = _load_result(await live_read_tools.word_live_get_comments(filename))
    elif action == "revisions":
        result = _load_result(await live_read_tools.word_live_list_revisions(filename))
    elif action == "paragraph_format":
        result = _load_result(await live_read_tools.word_live_get_paragraph_format(
            filename, start_paragraph, end_paragraph, include_runs
        ))
    elif action == "snapshot":
        result = _load_result(await live_read_tools.word_live_take_snapshot(filename))
    elif action == "diff":
        result = _load_result(await live_read_tools.word_live_get_diff(filename))
    else:
        return _dump({
            "error": "Invalid action",
            "valid_actions": ["text", "page_text", "info", "comments", "revisions", "paragraph_format", "snapshot", "diff"],
        })
    result["session_id"] = session_id
    return _dump(result)


async def word_v2_search(
    session_id: str,
    find_text: str,
    match_case: bool = False,
    whole_word: bool = False,
    use_wildcards: bool = False,
    context_chars: int = 80,
    max_results: int = 20,
) -> str:
    """Find text and return reusable handles for later v2 edit/format/comment calls."""
    try:
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    if len(find_text or "") > 255:
        return _dump(_long_find_usage(len(find_text or "")))

    result = _load_result(await live_read_tools.word_live_find_text(
        filename=filename,
        search_text=find_text,
        match_case=match_case,
        whole_word=whole_word,
        use_wildcards=use_wildcards,
        context_chars=context_chars,
        max_results=max_results,
    ))
    if result.get("error"):
        return _dump(result)

    matches = result.get("matches", [])
    result["matches"] = [_store_handle(session_id, match, i) for i, match in enumerate(matches)]
    result["session_id"] = session_id
    result["usage"] = "Pass a match handle to word_v2_edit, word_v2_format, or word_v2_comment."
    return _dump(result)


async def word_v2_edit(
    session_id: str,
    action: str,
    text: str = "",
    find_text: str = "",
    replace_text: str = "",
    handle: str = None,
    target: dict = None,
    start: int = None,
    end: int = None,
    position: str = "end",
    replace_all: bool = False,
    match_case: bool = False,
    whole_word: bool = False,
    use_wildcards: bool = False,
    paragraphs: list = None,
    paragraph_index: int = None,
    style: str = None,
    track_changes: bool = False,
    times: int = 1,
    url: str = "",
    footnote_index: int = None,
) -> str:
    """Perform live text edits. Actions: insert, replace, delete, insert_paragraphs, undo, add_hyperlink, add_footnote, delete_footnote."""
    try:
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    action = (action or "").lower()
    try:
        if action == "insert":
            if handle or target or (start is not None and end is not None):
                resolved = _resolve_target(session_id, handle, target, start, end)
                target_start, target_end = _target_range(resolved)
                position = str(target_start if position == "before" else target_end)
            result = _load_result(await live_tools.word_live_insert_text(
                filename, text, position, None, track_changes
            ))
        elif action == "replace":
            if find_text:
                if len(find_text) > 255:
                    return _dump(_long_find_usage(len(find_text)))
                result = _load_result(await live_tools.word_live_replace_text(
                    filename, find_text, replace_text or text, match_case,
                    whole_word, use_wildcards, replace_all, track_changes,
                ))
            else:
                if paragraph_index is not None and not (handle or target or (start is not None and end is not None)):
                    resolved = _resolve_paragraph_target(filename, paragraph_index)
                else:
                    resolved = _resolve_target(session_id, handle, target, start, end)
                target_start, target_end = _target_range(resolved)
                delete_result = _load_result(await live_tools.word_live_delete_text(
                    filename, target_start, target_end, track_changes,
                ))
                if delete_result.get("error"):
                    return _dump(delete_result)
                insert_result = _load_result(await live_tools.word_live_insert_text(
                    filename, replace_text or text, str(target_start), None, track_changes,
                ))
                result = {"success": not insert_result.get("error"), "delete": delete_result, "insert": insert_result}
        elif action == "delete":
            if find_text:
                if len(find_text) > 255:
                    return _dump(_long_find_usage(len(find_text)))
                result = _load_result(await live_tools.word_live_replace_text(
                    filename, find_text, "", match_case, whole_word,
                    use_wildcards, replace_all, track_changes,
                ))
            else:
                if paragraph_index is not None and not (handle or target or (start is not None and end is not None)):
                    resolved = _resolve_paragraph_target(filename, paragraph_index)
                else:
                    resolved = _resolve_target(session_id, handle, target, start, end)
                target_start, target_end = _target_range(resolved)
                result = _load_result(await live_tools.word_live_delete_text(
                    filename, target_start, target_end, track_changes,
                ))
        elif action == "insert_paragraphs":
            lower_paragraph_index = _lower_insert_paragraph_index(paragraph_index)
            result = _load_result(await live_tools.word_live_insert_paragraphs(
                filename, paragraphs, find_text or None, lower_paragraph_index,
                position, style, track_changes,
            ))
        elif action == "undo":
            result = _load_result(await live_tools.word_live_undo(filename, times))
        elif action == "add_hyperlink":
            if handle or target or (start is not None and end is not None):
                resolved = _resolve_target(session_id, handle, target, start, end)
                target_start, target_end = _target_range(resolved)
            else:
                target_start, target_end = None, None
            result = _load_result(await live_tools.word_live_add_hyperlink(
                filename=filename, url=url, text=text, start=target_start, end=target_end
            ))
        elif action == "add_footnote":
            if handle or target or (start is not None and end is not None):
                resolved = _resolve_target(session_id, handle, target, start, end)
                target_start, target_end = _target_range(resolved)
            else:
                target_start, target_end = None, None
            result = _load_result(await live_tools.word_live_add_footnote(
                filename=filename, text=text, start=target_start, end=target_end
            ))
        elif action == "delete_footnote":
            result = _load_result(await live_tools.word_live_delete_footnote(
                filename=filename, index=footnote_index
            ))
        else:
            return _dump({
                "error": "Invalid action",
                "valid_actions": ["insert", "replace", "delete", "insert_paragraphs", "undo", "add_hyperlink", "add_footnote", "delete_footnote"],
            })
    except ValueError as exc:
        return _dump({"error": str(exc)})

    result["session_id"] = session_id
    return _dump(result)


async def word_v2_format(
    session_id: str,
    action: str = "inline",
    handle: str = None,
    target: dict = None,
    start: int = None,
    end: int = None,
    start_paragraph: int = None,
    end_paragraph: int = None,
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    strikethrough: bool = None,
    font_name: str = None,
    font_size: float = None,
    font_color: str = None,
    highlight_color: int = None,
    style: str = None,
    alignment: str = None,
    page_break_before: bool = None,
    preserve_direct_formatting: bool = False,
    list_type: str = "bullet",
    level: int = 0,
    remove: bool = False,
    continue_previous: bool = False,
    track_changes: bool = False,
) -> str:
    """Apply live formatting. Actions: inline, paragraph, style, list."""
    try:
        filename = _resolve_filename(session_id=session_id)
        if handle or target:
            resolved = _resolve_target(session_id, handle, target, start, end)
            start, end = _target_range(resolved)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    action = (action or "inline").lower()
    if action in {"inline", "paragraph", "style"}:
        result = _load_result(await live_tools.word_live_format_text(
            filename, start, end, start_paragraph, end_paragraph,
            bold, italic, underline, strikethrough, font_name, font_size,
            font_color, highlight_color, style, alignment,
            page_break_before, preserve_direct_formatting, track_changes,
        ))
    elif action == "list":
        result = _load_result(await live_tools.word_live_apply_list(
            filename, start_paragraph, end_paragraph, list_type, level, remove,
            continue_previous, None, None, None, None, track_changes,
        ))
    else:
        return _dump({"error": "Invalid action", "valid_actions": ["inline", "paragraph", "style", "list"]})

    result["session_id"] = session_id
    return _dump(result)


async def word_v2_comment(
    session_id: str,
    action: str,
    comment_id: int = None,
    text: str = "",
    handle: str = None,
    target: dict = None,
    start: int = None,
    end: int = None,
    paragraph_index: int = None,
    author: str = DEFAULT_AUTHOR,
    resolve: bool = True,
) -> str:
    """Manage comments. Actions: create, reply, resolve, delete, list, get."""
    try:
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    action = (action or "").lower()
    try:
        if action == "create":
            if handle or target:
                resolved = _resolve_target(session_id, handle, target, start, end)
                start, end = _target_range(resolved)
            if start is None and end is None and paragraph_index is None:
                return _dump({
                    "error": "Comment target required",
                    "usage": (
                        "Provide one of: handle from word_v2_search, target, start/end offsets, "
                        "or paragraph_index from word_v2_get_content. Example: "
                        "word_v2_comment(session_id, action='create', paragraph_index=3, text='...')."
                    ),
                    "valid_targets": ["handle", "target", "start/end", "paragraph_index"],
                })
            result = _load_result(await live_read_tools.word_live_add_comment(
                filename, start, end, paragraph_index, text, author,
            ))
        elif action == "reply":
            result = _load_result(await live_read_tools.word_live_reply_to_comment(
                filename, comment_id, text, author,
            ))
        elif action == "resolve":
            result = _load_result(await live_read_tools.word_live_resolve_comment(
                filename, comment_id, resolve,
            ))
        elif action == "delete":
            result = _load_result(await live_read_tools.word_live_delete_comment(filename, comment_id))
        elif action in {"list", "get"}:
            result = _load_result(await live_read_tools.word_live_get_comments(filename))
            if action == "get" and comment_id is not None and not result.get("error"):
                comments = result.get("comments", [])
                result["comment"] = next((c for c in comments if c.get("index") == comment_id), None)
        else:
            return _dump({"error": "Invalid action", "valid_actions": ["create", "reply", "resolve", "delete", "list", "get"]})
    except ValueError as exc:
        return _dump({"error": str(exc)})

    result["session_id"] = session_id
    return _dump(result)


async def word_v2_track_changes(
    session_id: str,
    action: str,
    enable: bool = None,
    author: str = None,
    change_ids: list[int] = None,
    decision: str = "accept",
) -> str:
    """Manage tracked changes. Actions: toggle, list, accept, reject, decide."""
    try:
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    action = (action or "").lower()
    if action == "toggle":
        result = _load_result(await live_tools.word_live_toggle_track_changes(filename, enable))
    elif action == "list":
        result = _load_result(await live_read_tools.word_live_list_revisions(filename))
    elif action == "accept" or (action == "decide" and decision == "accept"):
        result = _load_result(await live_read_tools.word_live_accept_revisions(filename, author, change_ids))
    elif action == "reject" or (action == "decide" and decision == "reject"):
        result = _load_result(await live_read_tools.word_live_reject_revisions(filename, author, change_ids))
    else:
        return _dump({"error": "Invalid action", "valid_actions": ["toggle", "list", "accept", "reject", "decide"]})

    result["session_id"] = session_id
    return _dump(result)


async def word_v2_table(
    session_id: str,
    action: str,
    table_index: int = 1,
    rows: int = 2,
    cols: int = 2,
    position: str = "end",
    data: list = None,
    row: int = None,
    col: int = None,
    text: str = None,
    cells: list = None,
    start_row: int = None,
    start_col: int = None,
    end_row: int = None,
    end_col: int = None,
    style: str = "Table Grid",
    autofit: str = "window",
    border_style: str = None,
    cell_bold: list[list] = None,
    cell_alignment: list[list] = None,
    column_widths: list[float] = None,
    table_alignment: str = None,
    cell_shading: list[list] = None,
    track_changes: bool = False,
) -> str:
    """Create, edit, and format tables. Actions mirror common table intents."""
    try:
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    action = (action or "").lower()
    if action in {"create", "add"}:
        result = _load_result(await live_tools.word_live_add_table(
            filename, rows, cols, position, data, style, autofit, track_changes,
        ))
    elif action == "format":
        result = _load_result(await live_tools.word_live_format_table(
            filename, table_index, border_style, cell_bold, cell_alignment,
            column_widths, table_alignment, cell_shading, autofit,
        ))
    elif action in {
        "get_info", "set_cell", "set_row", "set_range", "add_row", "delete_row",
        "add_column", "delete_column", "merge_cells", "autofit", "delete_table",
    }:
        if action == "delete_column" and col is None:
            return _dump({
                "error": "delete_column requires col",
                "usage": "Pass the 1-based column number, e.g. word_v2_table(session_id, action='delete_column', table_index=1, col=2).",
            })
        result = _load_result(await live_tools.word_live_modify_table(
            filename, table_index, action, row, col, text, row, col, None, cells,
            start_row, start_col, end_row, end_col, autofit, False, track_changes,
        ))
    else:
        return _dump({
            "error": "Invalid action",
            "valid_actions": [
                "create", "format", "get_info", "set_cell", "set_row", "set_range",
                "add_row", "delete_row", "add_column", "delete_column",
                "merge_cells", "autofit", "delete_table",
            ],
        })

    result["session_id"] = session_id
    return _dump(result)


async def word_v2_media(
    session_id: str,
    action: str,
    path: str,
    paragraph_index: int = None,
    position: str = "end",
    width_inches: float = None,
    height_inches: float = None,
    width_pt: float = None,
    height_pt: float = None,
    alignment: str = None,
    wrapping: str = "inline",
    border_style: str = None,
    border_width_pt: float = None,
    border_color: str = None,
    link_to_file: bool = False,
    paragraph_after: bool = True,
    left_pt: float = None,
    top_pt: float = None,
    relative_horizontal_position: int = None,
    relative_vertical_position: int = None,
) -> str:
    """Insert media into a live session. Actions: insert_image."""
    action = (action or "").lower()
    if action not in {"insert_image", "image", "insert"}:
        return _dump({"error": "Invalid action", "valid_actions": ["insert_image"]})
    if not path:
        return _dump({"error": "path is required"})

    try:
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    result = _load_result(await live_tools.word_live_insert_image(
        filename=filename,
        image_path=path,
        paragraph_index=paragraph_index,
        position=position,
        width_inches=width_inches,
        height_inches=height_inches,
        width_pt=width_pt,
        height_pt=height_pt,
        alignment=alignment,
        wrapping=wrapping,
        border_style=border_style,
        border_width_pt=border_width_pt,
        border_color=border_color,
        link_to_file=link_to_file,
        paragraph_after=paragraph_after,
        left_pt=left_pt,
        top_pt=top_pt,
        relative_horizontal_position=relative_horizontal_position,
        relative_vertical_position=relative_vertical_position,
    ))
    result["session_id"] = session_id
    return _dump(result)


async def word_v2_mutations(
    session_id: str,
    action: str,
    operations: list[dict] = None,
) -> str:
    """Preview or apply multiple v2 operations.

    Each operation is {"tool": "edit|format|comment|track_changes|table|media|layout", ...args}.
    Preview validates shape only; apply runs operations in order.
    """
    operations = operations or []
    if action == "preview":
        return _dump({
            "success": True,
            "session_id": session_id,
            "operation_count": len(operations),
            "operations": operations,
            "message": "Preview only. Call action='apply' to execute.",
        })
    if action != "apply":
        return _dump({"error": "Invalid action", "valid_actions": ["preview", "apply"]})

    results = []
    for i, operation in enumerate(operations):
        requested_tool = operation.get("tool") or ""
        tool = _normalize_mutation_tool_name(requested_tool)
        args = {k: v for k, v in operation.items() if k != "tool"}
        args["session_id"] = session_id
        if tool == "edit":
            raw = await word_v2_edit(**args)
        elif tool == "format":
            raw = await word_v2_format(**args)
        elif tool == "comment":
            raw = await word_v2_comment(**args)
        elif tool == "track_changes":
            raw = await word_v2_track_changes(**args)
        elif tool == "table":
            raw = await word_v2_table(**args)
        elif tool == "media":
            raw = await word_v2_media(**args)
        elif tool == "layout":
            raw = await word_v2_layout(**args)
        else:
            raw = _dump({
                "error": f"Invalid operation tool at index {i}: {requested_tool}",
                "valid_tools": ["edit", "format", "comment", "track_changes", "table", "media", "layout"],
                "usage": "Use short names like 'edit' or full public names like 'word_v2_edit'.",
            })
        result = _load_result(raw)
        results.append({"index": i, "tool": tool, "requested_tool": requested_tool, "result": result})
        if result.get("error"):
            return _dump({"success": False, "session_id": session_id, "results": results})

    return _dump({"success": True, "session_id": session_id, "results": results})


async def word_v2_layout(
    session_id: str,
    action: str,
    page_size: str = "letter",
    orientation: str = "portrait",
    width: float = None,
    height: float = None,
    margins: dict = None,
    position: str = "end",
    paragraph_index: int = None,
    break_type: str = "next_page",
    title: str = None,
    subject: str = None,
    author: str = None,
    keywords: str = None,
    comments: str = None,
    category: str = None,
    manager: str = None,
    company: str = None,
    last_author: str = None,
) -> str:
    """Manage page setup, breaks, and document properties for a live session."""
    action = (action or "").lower()
    if action == "break":
        action = "page_break" if (break_type or "").lower() in {"", "page", "page_break", "manual_page"} else "section_break"
    valid_actions = ["page_setup", "page_break", "section_break", "properties", "break"]
    if action not in valid_actions:
        return _dump({"error": "Invalid action", "valid_actions": valid_actions})

    try:
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    if action == "page_setup":
        try:
            result = _layout_page_setup_live(
                filename,
                {
                    "size": page_size,
                    "orientation": orientation,
                    "width": width,
                    "height": height,
                    "margins": margins,
                },
            )
        except ValueError as exc:
            result = {"error": str(exc)}
    elif action == "page_break":
        result = _layout_insert_break_live(
            filename,
            "page",
            position=position,
            paragraph_index=paragraph_index,
        )
    elif action == "section_break":
        result = _layout_insert_break_live(
            filename,
            "section",
            position=position,
            paragraph_index=paragraph_index,
            break_type=break_type,
        )
    else:
        result = _load_result(await live_read_tools.word_live_set_core_properties(
            filename=filename,
            title=title,
            subject=subject,
            author=author,
            keywords=keywords,
            comments=comments,
            category=category,
            manager=manager,
            company=company,
            last_author=last_author,
        ))

    result["session_id"] = session_id
    return _dump(result)


async def _paragraph_count(session_id: str) -> int | None:
    result = _load_result(await word_v2_get_content(session_id=session_id, action="info"))
    if result.get("error"):
        return None
    value = result.get("paragraphs")
    return int(value) if value is not None else None


def _find_recent_paragraph_index(session_id: str, text: str, fallback: int | None = None) -> int | None:
    needle = (text or "").strip()
    if not needle or sys.platform != "win32":
        return fallback
    try:
        from word_document_server.core.word_com import find_document, get_word_app

        filename = _resolve_filename(session_id=session_id)
        doc = find_document(get_word_app(), filename)
        for index in range(int(doc.Paragraphs.Count), 0, -1):
            para_text = str(doc.Paragraphs(index).Range.Text).rstrip("\r\x07").strip()
            if para_text == needle:
                return index
    except Exception:
        pass
    return fallback


def _apply_shape_position(shape_obj, shape: dict[str, Any]) -> None:
    for key, attr in [
        ("relative_horizontal_position", "RelativeHorizontalPosition"),
        ("relative_vertical_position", "RelativeVerticalPosition"),
        ("left_pt", "Left"),
        ("top_pt", "Top"),
    ]:
        value = shape.get(key)
        if value is None:
            continue
        try:
            setattr(shape_obj, attr, int(value) if key.startswith("relative_") else float(value))
        except Exception:
            pass
    wrap_type = shape.get("wrap_type")
    if wrap_type is not None:
        try:
            shape_obj.WrapFormat.Type = int(wrap_type)
        except Exception:
            pass


def _apply_shape_visuals(shape_obj, shape: dict[str, Any]) -> None:
    fill_visible = shape.get("fill_visible")
    if fill_visible is not None:
        try:
            shape_obj.Fill.Visible = int(fill_visible)
        except Exception:
            pass
    fill_color = shape.get("fill_color")
    if fill_color is not None:
        try:
            shape_obj.Fill.ForeColor.RGB = int(fill_color)
        except Exception:
            pass
    line_visible = shape.get("line_visible")
    if line_visible is not None:
        try:
            shape_obj.Line.Visible = int(line_visible)
        except Exception:
            pass
    line_color = shape.get("line_color")
    if line_color is not None:
        try:
            shape_obj.Line.ForeColor.RGB = int(line_color)
        except Exception:
            pass
    line_weight = shape.get("line_weight")
    if line_weight is not None and float(line_weight) >= 0:
        try:
            shape_obj.Line.Weight = float(line_weight)
        except Exception:
            pass


def _apply_text_box_format(shape_obj, shape: dict[str, Any]) -> None:
    text_range = _safe_attr(_safe_attr(shape_obj, "TextFrame"), "TextRange")
    if text_range is None:
        return
    text_range.Text = str(shape.get("text") or "")
    font = _safe_attr(text_range, "Font")
    if font is None:
        return
    font_name = shape.get("font_name")
    if font_name:
        try:
            font.Name = str(font_name)
        except Exception:
            pass
    font_size = shape.get("font_size")
    if font_size is not None and 0 < float(font_size) < 500:
        try:
            font.Size = float(font_size)
        except Exception:
            pass
    bold = shape.get("bold")
    if bold in {-1, 0, 1, True, False}:
        try:
            font.Bold = int(bold)
        except Exception:
            pass
    font_color = shape.get("font_color")
    if font_color is not None and 0 <= int(font_color) <= 0xFFFFFF:
        try:
            font.Color = int(font_color)
        except Exception:
            pass


def _insert_title_shape_live(session_id: str, shape: dict[str, Any]) -> dict[str, Any]:
    if sys.platform != "win32":
        return {"error": "shape replay is only available on Windows"}
    if not isinstance(shape, dict) or shape.get("error"):
        return {"success": False, "skipped": True, "reason": "invalid shape"}
    name = str(shape.get("name") or "")
    is_text_box = bool(shape.get("text")) or name.lower().startswith("text box") or shape.get("shape_type") == 17
    is_rectangle = name.lower().startswith("rectangle") or shape.get("auto_shape_type") == 1
    if not is_text_box and not is_rectangle:
        return {"success": True, "skipped": True, "reason": "unsupported shape type", "name": name}

    try:
        from word_document_server.core.word_com import find_document, get_word_app, undo_record

        filename = _resolve_filename(session_id=session_id)
        app = get_word_app()
        doc = find_document(app, filename)
        anchor = doc.Range(0, 0)
        left = float(shape.get("left_pt") or 0)
        top = float(shape.get("top_pt") or 0)
        width = float(shape.get("width_pt") or 100)
        height = float(shape.get("height_pt") or 40)
        with undo_record(app, "MCP: Replay Title Shape"):
            if is_text_box:
                created = doc.Shapes.AddTextbox(Orientation=1, Left=left, Top=top, Width=width, Height=height, Anchor=anchor)
                _apply_text_box_format(created, shape)
            else:
                auto_shape_type = int(shape.get("auto_shape_type") or 1)
                created = doc.Shapes.AddShape(Type=auto_shape_type, Left=left, Top=top, Width=width, Height=height, Anchor=anchor)
            _apply_shape_position(created, shape)
            _apply_shape_visuals(created, shape)
        return {
            "success": True,
            "name": name,
            "text_box": is_text_box,
            "rectangle": is_rectangle and not is_text_box,
            "width_pt": width,
            "height_pt": height,
            "left_pt": left,
            "top_pt": top,
        }
    except Exception as exc:
        return {"error": str(exc), "name": name}


async def _apply_blueprint_block(session_id: str, block: dict[str, Any]) -> dict[str, Any]:
    block_type = block.get("type")
    if block_type == "title_page":
        title = block.get("title") or block.get("text") or ""
        if title:
            raw = await word_v2_edit(
                session_id=session_id,
                action="insert_paragraphs",
                paragraphs=[str(title)],
                position="end",
                style=block.get("style") or "Heading 3",
            )
            result = _load_result(raw)
            if result.get("error"):
                return {"block_type": block_type, "result": result}
        else:
            result = {"success": True, "message": "empty title_page block"}

        image_results = []
        for image in block.get("images") or []:
            path = image.get("path") or image.get("asset_path")
            if not path:
                continue
            image_raw = await word_v2_media(
                session_id=session_id,
                action="insert_image",
                path=path,
                position="end",
                width_pt=image.get("width_pt"),
                height_pt=image.get("height_pt"),
                alignment=image.get("alignment") or "center",
                wrapping=image.get("wrapping") or _word_wrap_name(image.get("wrap_type")) or "inline",
                paragraph_after=image.get("paragraph_after", True),
                left_pt=image.get("left_pt"),
                top_pt=image.get("top_pt"),
                relative_horizontal_position=image.get("relative_horizontal_position"),
                relative_vertical_position=image.get("relative_vertical_position"),
            )
            image_results.append(_load_result(image_raw))

        shape_results = []
        for shape in block.get("shapes") or []:
            if _is_picture_shape(shape):
                continue
            shape_results.append(_insert_title_shape_live(session_id, shape))

        break_result = None
        if block.get("page_break_after", True):
            break_result = _load_result(await word_v2_layout(
                session_id=session_id,
                action="page_break",
                position="end",
            ))
        response = {
            "block_type": block_type,
            "result": result,
            "image_results": image_results,
            "shape_results": shape_results,
            "break_result": break_result,
        }
        if any(not (image.get("path") or image.get("asset_path")) for image in block.get("images") or []):
            response["warning"] = "title_page creation replays simple text boxes/rectangles and remains approximate for pictures without asset paths."
        return response

    if block_type == "toc":
        try:
            filename = _resolve_filename(session_id=session_id)
        except ValueError as exc:
            return {"block_type": block_type, "result": {"error": str(exc)}}
        result = _insert_toc_live(
            filename=filename,
            title=block.get("title") or "Table of Contents",
            levels=int(block.get("levels") or 3),
            page_break_after=block.get("page_break_after", True),
        )
        return {"block_type": block_type, "result": result}

    if block_type == "paragraph":
        style = block.get("style") or "Normal"
        before = await _paragraph_count(session_id)
        raw = await word_v2_edit(
            session_id=session_id,
            action="insert_paragraphs",
            paragraphs=[str(block.get("text") or "")],
            position="end",
            style=style,
        )
        result = _load_result(raw)
        format_result = None
        list_format = _blueprint_list_format(block.get("numbering"))
        after = await _paragraph_count(session_id) if list_format and not result.get("error") else None
        if before is not None and after is not None and after > before:
            fallback_index = max(1, after - 1)
            target_paragraph = _find_recent_paragraph_index(
                session_id,
                str(block.get("text") or ""),
                fallback=fallback_index,
            )
            format_result = _load_result(await word_v2_format(
                session_id=session_id,
                action="list",
                start_paragraph=target_paragraph,
                end_paragraph=target_paragraph,
                list_type=list_format["list_type"],
                level=list_format["level"],
                continue_previous=list_format["continue_previous"],
            ))
        return {"block_type": block_type, "result": result, "format_result": format_result}

    if block_type == "heading":
        if block.get("page_break_before"):
            await word_v2_layout(session_id=session_id, action="page_break", position="end")
        level = int(block.get("level") or 1)
        style = block.get("style") or f"Heading {level}"
        raw = await word_v2_edit(
            session_id=session_id,
            action="insert_paragraphs",
            paragraphs=[str(block.get("text") or "")],
            position="end",
            style=style,
        )
        return {"block_type": block_type, "result": _load_result(raw)}

    if block_type == "list":
        items = [str(item) for item in block.get("items") or []]
        before = await _paragraph_count(session_id)
        raw = await word_v2_edit(
            session_id=session_id,
            action="insert_paragraphs",
            paragraphs=items,
            position="end",
            style=block.get("style") or "List Paragraph",
        )
        result = _load_result(raw)
        after = await _paragraph_count(session_id)
        format_result = None
        if before is not None and after is not None and after >= before + len(items):
            list_type = "number" if block.get("ordered") else "bullet"
            format_result = _load_result(await word_v2_format(
                session_id=session_id,
                action="list",
                start_paragraph=before + 1,
                end_paragraph=after,
                list_type=block.get("list_type") or list_type,
                level=int(block.get("level") or 0),
            ))
        return {"block_type": block_type, "result": result, "format_result": format_result}

    if block_type == "table":
        rows = block.get("rows") or block.get("data") or []
        row_count = len(rows)
        col_count = max(len(row) for row in rows) if rows else 0
        raw = await word_v2_table(
            session_id=session_id,
            action="create",
            rows=row_count,
            cols=col_count,
            position="end",
            data=rows,
            style=block.get("style") or "Table Grid",
            autofit=block.get("autofit") or "window",
        )
        return {"block_type": block_type, "result": _load_result(raw)}

    if block_type == "image":
        raw = await word_v2_media(
            session_id=session_id,
            action="insert_image",
            path=block.get("path"),
            paragraph_index=block.get("paragraph_index"),
            position=block.get("position") or "end",
            width_inches=block.get("width_inches"),
            height_inches=block.get("height_inches"),
            width_pt=block.get("width_pt"),
            height_pt=block.get("height_pt"),
            alignment=block.get("alignment"),
            wrapping=block.get("wrapping") or _word_wrap_name(block.get("wrap_type")) or "inline",
            border_style=block.get("border_style"),
            border_width_pt=block.get("border_width_pt"),
            border_color=block.get("border_color"),
            link_to_file=bool(block.get("link_to_file")),
            paragraph_after=block.get("paragraph_after", True),
            left_pt=block.get("left_pt"),
            top_pt=block.get("top_pt"),
            relative_horizontal_position=block.get("relative_horizontal_position"),
            relative_vertical_position=block.get("relative_vertical_position"),
        )
        return {"block_type": block_type, "result": _load_result(raw)}

    if block_type == "image_placeholder":
        label = block.get("label") or block.get("asset_id") or "image"
        raw = await word_v2_edit(
            session_id=session_id,
            action="insert_paragraphs",
            paragraphs=[f"[Image placeholder: {label}]"],
            position="end",
            style=block.get("style") or "Caption",
        )
        return {
            "block_type": block_type,
            "result": _load_result(raw),
            "warning": "image_placeholder creates editable placeholder text; use an image block or word_v2_media for real images.",
        }

    if block_type == "page_break":
        raw = await word_v2_layout(
            session_id=session_id,
            action="page_break",
            position=block.get("position") or "end",
            paragraph_index=block.get("paragraph_index"),
        )
        return {"block_type": block_type, "result": _load_result(raw)}

    if block_type == "section_break":
        raw = await word_v2_layout(
            session_id=session_id,
            action="section_break",
            position=block.get("position") or "end",
            paragraph_index=block.get("paragraph_index"),
            break_type=block.get("break_type") or "next_page",
        )
        return {"block_type": block_type, "result": _load_result(raw)}

    return {"block_type": block_type, "result": {"error": f"Unsupported block type: {block_type}"}}


async def word_v2_blueprint(
    action: str,
    session_id: str = None,
    blueprint: dict = None,
    out: str = None,
    visible: bool = True,
    asset_dir: str = None,
) -> str:
    """Create, inspect, validate, or export structured document blueprints."""
    action = (action or "").lower()
    valid_actions = ["create", "inspect", "validate", "export"]
    if action not in valid_actions:
        return _dump({"error": "Invalid action", "valid_actions": valid_actions})

    if action in {"inspect", "export"}:
        try:
            filename = _resolve_filename(session_id=session_id)
        except ValueError as exc:
            return _dump({"error": str(exc)})
        result = _inspect_blueprint_live(filename, asset_dir=asset_dir)
        result["session_id"] = session_id
        if not result.get("error") and session_id in _sessions:
            _sessions[session_id]["blueprint"] = result.get("session_blueprint")
        return _dump(result)

    errors = _validate_blueprint(blueprint)
    if action == "validate":
        result = {
            "success": not errors,
            "errors": errors,
            "expected": _blueprint_expected_counts(blueprint),
        }
        if session_id:
            info = _load_result(await word_v2_get_content(session_id=session_id, action="info"))
            result["current"] = {
                "paragraphs": info.get("paragraphs"),
                "tables": info.get("tables"),
                "pages": info.get("pages"),
            } if not info.get("error") else {"error": info.get("error")}
            mismatches = []
            expected = result["expected"]
            current = result["current"]
            if not current.get("error"):
                for key in ("paragraphs", "tables"):
                    if current.get(key) is not None and expected.get(key) != current.get(key):
                        mismatches.append({"field": key, "expected": expected.get(key), "actual": current.get(key)})
            result["mismatches"] = mismatches
            result["success"] = result["success"] and not mismatches
        return _dump(result)

    if errors:
        return _dump({"success": False, "errors": errors})

    normalized = _normalize_blueprint(blueprint)
    created = _load_result(await word_v2_open(action="new", visible=visible))
    if created.get("error"):
        return _dump(created)
    new_session_id = created["session_id"]

    page_setup = normalized.get("page_setup") or {}
    if page_setup:
        await word_v2_layout(
            session_id=new_session_id,
            action="page_setup",
            page_size=page_setup.get("size") or "letter",
            orientation=page_setup.get("orientation") or "portrait",
            width=page_setup.get("width"),
            height=page_setup.get("height"),
            margins=page_setup.get("margins"),
        )

    properties = normalized.get("properties") or {}
    if properties:
        await word_v2_layout(
            session_id=new_session_id,
            action="properties",
            title=properties.get("title"),
            subject=properties.get("subject"),
            author=properties.get("author"),
            keywords=properties.get("keywords"),
            comments=properties.get("comments"),
            category=properties.get("category"),
            manager=properties.get("manager"),
            company=properties.get("company"),
            last_author=properties.get("last_author"),
        )

    block_results = []
    warnings = []
    for index, block in enumerate(normalized["blocks"]):
        block_result = await _apply_blueprint_block(new_session_id, block)
        block_result["index"] = index
        block_results.append(block_result)
        if block_result.get("warning"):
            warnings.append({"index": index, "warning": block_result["warning"]})
        result = block_result.get("result") or {}
        if result.get("error"):
            return _dump({
                "success": False,
                "session_id": new_session_id,
                "created": created,
                "results": block_results,
                "warnings": warnings,
            })

    if any((block.get("type") == "toc") for block in normalized["blocks"]):
        try:
            filename = _resolve_filename(session_id=new_session_id)
            update_result = _update_tocs_live(filename)
            if update_result.get("error"):
                warnings.append({"index": None, "warning": update_result["error"]})
        except Exception as exc:
            warnings.append({"index": None, "warning": f"TOC update failed: {exc}"})

    save_result = None
    if out:
        save_result = _load_result(await word_v2_save(session_id=new_session_id, out=out))

    if new_session_id in _sessions:
        _sessions[new_session_id]["blueprint"] = normalized
    return _dump({
        "success": True,
        "session_id": new_session_id,
        "created": created,
        "block_count": len(normalized["blocks"]),
        "results": block_results,
        "warnings": warnings,
        "save_result": save_result,
    })


async def word_v2_protection(
    session_id: str,
    action: str,
    protection_type: str = "read_only",
    password: str = None,
) -> str:
    """Manage editing protection on a live session.

    Actions: protect, unprotect.
    Protection Types: read_only, comments, tracked_changes.
    """
    try:
        filename = _resolve_filename(session_id=session_id)
    except ValueError as exc:
        return _dump({"error": str(exc)})

    action = (action or "").lower()
    if action == "protect":
        result = _load_result(await live_tools.word_live_protect_document(
            filename=filename, protection_type=protection_type, password=password
        ))
    elif action == "unprotect":
        result = _load_result(await live_tools.word_live_unprotect_document(
            filename=filename, password=password
        ))
    else:
        return _dump({"error": "Invalid action. Supported actions: 'protect', 'unprotect'"})

    result["session_id"] = session_id
    return _dump(result)
