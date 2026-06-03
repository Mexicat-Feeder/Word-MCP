"""V2 live Word tools: grouped, agent-friendly facade over COM helpers.

The v2 surface is inspired by SuperDoc's MCP shape: a few lifecycle tools plus
grouped intent tools. It keeps the existing live COM implementations underneath
so v2 can improve tool ergonomics without rewriting Word automation logic.
"""

from __future__ import annotations

import json
import sys
import time
import uuid
from typing import Any

from word_document_server.defaults import DEFAULT_AUTHOR
from word_document_server.tools import live_read_tools, live_tools


_sessions: dict[str, dict[str, Any]] = {}
_handles: dict[str, dict[str, dict[str, Any]]] = {}

DEFAULT_TEMPLATE_ID = "default_plain"
SUPPORTED_BLUEPRINT_BLOCKS = {
    "paragraph",
    "heading",
    "list",
    "table",
    "image_placeholder",
    "page_break",
    "section_break",
}

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
    if normalized["size"] not in PAGE_SIZES:
        raise ValueError(f"Unsupported page size: {normalized['size']}. Use one of {sorted(PAGE_SIZES)}")
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

    raw_blocks = source.get("blocks") or []
    blocks = list(raw_blocks) if isinstance(raw_blocks, list) else raw_blocks
    if isinstance(blocks, list) and sections:
        for section in sections:
            if isinstance(section, dict):
                blocks.extend(section.get("blocks") or [])

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


def _inspect_blueprint_live(filename: str) -> dict[str, Any]:
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

    blocks = []
    for index in range(1, doc.Paragraphs.Count + 1):
        para = doc.Paragraphs(index)
        text = para.Range.Text.rstrip("\r\x07")
        style = str(para.Style) if para.Style else "Normal"
        block_type = "paragraph"
        level = None
        lowered = style.lower()
        if lowered.startswith("heading"):
            block_type = "heading"
            try:
                level = int(style.split()[-1])
            except Exception:
                level = 1
        block = {"type": block_type, "text": text, "style": style, "paragraph_index": index}
        if level is not None:
            block["level"] = level
        blocks.append(block)

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
                    "pages": doc.ComputeStatistics(2),
                    "words": doc.ComputeStatistics(0),
                    "paragraphs": doc.Paragraphs.Count,
                    "tables": doc.Tables.Count,
                    "inline_shapes": doc.InlineShapes.Count,
                    "shapes": doc.Shapes.Count,
                },
                "sections": [{"blocks": blocks}],
            },
        },
    }


def _blueprint_expected_counts(blueprint: dict | None) -> dict[str, int]:
    normalized = _normalize_blueprint(blueprint)
    paragraphs = 0
    tables = 0
    image_placeholders = 0
    if not isinstance(normalized["blocks"], list):
        return {"paragraphs": 0, "tables": 0, "image_placeholders": 0, "blocks": 0}
    for block in normalized["blocks"]:
        block_type = block.get("type")
        if block_type in {"paragraph", "heading", "image_placeholder"}:
            paragraphs += 1
            if block_type == "image_placeholder":
                image_placeholders += 1
        elif block_type == "list":
            paragraphs += len(block.get("items") or [])
        elif block_type == "table":
            tables += 1
        elif block_type in {"page_break", "section_break"}:
            pass
    return {
        "paragraphs": paragraphs,
        "tables": tables,
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
    return session.get("filename") or session.get("full_path")


def _touch_session(session_id: str = None, filename: str = None, full_path: str = None) -> None:
    if not session_id or session_id not in _sessions:
        return
    if filename:
        _sessions[session_id]["filename"] = filename
    if full_path:
        _sessions[session_id]["full_path"] = full_path
    _sessions[session_id]["updated_at"] = time.time()


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
            result["usage"] = (
                "Call word_v2_open() to attach to the active document, "
                "word_v2_open(action='attach', path='<name|full_path|index>') to attach a listed document, "
                "or word_v2_open(path='<file.docx>') to open a file."
            )
        return _dump(result)

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
            "valid_actions": ["open", "active", "attach", "list", "new"],
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

    Actions: text, page_text, info, comments, revisions, paragraph_format.
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
    else:
        return _dump({
            "error": "Invalid action",
            "valid_actions": ["text", "page_text", "info", "comments", "revisions", "paragraph_format"],
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
                result = _load_result(await live_tools.word_live_replace_text(
                    filename, find_text, replace_text or text, match_case,
                    whole_word, use_wildcards, replace_all, track_changes,
                ))
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
                result = _load_result(await live_tools.word_live_replace_text(
                    filename, find_text, "", match_case, whole_word,
                    use_wildcards, replace_all, track_changes,
                ))
            else:
                resolved = _resolve_target(session_id, handle, target, start, end)
                target_start, target_end = _target_range(resolved)
                result = _load_result(await live_tools.word_live_delete_text(
                    filename, target_start, target_end, track_changes,
                ))
        elif action == "insert_paragraphs":
            result = _load_result(await live_tools.word_live_insert_paragraphs(
                filename, paragraphs, find_text or None, paragraph_index,
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


async def word_v2_mutations(
    session_id: str,
    action: str,
    operations: list[dict] = None,
) -> str:
    """Preview or apply multiple v2 operations.

    Each operation is {"tool": "edit|format|comment|track_changes|table|layout", ...args}.
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
        tool = (operation.get("tool") or "").lower()
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
        elif tool == "layout":
            raw = await word_v2_layout(**args)
        else:
            raw = _dump({"error": f"Invalid operation tool at index {i}: {tool}"})
        result = _load_result(raw)
        results.append({"index": i, "tool": tool, "result": result})
        if result.get("error"):
            return _dump({"success": False, "session_id": session_id, "results": results})

    return _dump({"success": True, "session_id": session_id, "results": results})


async def word_v2_layout(
    session_id: str,
    action: str,
    page_size: str = "letter",
    orientation: str = "portrait",
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
    valid_actions = ["page_setup", "page_break", "section_break", "properties"]
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


async def _apply_blueprint_block(session_id: str, block: dict[str, Any]) -> dict[str, Any]:
    block_type = block.get("type")
    if block_type == "paragraph":
        style = block.get("style") or "Normal"
        raw = await word_v2_edit(
            session_id=session_id,
            action="insert_paragraphs",
            paragraphs=[str(block.get("text") or "")],
            position="end",
            style=style,
        )
        return {"block_type": block_type, "result": _load_result(raw)}

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
            "warning": "image_placeholder creates editable placeholder text; real image insertion will be added in word_v2_media.",
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
        result = _inspect_blueprint_live(filename)
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
