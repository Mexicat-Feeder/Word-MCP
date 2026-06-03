"""V2 live Word tools: grouped, agent-friendly facade over COM helpers.

The v2 surface is inspired by SuperDoc's MCP shape: a few lifecycle tools plus
grouped intent tools. It keeps the existing live COM implementations underneath
so v2 can improve tool ergonomics without rewriting Word automation logic.
"""

from __future__ import annotations

import json
import time
import uuid
from typing import Any

from word_document_server.defaults import DEFAULT_AUTHOR
from word_document_server.tools import live_read_tools, live_tools


_sessions: dict[str, dict[str, Any]] = {}
_handles: dict[str, dict[str, dict[str, Any]]] = {}


def _load_result(raw: str | dict) -> dict[str, Any]:
    if isinstance(raw, dict):
        return raw
    try:
        return json.loads(raw)
    except Exception:
        return {"success": False, "raw": raw}


def _dump(result: dict[str, Any]) -> str:
    return json.dumps(result, ensure_ascii=False)


def _new_id(prefix: str) -> str:
    return f"{prefix}_{uuid.uuid4().hex[:12]}"


def _register_session(open_result: dict[str, Any]) -> str:
    session_id = _new_id("word")
    filename = open_result.get("document") or open_result.get("name")
    _sessions[session_id] = {
        "session_id": session_id,
        "filename": filename,
        "full_path": open_result.get("full_path"),
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

    Each operation is {"tool": "edit|format|comment|track_changes|table", ...args}.
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
        else:
            raw = _dump({"error": f"Invalid operation tool at index {i}: {tool}"})
        result = _load_result(raw)
        results.append({"index": i, "tool": tool, "result": result})
        if result.get("error"):
            return _dump({"success": False, "session_id": session_id, "results": results})

    return _dump({"success": True, "session_id": session_id, "results": results})


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
