---
name: word-mcp
description: "Create, edit, inspect, redline, comment on, save, close, or recreate Microsoft Word documents through the Word MCP server. Use the configured MCP tools, not Python document scripts."
version: 1.3.0
author: word-mcp-live
tags: [word, mcp, documents, office]
---

# Word MCP Live Skill

Use this skill for Microsoft Word document work when the `word` MCP server is
configured in Hermes or OpenClaw.

## Critical Naming Rule

Use the exact Word MCP tool names exposed by the current agent runtime. Do not
invent shell commands or Python scripts to work around tool naming.

Hermes prefixes MCP tools as `mcp_{server}_{tool}`. With the server named
`word`, callable Hermes tools look like this:

OpenClaw commonly exposes MCP tools with a server prefix such as
`word__word_v2_open`. If OpenClaw shows a different provider-safe name in its
tool list, use the visible OpenClaw name exactly.

| Intent | Hermes tool | Typical OpenClaw tool |
| --- | --- | --- |
| Open, list, attach, or create a document | `mcp_word_word_v2_open` | `word__word_v2_open` |
| Save or export | `mcp_word_word_v2_save` | `word__word_v2_save` |
| Close | `mcp_word_word_v2_close` | `word__word_v2_close` |
| Read text, info, comments, revisions, pages, formatting | `mcp_word_word_v2_get_content` | `word__word_v2_get_content` |
| Search text and get handles | `mcp_word_word_v2_search` | `word__word_v2_search` |
| Insert, replace, delete, insert paragraphs, links, footnotes | `mcp_word_word_v2_edit` | `word__word_v2_edit` |
| Format text, paragraphs, styles, or lists | `mcp_word_word_v2_format` | `word__word_v2_format` |
| Create, list, reply, resolve, or delete comments | `mcp_word_word_v2_comment` | `word__word_v2_comment` |
| Enable/list/accept/reject tracked changes | `mcp_word_word_v2_track_changes` | `word__word_v2_track_changes` |
| Create, inspect, edit, or format tables | `mcp_word_word_v2_table` | `word__word_v2_table` |
| Insert images/media | `mcp_word_word_v2_media` | `word__word_v2_media` |
| Page setup, page breaks, section breaks, document properties | `mcp_word_word_v2_layout` | `word__word_v2_layout` |
| Inspect, validate, export, or create blueprints | `mcp_word_word_v2_blueprint` | `word__word_v2_blueprint` |
| Protect or unprotect documents | `mcp_word_word_v2_protection` | `word__word_v2_protection` |
| Batch multiple operations | `mcp_word_word_v2_mutations` | `word__word_v2_mutations` |

Do not try to call bare names like `word_v2_open`. Do not use Codex-style
names like `mcp__word.word_v2_open`. Those names are not the normal Hermes or
OpenClaw tool names.

The JSON examples below are arguments for the named Word MCP tool, not shell
commands. When running in OpenClaw, call the OpenClaw tool name from the table
with the same JSON arguments.

## Do Not Use Terminal For MCP Calls

Do not run shell commands such as:

```bash
hermes mcp call word word_v2_open '{"action":"list"}'
```

Hermes has no `mcp call` command. `hermes mcp test word` only verifies that the
server can start and list tool schemas; it does not execute Word MCP tools.
OpenClaw MCP registry commands are also setup/diagnostic commands, not a
replacement for native tool calls inside an agent session.

To open/list Word documents, make a native agent tool call to the open/list tool
for the current client, for example `mcp_word_word_v2_open` in Hermes or
`word__word_v2_open` in OpenClaw, with:

```json
{"action": "list"}
```

If the Word MCP tools are not available in the current toolset, stop and report
that the session must be restarted or reloaded with the `word` MCP server
enabled. Do not use the terminal tool to work around missing MCP tools.

## Toolset Requirement

The `word` MCP server must be in the active toolset. If Hermes says a Word MCP
tool is "not available in this toolset", the session was likely started with a
restricted toolset such as only `hermes-cli`. The session needs the `word`
toolset, for example `word` or `hermes-cli,word`. In OpenClaw, start a new
session after installing or changing MCP/skill configuration.

Do not work around a missing Word MCP tool by writing Python. Report the
toolset problem and ask for the session to be restarted with the `word` toolset.

## Hard Rules

- Use the current client's `word_v2_*` MCP tools for Word document operations.
- Do not write or run Python, `python-docx`, `pywin32`, PowerShell COM, OOXML
  zip edits, or filesystem mutation scripts to inspect, edit, comment, redline,
  save, close, or recreate `.docx` files.
- Always open, attach, or list first with the Word open/list tool; carry the
  returned `session_id`.
- Inspect before structural edits on existing documents.
- Re-inspect after edits or blueprint creation before final save.
- Use handles returned by `mcp_word_word_v2_search` for targeted edits,
  formatting, and comments.
- Close documents when finished. Use `save_changes="save"` only after the
  result was inspected and looks correct.

Only use Python for ordinary repo maintenance or when the user explicitly asks
to debug or modify the Word MCP server implementation itself.

## Prerequisites

The Word MCP server uses Microsoft Word automation underneath. `hermes mcp test
word` or `openclaw mcp doctor word --probe` verifies that the MCP server starts
and exposes schemas; the real runtime check is a Word MCP operation such as
opening or listing Word documents.

If the Word open/list tool hangs or fails, report the Word/MCP runtime failure
and stop. Do not verify or edit the document with Python as a fallback.

## Open Or Attach

If the document is already open, call `mcp_word_word_v2_open` with:

```json
{"action": "list"}
```

If you lost the current `session_id`, call:

```json
{"action": "sessions"}
```

Then attach by index, name, or full path with `mcp_word_word_v2_open`:

```json
{"action": "attach", "path": "1"}
```

If the user gives a file path, call `mcp_word_word_v2_open` with:

```json
{"path": "C:\\Docs\\Example.docx", "read_only": false, "visible": false}
```

For a new document, call `mcp_word_word_v2_open` with:

```json
{"action": "new", "visible": false}
```

## Simple Edit With Tracked Changes And Comments

1. Search with `mcp_word_word_v2_search`.
2. Replace with `mcp_word_word_v2_edit` using the returned handle.
3. Comment with `mcp_word_word_v2_comment` using the same handle if needed.
4. Inspect with `mcp_word_word_v2_get_content`.
5. Save with `mcp_word_word_v2_save`.

Call `mcp_word_word_v2_search`:

```json
{"session_id": "word_abc", "find_text": "old text"}
```

Call `mcp_word_word_v2_edit`:

```json
{
  "session_id": "word_abc",
  "action": "replace",
  "handle": "match_1",
  "text": "new text",
  "track_changes": true
}
```

Call `mcp_word_word_v2_comment`:

```json
{
  "session_id": "word_abc",
  "action": "create",
  "handle": "match_1",
  "text": "Explain why this changed."
}
```

For many edits, use `mcp_word_word_v2_mutations` instead of one call per
change. Each operation uses the same arguments as its grouped v2 tool. The
`tool` field accepts short names such as `edit` or full public names such as
`word_v2_edit` and `mcp_word_word_v2_edit`:

```json
{
  "session_id": "word_abc",
  "action": "apply",
  "operations": [
    {
      "tool": "edit",
      "action": "replace",
      "handle": "match_1",
      "text": "new text",
      "track_changes": true
    },
    {
      "tool": "comment",
      "action": "create",
      "handle": "match_1",
      "text": "Why this changed."
    }
  ]
}
```

For before/after checking, call `mcp_word_word_v2_get_content` with
`{"action":"snapshot"}` before edits and `{"action":"diff"}` after edits.

## Structured Creation Or Recreation

Use `mcp_word_word_v2_blueprint` for structured document creation and faithful
recreation from a reference document.

For reference recreation:

1. Open the reference read-only with `mcp_word_word_v2_open`.
2. Inspect it with `mcp_word_word_v2_blueprint` and pass `asset_dir` when image
   assets are available.
3. Create the new document with `mcp_word_word_v2_blueprint`.
4. Inspect the new document.
5. Save and close both sessions.

Call `mcp_word_word_v2_blueprint` to inspect:

```json
{
  "action": "inspect",
  "session_id": "word_ref",
  "asset_dir": "C:\\Docs\\assets\\images"
}
```

Use the returned `session_blueprint.document` as the source for creation:

```json
{
  "action": "create",
  "blueprint": {"document": "...preserved inspected document object..."},
  "out": "C:\\Docs\\generated.docx",
  "visible": false
}
```

Preserve these fields exactly when replaying inspected blueprints:

- `path` and `asset_path`
- `width_pt`, `height_pt`
- `wrapping`, `wrap_type`
- `left_pt`, `top_pt`
- `relative_horizontal_position`, `relative_vertical_position`
- `anchor_paragraph_index`
- paragraph `style`
- paragraph `numbering`
- table `rows`, `row_count`, `col_count`, `style`
- page setup: `size`, `width`, `height`, `orientation`, `margins`

Do not simplify values that look strange. Word positioning values such as
`left_pt: -999995.0` can be valid and should be replayed as-is.

## Pitfalls And Limits

- **`find_text` is capped at 255 characters** in `mcp_word_word_v2_edit` with
  `action='replace'` because Word's native Find API has that limit. Do not
  blindly split a paragraph into tiny chunks. Search a short unique anchor and
  edit with the returned `handle`, use `start`/`end` from
  `mcp_word_word_v2_get_content(action='page_text')`, or replace a whole
  paragraph with `paragraph_index`.
- **Comments require a target**: provide a `handle`, `target`, `start`/`end`,
  or `paragraph_index`. A `handle` from `mcp_word_word_v2_search` is valid and
  is usually the easiest target.
- **Paragraph indexes in v2 are 1-based**: use the `index` returned by
  `mcp_word_word_v2_get_content`; paragraph 1 is the first paragraph.
- **`insert_paragraphs` accepts strings or objects**: use strings for simple
  paragraphs, or `{ "text": "Heading", "style": "Heading 1" }` objects when
  each paragraph needs its own style.
- **Layout breaks are explicit**: prefer `page_break` or `section_break`. The
  generic `break` alias also works; pass `break_type="page"` for a page break.
- **Tracked changes are explicit**: pass `track_changes=true` on edit/format
  calls that support it, or call `mcp_word_word_v2_track_changes` with
  `action='toggle', enable=true` before a review batch. Edits made without one
  of those are applied silently.
- **Handle reuse**: a `handle` from `mcp_word_word_v2_search` can be reused
  across `edit`, `comment`, and `format` calls in the same session, but
  re-search if the document state may have changed between calls.

## If A Tool Call Fails

1. Check whether the called function name starts with `mcp_word_word_v2_`.
2. If the error says the tool is not in the toolset, the active session needs
   the `word` toolset; do not switch to Python.
3. If the session is unknown, call `mcp_word_word_v2_open` with
   `{"action": "list"}` and attach again.
4. If a target is missing, inspect with `mcp_word_word_v2_get_content` or search
   with `mcp_word_word_v2_search`.
5. If the Word MCP server hangs or fails, report the MCP/Word issue instead of
   editing the document by script.
