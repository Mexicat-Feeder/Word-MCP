---
name: word-mcp
description: "Create, edit, inspect, redline, comment on, save, close, or recreate Microsoft Word documents using the connected Word MCP live tools. Always prefer word_v2_* MCP tool calls over Python, python-docx, pywin32, PowerShell, OOXML zip edits, or manual scripts for Word document work."
version: 1.1.0
author: hermes-agent
tags: [word, mcp, documents, office]
---

# Word MCP Live Skill

Use this skill for Word document tasks when the Word MCP server is available.
It is written for local models that need strict tool-use patterns.

## First Choice: Word MCP Tools

For document work, call the connected Word MCP tools directly. Do not write or
run Python to inspect, create, edit, redline, comment, save, close, or rebuild a
Word document.

Do not use Python, `python-docx`, `pywin32`, PowerShell COM automation, OOXML
zip editing, or filesystem document mutation except when:

- the user explicitly asks to debug or modify the Word MCP server itself;
- the Word MCP server/tools are unavailable and the user approves a fallback;
- the task is ordinary repo maintenance, such as running tests.

If the task mentions a `.docx`, an open Word document, comments, tracked
changes, layout, images, tables, styles, headers, footers, or document
recreation, start with `word_v2_open` or `word_v2_blueprint`; do not start by
creating a Python script.

When the client exposes namespaced tools, use the Word MCP namespace directly
such as `mcp__word.word_v2_open`, `mcp__word.word_v2_edit`, and
`mcp__word.word_v2_blueprint`. The JSON examples below are tool arguments, not
shell commands.

## Core Rules

- Always open, attach, or list first with `word_v2_open`. Carry the returned
  `session_id`.
- Use `word_v2_*` MCP tools only for document operations. Do not call
  lower-level `word_live_*` tools.
- Do not use Python/manual COM/OOXML scripts as a substitute for a missing or
  unfamiliar MCP call. Use the closest `word_v2_*` tool and inspect after.
- Do not invent file paths. Use paths given by the user or returned by inspect.
- Inspect before making structural changes when the document already exists.
- Re-inspect after creation or edits before saving final work.
- Preserve layout metadata exactly when replaying a blueprint.
- Close documents when finished. Use `save_changes="save"` only when the result
  was inspected and looks correct.

## Tool Selection

- Open, attach, list, or create a default blank document: `word_v2_open`.
- Save in place, save as, or export PDF: `word_v2_save`.
- Close a document: `word_v2_close`.
- Read text, page text, comments, revisions, document info, or paragraph
  formatting: `word_v2_get_content`.
- Find target text and get reusable handles: `word_v2_search`.
- Insert, replace, delete, insert paragraphs, hyperlinks, or footnotes:
  `word_v2_edit`.
- Apply text formatting, paragraph formatting, styles, or lists:
  `word_v2_format`.
- Create, list, reply to, resolve, or delete comments: `word_v2_comment`.
- Enable, list, accept, reject, or decide tracked changes:
  `word_v2_track_changes`.
- Create, inspect, edit, or format tables: `word_v2_table`.
- Insert images/media with size, wrapping, and placement: `word_v2_media`.
- Page setup, breaks, and document properties: `word_v2_layout`.
- Inspect, validate, export, or create structured page-aware documents:
  `word_v2_blueprint`.
- Protect or unprotect documents: `word_v2_protection`.
- Batch multiple edits in order: `word_v2_mutations`.

## Open Or Attach

If the user says the document is already open:

Call `word_v2_open` with:

```json
{"action": "list"}
```

Then attach by index, name, or full path:

Call `word_v2_open` with:

```json
{"action": "attach", "path": "1"}
```

If the user gives a file path:

Call `word_v2_open` with:

```json
{"path": "C:\\Docs\\Example.docx", "read_only": false, "visible": false}
```

## Reference Recreation Recipe

Use this when the user provides a reference document and image assets.

1. Open the reference read-only.
2. Inspect with `asset_dir`.
3. Create from the inspected blueprint.
4. Inspect the new document.
5. Save and close both sessions.

```json
{
  "path": "C:\\Docs\\reference.docx",
  "read_only": true,
  "visible": false
}
```

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

## Preserve These Fields

When copying image, shape, paragraph, table, or page blocks from an inspected
blueprint, preserve these fields exactly:

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
`left_pt: -999995.0` are valid and should be replayed as-is.

## Asset Rules

- Put title-page picture assets in the same `asset_dir` as body images.
- Title-page asset filenames should contain `title`, for example:
  - `Title Page pic 1.png`
  - `Title page pic 2 logo.png`
- Body screenshots should not contain `title` in the filename.
- Inspect maps title assets separately from body assets.

Expected inspect fields:

```json
{
  "reference": {
    "asset_count": 23,
    "title_asset_count": 4,
    "body_asset_count": 19
  }
}
```

## Common Workflows

### Make Simple Edits With Comments

1. Search for text.
2. Use returned `handle`.
3. Replace or comment using the handle.
4. Inspect or get content.
5. Save.

```json
{"session_id": "word_abc", "find_text": "old text"}
```

Call `word_v2_edit` with:

```json
{
  "session_id": "word_abc",
  "action": "replace",
  "handle": "match_1",
  "text": "new text",
  "track_changes": true
}
```

Call `word_v2_comment` with:

```json
{
  "session_id": "word_abc",
  "action": "create",
  "handle": "match_1",
  "text": "Explain why this changed."
}
```

### Create A New Structured Document

Use `word_v2_blueprint(action="create")` for structured output.

```json
{
  "action": "create",
  "visible": false,
  "out": "C:\\Docs\\new.docx",
  "blueprint": {
    "document": {
      "page_setup": {
        "size": "custom",
        "width": 612,
        "height": 792,
        "orientation": "portrait",
        "margins": {"top": 72, "bottom": 72, "left": 54, "right": 54}
      },
      "blocks": [
        {"type": "title_page", "text": "Document Title", "page_break_after": true},
        {"type": "toc", "title": "Table of Contents", "levels": 3},
        {"type": "heading", "text": "1. Overview", "level": 1},
        {"type": "paragraph", "text": "Body text.", "style": "Normal"}
      ]
    }
  }
}
```

## Warning Meanings

- `warnings: []`: no known structural warning.
- `title_page contains pictures...without mapped assets`: inspect or recreate
  was missing external title image paths. Provide `asset_dir` with title assets.
- Asset count mismatch: the asset folder does not match the reference image
  count. Check filenames and ensure title images contain `title`.
- Unknown session: open or attach again and use the new `session_id`.

## Local Model Failure Modes

- Do not manually retype large blueprints. Copy returned blocks forward.
- Do not call Python because you need to read or edit a `.docx`; use
  `word_v2_get_content`, `word_v2_search`, `word_v2_edit`, or
  `word_v2_blueprint`.
- Do not create a temporary script to add comments or tracked changes; use
  `word_v2_comment` and `word_v2_track_changes`.
- Do not drop `numbering`; it is needed to replay real Word lists.
- Do not drop title-page `images` or `shapes`; they are needed for cover pages.
- Do not replace absolute Windows paths with guessed relative paths.
- Do not save if inspect shows obvious missing images, warnings, or bad page
  roles.

## If A Tool Call Fails

1. Read the error and retry with corrected parameters.
2. If the session is unknown, call `word_v2_open(action="list")` and attach
   again.
3. If a target is missing, inspect with `word_v2_get_content` or search with
   `word_v2_search`.
4. Do not switch to Python unless the user agrees the MCP path is unavailable.
