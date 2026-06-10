# Live V2 Agent Guide

The `word_v2_*` tools are the only public MCP interface for agents that edit
documents through Microsoft Word COM. They are intentionally grouped by intent
so local models have fewer tool names to choose from.

Lower-level `word_live_*` helpers still exist inside the codebase as
implementation details, but they are not registered as MCP tools.

## Core Workflow

1. `word_v2_open()` to attach to the active Word document, or
   `word_v2_open(path="contract.docx")` to open a file.
2. Keep the returned `session_id`.
3. Read context with `word_v2_get_content(session_id, action="info")` or
   `word_v2_get_content(session_id, action="text")`.
4. Find edit targets with `word_v2_search(session_id, find_text="...")`.
5. Use returned `handle` values for edits, formatting, or comments.
6. Validate with `word_v2_get_content` or `word_v2_search`.
7. Save with `word_v2_save(session_id)`.
8. Close with `word_v2_close(session_id, save_changes="save")`.

## Tools

### `word_v2_open`

Lists, attaches, opens, or creates Word documents.

Common calls:

- `word_v2_open()` attaches to the active Word document and returns `session_id`.
- `word_v2_open(action="list")` lists open Word documents and current MCP
  sessions without creating a new session.
- `word_v2_open(action="sessions")` lists current MCP `session_id` values when
  the agent lost track of them.
- `word_v2_open(action="attach", path="2")` attaches to a listed document by
  `index`, `name`, or `full_path`.
- `word_v2_open(path="contract.docx")` opens a file. Relative paths resolve
  from the server working directory or the provided `directory`.
- `word_v2_open(action="new")` creates a new document from the built-in
  `default_plain` template profile.

Do not call `word_v2_open()` when you intend to create a new document. Use
`action="new"` so agents do not accidentally ignore a user's active document.

### `word_v2_get_content`

Reads from the open session.

Actions:

- `text`: paragraph text.
- `page_text`: text for page ranges, including character offsets.
- `info`: document metadata.
- `comments`: all comments.
- `revisions`: tracked changes.
- `paragraph_format`: paragraph/run formatting diagnostics.
- `snapshot`: stores a baseline for later diffing.
- `diff`: returns paragraphs changed since the last snapshot.

### `word_v2_search`

Finds text and returns mutation-ready handles:

```json
{
  "handle": "match_1",
  "target": { "kind": "selection", "start": 125, "end": 134 },
  "text": "ACME Corp",
  "context": "..."
}
```

Prefer passing `handle` to later tools instead of manually copying `start` and
`end`.

### `word_v2_edit`

Actions:

- `insert`: insert text at `position`, `start`, `end`, `target`, or `handle`.
- `replace`: replace `find_text` or a `handle`/`target`.
- `delete`: delete `find_text` or a `handle`/`target`.
- `insert_paragraphs`: insert paragraph list at start/end or near text/index.
- `undo`: undo recent Word operations.

For `insert_paragraphs`, `paragraphs` may be a list of strings or a list of
objects such as `{ "text": "Heading", "style": "Heading 1" }`.

Use `track_changes=true` for suggested edits. Public v2 `paragraph_index`
values are 1-based, matching `word_v2_get_content` output.

For long grammar edits, avoid `find_text` strings over 255 characters because
Word's native Find API rejects them. Use one of these patterns instead:

- Search a short unique anchor, then edit the returned `handle`.
- Replace a whole paragraph with
  `word_v2_edit(action="replace", paragraph_index=3, text="...", track_changes=true)`.
- Use `word_v2_get_content(action="page_text")` to get `start`/`end` offsets.

### `word_v2_format`

Actions:

- `inline`: bold, italic, underline, font, color, highlight.
- `paragraph`: paragraph alignment/page-break formatting.
- `style`: apply a Word style.
- `list`: apply or remove list formatting over paragraph ranges.

For text ranges, pass a `handle` from `word_v2_search`.

### `word_v2_comment`

Actions:

- `create`: add a comment on a `handle`, `target`, `start/end`, or paragraph.
- `reply`: reply to a comment thread.
- `resolve`: resolve/unresolve a comment if Word supports it.
- `delete`: delete a comment.
- `list`: list comments.
- `get`: get one comment by `comment_index`.

Use `text` for comment bodies.

### `word_v2_track_changes`

Actions:

- `toggle`: set or toggle Track Changes mode.
- `list`: list revisions.
- `accept`: accept revisions.
- `reject`: reject revisions.
- `decide`: accept/reject based on `decision`.

Revision IDs are volatile. Call `list` immediately before accepting or
rejecting specific IDs.

### `word_v2_table`

Actions:

- `create`: add a table.
- `format`: table borders, alignment, shading, column widths.
- `get_info`, `set_cell`, `set_row`, `set_range`, `add_row`, `delete_row`,
  `add_column`, `delete_column`, `merge_cells`, `autofit`, `delete_table`.

Table row and column indexes follow the underlying live COM table tools:
generally 1-based for live table modification.
For `delete_column`, pass `col` with the 1-based column number.

### `word_v2_media`

Actions:

- `insert_image`: insert a PNG/JPG/BMP/etc. into the live document.

Use `path` for the image file path. Use `position="end"` for append-style
generation, or `paragraph_index` to insert before a specific paragraph. Width
can be controlled with `width_pt` or `width_inches`; if only width or height is
provided, Word keeps the image aspect ratio.
By default `paragraph_after=true`, so the next block starts below the image.

Common example:

```json
{
  "session_id": "word_abc123",
  "action": "insert_image",
  "path": "C:\\Docs\\assets\\images\\2.1 Initial Setup pic 1.png",
  "position": "end",
  "width_pt": 504,
  "alignment": "center",
  "wrapping": "inline"
}
```

For floating images copied from an inspected blueprint, preserve placement
fields instead of simplifying them. Use `wrapping` or `wrap_type`, plus
`left_pt`, `top_pt`, `relative_horizontal_position`, and
`relative_vertical_position` when they are present. Values such as
`left_pt: -999995.0` are valid Word positioning values and should be replayed
as-is.

### `word_v2_layout`

Actions:

- `page_setup`: set page size, orientation, and margins.
- `page_break`: insert a manual page break.
- `section_break`: insert a section break.
- `break`: compatibility alias. Use `break_type="page"` for a page break, or
  a Word section break type such as `next_page`, `continuous`, `even_page`, or
  `odd_page` for a section break.
- `properties`: set document properties such as title, subject, author, or company.

Use this when document structure matters more than text edits alone.

### `word_v2_blueprint`

Actions:

- `create`: create a new `default_plain` document from structured blocks.
- `inspect`: inspect an open document into a page-aware blueprint.
- `validate`: compare a blueprint's expected structure against a live session.
- `export`: return the current inspected blueprint JSON.

For reference-driven recreation, call `inspect` with `asset_dir` pointing to the
folder of extracted screenshots. The inspector returns `document.pages`,
`document.blocks`, `document.images`, `document.tables`, and `document.shapes`.
Page roles distinguish `title_page`, `toc`, and `body`; body images include
exact `width_pt`/`height_pt`, mapped `path` values when assets are available,
and floating placement metadata such as `left_pt`, `top_pt`, `wrap_type`, and
relative positioning constants.

Title-page blocks can include `shapes`. Creation replays simple text boxes and
rectangles from those records, including position, size, fill/line visibility,
and basic text formatting. Title-page pictures still need mapped image assets;
without them, recreation remains approximate.
Put title-page picture assets in the same `asset_dir` with `title` in the
filename, such as `Title Page pic 1.png`, `Title page pic 2 logo.png`, etc. The
inspector maps those title assets to page-1 picture shapes separately from body
screenshots, then replays them as positioned floating images.

Supported first-pass block types:

- `title_page`: `{ "type": "title_page", "text": "...", "page_break_after": true }`
- shaped `title_page`: `{ "type": "title_page", "shapes": [{ "name": "Text Box 1", "text": "Title", "left_pt": 24, "top_pt": 48, "width_pt": 420, "height_pt": 90 }] }`
- `toc`: `{ "type": "toc", "title": "Table of Contents", "levels": 3 }`
- `heading`: `{ "type": "heading", "text": "...", "level": 1 }`
- `paragraph`: `{ "type": "paragraph", "text": "...", "style": "Normal" }`
- `list`: `{ "type": "list", "items": ["One", "Two"], "ordered": false }`
- `table`: `{ "type": "table", "rows": [["Metric", "Value"]] }`
- `image`: `{ "type": "image", "path": "C:\\Docs\\screen.png", "width_pt": 504, "alignment": "center" }`
- floating `image`: `{ "type": "image", "path": "C:\\Docs\\screen.png", "width_pt": 396, "height_pt": 231.1, "wrapping": "infront", "left_pt": -999995.0, "top_pt": 55.5, "relative_horizontal_position": 0, "relative_vertical_position": 2 }`
- `page_break`: `{ "type": "page_break" }`
- `section_break`: `{ "type": "section_break", "break_type": "next_page" }`
- `image_placeholder`: editable placeholder text when the asset is unavailable.

Use blueprint mode for high-fidelity generation. Use simple edit/format/table
tools for quick spontaneous documents.

When replaying an inspected blueprint, keep paragraph `numbering` objects. The
creator uses them to reapply Word list formatting to paragraphs that were real
bullets or numbered list items in the reference.

### `word_v2_mutations`

Use `action="preview"` to inspect a batch shape without changing the document.
Use `action="apply"` to execute operations in order.
Operation `tool` may use short names (`edit`, `comment`, `layout`) or public
tool names (`word_v2_edit`, `mcp_word_word_v2_edit`).

Example:

```json
{
  "session_id": "word_abc123",
  "action": "apply",
  "operations": [
    { "tool": "edit", "action": "replace", "handle": "match_1", "text": "NewCo Inc.", "track_changes": true },
    { "tool": "comment", "action": "create", "handle": "match_1", "text": "Updated party name." }
  ]
}
```

For review workflows, call `word_v2_get_content` to take a snapshot before
edits, apply a batch, then call `word_v2_get_content` again to inspect the
diff:

```json
{ "session_id": "word_abc123", "action": "snapshot" }
```

```json
{ "session_id": "word_abc123", "action": "diff" }
```

## Agent Rules

- Always open or attach first and carry `session_id`.
- If the user says a document is already open, start with `word_v2_open()` or
  `word_v2_open(action="list")`; do not create a blank document.
- If the user asks for a new document, use `word_v2_open(action="new")` or
  `word_v2_blueprint(action="create")`.
- Search before editing unless the user gives exact `start/end`.
- Prefer `handle` over offsets.
- Re-search after edits because offsets can move.
- Use `word_v2_mutations` for batches of edits/comments instead of one tool
  call per change.
- Use `track_changes=true` for user-reviewable edits.
- Save only after validation.
- Close with `save_changes="discard"` if validation shows a bad edit and undo
  is not enough.
