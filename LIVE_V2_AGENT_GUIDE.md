# Live V2 Agent Guide

The `word_v2_*` tools are the preferred interface for agents that edit documents
through Microsoft Word COM. They are intentionally grouped by intent so local
models have fewer tool names to choose from.

The older `word_live_*` tools still exist for compatibility and advanced escape
hatches. Prefer v2 first.

## Core Workflow

1. `word_v2_open(path="contract.docx")`
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

Opens a document in Word and returns a `session_id`.

Use `path` for the document filename/path. Relative paths resolve from the
server working directory or the provided `directory`.

### `word_v2_get_content`

Reads from the open session.

Actions:

- `text`: paragraph text.
- `page_text`: text for page ranges, including character offsets.
- `info`: document metadata.
- `comments`: all comments.
- `revisions`: tracked changes.
- `paragraph_format`: paragraph/run formatting diagnostics.

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

Use `track_changes=true` for suggested edits.

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

Use `comment_text` for comment bodies.

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

### `word_v2_mutations`

Use `action="preview"` to inspect a batch shape without changing the document.
Use `action="apply"` to execute operations in order.

Example:

```json
{
  "session_id": "word_abc123",
  "action": "apply",
  "operations": [
    { "tool": "edit", "action": "replace", "handle": "match_1", "text": "NewCo Inc.", "track_changes": true },
    { "tool": "comment", "action": "create", "handle": "match_1", "comment_text": "Updated party name." }
  ]
}
```

## Agent Rules

- Always open first and carry `session_id`.
- Search before editing unless the user gives exact `start/end`.
- Prefer `handle` over offsets.
- Re-search after edits because offsets can move.
- Use `track_changes=true` for user-reviewable edits.
- Save only after validation.
- Close with `save_changes="discard"` if validation shows a bad edit and undo
  is not enough.
