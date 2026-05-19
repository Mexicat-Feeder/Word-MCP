# Live Tools Agent Guide

This guide is written for LLM agents using the `word_live_*` tools. It focuses
on the operational details that matter when editing real Word documents through
Word COM/JXA instead of manipulating `.docx` files offline.

## Mental Model

- File-based tools edit closed `.docx` files with python-docx.
- Live tools edit a document that is open in Microsoft Word.
- Live tools can preserve Word behavior: tracked changes, threaded comments,
  fields, tables, page layout, screenshots, and per-tool undo records.
- Each live mutation is normally wrapped as one Word undo entry named `MCP: ...`.

Prefer live tools whenever the document is open in Word, when the user needs
tracked changes/comments, or when layout/fields/tables must be handled by Word
itself.

## Safe Default Workflow

1. Call `word_live_list_open`.
2. If the target document is not open, call `word_live_open_document`.
3. Call `word_live_get_text`, `word_live_get_info`, or `word_live_get_page_text`
   to understand the current document.
4. Use `word_live_find_text` or page-text character offsets before range edits.
5. Make one focused mutation per tool call.
6. Validate with a read tool.
7. Save with `word_live_save`.
8. Close with `word_live_close_document` when done.

For long documents, use `word_live_get_page_text` instead of reading the entire
document. It returns `char_start` and `char_end` offsets that can be passed to
formatting and delete tools.

## Parameter Conventions

Agents should prefer these names across tools:

- `start` and `end` for character offsets.
- `paragraph_index` for a paragraph number. Check each tool description for
  whether it is 0-based or 1-based.
- `row` and `col` for a single table cell.
- `find_text` and `replace_text` for search/replace operations.
- `target_text` for text used as an anchor for comments, footnotes, or nearby
  insertion.
- `comment_text` for the body of a comment or comment reply.
- `position` for insertion placement such as `start`, `end`, `before`, or
  `after`.

Avoid older names such as `start_pos`, `end_pos`, `row_index`, `col_index`,
`text_to_find`, `text_content`, `target_paragraph_index`, `search_text`, and
`insert_position` when calling MCP tools. Those names may still appear inside
implementation code or older examples, but the public tool schema uses the
simpler names above.

## File Paths And Roots

- Relative filenames resolve from the server process working directory.
- The server commonly runs from a document folder. In that setup, relative names
  like `contract.docx` are the most reliable and user-friendly.
- Absolute paths are supported when the server process can see that path.
- A client machine path is not automatically visible to a remote/server-side
  Word process. If client-side file access is added later, treat it as an
  explicit import/export or allowlisted-root feature, not as arbitrary local path
  access.
- After `word_live_save(save_as=...)`, the open document name changes to the
  saved-as basename. Use the new filename for later live calls.

## Open, Save, Close

Use:

- `word_live_open_document(filename, visible=true)` to open a target document.
- `word_live_save(filename)` to save in place.
- `word_live_save(filename, save_as="new_name.docx")` to save as a new document.
- `word_live_close_document(filename, save_changes="yes")` to save and close.
- `word_live_close_document(filename, save_changes="no")` to discard and close.

Quirks:

- After save-as, calls using the old filename will fail because Word now has the
  new document open.
- If close returns an error, call `word_live_list_open` before retrying. Word may
  already have closed the document despite a stale COM error.

## Text Editing

Use:

- `word_live_insert_text` for simple insertion at start, end, cursor, bookmark,
  or character offset.
- `word_live_insert_paragraphs` for paragraph-level insertion near a target
  paragraph.
- `word_live_replace_text` for find/replace, especially across tracked-change
  boundaries.
- `word_live_delete_text` for deleting a character range.
- `word_live_format_text` for formatting without changing text content.

Quirks:

- `word_live_replace_text` is case-insensitive by default. Set
  `match_case=true` and/or `match_whole_word=true` when replacing marker text.
- Do not use `replace_all=true` with `track_changes=true`; the tool rejects this
  because tracked deletions remain findable and can cause replacement loops.
- Character offsets move after edits. Re-find text or re-read page text after
  each mutation if you need another range operation.
- Control bytes are rejected. Do not try to insert or search for Word table cell
  separator bytes.
- Use `word_live_format_text` for visual changes; tracked text tools only change
  text content.

## Tracked Changes

Use:

- `word_live_toggle_track_changes(enable=true/false)` to control document state.
- Most mutation tools also have `track_changes`.
- `word_live_list_revisions` to inspect revision IDs.
- `word_live_accept_revisions` and `word_live_reject_revisions` to resolve
  changes.

Quirks:

- Revision IDs are volatile. Accepting or rejecting one revision can renumber the
  remaining revisions.
- If you use specific revision IDs, call `word_live_list_revisions` immediately
  before accept/reject.
- Accept/reject responses include `requested` and `missing_ids` for stale IDs.
  If `missing_ids` is not empty, re-list revisions before continuing.
- Word may coalesce adjacent tracked insertions into one revision. Do not assume
  one tool call equals one revision ID.
- Author names for tracked revisions are controlled by Word's current user
  identity and `MCP_AUTHOR`; Microsoft 365 may still normalize identities.

## Comments

Use:

- `word_live_add_comment` with `start`/`end` or `paragraph_index`.
- `word_live_reply_to_comment` for threaded replies.
- `word_live_get_comments` to inspect top-level comments and nested replies.
- `word_live_delete_comment` to delete a comment. If the comment has replies,
  the tool deletes the full thread recursively.

Quirks:

- `word_live_get_comments` reports top-level comments only in `comments`; replies
  are nested under `replies`.
- `word_live_add_comment` and `word_live_reply_to_comment` return
  `actual_author` and `author_applied`. If `author_applied=false`, Word forced a
  different Office identity.
- `word_live_resolve_comment` is limited by Word Modern Comments. On Microsoft
  365 it may return a known COM limitation instead of resolving the thread.
- File-based `get_all_comments` may list replies as separate comments because it
  reads raw OOXML. Prefer live `word_live_get_comments` for active Word threads.

## Tables

Use:

- `word_live_add_table` to insert a table.
- `word_live_modify_table` for data and structure:
  `get_info`, `set_cell`, `set_row`, `set_range`, `add_row`, `delete_row`,
  `add_column`, `delete_column`, `merge_cells`, `autofit`, `delete_table`.
- `word_live_format_table` for borders, alignment, bold cells, shading, column
  widths, and autofit.

Quirks:

- Table row/column indexes are 1-based.
- `word_live_format_table` expects real nested arrays, for example:
  `cell_bold=[[1, 1, true]]`,
  `cell_alignment=[[0, 0, "center"]]`,
  `cell_shading=[[1, 0, "#DDDDDD"]]`.
- Do not pass stringified list literals when the MCP schema advertises nested
  arrays. Some clients reject them before the tool runs.
- After merging cells, Word may report mixed cell widths. Some column operations
  can fail on mixed-width tables. Use table info and prefer row/cell operations
  after merges.
- `delete_table` has orphan-scrubbing behavior for Word cell separator bytes.

## Layout, Fields, References

Use:

- `word_live_set_page_layout` for margins, orientation, and page size.
- `word_live_add_page_numbers` for page fields.
- `word_live_add_watermark` for a text watermark.
- `word_live_list_cross_reference_items` before
  `word_live_insert_cross_reference`.
- `word_live_insert_equation` for UnicodeMath equations.
- `word_live_apply_list` for bullet/numbered/multilevel lists.
- `word_live_setup_heading_numbering` for numbered Heading 1/Heading 2 styles.

Quirks:

- Cross-reference target indexes come from
  `word_live_list_cross_reference_items`; do not invent them.
- `word_live_setup_heading_numbering` paragraph numbers are 1-based. String
  numbers such as `"1"` are coerced, but numeric arrays are preferred.
- Applying styles or fields can alter paragraph ranges. Re-read formatting or
  page text before making range-sensitive follow-up edits.

## Screenshots And Images

Use:

- `word_screen_capture` to capture the active Word window.
- `word_live_insert_image` to insert PNG/JPG/BMP images.

Quirks:

- On Windows, screenshot capture requires Pillow (`PIL`) and pywin32.
- `word_screen_capture` captures the Word window, not just the page content.
- `word_live_insert_image` needs an image path visible to the server process.
- For non-inline wrapping, Word converts the image into a floating shape. Inline
  wrapping is usually safer for automated document edits.

## Undo And Recovery

Use:

- `word_live_get_undo_history` to inspect Word undo entries.
- `word_live_undo(times=N)` to undo the last N live operations.

Quirks:

- One MCP mutation is usually one undo entry, but Word can group internal edits.
- If a tool returns an error after a COM operation, inspect the document before
  retrying. Word sometimes completes part of an operation before COM reports an
  error.
- Prefer saving only after validation. If a document is in a bad state, use undo
  or close with `save_changes="no"`.

## Recommended Agent Patterns

For a tracked edit:

1. `word_live_find_text`
2. `word_live_replace_text(..., match_whole_word=true, track_changes=true)`
3. `word_live_list_revisions`
4. `word_live_get_text` or `word_live_find_text` to verify result

For a comment:

1. `word_live_find_text`
2. `word_live_add_comment(start=..., end=..., comment_text=..., author=...)`
3. `word_live_get_comments`

For table updates:

1. `word_live_modify_table(operation="get_info")`
2. `word_live_modify_table(operation="set_cell" | "set_row" | "set_range")`
3. `word_live_format_table(...)`
4. `word_live_modify_table(operation="get_info")`

For finalization:

1. `word_live_list_revisions`
2. `word_live_get_comments`
3. `word_live_get_info`
4. `word_live_save`
5. `word_live_close_document`
6. `word_live_list_open`

## Common Failure Responses

- `Document 'x.docx' is not open`: call `word_live_list_open`; if the document
  was saved-as, use the new basename.
- `No module named 'PIL'`: install Pillow in the server environment.
- `Comment.Done is not available`: Word Modern Comments limitation; treat
  resolve/unresolve as unsupported for that environment.
- `missing_ids` on revision accept/reject: revision IDs changed; re-list and
  retry with current IDs.
- `Cannot access individual columns ... mixed cell widths`: table has merged or
  irregular cells; avoid column-level operations or normalize the table first.
