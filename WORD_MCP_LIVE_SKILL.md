# Word MCP Live Skill

Use this skill when working with the Word MCP live COM tools. It is written for
local models that need strict tool-use patterns.

## Core Rules

- Always open, attach, or list first. Carry the returned `session_id`.
- Prefer `word_v2_*` tools only. Do not call lower-level `word_live_*` tools.
- Do not invent file paths. Use paths given by the user or returned by inspect.
- Inspect before making structural changes when the document already exists.
- Re-inspect after creation or edits before saving final work.
- Preserve layout metadata exactly when replaying a blueprint.
- Close documents when finished. Use `save_changes="save"` only when the result
  was inspected and looks correct.

## Open Or Attach

If the user says the document is already open:

```json
{"action": "list"}
```

Then attach by index, name, or full path:

```json
{"action": "attach", "path": "1"}
```

If the user gives a file path:

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

```json
{
  "session_id": "word_abc",
  "action": "replace",
  "handle": "match_1",
  "text": "new text",
  "track_changes": true
}
```

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
- Do not drop `numbering`; it is needed to replay real Word lists.
- Do not drop title-page `images` or `shapes`; they are needed for cover pages.
- Do not replace absolute Windows paths with guessed relative paths.
- Do not save if inspect shows obvious missing images, warnings, or bad page
  roles.
