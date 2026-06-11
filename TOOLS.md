# Public Tool Reference

This branch exposes only grouped `word_*` MCP tools. Lower-level
`word_live_*` functions are implementation helpers and are not public MCP tools.

## Lifecycle

| Tool | Description |
| --- | --- |
| `word_open` | Open files or create new visible documents by default; list sessions/documents or explicitly attach to already-open documents. |
| `word_save` | Save the current session in place, save to a new file, or export PDF. |
| `word_close` | Close the Word document associated with a session. |

## Read And Search

| Tool | Description |
| --- | --- |
| `word_get_content` | Read text, page text, document info, comments, revisions, paragraph formatting, snapshots, or diffs. |
| `word_search` | Search text and return match handles for later edits, formatting, or comments. |

## Mutations

| Tool | Description |
| --- | --- |
| `word_edit` | Insert, replace, delete, insert paragraphs, undo, add hyperlinks, add footnotes, or delete footnotes. `insert_paragraphs` accepts strings or `{text, style}` objects. |
| `word_format` | Format inline text, paragraphs, styles, and lists. |
| `word_comment` | Create, reply to, resolve, delete, list, or get comments. |
| `word_track_changes` | Toggle, list, accept, reject, or decide tracked changes. |
| `word_table` | Create, inspect, edit, or format tables. |
| `word_media` | Insert images with sizing, alignment, wrapping, floating placement, and optional borders. |
| `word_layout` | Set page setup, insert page/section breaks, accept `break` as an alias, or set document properties. |
| `word_blueprint` | Create, inspect, validate, or export structured/page-aware document blueprints. |
| `word_protection` | Protect or unprotect a document. |
| `word_mutations` | Preview or apply multiple operations in order. Operation tools can be short names or public `word_*` names. |

## Compatibility Policy

Do not register duplicate public tools for the same behavior. If a lower-level
helper is needed by agents, add it as an action or parameter on the relevant
`word_*` grouped tool.
