# Public Tool Reference

This branch exposes only grouped `word_v2_*` MCP tools. Lower-level
`word_live_*` functions are implementation helpers and are not public MCP tools.

## Lifecycle

| Tool | Description |
| --- | --- |
| `word_v2_open` | List open documents, list MCP sessions, attach to active/listed documents, open files, or create default-template sessions. |
| `word_v2_save` | Save the current session in place, save to a new file, or export PDF. |
| `word_v2_close` | Close the Word document associated with a session. |

## Read And Search

| Tool | Description |
| --- | --- |
| `word_v2_get_content` | Read text, page text, document info, comments, revisions, paragraph formatting, snapshots, or diffs. |
| `word_v2_search` | Search text and return match handles for later edits, formatting, or comments. |

## Mutations

| Tool | Description |
| --- | --- |
| `word_v2_edit` | Insert, replace, delete, insert paragraphs, undo, add hyperlinks, add footnotes, or delete footnotes. |
| `word_v2_format` | Format inline text, paragraphs, styles, and lists. |
| `word_v2_comment` | Create, reply to, resolve, delete, list, or get comments. |
| `word_v2_track_changes` | Toggle, list, accept, reject, or decide tracked changes. |
| `word_v2_table` | Create, inspect, edit, or format tables. |
| `word_v2_media` | Insert images with sizing, alignment, wrapping, floating placement, and optional borders. |
| `word_v2_layout` | Set page setup, insert page/section breaks, or set document properties. |
| `word_v2_blueprint` | Create, inspect, validate, or export structured/page-aware document blueprints. |
| `word_v2_protection` | Protect or unprotect a document. |
| `word_v2_mutations` | Preview or apply multiple v2 operations in order. |

## Compatibility Policy

Do not register duplicate public tools for the same behavior. If a lower-level
helper is needed by agents, add it as an action or parameter on the relevant
`word_v2_*` grouped tool.
