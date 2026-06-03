# Word MCP Live

Live Microsoft Word editing through a small, agent-friendly MCP interface.

This dev branch exposes only grouped `word_v2_*` MCP tools. The lower-level
`word_live_*` modules remain in the codebase as internal COM/JXA helpers, but
they are not registered as public tool calls. This keeps local models from
choosing between duplicate ways to open, read, edit, save, or close documents.

## Public Tool Surface

The server registers 12 public MCP tools:

| Tool | Purpose |
| --- | --- |
| `word_v2_open` | Open or create a live Word session and return `session_id`. |
| `word_v2_save` | Save in place, save as another document, or export PDF. |
| `word_v2_close` | Close a live session. |
| `word_v2_get_content` | Read text, page text, info, comments, revisions, or paragraph formatting. |
| `word_v2_search` | Find text and return reusable match handles. |
| `word_v2_edit` | Insert, replace, delete, insert paragraphs, undo, add hyperlinks, and manage footnotes. |
| `word_v2_format` | Apply inline formatting, paragraph formatting, styles, and lists. |
| `word_v2_comment` | Create, reply to, resolve, delete, list, or get comments. |
| `word_v2_track_changes` | Toggle, list, accept, reject, or decide revisions. |
| `word_v2_table` | Create, inspect, edit, and format tables. |
| `word_v2_mutations` | Preview or apply grouped v2 operations in order. |
| `word_v2_protection` | Protect or unprotect the active document. |

See [LIVE_V2_AGENT_GUIDE.md](LIVE_V2_AGENT_GUIDE.md) for agent workflow and
quirks.

## Requirements

- Python 3.11+
- Microsoft Word for live editing
- Windows for COM automation, or macOS for the subset supported by JXA

## Install From Source

```powershell
git clone https://github.com/Mexicat-Feeder/Word-MCP.git
cd Word-MCP
uv sync
```

## Run With Stdio

```powershell
uv run python -m word_document_server.main
```

For MCP clients, use `uv` as the command and split arguments by token:

```json
{
  "mcpServers": {
    "word": {
      "command": "uv",
      "args": [
        "--directory",
        "C:/path/to/Word-MCP",
        "run",
        "python",
        "-m",
        "word_document_server.main"
      ],
      "env": {
        "MCP_TRANSPORT": "stdio",
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

## Run With HTTP

```powershell
$env:MCP_TRANSPORT = "streamable-http"
$env:MCP_HOST = "0.0.0.0"
$env:MCP_PORT = "8000"
uv run python -m word_document_server.main
```

## Configuration

| Variable | Default | Description |
| --- | --- | --- |
| `MCP_AUTHOR` | `Author` | Author name for tracked changes and comments. |
| `MCP_AUTHOR_INITIALS` | empty | Initials for comments where Word supports them. |
| `MCP_TRANSPORT` | `stdio` | `stdio`, `streamable-http`, or `sse`. |
| `MCP_HOST` | `0.0.0.0` | Host for HTTP/SSE transports. |
| `MCP_PORT` | `8000` | Port for HTTP/SSE transports. |

## Typical Agent Flow

1. Call `word_v2_open`.
2. Keep the returned `session_id`.
3. Inspect with `word_v2_get_content`.
4. Locate text with `word_v2_search`.
5. Edit, format, comment, or manage revisions using the same `session_id`.
6. Save with `word_v2_save`.
7. Close with `word_v2_close`.

## Development

```powershell
uv run python -m pytest -q
uv run python -m py_compile word_document_server/main.py word_document_server/tools/live_v2_tools.py
```

Public tool shape is guarded by tests. Adding a new MCP tool should be done as a
`word_v2_*` grouped tool, not by re-registering the lower-level `word_live_*`
helpers.
