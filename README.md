# Word MCP Live

Live Microsoft Word editing through a small, agent-friendly MCP interface.

This dev branch exposes only grouped `word_v2_*` MCP tools. The lower-level
`word_live_*` modules remain in the codebase as internal COM/JXA helpers, but
they are not registered as public tool calls. This keeps local models from
choosing between duplicate ways to open, read, edit, save, or close documents.

## Public Tool Surface

The server registers 15 public MCP tools:

| Tool | Purpose |
| --- | --- |
| `word_v2_open` | List open documents, attach to an open document, open a file, or create a new default-template session. |
| `word_v2_save` | Save in place, save as another document, or export PDF. |
| `word_v2_close` | Close a live session. |
| `word_v2_get_content` | Read text, page text, info, comments, revisions, or paragraph formatting. |
| `word_v2_search` | Find text and return reusable match handles. |
| `word_v2_edit` | Insert, replace, delete, insert paragraphs, undo, add hyperlinks, and manage footnotes. |
| `word_v2_format` | Apply inline formatting, paragraph formatting, styles, and lists. |
| `word_v2_comment` | Create, reply to, resolve, delete, list, or get comments. |
| `word_v2_track_changes` | Toggle, list, accept, reject, or decide revisions. |
| `word_v2_table` | Create, inspect, edit, and format tables. |
| `word_v2_media` | Insert images with sizing, alignment, wrapping, floating placement, and optional borders. |
| `word_v2_mutations` | Preview or apply grouped v2 operations in order. |
| `word_v2_layout` | Set page setup, insert page/section breaks, and set document properties. |
| `word_v2_blueprint` | Create, inspect, validate, or export structured document blueprints. |
| `word_v2_protection` | Protect or unprotect the active document. |

See [LIVE_V2_AGENT_GUIDE.md](LIVE_V2_AGENT_GUIDE.md) for the full agent
workflow reference. For local models, pass
[WORD_MCP_LIVE_SKILL.md](WORD_MCP_LIVE_SKILL.md) as the compact operational
skill.

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

## One-Step Windows Install

For a fresh Windows machine, run the installer script from the repo root:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\install-word-mcp.ps1
```

The script:

- Installs `uv` if it is missing.
- Installs the Python version from `.python-version` with `uv`.
- Runs `uv sync`.
- Creates a default `.env` if one does not already exist.
- Runs smoke tests unless `-SkipTests` is passed.
- Asks whether to configure Hermes, OpenClaw, or a custom MCP client.
- Writes and prints a target-specific `word-mcp-*.json` config file.
- Registers the server with the selected client CLI when available.

Useful options:

```powershell
.\scripts\install-word-mcp.ps1 -Target hermes
.\scripts\install-word-mcp.ps1 -Target openclaw
.\scripts\install-word-mcp.ps1 -Target custom
.\scripts\install-word-mcp.ps1 -Author "Your Name" -AuthorInitials "YN"
.\scripts\install-word-mcp.ps1 -SkipTests
.\scripts\install-word-mcp.ps1 -SkipRegister
.\scripts\install-word-mcp.ps1 -Transport streamable-http -HostAddress 127.0.0.1 -Port 8000
```

Targets:

| Target | Behavior |
| --- | --- |
| `hermes` | Generates Hermes config and runs `hermes mcp add word ...` when the CLI is on `PATH`. |
| `openclaw` | Generates OpenClaw config and runs `openclaw mcp set word ...`, then `openclaw mcp doctor word --probe` unless `-SkipProbe` is passed. |
| `custom` | Only writes and prints the config object. |

Hermes stdio config output:

```json
{
  "command": "C:\\path\\to\\Word-MCP\\.venv\\Scripts\\python.exe",
  "args": [
    "-m",
    "word_document_server.main"
  ],
  "env": {
    "MCP_TRANSPORT": "stdio",
    "PYTHONPATH": "C:\\path\\to\\Word-MCP"
  },
  "timeout": 180,
  "connect_timeout": 30
}
```

OpenClaw stdio config uses `cwd` instead of `PYTHONPATH`:

```json
{
  "command": "C:\\path\\to\\Word-MCP\\.venv\\Scripts\\python.exe",
  "args": [
    "-m",
    "word_document_server.main"
  ],
  "cwd": "C:\\path\\to\\Word-MCP",
  "env": {
    "MCP_TRANSPORT": "stdio"
  },
  "timeout": 180,
  "connectTimeout": 30
}
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

1. Call `word_v2_open()` to attach to the active Word document, or
   `word_v2_open(action="list")` if you need to choose from open documents.
2. Keep the returned `session_id`.
3. Inspect with `word_v2_get_content`.
4. Locate text with `word_v2_search`.
5. Edit, format, comment, or manage revisions using the same `session_id`.
6. Save with `word_v2_save`.
7. Close with `word_v2_close`.

For spontaneous document creation, use `word_v2_open(action="new")`. New
documents are initialized with the built-in `default_plain` template profile so
simple edits start from consistent page setup and professional default styles.

For precision document generation, use `word_v2_blueprint(action="create",
blueprint={...})` to create from a structured block list, then validate or
inspect with `word_v2_blueprint(action="validate"|"inspect")`.
For reference-driven recreation, pass `asset_dir` to
`word_v2_blueprint(action="inspect", asset_dir="C:/path/to/images")`. The
returned blueprint is page-aware and includes title-page, TOC, body-image,
table, and shape metadata for higher-fidelity rebuilds. Floating image blocks
can carry `wrapping`, `left_pt`, `top_pt`, and relative positioning constants;
preserve those values when recreating from a reference.
Paragraph `numbering` metadata from inspected blueprints is also replayed for
real Word bullet/numbered list paragraphs.
Title-page text boxes and simple rectangles are replayed from inspected shape
records; title-page pictures still require mapped assets for faithful rebuilds.
Title-page image assets are detected when their filenames contain `title` and
are mapped separately from body screenshots.

## Development

```powershell
uv run python -m pytest -q
uv run python -m py_compile word_document_server/main.py word_document_server/tools/live_v2_tools.py
```

Public tool shape is guarded by tests. Adding a new MCP tool should be done as a
`word_v2_*` grouped tool, not by re-registering the lower-level `word_live_*`
helpers.
