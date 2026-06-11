# Word MCP Live

Word MCP Live is a Windows-only MCP server for editing Microsoft Word documents
that are open on your machine.

It lets an AI agent open or attach to Word documents, read content, make edits,
turn tracked changes on or off, add comments, work with tables/images, save, and
close documents through a small `word_v2_*` tool surface.

This project is focused on live Word automation through Microsoft Word COM, so
it requires Windows and Microsoft Word.

## Install

Open PowerShell, clone the repo, and run the installer:

```powershell
git clone https://github.com/Mexicat-Feeder/Word-MCP.git
cd Word-MCP
powershell -ExecutionPolicy Bypass -File .\scripts\install-word-mcp.ps1
```

The installer will:

- Install `uv` if needed.
- Install the pinned Python version.
- Create the local virtual environment.
- Install dependencies.
- Run smoke tests.
- Ask whether to configure `hermes`, `openclaw`, or `custom`.
- Install the Word MCP agent skill for Hermes or OpenClaw.

For a specific client, skip the selector:

```powershell
.\scripts\install-word-mcp.ps1 -Target hermes
.\scripts\install-word-mcp.ps1 -Target openclaw
.\scripts\install-word-mcp.ps1 -Target custom
```

Use `custom` if your MCP client does not have a supported CLI installer. The
script will print the config object you can paste into your client.

## Useful Options

```powershell
.\scripts\install-word-mcp.ps1 -Author "Your Name" -AuthorInitials "YN"
.\scripts\install-word-mcp.ps1 -SkipTests
.\scripts\install-word-mcp.ps1 -SkipRegister
.\scripts\install-word-mcp.ps1 -SkipSkillInstall
```

`MCP_AUTHOR` and `MCP_AUTHOR_INITIALS` control the author shown on Word comments
and tracked changes.

## Run Manually

After installation, the server can be started with:

```powershell
uv run python -m word_document_server.main
```

The default transport is `stdio`, which is what most local MCP clients expect.

## Development

```powershell
uv run python -m pytest -q
```

See `LIVE_V2_AGENT_GUIDE.md` for the detailed tool workflow and
`WORD_MCP_LIVE_SKILL.md` for a compact agent-facing skill.
