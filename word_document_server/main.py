"""Entry point for the live-only Word MCP server.

The public MCP surface intentionally exposes only the grouped ``word_v2_*``
tools. Lower-level ``word_live_*`` helpers still exist under
``word_document_server.tools`` as implementation details, but they are not
registered as public MCP tools.
"""

from __future__ import annotations

import os
import sys

from dotenv import load_dotenv
from fastmcp import FastMCP
from mcp.types import ToolAnnotations

from word_document_server.defaults import DEFAULT_AUTHOR
from word_document_server.tools import live_v2_tools


print("Loading configuration from .env file...", file=sys.stderr)
load_dotenv()
os.environ.setdefault("FASTMCP_LOG_LEVEL", "INFO")

mcp = FastMCP("Word Document Server")


def get_transport_config() -> dict[str, object]:
    """Read transport configuration from environment variables."""
    transport = os.getenv("MCP_TRANSPORT", "stdio").lower()
    valid_transports = {"stdio", "streamable-http", "sse"}
    if transport not in valid_transports:
        print(f"Warning: Invalid transport '{transport}'. Falling back to 'stdio'.", file=sys.stderr)
        transport = "stdio"

    return {
        "transport": transport,
        "host": os.getenv("MCP_HOST", "0.0.0.0"),
        "port": int(os.getenv("PORT", os.getenv("MCP_PORT", "8000"))),
        "path": os.getenv("MCP_PATH", "/mcp"),
        "sse_path": os.getenv("MCP_SSE_PATH", "/sse"),
    }


def register_tools() -> None:
    """Register the public MCP tools."""

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Open",
            destructiveHint=False,
        ),
    )
    async def word_v2_open(
        path: str = None,
        directory: str = ".",
        visible: bool = True,
        read_only: bool = False,
        password: str | None = None,
        action: str = "open",
    ):
        """Open, attach to, list, list sessions, or create Word documents in live mode."""
        return await live_v2_tools.word_v2_open(
            path=path,
            directory=directory,
            visible=visible,
            read_only=read_only,
            password=password,
            action=action,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Save",
            destructiveHint=True,
        ),
    )
    async def word_v2_save(session_id: str, out: str = None):
        """Save a live Word session in place, to a new path, or as PDF."""
        return await live_v2_tools.word_v2_save(session_id=session_id, out=out)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Close",
            destructiveHint=True,
        ),
    )
    async def word_v2_close(session_id: str, save_changes: str = "save"):
        """Close a live Word session."""
        return await live_v2_tools.word_v2_close(session_id=session_id, save_changes=save_changes)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Get Content",
            readOnlyHint=True,
        ),
    )
    async def word_v2_get_content(
        session_id: str,
        action: str = "text",
        page: int = 1,
        end_page: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        include_runs: bool = False,
    ):
        """Read text, page text, info, comments, revisions, paragraph formatting, snapshots, or diffs."""
        return await live_v2_tools.word_v2_get_content(
            session_id=session_id,
            action=action,
            page=page,
            end_page=end_page,
            start_paragraph=start_paragraph,
            end_paragraph=end_paragraph,
            include_runs=include_runs,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Search",
            readOnlyHint=True,
        ),
    )
    async def word_v2_search(
        session_id: str,
        find_text: str,
        match_case: bool = False,
        whole_word: bool = False,
        use_wildcards: bool = False,
        context_chars: int = 50,
        max_results: int = 100,
    ):
        """Search text and return reusable match handles."""
        return await live_v2_tools.word_v2_search(
            session_id=session_id,
            find_text=find_text,
            match_case=match_case,
            whole_word=whole_word,
            use_wildcards=use_wildcards,
            context_chars=context_chars,
            max_results=max_results,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Edit",
            destructiveHint=True,
        ),
    )
    async def word_v2_edit(
        session_id: str,
        action: str,
        text: str = "",
        find_text: str = "",
        replace_text: str = "",
        handle: str = None,
        target: dict = None,
        start: int = None,
        end: int = None,
        position: str = "end",
        replace_all: bool = False,
        match_case: bool = False,
        whole_word: bool = False,
        use_wildcards: bool = False,
        paragraphs: list = None,
        paragraph_index: int = None,
        style: str = None,
        track_changes: bool = False,
        times: int = 1,
        url: str = "",
        footnote_index: int = None,
    ):
        """Edit live text. Actions: insert, replace, delete, insert_paragraphs, undo, add_hyperlink, add_footnote, delete_footnote."""
        return await live_v2_tools.word_v2_edit(
            session_id=session_id,
            action=action,
            text=text,
            find_text=find_text,
            replace_text=replace_text,
            handle=handle,
            target=target,
            start=start,
            end=end,
            position=position,
            replace_all=replace_all,
            match_case=match_case,
            whole_word=whole_word,
            use_wildcards=use_wildcards,
            paragraphs=paragraphs,
            paragraph_index=paragraph_index,
            style=style,
            track_changes=track_changes,
            times=times,
            url=url,
            footnote_index=footnote_index,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Format",
            destructiveHint=True,
        ),
    )
    async def word_v2_format(
        session_id: str,
        action: str = "inline",
        handle: str = None,
        target: dict = None,
        start: int = None,
        end: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        bold: bool = None,
        italic: bool = None,
        underline: bool = None,
        strikethrough: bool = None,
        font_name: str = None,
        font_size: float = None,
        font_color: str = None,
        highlight_color: int = None,
        style: str = None,
        alignment: str = None,
        page_break_before: bool = None,
        preserve_direct_formatting: bool = False,
        list_type: str = "bullet",
        level: int = 0,
        remove: bool = False,
        continue_previous: bool = False,
        track_changes: bool = False,
    ):
        """Format text, paragraphs, styles, or lists."""
        return await live_v2_tools.word_v2_format(
            session_id=session_id,
            action=action,
            handle=handle,
            target=target,
            start=start,
            end=end,
            start_paragraph=start_paragraph,
            end_paragraph=end_paragraph,
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            font_name=font_name,
            font_size=font_size,
            font_color=font_color,
            highlight_color=highlight_color,
            style=style,
            alignment=alignment,
            page_break_before=page_break_before,
            preserve_direct_formatting=preserve_direct_formatting,
            list_type=list_type,
            level=level,
            remove=remove,
            continue_previous=continue_previous,
            track_changes=track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Comment",
            destructiveHint=True,
        ),
    )
    async def word_v2_comment(
        session_id: str,
        action: str,
        comment_id: int = None,
        text: str = "",
        handle: str = None,
        target: dict = None,
        start: int = None,
        end: int = None,
        paragraph_index: int = None,
        author: str = DEFAULT_AUTHOR,
        resolve: bool = True,
    ):
        """Create, reply to, resolve, delete, list, or get comments."""
        return await live_v2_tools.word_v2_comment(
            session_id=session_id,
            action=action,
            comment_id=comment_id,
            text=text,
            handle=handle,
            target=target,
            start=start,
            end=end,
            paragraph_index=paragraph_index,
            author=author,
            resolve=resolve,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Track Changes",
            destructiveHint=True,
        ),
    )
    async def word_v2_track_changes(
        session_id: str,
        action: str,
        enable: bool = None,
        author: str = None,
        change_ids: list[int] = None,
        decision: str = "accept",
    ):
        """Toggle, list, accept, reject, or decide tracked changes."""
        return await live_v2_tools.word_v2_track_changes(
            session_id=session_id,
            action=action,
            enable=enable,
            author=author,
            change_ids=change_ids,
            decision=decision,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Table",
            destructiveHint=True,
        ),
    )
    async def word_v2_table(
        session_id: str,
        action: str,
        table_index: int = 1,
        rows: int = 2,
        cols: int = 2,
        position: str = "end",
        data: list = None,
        row: int = None,
        col: int = None,
        text: str = None,
        cells: list = None,
        start_row: int = None,
        start_col: int = None,
        end_row: int = None,
        end_col: int = None,
        style: str = "Table Grid",
        autofit: str = "window",
        border_style: str = None,
        cell_bold: list[list] = None,
        cell_alignment: list[list] = None,
        column_widths: list[float] = None,
        table_alignment: str = None,
        cell_shading: list[list] = None,
        track_changes: bool = False,
    ):
        """Create, inspect, edit, or format tables."""
        return await live_v2_tools.word_v2_table(
            session_id=session_id,
            action=action,
            table_index=table_index,
            rows=rows,
            cols=cols,
            position=position,
            data=data,
            row=row,
            col=col,
            text=text,
            cells=cells,
            start_row=start_row,
            start_col=start_col,
            end_row=end_row,
            end_col=end_col,
            style=style,
            autofit=autofit,
            border_style=border_style,
            cell_bold=cell_bold,
            cell_alignment=cell_alignment,
            column_widths=column_widths,
            table_alignment=table_alignment,
            cell_shading=cell_shading,
            track_changes=track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Media",
            destructiveHint=True,
        ),
    )
    async def word_v2_media(
        session_id: str,
        action: str,
        path: str,
        paragraph_index: int = None,
        position: str = "end",
        width_inches: float = None,
        height_inches: float = None,
        width_pt: float = None,
        height_pt: float = None,
        alignment: str = None,
        wrapping: str = "inline",
        border_style: str = None,
        border_width_pt: float = None,
        border_color: str = None,
        link_to_file: bool = False,
        paragraph_after: bool = True,
        left_pt: float = None,
        top_pt: float = None,
        relative_horizontal_position: int = None,
        relative_vertical_position: int = None,
    ):
        """Insert media into a live session."""
        return await live_v2_tools.word_v2_media(
            session_id=session_id,
            action=action,
            path=path,
            paragraph_index=paragraph_index,
            position=position,
            width_inches=width_inches,
            height_inches=height_inches,
            width_pt=width_pt,
            height_pt=height_pt,
            alignment=alignment,
            wrapping=wrapping,
            border_style=border_style,
            border_width_pt=border_width_pt,
            border_color=border_color,
            link_to_file=link_to_file,
            paragraph_after=paragraph_after,
            left_pt=left_pt,
            top_pt=top_pt,
            relative_horizontal_position=relative_horizontal_position,
            relative_vertical_position=relative_vertical_position,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Mutations",
            readOnlyHint=False,
        ),
    )
    async def word_v2_mutations(
        session_id: str,
        action: str,
        operations: list[dict] = None,
    ):
        """Preview or apply multiple v2 operations in order."""
        return await live_v2_tools.word_v2_mutations(
            session_id=session_id,
            action=action,
            operations=operations,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Layout",
            destructiveHint=True,
        ),
    )
    async def word_v2_layout(
        session_id: str,
        action: str,
        page_size: str = "letter",
        orientation: str = "portrait",
        width: float = None,
        height: float = None,
        margins: dict = None,
        position: str = "end",
        paragraph_index: int = None,
        break_type: str = "next_page",
        title: str = None,
        subject: str = None,
        author: str = None,
        keywords: str = None,
        comments: str = None,
        category: str = None,
        manager: str = None,
        company: str = None,
        last_author: str = None,
    ):
        """Manage page setup, breaks, and document properties."""
        return await live_v2_tools.word_v2_layout(
            session_id=session_id,
            action=action,
            page_size=page_size,
            orientation=orientation,
            width=width,
            height=height,
            margins=margins,
            position=position,
            paragraph_index=paragraph_index,
            break_type=break_type,
            title=title,
            subject=subject,
            author=author,
            keywords=keywords,
            comments=comments,
            category=category,
            manager=manager,
            company=company,
            last_author=last_author,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Blueprint",
            destructiveHint=True,
        ),
    )
    async def word_v2_blueprint(
        action: str,
        session_id: str = None,
        blueprint: dict = None,
        out: str = None,
        visible: bool = True,
        asset_dir: str = None,
    ):
        """Create, inspect, validate, or export structured document blueprints."""
        return await live_v2_tools.word_v2_blueprint(
            action=action,
            session_id=session_id,
            blueprint=blueprint,
            out=out,
            visible=visible,
            asset_dir=asset_dir,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Protection",
            destructiveHint=True,
        ),
        description=live_v2_tools.word_v2_protection.__doc__,
    )
    async def word_v2_protection(
        session_id: str,
        action: str,
        protection_type: str = "read_only",
        password: str = None,
    ):
        """Protect or unprotect a live session."""
        return await live_v2_tools.word_v2_protection(
            session_id=session_id,
            action=action,
            protection_type=protection_type,
            password=password,
        )


def run_server():
    """Run the Word Document MCP server with the configured transport."""
    config = get_transport_config()

    from word_document_server.utils.save_utils import install_save_hook
    from word_document_server.utils.path_utils import install_path_hook

    install_save_hook()
    install_path_hook()
    register_tools()

    transport_type = config["transport"]
    print(f"Starting Word Document MCP Server with {transport_type} transport...", file=sys.stderr)

    try:
        if transport_type == "stdio":
            print("Server running on stdio transport", file=sys.stderr)
            mcp.run(transport="stdio")
        elif transport_type == "streamable-http":
            print(
                f"Server running on streamable-http transport at "
                f"http://{config['host']}:{config['port']}{config['path']}",
                file=sys.stderr,
            )
            mcp.run(
                transport="streamable-http",
                host=config["host"],
                port=config["port"],
                path=config["path"],
            )
        elif transport_type == "sse":
            print(
                f"Server running on SSE transport at "
                f"http://{config['host']}:{config['port']}{config['sse_path']}",
                file=sys.stderr,
            )
            mcp.run(
                transport="sse",
                host=config["host"],
                port=config["port"],
                path=config["sse_path"],
            )
    except KeyboardInterrupt:
        print("\nShutting down server...", file=sys.stderr)
    except Exception as exc:
        print(f"Error starting server: {exc}", file=sys.stderr)
        sys.exit(1)

    return mcp


def main():
    """Console entry point."""
    run_server()


if __name__ == "__main__":
    main()
