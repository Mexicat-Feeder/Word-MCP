"""
Main entry point for the Word Document MCP Server.
Acts as the central controller for the MCP server that handles Word document operations.
Supports multiple transports: stdio, sse, and streamable-http using standalone FastMCP.
"""

import os
import sys
from dotenv import load_dotenv

# Load environment variables from .env file
print("Loading configuration from .env file...", file=sys.stderr)
load_dotenv()
from word_document_server.defaults import DEFAULT_AUTHOR, DEFAULT_INITIALS
# Set required environment variable for FastMCP 2.8.1+
os.environ.setdefault('FASTMCP_LOG_LEVEL', 'INFO')
from fastmcp import FastMCP
from mcp.types import ToolAnnotations
from word_document_server.tools import (
    live_tools,
    live_read_tools,
    live_v2_tools,
    live_layout_tools,
    screen_capture_tools,
)

def get_transport_config():
    """
    Get transport configuration from environment variables.
    
    Returns:
        dict: Transport configuration with type, host, port, and other settings
    """
    # Default configuration
    config = {
        'transport': 'stdio',  # Default to stdio for backward compatibility
        'host': '0.0.0.0',
        'port': 8000,
        'path': '/mcp',
        'sse_path': '/sse'
    }
    
    # Override with environment variables if provided
    transport = os.getenv('MCP_TRANSPORT', 'stdio').lower()
    print(f"Transport: {transport}", file=sys.stderr)
    # Validate transport type
    valid_transports = ['stdio', 'streamable-http', 'sse']
    if transport not in valid_transports:
        print(f"Warning: Invalid transport '{transport}'. Falling back to 'stdio'.", file=sys.stderr)
        transport = 'stdio'
    
    config['transport'] = transport
    config['host'] = os.getenv('MCP_HOST', config['host'])
    # Use PORT from Render if available, otherwise fall back to MCP_PORT or default
    config['port'] = int(os.getenv('PORT', os.getenv('MCP_PORT', config['port'])))
    config['path'] = os.getenv('MCP_PATH', config['path'])
    config['sse_path'] = os.getenv('MCP_SSE_PATH', config['sse_path'])
    
    return config


def setup_logging(debug_mode):
    """
    Setup logging based on debug mode.
    
    Args:
        debug_mode (bool): Whether to enable debug logging
    """
    import logging
    
    if debug_mode:
        logging.basicConfig(
            level=logging.DEBUG,
            format='%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        print("Debug logging enabled", file=sys.stderr)
    else:
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )


# Initialize FastMCP server
mcp = FastMCP("Word Document Server")


def register_tools():
    """Register all tools with the MCP server using FastMCP decorators."""
    
    # --- Live editing tools (Windows only, requires Word running) ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Screen Capture",
            readOnlyHint=True,
        ),
    )
    async def word_screen_capture(filename: str = None, output_path: str = None):
        """[Windows only] Capture a screenshot of a Word document window.
        Returns the path to the saved PNG image. Requires Word to be running."""
        return await screen_capture_tools.word_screen_capture(filename, output_path)

    # --- V2 live-COM tools: grouped, agent-friendly primary surface ---

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
    ):
        """Open a Word document in live COM mode and return a session_id.
        Use this before all other word_v2_* tools."""
        return await live_v2_tools.word_v2_open(
            path=path,
            directory=directory,
            visible=visible,
            read_only=read_only,
            password=password,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Save",
            destructiveHint=True,
        ),
    )
    async def word_v2_save(session_id: str, out: str = None):
        """Save a v2 live session in place, or save as out."""
        return await live_v2_tools.word_v2_save(session_id=session_id, out=out)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word V2 Close",
            destructiveHint=True,
        ),
    )
    async def word_v2_close(session_id: str, save_changes: str = "save"):
        """Close a v2 live session."""
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
        """Retrieve content or properties from a live session.
        Actions: text, page_text, info, comments, revisions, paragraph_format."""
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
        """Search text in a live session."""
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
        """Format text or paragraphs in a live session."""
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
        """Manage comments on a live session."""
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
        """Manage tracked changes/revisions on a live session."""
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
        """Manage tables on a live session."""
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
            title="Word V2 Mutations",
            readOnlyHint=False,
        ),
    )
    async def word_v2_mutations(
        session_id: str,
        action: str,
        operations: list[dict] = None,
    ):
        """Preview or apply multiple v2 operations."""
        return await live_v2_tools.word_v2_mutations(
            session_id=session_id,
            action=action,
            operations=operations,
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
        return await live_v2_tools.word_v2_protection(
            session_id=session_id,
            action=action,
            protection_type=protection_type,
            password=password,
        )

    # --- Live basic tools block ---
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Text",
            destructiveHint=True,
        ),
    )
    async def word_live_insert_text(
        filename: str = None,
        text: str = "",
        position: str = "end",
        bookmark: str = None,
        track_changes: bool = False,
    ):
        """[Windows only] Insert text into a Word document that is open in Word.
        Position: 'start', 'end', 'cursor', or character offset. Requires Word running."""
        return await live_tools.word_live_insert_text(
            filename, text, position, bookmark, track_changes
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Format Text",
            destructiveHint=True,
        ),
        description=live_tools.word_live_format_text.__doc__,
    )
    async def word_live_format_text(
        filename: str = None,
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
        style_name: str = None,
        paragraph_alignment: str = None,
        page_break_before: bool = None,
        preserve_direct_formatting: bool = False,
        track_changes: bool = False,
    ):
        return await live_tools.word_live_format_text(
            filename, start, end, start_paragraph, end_paragraph,
            bold, italic, underline, strikethrough,
            font_name, font_size, font_color, highlight_color,
            style_name, paragraph_alignment, page_break_before,
            preserve_direct_formatting, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Replace Text",
            destructiveHint=True,
        ),
        description=live_tools.word_live_replace_text.__doc__,
    )
    async def word_live_replace_text(
        filename: str = None,
        find_text: str = "",
        replace_text: str = "",
        match_case: bool = False,
        match_whole_word: bool = False,
        use_wildcards: bool = False,
        replace_all: bool = True,
        track_changes: bool = False,
    ):
        return await live_tools.word_live_replace_text(
            filename, find_text, replace_text, match_case,
            match_whole_word, use_wildcards, replace_all, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Paragraphs",
            destructiveHint=True,
        ),
        description=live_tools.word_live_insert_paragraphs.__doc__,
    )
    async def word_live_insert_paragraphs(
        filename: str = None,
        paragraphs: list = None,
        target_text: str = None,
        paragraph_index: int = None,
        position: str = "after",
        style: str = None,
        track_changes: bool = False,
    ):
        return await live_tools.word_live_insert_paragraphs(
            filename, paragraphs, target_text, paragraph_index,
            position, style, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Table",
            destructiveHint=True,
        ),
    )
    async def word_live_add_table(
        filename: str = None,
        rows: int = 2,
        cols: int = 2,
        position: str = "end",
        data: list = None,
        style: str = "Table Grid",
        autofit: str = "window",
        track_changes: bool = False,
    ):
        """[Windows only] Add a table to a Word document open in Word.
        Optionally provide data as 2D list. Default style is 'Table Grid' with
        autofit to window width. Set style=None for no style, autofit=None for
        legacy fixed behavior. Requires Word running."""
        return await live_tools.word_live_add_table(
            filename, rows, cols, position, data, style, autofit, track_changes
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Format Table",
            destructiveHint=True,
        ),
        description=live_tools.word_live_format_table.__doc__,
    )
    async def word_live_format_table(
        filename: str = None,
        table_index: int = -1,
        border_style: str = None,
        cell_bold: list[list] = None,
        cell_alignment: list[list] = None,
        column_widths: list[float] = None,
        table_alignment: str = None,
        cell_shading: list[list] = None,
        autofit: str = None,
    ):
        return await live_tools.word_live_format_table(
            filename, table_index, border_style, cell_bold, cell_alignment,
            column_widths, table_alignment, cell_shading, autofit
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Modify Table",
            destructiveHint=True,
        ),
        description=live_tools.word_live_modify_table.__doc__,
    )
    async def word_live_modify_table(
        filename: str = None,
        table_index: int = 1,
        operation: str = "get_info",
        row: int = None,
        col: int = None,
        text: str = None,
        before_row: int = None,
        before_col: int = None,
        header: str = None,
        cells: list = None,
        start_row: int = None,
        start_col: int = None,
        end_row: int = None,
        end_col: int = None,
        autofit_mode: str = "content",
        accept_revisions: bool = False,
        track_changes: bool = False,
    ):
        return await live_tools.word_live_modify_table(
            filename, table_index, operation, row, col, text,
            before_row, before_col, header, cells,
            start_row, start_col, end_row, end_col,
            autofit_mode, accept_revisions, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Delete Text",
            destructiveHint=True,
        ),
    )
    async def word_live_delete_text(
        filename: str = None,
        start: int = None,
        end: int = None,
        track_changes: bool = False,
    ):
        """[Windows only] Delete text from a Word document open in Word.
        Specify start/end character positions. Requires Word running."""
        return await live_tools.word_live_delete_text(
            filename, start, end, track_changes
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Apply List",
            destructiveHint=True,
        ),
        description=live_tools.word_live_apply_list.__doc__,
    )
    async def word_live_apply_list(
        filename: str = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        list_type: str = "bullet",
        level: int = 0,
        remove: bool = False,
        continue_previous: bool = False,
        number_format: dict = None,
        number_style: dict = None,
        start_at: dict = None,
        level_map: dict = None,
        track_changes: bool = False,
    ):
        return await live_tools.word_live_apply_list(
            filename, start_paragraph, end_paragraph, list_type,
            level, remove, continue_previous, number_format,
            number_style, start_at, level_map, track_changes,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Setup Heading Numbering",
            destructiveHint=True,
        ),
        description=live_tools.word_live_setup_heading_numbering.__doc__,
    )
    async def word_live_setup_heading_numbering(
        filename: str = None,
        h1_paragraphs: list[int] = None,
        h2_paragraphs: list[int] = None,
        strip_manual_numbers: bool = True,
        h1_number_format: str = None,
        h2_number_format: str = None,
        font_name: str = None,
        h1_size: float = None,
        h2_size: float = None,
        bold: bool = None,
        alignment: str = None,
        font_color: str = None,
        h1_space_before: float = None,
        h1_space_after: float = None,
        h2_space_before: float = None,
        h2_space_after: float = None,
        line_spacing: float = None,
    ):
        return await live_tools.word_live_setup_heading_numbering(
            filename, h1_paragraphs, h2_paragraphs, strip_manual_numbers,
            h1_number_format, h2_number_format,
            font_name, h1_size, h2_size, bold, alignment, font_color,
            h1_space_before, h1_space_after, h2_space_before, h2_space_after,
            line_spacing,
        )

    # --- Live read tools (Windows only, requires Word running) ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Text",
            readOnlyHint=True,
        ),
    )
    async def word_live_get_text(filename: str = None):
        """[Windows only] Get all text from a Word document open in Word, paragraph by paragraph. For large documents (200+ paragraphs), automatically returns only the first 3 pages — use word_live_get_page_text to read specific pages."""
        return await live_read_tools.word_live_get_text(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Take Snapshot",
            readOnlyHint=True,
        ),
    )
    async def word_live_take_snapshot(filename: str = None):
        """[Windows only] Store a snapshot of the current document text for later diffing without returning the full text. Use word_live_get_diff afterwards to see what changed."""
        return await live_read_tools.word_live_take_snapshot(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Diff",
            readOnlyHint=True,
        ),
    )
    async def word_live_get_diff(filename: str = None):
        """[Windows only] Return only paragraphs that changed since the last snapshot. Compares current document against snapshot from word_live_take_snapshot. Returns added, modified, deleted paragraphs. Automatically updates snapshot after diffing."""
        return await live_read_tools.word_live_get_diff(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Snapshot Status",
            readOnlyHint=True,
        ),
    )
    async def word_live_snapshot_status(filename: str = None):
        """[Windows only] Check whether a snapshot exists for the document and how old it is. Returns has_snapshot, age_seconds, and paragraph_count."""
        return await live_read_tools.word_live_snapshot_status(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Paragraph Format",
            readOnlyHint=True,
        ),
        description=live_read_tools.word_live_get_paragraph_format.__doc__,
    )
    async def word_live_get_paragraph_format(
        filename: str = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        include_runs: bool = False,
    ):
        return await live_read_tools.word_live_get_paragraph_format(
            filename, start_paragraph, end_paragraph, include_runs,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Info",
            readOnlyHint=True,
        ),
    )
    async def word_live_get_info(filename: str = None):
        """[Windows only] Get document info (pages, words, sections, etc.) from a Word document open in Word. Requires Word running."""
        return await live_read_tools.word_live_get_info(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Set Core Properties",
            destructiveHint=True,
        ),
        description=live_read_tools.word_live_set_core_properties.__doc__,
    )
    async def word_live_set_core_properties(
        filename: str = None,
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
        return await live_read_tools.word_live_set_core_properties(
            filename=filename,
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
            title="Word Live List Open",
            readOnlyHint=True,
        ),
    )
    async def word_live_list_open():
        """[Windows only] List all documents currently open in Word with name, path, pages, and saved status."""
        return await live_read_tools.word_live_list_open()

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Find Text",
            readOnlyHint=True,
        ),
    )
    async def word_live_find_text(
        filename: str = None,
        find_text: str = "",
        match_case: bool = False,
        whole_word: bool = False,
        use_wildcards: bool = False,
        context_chars: int = 60,
        max_results: int = 50,
    ):
        """[Windows only] Find text in a Word document open in Word. Returns positions and context.
        With use_wildcards=True, supports ^m (page break), ^t (tab), ^p (paragraph mark) and Word wildcards.
        context_chars controls how many characters of surrounding context to return (default 60). Requires Word running."""
        return await live_read_tools.word_live_find_text(
            filename, find_text, match_case, whole_word,
            use_wildcards, context_chars, max_results,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Comments",
            readOnlyHint=True,
        ),
    )
    async def word_live_get_comments(filename: str = None):
        """[Windows only] Get all comments from a Word document open in Word. Requires Word running."""
        return await live_read_tools.word_live_get_comments(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Comment",
            destructiveHint=True,
        ),
    )
    async def word_live_add_comment(
        filename: str = None,
        start: int = None,
        end: int = None,
        paragraph_index: int = None,
        comment_text: str = "",
        author: str = DEFAULT_AUTHOR,
    ):
        """[Windows only] Add a comment to a Word document open in Word.
        Specify start/end character positions or paragraph_index (1-indexed). Requires Word running."""
        return await live_read_tools.word_live_add_comment(
            filename, start, end, paragraph_index, comment_text, author
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Reply to Comment",
            destructiveHint=True,
        ),
    )
    async def word_live_reply_to_comment(
        filename: str = None,
        comment_index: int = None,
        comment_text: str = "",
        author: str = DEFAULT_AUTHOR,
    ):
        """[Windows only] Reply to an existing comment in a Word document open in Word.
        Adds a threaded reply. Use word_live_get_comments to find the comment_index.
        Requires Word 2016+ running."""
        return await live_read_tools.word_live_reply_to_comment(
            filename, comment_index, comment_text, author
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Resolve Comment",
            destructiveHint=True,
        ),
    )
    async def word_live_resolve_comment(
        filename: str = None,
        comment_index: int = None,
        resolve: bool = True,
    ):
        """[Windows only] Resolve or unresolve a comment in a Word document open in Word.
        Sets the comment's Done property. Use word_live_get_comments to find comment_index.
        Requires Word 2016+ running."""
        return await live_read_tools.word_live_resolve_comment(
            filename, comment_index, resolve
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Delete Comment",
            destructiveHint=True,
        ),
    )
    async def word_live_delete_comment(
        filename: str = None,
        comment_index: int = None,
    ):
        """[Windows only] Delete a comment from a Word document open in Word.
        Permanently removes the comment. Use word_live_get_comments to find comment_index.
        Requires Word running."""
        return await live_read_tools.word_live_delete_comment(
            filename, comment_index
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live List Revisions",
            readOnlyHint=True,
        ),
    )
    async def word_live_list_revisions(filename: str = None):
        """[Windows only] List all tracked changes (revisions) in a Word document open in Word. Requires Word running."""
        return await live_read_tools.word_live_list_revisions(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Accept Revisions",
            destructiveHint=True,
        ),
    )
    async def word_live_accept_revisions(
        filename: str = None,
        author: str = None,
        revision_ids: list[int] = None,
    ):
        """[Windows only] Accept tracked changes in a Word document open in Word.
        Filter by author or specific revision IDs. Requires Word running."""
        return await live_read_tools.word_live_accept_revisions(
            filename, author, revision_ids
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Reject Revisions",
            destructiveHint=True,
        ),
    )
    async def word_live_reject_revisions(
        filename: str = None,
        author: str = None,
        revision_ids: list[int] = None,
    ):
        """[Windows only] Reject tracked changes in a Word document open in Word.
        Filter by author or specific revision IDs. Requires Word running."""
        return await live_read_tools.word_live_reject_revisions(
            filename, author, revision_ids
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Page Text",
            readOnlyHint=True,
        ),
        description=live_read_tools.word_live_get_page_text.__doc__,
    )
    async def word_live_get_page_text(
        filename: str = None,
        page: int = 1,
        end_page: int = None,
    ):
        return await live_read_tools.word_live_get_page_text(
            filename, page, end_page,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Get Undo History",
            readOnlyHint=True,
        ),
    )
    async def word_live_get_undo_history(filename: str = None):
        """[Windows only] Get the undo stack from a Word document open in Word.
        Shows MCP tool operations as named entries. Requires Word running."""
        return await live_read_tools.word_live_get_undo_history(filename)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Undo",
            destructiveHint=True,
        ),
    )
    async def word_live_undo(
        filename: str = None,
        times: int = 1,
    ):
        """[Windows only] Undo the last N operations in a Word document open in Word.
        Each MCP tool call is one undo entry. Requires Word running."""
        return await live_tools.word_live_undo(filename, times)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Save",
            destructiveHint=True,
        ),
    )
    async def word_live_save(
        filename: str = None,
        save_as: str = None,
    ):
        """[Windows only] Save a Word document open in Word.
        Optionally save to a new path with save_as. Requires Word running."""
        return await live_tools.word_live_save(filename, save_as)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Toggle Track Changes",
            destructiveHint=True,
        ),
    )
    async def word_live_toggle_track_changes(
        filename: str = None,
        enable: bool = None,
    ):
        """[Windows only] Toggle or set Track Changes mode on a Word document.
        If enable is omitted, toggles current state. Requires Word running."""
        return await live_tools.word_live_toggle_track_changes(filename, enable)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Image",
            destructiveHint=True,
        ),
        description=live_tools.word_live_insert_image.__doc__,
    )
    async def word_live_insert_image(
        filename: str = None,
        image_path: str = "",
        paragraph_index: int = None,
        position: str = "end",
        width_inches: float = None,
        height_inches: float = None,
        width_pt: float = None,
        height_pt: float = None,
        alignment: str = None,
        wrapping: str = None,
        border_style: str = None,
        border_width_pt: float = None,
        border_color: str = None,
        link_to_file: bool = False,
    ):
        return await live_tools.word_live_insert_image(
            filename, image_path, paragraph_index, position,
            width_inches, height_inches, width_pt, height_pt,
            alignment, wrapping, border_style, border_width_pt,
            border_color, link_to_file
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Cross Reference",
            destructiveHint=True,
        ),
    )
    async def word_live_insert_cross_reference(
        filename: str = None,
        ref_type: str = "heading",
        ref_item: int = 1,
        ref_kind: str = "text",
        position: str = "end",
        paragraph_index: int = None,
        insert_as_hyperlink: bool = True,
    ):
        """[Windows only] Insert a cross-reference to a heading, bookmark, figure, table, etc.
        First use word_live_list_cross_reference_items to discover available targets.
        ref_type: heading, bookmark, figure, table, equation, footnote, endnote.
        ref_kind: text, number, number_no_context, page, above_below.
        Requires Word running."""
        return await live_tools.word_live_insert_cross_reference(
            filename, ref_type, ref_item, ref_kind,
            position, paragraph_index, insert_as_hyperlink
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live List Cross Reference Items",
            readOnlyHint=True,
        ),
    )
    async def word_live_list_cross_reference_items(
        filename: str = None,
        ref_type: str = "heading",
    ):
        """[Windows only] List available cross-reference targets in a Word document.
        Returns items that can be referenced with word_live_insert_cross_reference.
        ref_type: heading, bookmark, figure, table, equation, footnote, endnote.
        Requires Word running."""
        return await live_tools.word_live_list_cross_reference_items(filename, ref_type)

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Insert Equation",
            destructiveHint=True,
        ),
    )
    async def word_live_insert_equation(
        filename: str = None,
        equation: str = "",
        paragraph_index: int = None,
        position: str = "end",
        display_mode: bool = False,
    ):
        """[Windows only] Insert a mathematical equation into a Word document using UnicodeMath syntax.
        Examples: "x^2 + y^2 = z^2", "(a+b)/(c+d)" (fraction), "\\sqrt(x^2+y^2)" (root),
        "\\alpha + \\beta" (Greek), "\\int_0^\\infty e^(-x^2) dx" (integral),
        "\\sum_(i=1)^n i^2" (summation), "\\matrix(a&b@c&d)" (matrix).
        display_mode=True centers the equation on its own line.
        Requires Word running."""
        return await live_tools.word_live_insert_equation(
            filename, equation, paragraph_index, position, display_mode
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Diagnose Layout",
            readOnlyHint=True,
        ),
        description=live_read_tools.word_live_diagnose_layout.__doc__,
    )
    async def word_live_diagnose_layout(filename: str = None):
        return await live_read_tools.word_live_diagnose_layout(filename)

    # --- Live layout tools (Windows only, requires Word running) ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Set Page Layout",
            destructiveHint=True,
        ),
    )
    async def word_live_set_page_layout(
        filename: str = None,
        section_index: int = 1,
        orientation: str = None,
        page_width_inches: float = None,
        page_height_inches: float = None,
        margin_top_inches: float = None,
        margin_bottom_inches: float = None,
        margin_left_inches: float = None,
        margin_right_inches: float = None,
    ):
        """[Windows only] Set page layout (orientation, size, margins) for a section in a Word document open in Word. Requires Word running."""
        return await live_layout_tools.word_live_set_page_layout(
            filename, section_index, orientation,
            page_width_inches, page_height_inches,
            margin_top_inches, margin_bottom_inches,
            margin_left_inches, margin_right_inches,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Header/Footer",
            destructiveHint=True,
        ),
    )
    async def word_live_add_header_footer(
        filename: str = None,
        section_index: int = 1,
        header_text: str = None,
        footer_text: str = None,
        header_alignment: str = "center",
        footer_alignment: str = "center",
    ):
        """[Windows only] Add header and/or footer to a section in a Word document open in Word. Requires Word running."""
        return await live_layout_tools.word_live_add_header_footer(
            filename, section_index, header_text, footer_text,
            header_alignment, footer_alignment,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Page Numbers",
            destructiveHint=True,
        ),
    )
    async def word_live_add_page_numbers(
        filename: str = None,
        section_index: int = 1,
        position: str = "footer",
        alignment: str = "center",
        prefix: str = "",
        suffix: str = "",
        include_total: bool = False,
    ):
        """[Windows only] Add page numbers to header or footer in a Word document open in Word. Requires Word running."""
        return await live_layout_tools.word_live_add_page_numbers(
            filename, section_index, position, alignment,
            prefix, suffix, include_total,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Section Break",
            destructiveHint=True,
        ),
    )
    async def word_live_add_section_break(
        filename: str = None,
        break_type: str = "new_page",
    ):
        """[Windows only] Add a section break (new_page, continuous, even_page, odd_page) to a Word document open in Word. Requires Word running."""
        return await live_layout_tools.word_live_add_section_break(
            filename, break_type,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Set Paragraph Spacing",
            destructiveHint=True,
        ),
    )
    async def word_live_set_paragraph_spacing(
        filename: str = None,
        paragraph_index: int = None,
        start_paragraph: int = None,
        end_paragraph: int = None,
        space_before_pt: float = None,
        space_after_pt: float = None,
        line_spacing: float = None,
        line_spacing_rule: str = None,
        keep_with_next: bool = None,
        keep_together: bool = None,
        alignment: str = None,
    ):
        """[Windows only] Set paragraph spacing and layout properties in a Word document open in Word. Paragraphs are 1-indexed. Requires Word running."""
        return await live_layout_tools.word_live_set_paragraph_spacing(
            filename, paragraph_index, start_paragraph, end_paragraph,
            space_before_pt, space_after_pt, line_spacing, line_spacing_rule,
            keep_with_next, keep_together, alignment,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Bookmark",
            destructiveHint=True,
        ),
    )
    async def word_live_add_bookmark(
        filename: str = None,
        paragraph_index: int = 1,
        bookmark_name: str = "",
    ):
        """[Windows only] Add a named bookmark at a paragraph in a Word document open in Word.
        Paragraph is 1-indexed. Requires Word running."""
        return await live_layout_tools.word_live_add_bookmark(
            filename, paragraph_index, bookmark_name,
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Watermark",
            destructiveHint=True,
        ),
    )
    async def word_live_add_watermark(
        filename: str = None,
        text: str = "TASLAK",
        font_size: int = 72,
        font_color: str = "C0C0C0",
        rotation: int = -45,
        section_index: int = 1,
    ):
        """[Windows only] Add a diagonal text watermark to a Word document open in Word. Requires Word running."""
        return await live_layout_tools.word_live_add_watermark(
            filename, text, font_size, font_color, rotation, section_index,
        )

    # --- Open / Close document (live) ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Open Document",
        ),
    )

    async def word_live_open_document(
        filename: str = None,
        directory: str = ".",
        visible: bool = True,
        read_only: bool = False,
        password: str | None = None,
    ):
        """[Windows/macOS] Open a Word document so live tools can operate on it. Requires Word running."""
        return await live_tools.word_live_open_document(
            filename, directory, visible, read_only, password,
        )
    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Close Document",
            destructiveHint=True,
        ),
    )
    async def word_live_close_document(
        filename: str = None,
        save_changes: str = "save",
    ):
        """[Windows/macOS] Close a document that is currently open in Word. Requires Word running."""
        return await live_tools.word_live_close_document(
            filename, save_changes,
        )

    # --- List open documents (live) ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live List Open Documents",
            readOnlyHint=True,
        ),
    )
    async def word_live_list_open_documents():
        """[Windows/macOS] List all documents currently open in Word.

        Returns:
            JSON with list of open documents, their paths, and modification status.
        """
        return await live_tools.word_live_list_open_documents()

    # --- New Live Tool Extensions ---

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Hyperlink",
            destructiveHint=True,
        ),
        description=live_tools.word_live_add_hyperlink.__doc__,
    )
    async def word_live_add_hyperlink(
        filename: str = None,
        url: str = "",
        text: str = "",
        start: int = None,
        end: int = None,
    ):
        return await live_tools.word_live_add_hyperlink(
            filename, url, text, start, end
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Add Footnote",
            destructiveHint=True,
        ),
        description=live_tools.word_live_add_footnote.__doc__,
    )
    async def word_live_add_footnote(
        filename: str = None,
        text: str = "",
        start: int = None,
        end: int = None,
    ):
        return await live_tools.word_live_add_footnote(
            filename, text, start, end
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Delete Footnote",
            destructiveHint=True,
        ),
        description=live_tools.word_live_delete_footnote.__doc__,
    )
    async def word_live_delete_footnote(
        filename: str = None,
        index: int = None,
    ):
        return await live_tools.word_live_delete_footnote(
            filename, index
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Protect Document",
            destructiveHint=True,
        ),
        description=live_tools.word_live_protect_document.__doc__,
    )
    async def word_live_protect_document(
        filename: str = None,
        protection_type: str = "read_only",
        password: str = None,
    ):
        return await live_tools.word_live_protect_document(
            filename, protection_type, password
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Unprotect Document",
            destructiveHint=True,
        ),
        description=live_tools.word_live_unprotect_document.__doc__,
    )
    async def word_live_unprotect_document(
        filename: str = None,
        password: str = None,
    ):
        return await live_tools.word_live_unprotect_document(
            filename, password
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Convert To PDF",
            destructiveHint=False,
        ),
        description=live_tools.word_live_convert_to_pdf.__doc__,
    )
    async def word_live_convert_to_pdf(
        filename: str = None,
        pdf_path: str = "",
    ):
        return await live_tools.word_live_convert_to_pdf(
            filename, pdf_path
        )

    @mcp.tool(
        annotations=ToolAnnotations(
            title="Word Live Create Document",
            destructiveHint=False,
        ),
        description=live_tools.word_live_create_document.__doc__,
    )
    async def word_live_create_document(
        visible: bool = True,
    ):
        return await live_tools.word_live_create_document(
            visible
        )


def run_server():
    """Run the Word Document MCP Server with configurable transport."""
    # Get transport configuration
    config = get_transport_config()
    
    # Setup logging
    # setup_logging(config['debug'])
    
    # Monkey-patch Document.save() to preserve comments.xml and other custom parts
    from word_document_server.utils.save_utils import install_save_hook
    install_save_hook()

    # Monkey-patch PhysPkgReader to detect Word-locked files
    from word_document_server.utils.path_utils import install_path_hook
    install_path_hook()

    # Register all tools
    register_tools()
    
    # Print startup information
    transport_type = config['transport']
    print(f"Starting Word Document MCP Server with {transport_type} transport...", file=sys.stderr)
    
    # if config['debug']:
    #     print(f"Configuration: {config}")
    
    try:
        if transport_type == 'stdio':
            # Run with stdio transport (default, backward compatible)
            print("Server running on stdio transport", file=sys.stderr)
            mcp.run(transport='stdio')
            
        elif transport_type == 'streamable-http':
            # Run with streamable HTTP transport
            print(f"Server running on streamable-http transport at http://{config['host']}:{config['port']}{config['path']}", file=sys.stderr)
            mcp.run(
                transport='streamable-http',
                host=config['host'],
                port=config['port'],
                path=config['path']
            )
            
        elif transport_type == 'sse':
            # Run with SSE transport
            print(f"Server running on SSE transport at http://{config['host']}:{config['port']}{config['sse_path']}", file=sys.stderr)
            mcp.run(
                transport='sse',
                host=config['host'],
                port=config['port'],
                path=config['sse_path']
            )
            
    except KeyboardInterrupt:
        print("\nShutting down server...", file=sys.stderr)
    except Exception as e:
        print(f"Error starting server: {e}", file=sys.stderr)
        if config['debug']:
            import traceback
            traceback.print_exc()
        sys.exit(1)
    
    return mcp


def main():
    """Main entry point for the server."""
    run_server()


if __name__ == "__main__":
    main()
