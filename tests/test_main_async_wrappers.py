import ast
from pathlib import Path


ASYNC_TOOL_MODULES = {
    "live_api_tools",
}

EXPECTED_PUBLIC_TOOLS = {
    "word_open",
    "word_save",
    "word_close",
    "word_get_content",
    "word_search",
    "word_edit",
    "word_format",
    "word_comment",
    "word_track_changes",
    "word_table",
    "word_media",
    "word_mutations",
    "word_layout",
    "word_blueprint",
    "word_protection",
}


def _decorated_word_tools():
    tree = ast.parse(Path("word_document_server/main.py").read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
            continue
        if not node.name.startswith("word_"):
            continue
        if any(
            isinstance(decorator, ast.Call)
            and isinstance(decorator.func, ast.Attribute)
            and decorator.func.attr == "tool"
            for decorator in node.decorator_list
        ):
            yield node


def _module_call_name(value):
    if not isinstance(value, ast.Call):
        return None
    if not isinstance(value.func, ast.Attribute):
        return None
    if not isinstance(value.func.value, ast.Name):
        return None
    module_name = value.func.value.id
    if module_name not in ASYNC_TOOL_MODULES:
        return None
    return f"{module_name}.{value.func.attr}"


def test_registered_word_tools_are_async_wrappers():
    tools = list(_decorated_word_tools())
    assert tools, "No decorated word_* tools found in main.py"

    sync_tools = [node.name for node in tools if not isinstance(node, ast.AsyncFunctionDef)]
    assert not sync_tools, f"Registered wrappers must be async: {sync_tools}"


def test_public_tool_surface_is_grouped_only():
    tool_names = {node.name for node in _decorated_word_tools()}

    assert tool_names == EXPECTED_PUBLIC_TOOLS
    assert not any(name.startswith("word_live_") for name in tool_names)
    assert "word_screen_capture" not in tool_names


def test_registered_word_tools_await_async_implementations():
    failures = []
    for node in _decorated_word_tools():
        for child in ast.walk(node):
            if not isinstance(child, ast.Return) or child.value is None:
                continue
            if _module_call_name(child.value):
                failures.append(f"{node.name} returns un-awaited {_module_call_name(child.value)}")
            if isinstance(child.value, ast.Await):
                target = _module_call_name(child.value.value)
                if target is None:
                    continue

    assert not failures, "\n".join(failures)
