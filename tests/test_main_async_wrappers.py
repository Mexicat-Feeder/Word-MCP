import ast
from pathlib import Path


ASYNC_TOOL_MODULES = {
    "live_v2_tools",
    "live_tools",
    "live_read_tools",
    "live_layout_tools",
    "screen_capture_tools",
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
