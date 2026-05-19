import ast
from pathlib import Path


def _registered_tools():
    main_py = Path("word_document_server/main.py")
    tree = ast.parse(main_py.read_text(encoding="utf-8"))

    for node in ast.walk(tree):
        if not isinstance(node, ast.FunctionDef):
            continue

        for decorator in node.decorator_list:
            if (
                isinstance(decorator, ast.Call)
                and isinstance(decorator.func, ast.Attribute)
                and decorator.func.attr == "tool"
            ):
                yield node.name, [arg.arg for arg in node.args.args]
                break


def test_public_tool_parameter_names_are_agent_friendly():
    confusing_names = {
        "start_pos": "start",
        "end_pos": "end",
        "row_index": "row",
        "col_index": "col",
        "text_to_find": "find_text",
        "text_content": "text",
        "target_paragraph_index": "paragraph_index",
        "search_text": "target_text or find_text",
        "insert_position": "position",
    }

    failures = []
    for tool_name, params in _registered_tools():
        for confusing_name, preferred_name in confusing_names.items():
            if confusing_name in params:
                failures.append(f"{tool_name}: use {preferred_name!r} instead of {confusing_name!r}")

    assert not failures, "\n".join(failures)
