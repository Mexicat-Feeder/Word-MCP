import ast
import inspect
from pathlib import Path
from word_document_server.tools import live_v2_tools


def _get_v2_wrappers_ast():
    main_py = Path("word_document_server/main.py")
    tree = ast.parse(main_py.read_text(encoding="utf-8"))
    for parent in ast.walk(tree):
        for child in ast.iter_child_nodes(parent):
            child.parent = parent
    
    wrappers = {}
    for node in ast.walk(tree):
        if isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)) and node.name.startswith("word_v2_"):
            # Extract arguments
            args = [arg.arg for arg in node.args.args]
            wrappers[node.name] = {
                "params": args,
                "node": node
            }
    return wrappers


def test_v2_wrapper_signatures_match_implementations():
    wrappers = _get_v2_wrappers_ast()
    assert len(wrappers) > 0, "No word_v2_* wrappers found in main.py"
    
    for tool_name, info in wrappers.items():
        assert isinstance(info["node"], ast.AsyncFunctionDef), f"{tool_name} wrapper must be async"

        # 1. Verify implementation exists
        assert hasattr(live_v2_tools, tool_name), f"live_v2_tools is missing implementation for {tool_name}"
        impl_func = getattr(live_v2_tools, tool_name)
        
        # 2. Get implementation signature
        impl_sig = inspect.signature(impl_func)
        impl_params = list(impl_sig.parameters.keys())
        
        wrapper_params = info["params"]
        
        # 3. The public MCP wrapper and implementation must use the same
        # names in the same order. v2 mutations call implementations directly,
        # so any alias or order drift creates a second hidden schema.
        assert wrapper_params == impl_params, (
            f"Wrapper signature for {tool_name} does not match implementation.\n"
            f"wrapper: {wrapper_params}\n"
            f"impl:    {impl_params}"
        )


def test_v2_wrappers_use_only_keyword_calls():
    wrappers = _get_v2_wrappers_ast()
    
    for tool_name, info in wrappers.items():
        node = info["node"]
        
        # Find the call to live_v2_tools.<tool_name>
        calls = []
        for child in ast.walk(node):
            if isinstance(child, ast.Call):
                # Check if it calls live_v2_tools.<name> or similar
                if (
                    isinstance(child.func, ast.Attribute)
                    and child.func.attr == tool_name
                ):
                    calls.append(child)
        
        assert len(calls) == 1, f"Expected exactly 1 call to live_v2_tools.{tool_name} in {tool_name} wrapper, found {len(calls)}"
        call = calls[0]
        assert isinstance(getattr(call, "parent", None), ast.Await), f"{tool_name} must await live_v2_tools.{tool_name}"
        
        # Verify that all arguments are passed as keywords (i.e. no positional arguments)
        assert len(call.args) == 0, (
            f"Wrapper {tool_name} passes positional arguments to implementation: {[ast.unparse(a) for a in call.args]}. "
            f"Must use keyword arguments exclusively to prevent parameter shifting bugs."
        )
