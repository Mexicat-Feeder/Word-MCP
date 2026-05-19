"""Default configuration values, overridable via environment variables."""

import os

try:
    from dotenv import load_dotenv

    load_dotenv()
except Exception:
    # Keep defaults usable even if python-dotenv is unavailable in a minimal env.
    pass

DEFAULT_AUTHOR = os.environ.get("MCP_AUTHOR", "Author")
DEFAULT_INITIALS = os.environ.get("MCP_AUTHOR_INITIALS", "")
