import os
import sys

os.environ.setdefault("FASTMCP_LOG_LEVEL", "WARNING")

from fastmcp import FastMCP

mcp = FastMCP("visio-mcp")


def register_tools():
    pass


def run():
    if sys.platform != "win32":
        print("visio-mcp only runs on Windows", file=sys.stderr)
        raise SystemExit(1)
    register_tools()
    mcp.run(transport="stdio", show_banner=False)
