import os
import sys

os.environ.setdefault("FASTMCP_LOG_LEVEL", "WARNING")

from fastmcp import FastMCP
from mcp.types import ToolAnnotations

mcp = FastMCP("visio-mcp")
_registered = False


def _add(tool, title, destructive=False):
    if tool.mutates:
        hints = {"readOnlyHint": False, "destructiveHint": destructive}
    else:
        hints = {"readOnlyHint": True}
    mcp.tool(annotations=ToolAnnotations(title=title, **hints))(tool)


def register_tools():
    global _registered
    if _registered:
        return
    _registered = True
    from visio_mcp.tools import diagnose, files, generate, image, masters, read

    _add(generate.visio_create_diagram, "Create Visio Diagram")
    _add(generate.visio_layout_page, "Lay Out Visio Page")
    _add(diagnose.visio_diagnose_page, "Diagnose Visio Page")
    _add(masters.visio_find_masters, "Find Visio Masters")
    _add(masters.visio_list_templates, "List Visio Templates")
    _add(read.visio_list_open, "List Open Visio Drawings")
    _add(read.visio_get_page, "Read Visio Page")
    _add(read.visio_get_shape, "Read Visio Shape")
    _add(read.visio_find_shapes, "Find Visio Shapes")
    _add(read.visio_get_selection, "Read Visio Selection")
    _add(image.visio_get_page_image, "Get Visio Page Image")
    _add(files.visio_new_drawing, "New Visio Drawing")
    _add(files.visio_open_drawing, "Open Visio Drawing")
    _add(files.visio_save, "Save Visio Drawing", destructive=True)
    _add(files.visio_export, "Export Visio Drawing", destructive=True)
    _add(files.visio_close_drawing, "Close Visio Drawing", destructive=True)


def run():
    if sys.platform != "win32":
        print("visio-mcp only runs on Windows", file=sys.stderr)
        raise SystemExit(1)
    register_tools()
    mcp.run(transport="stdio", show_banner=False)
