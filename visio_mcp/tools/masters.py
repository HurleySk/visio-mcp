import json

from visio_mcp import stencil_index
from visio_mcp.live_tool import offline_tool


@offline_tool
def visio_find_masters(query: str = "", stencil: str | None = None, limit: int = 25):
    """Search Visio's installed shapes for masters to use in a diagram.

    Matches master names first, then stencil titles and keywords, so "flowchart decision"
    finds BASFLO:Decision. Pass stencil alone to list one stencil's masters. Use the returned
    "master" value in visio_create_diagram and visio_add_elements. Does not need Visio running.
    """
    return json.dumps(stencil_index.find_masters(query, stencil, limit), ensure_ascii=False)


@offline_tool
def visio_list_templates(query: str | None = None):
    """List Visio's installed templates, optionally filtered by words in their title or keywords.

    Pass the returned "template" value to visio_new_drawing or visio_create_diagram.
    Does not need Visio running.
    """
    return json.dumps({"templates": stencil_index.list_templates(query)}, ensure_ascii=False)
