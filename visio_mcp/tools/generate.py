import json
from typing import Literal

from visio_mcp.addressing import ShapeFinder, find_page, summarize
from visio_mcp.building import MasterResolver, StencilPool, drawing_units, drop_node, glue_edge, lay_out
from visio_mcp.com_runtime import current_session, edit_doc, find_document
from visio_mcp.live_tool import live_tool
from visio_mcp.models import DiagramSpec, check_graph, check_spacing, resolve_layout
from visio_mcp.tools.diagnose import report
from visio_mcp.tools.files import new_drawing, save_target, template_source
from visio_mcp.tools.image import render_page
from visio_mcp.tools.read import page_info


def _check_masters_early(app, spec):
    """Fail before the drawing exists where possible. A template's stencils aren't open until it
    is, so with a template only names that say their stencil can be checked yet."""
    resolver = MasterResolver(app, units=spec.units or "us", within=spec.stencils)
    nodes = spec.nodes
    if spec.template:
        nodes = [node for node in nodes if ":" in node.master]
    resolver.resolve_all(nodes)


def _close_new_drawing(session, document):
    name = document.Name
    try:
        document.Saved = True
        document.Close()
    except Exception:
        session.warnings.append(
            {"code": "new_drawing_left_open", "message": f"Couldn't close the new drawing '{name}'. It's empty, and safe to close."}
        )
    else:
        session.warnings.append({"code": "new_drawing_closed", "message": f"Closed the new drawing '{name}', so nothing is left over."})


def _build(document, page, spec, layout):
    refs = MasterResolver(document.Application, document, drawing_units(document), spec.stencils).resolve_all(spec.nodes)
    pool = StencilPool(document.Application, document)
    try:
        placeable = layout.style != "none"
        shapes = {
            node.key: drop_node(page, pool, node, refs[node.master], index, placeable)
            for index, node in enumerate(spec.nodes)
        }
        warning = pool.skipped_warning()
        if warning:
            current_session().warnings.append(warning)
    finally:
        pool.close()
    connectors = [
        {"id": glue_edge(page, shapes[edge.from_], shapes[edge.to], edge).ID, "from": edge.from_, "to": edge.to}
        for edge in spec.edges
    ]
    if placeable:
        lay_out(page, layout.style, layout.route, layout.spacing)
    fit = layout.fit_page if layout.fit_page is not None else placeable
    if fit:
        page.ResizeToFitContents()
    return {key: shape.ID for key, shape in shapes.items()}, connectors


@live_tool(mutates=True, start="if_missing", timeout=120.0)
def visio_create_diagram(spec: DiagramSpec):
    """Build a whole diagram in a new drawing in one call: shapes, glued connectors, layout and a picture.

    This is the way to make a diagram. Find masters with visio_find_masters first, then pass
    them all at once. It starts a visible Visio if none is running. A bad spec changes nothing,
    and one Ctrl+Z removes everything it adds. To add to an existing drawing, use visio_add_elements.

    Returns the drawing, its page, ids mapping your keys to shape IDs, the connectors, and
    issues found on the page, with a PNG to check the result by eye.
    """
    check_graph(spec.nodes, spec.edges)
    layout = resolve_layout(spec.layout, spec.nodes)
    source = template_source(spec.template, spec.units)
    target = save_target(spec.save_as, spec.overwrite) if spec.save_as else None
    session = current_session()
    app = session.app
    _check_masters_early(app, spec)

    document = new_drawing(app, source, spec.units)
    session.on_failure(lambda: _close_new_drawing(session, document))
    session.open_scope(app)
    page = document.Pages(1)
    ids, connectors = _build(document, page, spec, layout)
    issues = report(page)
    if target:
        document.SaveAs(target)
    summary = summarize(document)
    payload = {"drawing": summary["drawing"], "path": summary["path"], "page": page.Name, "ids": ids, "connectors": connectors, **issues}
    if not spec.return_image:
        return json.dumps(payload, ensure_ascii=False)
    from fastmcp.utilities.types import Image

    _, picture = render_page(document, page)
    return [payload, Image(data=picture, format="png")]


@live_tool(mutates=True)
def visio_layout_page(
    style: Literal["top_down", "left_right", "radial", "circular"] = "top_down",
    route: Literal["flowchart", "tree", "right_angle", "straight"] | None = None,
    spacing: float | None = None,
    shapes: list[int | str] | None = None,
    fit: bool | None = None,
    drawing: str | None = None,
    page: int | str | None = None,
):
    """Lay out a page again, or only some of its shapes. Connectors stay glued, and one Ctrl+Z undoes it.

    top_down and left_right give the cleanest results. radial and circular are weaker. route
    defaults to flowchart, or straight for radial and circular. spacing is the gap in inches.
    shapes limits it to those shapes, as IDs, names or exact text. fit resizes the page to its
    shapes, by default only when laying out the whole page. Shapes with no connections stay where
    they are unless they're marked placeable, so check the returned issues for overlaps.
    """
    check_spacing(spacing)
    target = find_page(find_document(current_session().app, drawing), page)
    finder = ShapeFinder(target)
    found = [finder.find(reference) for reference in shapes] if shapes else None
    edit_doc(drawing)
    lay_out(target, style, route, spacing, found)
    if fit if fit is not None else not shapes:
        target.ResizeToFitContents()
    return json.dumps({"page": page_info(target), **report(target)}, ensure_ascii=False)
