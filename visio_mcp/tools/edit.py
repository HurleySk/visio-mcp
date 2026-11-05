"""Edit tools. Each call is one Ctrl+Z step, and a failure part-way rolls the whole call back.
Everything a call names is found before the first change, so a bad reference changes nothing."""
import difflib
import json

from visio_mcp.addressing import ShapeFinder, find_page, identity
from visio_mcp.building import (
    VIS_SEL_EMPTY,
    VIS_SELECT,
    MasterResolver,
    StencilPool,
    add_to_page,
    drawing_units,
    page_style,
    set_data,
    set_fill,
)
from visio_mcp.com_runtime import classify_live, current_session, drawings_in, edit_doc, find_document
from visio_mcp.errors import VisioError
from visio_mcp.live_tool import live_tool
from visio_mcp.models import EdgeSpec, LayoutSpec, NodeSpec, ShapeUpdate, check_graph, check_updates, resolve_layout
from visio_mcp.tools.diagnose import report
from visio_mcp.tools.read import container_members, page_info

VIS_GLUED_SHAPES_ALL_1D = 0
VIS_TYPE_PAGE = 1
VIS_DELETE_NO_RENUMBER = 0
CONTAINER_MASTER = "SDCONT:Plain"
MAX_STEPS = 99


def _dumps(payload):
    return json.dumps(payload, ensure_ascii=False)


def _target(drawing, page):
    document = find_document(current_session().app, drawing)
    return document, find_page(document, page)


def _find_all(page, references):
    """Each shape once, in the order first named."""
    if not references:
        raise ValueError("Pass at least one shape.")
    finder = ShapeFinder(page)
    found = {}
    for reference in references:
        shape = finder.find(reference)
        found.setdefault(shape.ID, shape)
    return list(found.values())


def _top_level(shapes, doing):
    for shape in shapes:
        parent = shape.ContainingShape
        if parent.Type != VIS_TYPE_PAGE:
            raise VisioError(
                "inside_group",
                f"Shape {shape.ID} is part of group {parent.ID}, so it can't be {doing} on its own",
                hint="Use the group's ID, or ungroup it in Visio first.",
            )


def _page_names(document):
    pages = document.Pages
    return [pages(index).Name for index in range(1, pages.Count + 1)]


def _free_name(document, name, keep=None):
    """A page name no other page has. Visio refuses a taken name with a vague error."""
    name = name.strip()
    if not name:
        raise ValueError("A page name can't be empty.")
    pages = document.Pages
    for index in range(1, pages.Count + 1):
        page = pages(index)
        if keep is not None and page.Index == keep.Index:
            continue
        if name.casefold() in (page.Name.casefold(), page.NameU.casefold()):
            raise VisioError(
                "page_exists",
                f"Drawing '{document.Name}' already has a page called '{page.Name}'",
                hint="Pick another name.",
            )
    return name


def _existing_end(finder, end, keys):
    try:
        return finder.find(end)
    except VisioError as exc:
        close = difflib.get_close_matches(end, list(keys), n=5, cutoff=0.5)
        if exc.code != "shape_not_found" or not close:
            raise
        raise VisioError(
            "unknown_node",
            f"An edge refers to '{end}', which is neither a new node's key nor a shape on the page",
            details=close,
        ) from exc


@live_tool(mutates=True, timeout=120.0)
def visio_add_elements(
    nodes: list[NodeSpec] | None = None,
    edges: list[EdgeSpec] | None = None,
    layout: LayoutSpec | None = None,
    drawing: str | None = None,
    page: int | str | None = None,
):
    """Add shapes and connectors to an existing page, as one Ctrl+Z step.

    nodes are as in visio_create_diagram. An edge's from and to name a new node's key, else an
    existing shape's ID, name or exact text, so new nodes can join what's already there. Keys
    win over existing shapes.

    Placing: with x and y on every node, nothing else moves. Without them, the whole page is laid
    out again in the style it was last laid out with (else top_down), which moves existing shapes
    too. Edges alone move nothing unless layout is passed.

    Returns ids mapping your keys to shape IDs, the connectors, and issues found on the page.
    Look at the result with visio_get_page_image.
    """
    nodes = nodes or []
    edges = edges or []
    if not nodes and not edges:
        raise ValueError("Pass nodes, edges or both.")
    check_graph(nodes, edges, existing=True)
    session = current_session()
    document, target = _target(drawing, page)
    if layout is None and not nodes:
        layout = LayoutSpec(style="none")
    elif layout is None and all(node.x is None and node.y is None for node in nodes):
        layout = LayoutSpec(style=page_style(target))
    layout = resolve_layout(layout, nodes)
    keys = {node.key for node in nodes}
    finder = ShapeFinder(target)
    existing = {}
    for edge in edges:
        for end in (edge.from_, edge.to):
            if end not in keys and end not in existing:
                existing[end] = _existing_end(finder, end, keys)
    app = document.Application
    refs = MasterResolver(app, document, drawing_units(document)).resolve_all(nodes)
    edit_doc(drawing)
    ids, connectors = add_to_page(document, target, nodes, edges, refs, layout, session.warnings, existing)
    return _dumps({"drawing": document.Name, "page": page_info(target), "ids": ids, "connectors": connectors, **report(target)})


def _apply(shape, update):
    for cell, value in (("PinX", update.x), ("PinY", update.y), ("Width", update.w), ("Height", update.h)):
        if value is not None:
            shape.CellsU(cell).ResultIU = value
    if update.text is not None:
        shape.Text = update.text
    if update.data:
        set_data(shape, update.data)
    if update.fill:
        set_fill(shape, update.fill)


@live_tool(mutates=True)
def visio_update_shapes(updates: list[ShapeUpdate], drawing: str | None = None, page: int | str | None = None):
    """Change shapes' text, Shape Data, position, size or fill, as one Ctrl+Z step.

    Each update names a shape by ID, else exact name, else exact text, and sets only the fields
    given. x and y are the centre in inches from the page's bottom-left, and glued connectors
    follow a moved shape. Every shape is found before anything changes.

    Returns each shape's ID, name, text and master afterwards. Check the ID and text are the ones
    you meant, since Visio can reuse an ID after a delete.
    """
    check_updates(updates)
    document, target = _target(drawing, page)
    finder = ShapeFinder(target)
    found = [finder.find(update.shape) for update in updates]
    edit_doc(drawing)
    for shape, update in zip(found, updates):
        _apply(shape, update)
    once = {shape.ID: shape for shape in found}
    return _dumps({"drawing": document.Name, "page": target.Name, "updated": [identity(shape) for shape in once.values()]})


@live_tool(mutates=True)
def visio_delete_shapes(shapes: list[int | str], drawing: str | None = None, page: int | str | None = None):
    """Delete shapes, and the connectors glued to them, as one Ctrl+Z step.

    shapes are IDs, else exact names, else exact text. Deleting a container keeps the shapes in
    it, and deleting a group deletes the group's own members. Returns what was deleted, with the
    connectors that went with it listed apart.
    """
    document, target = _target(drawing, page)
    found = _find_all(target, shapes)
    _top_level(found, "deleted")
    chosen = {shape.ID for shape in found}
    glued = {}
    for shape in found:
        if shape.OneD:
            continue
        for connector_id in shape.GluedShapes(VIS_GLUED_SHAPES_ALL_1D, "", None) or ():
            if connector_id not in chosen and connector_id not in glued:
                glued[connector_id] = target.Shapes.ItemFromID(connector_id)
    deleted = [identity(shape) for shape in found]
    connectors = [identity(shape) for shape in glued.values()]
    edit_doc(drawing)
    selection = target.CreateSelection(VIS_SEL_EMPTY)
    for shape in [*found, *glued.values()]:
        selection.Select(shape, VIS_SELECT)
    selection.Delete()
    return _dumps({"drawing": document.Name, "page": page_info(target), "deleted": deleted, "connectors_deleted": connectors})


@live_tool(mutates=True)
def visio_add_container(
    shapes: list[int | str],
    title: str,
    master: str | None = None,
    drawing: str | None = None,
    page: int | str | None = None,
):
    """Put a titled container around shapes, as one Ctrl+Z step.

    shapes are IDs, else exact names, else exact text. The container is sized to fit them, and
    moving it in Visio moves them too. master defaults to 'SDCONT:Plain'; find other styles with
    visio_find_masters("container"). Returns the container and the shape IDs it holds.
    """
    document, target = _target(drawing, page)
    found = _find_all(target, shapes)
    _top_level(found, "put in a container")
    app = document.Application
    ref = MasterResolver(app, document, drawing_units(document)).resolve(master or CONTAINER_MASTER)
    session = current_session()
    edit_doc(drawing)
    pool = StencilPool(app, document)
    try:
        selection = target.CreateSelection(VIS_SEL_EMPTY)
        for shape in found:
            selection.Select(shape, VIS_SELECT)
        container = pool.drop_container(target, ref, selection)
        warning = pool.skipped_warning()
        if warning:
            session.warnings.append(warning)
    finally:
        pool.close()
    contains = container_members(container)
    if contains is None:
        raise VisioError(
            "not_a_container",
            f"Master '{ref.label}' isn't a container",
            hint='Find container masters with visio_find_masters("container").',
        )
    container.Text = title
    return _dumps(
        {
            "drawing": document.Name,
            "page": page_info(target),
            "container": identity(container),
            "contains": contains,
            **report(target),
        }
    )


@live_tool(mutates=True)
def visio_add_page(name: str | None = None, drawing: str | None = None):
    """Add a page to a drawing, as one Ctrl+Z step. Pass its name as page in later calls."""
    document = find_document(current_session().app, drawing)
    if name is not None:
        name = _free_name(document, name)
    edit_doc(drawing)
    added = document.Pages.Add()
    if name:
        added.Name = name
    return _dumps({"drawing": document.Name, "page": page_info(added), "index": added.Index, "pages": _page_names(document)})


def _theme_value(theme):
    if theme is None:
        return None
    if isinstance(theme, bool) or (isinstance(theme, int) and theme < 0):
        raise ValueError(f"theme must be a name or an index of 0 or more, not {theme!r}.")
    if isinstance(theme, str):
        theme = theme.strip()
        if not theme:
            raise ValueError("theme can't be empty. Use 'None' to remove the page's theme.")
    return theme


def _set_theme(page, theme):
    try:
        page.SetTheme(theme)
    except Exception as exc:
        if classify_live(exc).code in ("busy", "disconnected"):
            raise
        raise VisioError(
            "unknown_theme",
            f"Visio has no theme {theme!r}",
            hint="Use a name such as Office, Linear or Facet, or an index from 0 to 60, where 0 or 'None' removes the theme.",
        ) from exc


def _theme_index(page):
    sheet = page.PageSheet
    if not sheet.CellExistsU("ThemeIndex", 0):
        return None
    return int(sheet.CellsU("ThemeIndex").ResultIU)


@live_tool(mutates=True)
def visio_update_page(
    page: int | str | None = None,
    name: str | None = None,
    width: float | None = None,
    height: float | None = None,
    fit: bool | None = None,
    theme: str | int | None = None,
    drawing: str | None = None,
):
    """Rename a page, set its size or theme, or fit it to its shapes, as one Ctrl+Z step.

    width and height are in inches. fit sets the size from the shapes, so it can't go with them.
    theme is a name such as Office, Linear, Simple, Facet, Ion, Slice, Retrospect, Whisp, Organic,
    Pencil or Whiteboard, ignoring case, or an index from Visio's gallery. 'None' or 0 removes it.
    Returns the page with its theme index, and issues found on it.
    """
    if name is None and width is None and height is None and not fit and theme is None:
        raise ValueError("Pass at least one of name, width, height, fit or theme.")
    if fit and (width is not None or height is not None):
        raise ValueError("fit sets the size from the shapes, so it can't go with width or height.")
    for size in (width, height):
        if size is not None and size <= 0:
            raise ValueError("width and height must be greater than 0.")
    theme = _theme_value(theme)
    document, target = _target(drawing, page)
    if name is not None:
        name = _free_name(document, name, keep=target)
    edit_doc(drawing)
    if name is not None:
        target.Name = name
    sheet = target.PageSheet
    if width is not None:
        sheet.CellsU("PageWidth").ResultIU = width
    if height is not None:
        sheet.CellsU("PageHeight").ResultIU = height
    if fit:
        target.ResizeToFitContents()
    if theme is not None:
        _set_theme(target, theme)
    return _dumps({"drawing": document.Name, "page": {**page_info(target), "theme": _theme_index(target)}, **report(target)})


@live_tool(mutates=True)
def visio_delete_page(page: int | str, drawing: str | None = None):
    """Delete a page and everything on it, as one Ctrl+Z step.

    page is a name or a 1-based number. A drawing's only foreground page is refused, because
    Visio would wipe it and make a blank one.
    """
    document, target = _target(drawing, page)
    if not target.Background:
        pages = document.Pages
        foreground = sum(1 for index in range(1, pages.Count + 1) if not pages(index).Background)
        if foreground == 1:
            raise VisioError(
                "last_page",
                f"Page '{target.Name}' is the only page in '{document.Name}'",
                hint="Add another page first, or delete its shapes with visio_delete_shapes.",
            )
    name = target.Name
    edit_doc(drawing)
    target.Delete(VIS_DELETE_NO_RENUMBER)
    return _dumps({"drawing": document.Name, "deleted": name, "pages": _page_names(document)})


def _counts(app):
    rows = {}
    for document in drawings_in(app):
        pages = document.Pages
        rows[document.FullName] = {
            "drawing": document.Name,
            "pages": pages.Count,
            "shapes": sum(pages(index).Shapes.Count for index in range(1, pages.Count + 1)),
        }
    return rows


def _step(action, steps):
    if isinstance(steps, bool) or not 1 <= steps <= MAX_STEPS:
        raise ValueError(f"steps must be from 1 to {MAX_STEPS}.")
    session = current_session()
    app = session.app
    # Undoing inside someone else's open scope would fold the undo into it.
    session.close_leftover_scope(app)
    before = _counts(app)
    for _ in range(steps):
        getattr(app, action)()
    rows = []
    for key, now in _counts(app).items():
        then = before.get(key, {"pages": None, "shapes": None})
        rows.append(
            {
                "drawing": now["drawing"],
                "shapes_before": then["shapes"],
                "shapes_after": now["shapes"],
                "pages_before": then["pages"],
                "pages_after": now["pages"],
            }
        )
    return _dumps({"steps": steps, "drawings": rows})


@live_tool(mutates=True)
def visio_undo(steps: int = 1):
    """Undo the last steps in Visio, like pressing Ctrl+Z. Each edit tool call is one step.

    Visio keeps one history for all its drawings, and the user's own edits are in it, so a step
    can land in another drawing or take back something the user did. The shape and page counts
    per drawing, before and after, show where it landed. They don't show text or position
    changes, so check with visio_get_page. Steps beyond the start of the history do nothing.
    """
    return _step("Undo", steps)


@live_tool(mutates=True)
def visio_redo(steps: int = 1):
    """Redo steps undone in Visio, like pressing Ctrl+Y.

    Visio keeps one history for all its drawings, so a step can land in another drawing. The
    shape and page counts per drawing, before and after, show where it landed. They don't show
    text or position changes, so check with visio_get_page. Steps with nothing to redo do nothing.
    """
    return _step("Redo", steps)
