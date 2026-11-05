import json

from visio_mcp.addressing import ShapeFinder, find_page, identity, master_name, summarize
from visio_mcp.com_runtime import VIS_TYPE_STENCIL, classify_live, current_session, drawings_in, find_document
from visio_mcp.live_tool import live_tool

VIS_GET_GUID = 0
VIS_SECTION_PROP = 243
VIS_CUST_PROP_VALUE = 0
VIS_CUST_PROP_LABEL = 2
VIS_TYPE_GROUP = 2
VIS_CONTAINER_FLAGS_DEFAULT = 0
VIS_CONNECTED_INCOMING = 1
VIS_CONNECTED_OUTGOING = 2
MAX_PAGE_ROWS = 500
MAX_FOUND = 200


def _dumps(payload):
    return json.dumps(payload, ensure_ascii=False)


def _inches(shape, cell):
    return round(shape.CellsU(cell).ResultIU, 4)


def page_info(page):
    sheet = page.PageSheet
    return {
        "name": page.Name,
        "width": _inches(sheet, "PageWidth"),
        "height": _inches(sheet, "PageHeight"),
        "shapes": page.Shapes.Count,
    }


def connector_ends(shape):
    ends = {"from": None, "to": None}
    connects = shape.Connects
    for index in range(1, connects.Count + 1):
        connect = connects(index)
        end = {"BeginX": "from", "EndX": "to"}.get(connect.FromCell.Name)
        if end:
            ends[end] = connect.ToSheet.ID
    return ends


def _geometry(shape):
    return {
        "x": _inches(shape, "PinX"),
        "y": _inches(shape, "PinY"),
        "w": _inches(shape, "Width"),
        "h": _inches(shape, "Height"),
    }


def container_members(shape):
    """The IDs a container holds, or None if it isn't one. A container is a group to Visio."""
    properties = shape.ContainerProperties
    if properties is None:
        return None
    return list(properties.GetMemberShapes(VIS_CONTAINER_FLAGS_DEFAULT) or ())


def _listing(shape):
    row = identity(shape)
    if shape.OneD:
        row["kind"] = "connector"
        row.update(connector_ends(shape))
        return row
    contains = container_members(shape)
    if contains is not None:
        row["kind"] = "container"
        row["contains"] = contains
        row.update(_geometry(shape))
    elif shape.Type == VIS_TYPE_GROUP:
        row["kind"] = "group"
        row["members"] = shape.Shapes.Count
        row.update(_geometry(shape))
    else:
        row["kind"] = "shape"
        row.update(_geometry(shape))
    return row


def shape_data(shape):
    """Shape Data rows as (label or row name, row name, value). An unset row has value None.

    Visio reads a row with no formula back as "0.0000", and masters ship some rows holding an
    empty string, so both count as unset."""
    if not shape.SectionExists(VIS_SECTION_PROP, 0):
        return []
    rows = []
    for row in range(shape.RowCount(VIS_SECTION_PROP)):
        value = shape.CellsSRC(VIS_SECTION_PROP, row, VIS_CUST_PROP_VALUE)
        name = value.RowNameU
        label = shape.CellsSRC(VIS_SECTION_PROP, row, VIS_CUST_PROP_LABEL).ResultStr("")
        text = None if value.FormulaU == "" else value.ResultStr("")
        rows.append((label or name, name, text or None))
    return rows


def _cell(shape, name):
    if not shape.CellExistsU(name, 0):
        return None
    cell = shape.CellsU(name)
    return {"formula": cell.FormulaU, "value": cell.ResultIU, "text": cell.ResultStr("")}


def _ids(shapes):
    return list(shapes) if shapes else []


@live_tool(mutates=False, start="report")
def visio_list_open():
    """List the drawings open in Visio, saying which is active. Never starts Visio.

    Reports running=false if Visio isn't running. Includes saved drawings in other Visio
    windows where Windows lets them be reached.
    """
    session = current_session()
    error = session.attach_error
    if error is not None:
        payload = {"running": error.code != "visio_not_running", "reachable": False, "drawings": []}
        if error.code != "visio_not_running":
            payload["hint"] = error.hint
        return _dumps(payload)
    app = session.app
    try:
        active = app.ActiveDocument.FullName
    except Exception as exc:
        if classify_live(exc).code in ("busy", "disconnected"):
            raise
        active = None
    seen = set()
    rows = []
    for document in drawings_in(app) + session.connector.other_documents():
        key = document.FullName.casefold()
        if key not in seen and document.Type != VIS_TYPE_STENCIL:
            seen.add(key)
            rows.append({**summarize(document), "active": document.FullName == active})
    return _dumps({"running": True, "reachable": True, "drawings": rows})


@live_tool(mutates=False)
def visio_get_page(offset: int = 0, limit: int = 100, drawing: str | None = None, page: int | str | None = None):
    """List a page's top-level shapes and connectors, one page of results at a time.

    Shapes have centre x, y and size w, h in inches, from the bottom-left of the page.
    Connectors say which shape IDs they join in "from" and "to", where null means not glued.
    Containers list the shape IDs they hold in "contains".
    Pass next_offset back as offset to continue. Use visio_get_shape for a shape's data.
    """
    if offset < 0 or limit < 1:
        raise ValueError("offset can't be negative and limit must be at least 1.")
    limit = min(limit, MAX_PAGE_ROWS)
    document = find_document(current_session().app, drawing)
    target = find_page(document, page)
    shapes = target.Shapes
    total = shapes.Count
    end = min(offset + limit, total)
    return _dumps(
        {
            "drawing": document.Name,
            "page": page_info(target),
            "total": total,
            "offset": offset,
            "next_offset": end if end < total else None,
            "shapes": [_listing(shapes(index + 1)) for index in range(offset, end)],
        }
    )


@live_tool(mutates=False)
def visio_get_shape(
    shape: int | str,
    cells: list[str] | None = None,
    drawing: str | None = None,
    page: int | str | None = None,
):
    """Read one shape in detail: position, Shape Data, connections, containers it's in, what it holds.

    shape is an ID, else an exact name, else exact text. cells adds ShapeSheet cells by name,
    such as "FillForegnd", with null for a cell the shape doesn't have.
    """
    target = ShapeFinder(find_page(find_document(current_session().app, drawing), page)).find(shape)
    result = _listing(target)
    guid = target.UniqueID(VIS_GET_GUID)
    if guid:
        result["guid"] = guid
    data = shape_data(target)
    result["data"] = {label: value for label, _, value in data if value is not None}
    result["unset_data"] = [label for label, _, value in data if value is None]
    if not target.OneD:
        result["incoming"] = _ids(target.ConnectedShapes(VIS_CONNECTED_INCOMING, ""))
        result["outgoing"] = _ids(target.ConnectedShapes(VIS_CONNECTED_OUTGOING, ""))
    result["containers"] = _ids(target.MemberOfContainers)
    if result["kind"] == "group":
        members = target.Shapes
        result["member_shapes"] = [identity(members(index)) for index in range(1, members.Count + 1)]
    if cells:
        result["cells"] = {name: _cell(target, name) for name in cells}
    return _dumps(result)


def _data_matches(rows, wanted):
    for key, value in wanted.items():
        hit = any(
            key.casefold() in (label.casefold(), name.casefold())
            and found is not None
            and str(value).casefold() == found.casefold()
            for label, name, found in rows
        )
        if not hit:
            return False
    return True


def _master_matches(shape, wanted):
    name = master_name(shape)
    return name is not None and wanted.rpartition(":")[2].casefold() == name.casefold()


@live_tool(mutates=False)
def visio_find_shapes(
    text: str | None = None,
    master: str | None = None,
    data: dict[str, str] | None = None,
    drawing: str | None = None,
    page: int | str | None = None,
):
    """Find a page's top-level shapes by text (part of it, ignoring case), master name, or Shape Data.

    Filters combine. data matches a field's label or row name against an exact value,
    ignoring case. Returns at most 200 shapes with the total found.
    """
    if not text and not master and not data:
        raise ValueError("Pass at least one of text, master or data.")
    target = find_page(find_document(current_session().app, drawing), page)
    needle = text.casefold() if text else None
    shapes = target.Shapes
    found = []
    total = 0
    for index in range(1, shapes.Count + 1):
        shape = shapes(index)
        if needle and needle not in shape.Text.casefold():
            continue
        if master and not _master_matches(shape, master):
            continue
        if data and not _data_matches(shape_data(shape), data):
            continue
        total += 1
        if len(found) < MAX_FOUND:
            found.append(identity(shape))
    return _dumps({"page": target.Name, "total": total, "shapes": found})


@live_tool(mutates=False)
def visio_get_selection():
    """Read what's selected in the active Visio window, so you can act on what the user points at."""
    app = current_session().app
    try:
        window = app.ActiveWindow
        selection = window.Selection
        count = selection.Count
        document = window.Document.Name
        page = window.Page.Name
    except Exception as exc:
        if classify_live(exc).code in ("busy", "disconnected"):
            raise
        return _dumps({"drawing": None, "page": None, "shapes": []})
    return _dumps(
        {
            "drawing": document,
            "page": page,
            "shapes": [identity(selection.Item(index)) for index in range(1, count + 1)],
        }
    )
