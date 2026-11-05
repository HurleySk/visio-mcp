"""Find masters, drop shapes, glue connectors and lay out a page. Runs on the COM thread."""
import difflib
import logging
import re

from visio_mcp import stencil_index
from visio_mcp.com_runtime import VIS_TYPE_STENCIL
from visio_mcp.errors import VisioError
from visio_mcp.events import events_off
from visio_mcp.stencil_index import MasterRef, stencil_stem
from visio_mcp.tools.read import VIS_SECTION_PROP, shape_data

logger = logging.getLogger(__name__)

VIS_OPEN_RO_HIDDEN = 2 + 64
VIS_TAG_DEFAULT = 0
VIS_SEL_EMPTY = 0
VIS_SELECT = 2
METRIC_UNITS = (69, 70, 71, 72)
FILLED_ARROW = "13"
NO_ARROW = "0"
PLACEABLE = "1"
PLACE_STYLES = {"top_down": 1, "left_right": 2, "radial": 3, "circular": 6}
NORTH_SOUTH = {"flowchart": 5, "tree": 7, "right_angle": 1, "straight": 2}
WEST_EAST = {"flowchart": 6, "tree": 8, "right_angle": 1, "straight": 2}
# Layout orders siblings by where they start, so a diagonal keeps them in the caller's order.
START_X = 1.0
START_Y = 10.0
START_STEP = 0.25
VIS_TYPE_GROUP = 2
# Drop handlers that can open something and block the call. Others, such as
# SETF(GetRef(HideText),FALSE), finish setting the shape up, so they have to keep running.
INTERACTIVE_HANDLER = re.compile(
    "RUNADDON|DOCMD|OPENTEXTWIN|OPENSHEET|OPENGROUPWIN|OPENFILE|OPENPAGE|CALLTHIS|RUNMACRO|GOTOPAGE|HYPERLINK",
    re.IGNORECASE,
)


def drawing_units(document):
    units = document.Pages(1).PageSheet.CellsU("PageWidth").Units
    return "metric" if units in METRIC_UNITS else "us"


def _master_names(document):
    masters = document.Masters
    rows = []
    for index in range(1, masters.Count + 1):
        master = masters(index)
        rows.append((master.NameU, master.Name))
    return rows


def _named(rows, wanted):
    wanted = wanted.casefold()
    for name, local_name in rows:
        if wanted in (name.casefold(), local_name.casefold()):
            return name
    return None


class MasterResolver:
    """Turn the masters a caller names into places to get them from, and check they exist.

    'STEM:Name' means that stencil, installed or open. A bare name looks in the drawing's own
    masters, then open stencils, then installed ones, so what's already in use wins. Naming
    stencils limits a bare name's search to those.
    """

    def __init__(self, app, document=None, units="us", within=None):
        self._app = app
        self._document = document
        self._units = units
        self._within = [stencil_stem(item) for item in within] if within else None
        self._open = None
        self._found = {}

    def resolve_all(self, nodes):
        return {node.master: self.resolve(node.master) for node in nodes}

    def resolve(self, reference):
        if reference not in self._found:
            self._found[reference] = self._resolve(reference)
        return self._found[reference]

    def _resolve(self, reference):
        if not isinstance(reference, str) or not reference.strip():
            raise ValueError("A master must be a non-empty string like 'BASFLO:Decision'.")
        stem, separator, name = reference.partition(":")
        if separator:
            return self._qualified(stencil_stem(stem), reference, name.strip())
        return self._bare(reference.strip())

    def _open_stencils(self):
        if self._open is None:
            self._open = {}
            documents = self._app.Documents
            for index in range(1, documents.Count + 1):
                document = documents(index)
                if document.Type == VIS_TYPE_STENCIL:
                    self._open.setdefault(stencil_stem(document.Name), document)
        return self._open

    def _is_custom_open(self, stem):
        return stem not in stencil_index.stencils() and stem in self._open_stencils()

    def _qualified(self, stem, reference, name):
        if not self._is_custom_open(stem):
            return stencil_index.resolve_master(reference, self._units)
        stencil = self._open_stencils()[stem]
        rows = _master_names(stencil)
        found = _named(rows, name)
        if found is None:
            raise VisioError(
                "unknown_master",
                f"Open stencil {stem} has no master called '{name}'",
                details=[f"{stem}:{option}" for option in difflib.get_close_matches(name, [row[0] for row in rows], n=5, cutoff=0.5)] or None,
            )
        return MasterRef(stem, found, stencil.FullName)

    def _bare(self, name):
        if self._document is not None and self._within is None:
            found = _named(_master_names(self._document), name)
            if found is not None:
                return MasterRef("", found, None)
        hits = []
        for stem, stencil in self._open_stencils().items():
            if self._within is None or stem in self._within:
                found = _named(_master_names(stencil), name)
                if found is not None:
                    hits.append((stem, found, stencil.FullName))
        if len(hits) > 1:
            raise VisioError(
                "ambiguous_master",
                f"{len(hits)} open stencils have a master called '{name}'",
                hint="Use one of the listed names, like 'BASFLO:Decision'.",
                details=[f"{stem}:{found}" for stem, found, _ in hits],
            )
        if hits:
            stem, found, path = hits[0]
            if stem in stencil_index.stencils():
                # Open for some other drawing, perhaps in the other units.
                return stencil_index.resolve_master(f"{stem}:{found}", self._units)
            return MasterRef(stem, found, path)
        return self._installed(name)

    def _installed(self, name):
        if self._within is None:
            return stencil_index.resolve_master(name, self._units)
        installed = [stem for stem in self._within if not self._is_custom_open(stem)]
        if not installed:
            raise VisioError("unknown_master", f"None of the named stencils has a master called '{name}'")
        return stencil_index.resolve_master(name, self._units, within=installed)


def _opens_something_on_drop(shapes):
    for index in range(1, shapes.Count + 1):
        shape = shapes(index)
        if shape.CellExistsU("EventDrop", 0) and INTERACTIVE_HANDLER.search(shape.CellsU("EventDrop").FormulaU):
            return True
        if shape.Type == VIS_TYPE_GROUP and _opens_something_on_drop(shape.Shapes):
            return True
    return False


class StencilPool:
    """Gets masters for one call and drops them. Reuses an open stencil, otherwise opens it hidden
    and closes it again afterwards, because a hidden stencil left open outlives its drawing."""

    def __init__(self, app, document):
        self._app = app
        self._document = document
        self._stencils = {}
        self._opened = []
        self._interactive = {}
        self._skipped = []

    def drop(self, page, ref, x, y):
        """Some masters run an add-on when dropped, and its dialog would block this call until
        it timed out. Those alone drop with events off, so the rest still set themselves up."""
        master = self.master(ref)
        key = (ref.path, ref.name)
        if key not in self._interactive:
            self._interactive[key] = _opens_something_on_drop(master.Shapes)
        if not self._interactive[key]:
            return page.Drop(master, x, y)
        label = f"{ref.stem}:{ref.name}" if ref.stem else ref.name
        if label not in self._skipped:
            self._skipped.append(label)
        with events_off(self._app):
            return page.Drop(master, x, y)

    def skipped_warning(self):
        if not self._skipped:
            return None
        return {
            "code": "drop_handler_skipped",
            "message": f"Skipped what these masters normally run when dropped, because it opens a dialog: {', '.join(self._skipped)}.",
        }

    def _find_open(self, key):
        documents = self._app.Documents
        for index in range(1, documents.Count + 1):
            document = documents(index)
            if document.Type == VIS_TYPE_STENCIL and document.FullName.casefold() == key:
                return document
        return None

    def master(self, ref):
        if ref.path is None:
            return self._document.Masters.ItemU(ref.name)
        key = ref.path.casefold()
        stencil = self._stencils.get(key)
        if stencil is None:
            stencil = self._find_open(key)
            if stencil is None:
                stencil = self._app.Documents.OpenEx(ref.path, VIS_OPEN_RO_HIDDEN)
                self._opened.append(stencil)
            self._stencils[key] = stencil
        return stencil.Masters.ItemU(ref.name)

    def close(self):
        for stencil in self._opened:
            try:
                stencil.Close()
            except Exception:
                logger.warning("could not close a hidden stencil", exc_info=True)
        self._opened = []


def _quoted(text):
    return '"' + text.replace('"', '""') + '"'


def _new_row(shape, label):
    """Row names only allow a-zA-Z0-9_ and can't start with a digit, so the label holds the real name."""
    base = re.sub(r"[^a-zA-Z0-9_]", "_", label).strip("_") or "Field"
    if base[0].isdigit():
        base = f"F_{base}"
    name = base
    suffix = 1
    while shape.CellExistsU(f"Prop.{name}", 0):
        suffix += 1
        name = f"{base}_{suffix}"
    if not shape.SectionExists(VIS_SECTION_PROP, 0):
        shape.AddSection(VIS_SECTION_PROP)
    shape.AddNamedRow(VIS_SECTION_PROP, name, VIS_TAG_DEFAULT)
    shape.CellsU(f"Prop.{name}.Label").FormulaU = _quoted(label)
    return name


def set_data(shape, data):
    rows = {}
    for label, name, _ in shape_data(shape):
        rows.setdefault(label.casefold(), name)
        rows.setdefault(name.casefold(), name)
    for label, value in data.items():
        name = rows.get(label.casefold())
        if name is None:
            name = _new_row(shape, label)
            rows[label.casefold()] = name
        shape.CellsU(f"Prop.{name}").FormulaU = _quoted(value)


def set_fill(shape, colour):
    red, green, blue = (int(colour[start:start + 2], 16) for start in (1, 3, 5))
    shape.CellsU("FillForegnd").FormulaU = f"THEMEGUARD(RGB({red},{green},{blue}))"
    shape.CellsU("FillPattern").FormulaU = "1"


def drop_node(page, pool, node, ref, index, placeable):
    if node.x is None:
        x, y = START_X + index * START_STEP, START_Y - index * START_STEP
    else:
        x, y = node.x, node.y
    shape = pool.drop(page, ref, x, y)
    if node.w is not None:
        shape.CellsU("Width").ResultIU = node.w
    if node.h is not None:
        shape.CellsU("Height").ResultIU = node.h
    if node.text is not None:
        shape.Text = node.text
    if node.data:
        set_data(shape, node.data)
    if node.fill:
        set_fill(shape, node.fill)
    if placeable:
        # Layout leaves a shape with no connections where it is, on top of whatever
        # lands there, unless the shape itself says it can be placed.
        shape.CellsU("ObjType").FormulaU = PLACEABLE
    return shape


def glue_edge(page, begin, end, edge):
    connector = page.Drop(page.Application.ConnectorToolDataObject, 0, 0)
    connector.CellsU("BeginX").GlueTo(begin.CellsU("PinX"))
    connector.CellsU("EndX").GlueTo(end.CellsU("PinX"))
    if edge.arrow == "none":
        connector.CellsU("EndArrow").FormulaU = NO_ARROW
    else:
        connector.CellsU("EndArrow").FormulaU = FILLED_ARROW
    if edge.arrow == "both":
        connector.CellsU("BeginArrow").FormulaU = FILLED_ARROW
    if edge.text:
        connector.Text = edge.text
    return connector


def _route_style(style, route):
    if route is None:
        route = "straight" if style in ("radial", "circular") else "flowchart"
    return (WEST_EAST if style == "left_right" else NORTH_SOUTH)[route]


def lay_out(page, style, route=None, spacing=None, shapes=None):
    """Lay out the page, or only the given shapes. The styles stay on the page, as if set in Visio."""
    sheet = page.PageSheet
    sheet.CellsU("PlaceStyle").FormulaU = str(PLACE_STYLES[style])
    sheet.CellsU("RouteStyle").FormulaU = str(_route_style(style, route))
    if spacing is not None:
        sheet.CellsU("AvenueSizeX").ResultIU = spacing
        sheet.CellsU("AvenueSizeY").ResultIU = spacing
    if not shapes:
        page.Layout()
        return
    selection = page.CreateSelection(VIS_SEL_EMPTY)
    for shape in shapes:
        selection.Select(shape, VIS_SELECT)
    selection.Layout()
