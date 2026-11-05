"""Turn the names a caller uses into live pages and shapes. Runs on the COM thread."""
import difflib

from visio_mcp.com_runtime import classify_live
from visio_mcp.errors import VisioError


def summarize(document):
    pages = document.Pages
    return {
        "drawing": document.Name,
        "path": document.FullName if document.Path else None,
        "read_only": bool(document.ReadOnly),
        "saved": bool(document.Saved),
        "pages": [pages(index).Name for index in range(1, pages.Count + 1)],
    }


def _active_page(document):
    try:
        window = document.Application.ActiveWindow
        if window is not None and window.Document.ID == document.ID:
            return window.Page
    except Exception as exc:
        if classify_live(exc).code in ("busy", "disconnected"):
            raise
    return document.Pages(1)


def find_page(document, page=None):
    """A page by 1-based index or name. With neither, the one on screen, else the first."""
    if page is None:
        return _active_page(document)
    if isinstance(page, bool) or not isinstance(page, (int, str)):
        raise ValueError(f"A page must be a name or a number, not {page!r}.")
    pages = document.Pages
    count = pages.Count
    if isinstance(page, int):
        if 1 <= page <= count:
            return pages(page)
    else:
        wanted = page.casefold()
        for index in range(1, count + 1):
            candidate = pages(index)
            if wanted in (candidate.NameU.casefold(), candidate.Name.casefold()):
                return candidate
    raise VisioError(
        "page_not_found",
        f"Drawing '{document.Name}' has no page {page!r}",
        details=[pages(index).Name for index in range(1, count + 1)],
    )


def master_name(shape):
    master = shape.Master
    return master.NameU if master is not None else None


def identity(shape):
    return {"id": shape.ID, "name": shape.Name, "text": shape.Text, "master": master_name(shape)}


class ShapeFinder:
    """Finds top-level shapes by ID, then exact name, then exact text.

    IDs also reach shapes inside groups. Names and text are read once per finder, and
    shapes are looked up afresh each call because references go stale after an undo.
    """

    def __init__(self, page):
        self._page = page
        self._shapes = None

    def _all(self):
        if self._shapes is None:
            shapes = self._page.Shapes
            self._shapes = [shapes(index) for index in range(1, shapes.Count + 1)]
        return self._shapes

    def _by_id(self, shape_id):
        try:
            return self._page.Shapes.ItemFromID(shape_id)
        except Exception as exc:
            if classify_live(exc).code in ("busy", "disconnected"):
                raise
            return None

    def find(self, reference):
        if isinstance(reference, bool) or not isinstance(reference, (int, str)) or reference == "":
            raise ValueError(f"A shape must be an ID, a name or its text, not {reference!r}.")
        if isinstance(reference, int) or reference.isdigit():
            shape = self._by_id(int(reference))
            if shape is not None:
                return shape
            if isinstance(reference, int):
                raise self._not_found(reference)
        named = [shape for shape in self._all() if reference in (shape.NameU, shape.Name)]
        if len(named) == 1:
            return named[0]
        matches = named or [shape for shape in self._all() if shape.Text == reference]
        if len(matches) == 1:
            return matches[0]
        if matches:
            raise VisioError(
                "ambiguous_shape",
                f"{len(matches)} shapes on page '{self._page.Name}' match '{reference}'",
                hint="Use one of the listed IDs.",
                details=[identity(shape) for shape in matches[:20]],
            )
        raise self._not_found(reference)

    def _not_found(self, reference):
        by_label = {}
        for shape in self._all():
            for label in (shape.Name, shape.Text):
                if label:
                    by_label.setdefault(label, shape)
        close = difflib.get_close_matches(str(reference), list(by_label), n=5, cutoff=0.6)
        return VisioError(
            "shape_not_found",
            f"No shape on page '{self._page.Name}' matches {reference!r}",
            hint="List shapes with visio_get_page or visio_find_shapes.",
            details=[identity(by_label[label]) for label in close] or None,
        )
