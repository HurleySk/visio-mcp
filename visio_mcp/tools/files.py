import json
import os
from typing import Literal

from visio_mcp import stencil_index
from visio_mcp.addressing import find_page, summarize
from visio_mcp.com_runtime import current_session, find_document
from visio_mcp.errors import VisioError
from visio_mcp.live_tool import live_tool

VIS_MS_DEFAULT = 0
VIS_MS_METRIC = 1
VIS_MS_US = 2
VIS_OPEN_RO = 2
VIS_FIXED_FORMAT_PDF = 1
VIS_DOC_EX_INTENT_PRINT = 1
VIS_PRINT_ALL = 0
VIS_PRINT_FROM_TO = 1

OPEN_SUFFIXES = (".vsdx", ".vsdm", ".vsd", ".vdx")
SAVE_SUFFIXES = (".vsdx", ".vsdm", ".vsd")
TEMPLATE_SUFFIXES = (".vstx", ".vstm", ".vst")
EXPORT_SUFFIXES = {"png": ".png", "jpg": ".jpg", "svg": ".svg", "pdf": ".pdf"}


MEASURES = {None: VIS_MS_DEFAULT, "us": VIS_MS_US, "metric": VIS_MS_METRIC}


def _absolute(path):
    if not isinstance(path, str) or not os.path.isabs(path):
        raise ValueError(f"Path must be absolute, because Visio has its own working folder: {path!r}")
    return os.path.normpath(path)


def _write_target(path, overwrite):
    """Check a path before Visio writes to it. Visio replaces files without asking, and
    its error for a missing folder wrongly complains about the file name."""
    target = _absolute(path)
    folder = os.path.dirname(target)
    if not os.path.isdir(folder):
        raise VisioError("folder_not_found", f"Folder does not exist: {folder}", hint="Create the folder first.")
    if os.path.exists(target) and not overwrite:
        raise VisioError(
            "file_exists",
            f"File already exists: {target}",
            hint="Pass overwrite=true to replace it, or choose another path.",
        )
    return target


def _suffix(path):
    return os.path.splitext(path)[1].lower()


def _same_path(document, target):
    return bool(document.Path) and os.path.normcase(os.path.normpath(document.FullName)) == os.path.normcase(target)


def template_source(template, units):
    """What to pass Visio for a new drawing: a template file's path, or "" for a blank one."""
    if not template:
        return ""
    if not os.path.isabs(template):
        return stencil_index.resolve_template(template, units or "us")
    source = os.path.normpath(template)
    if not os.path.isfile(source):
        raise VisioError("file_not_found", f"Template not found: {source}")
    if _suffix(source) not in TEMPLATE_SUFFIXES:
        raise ValueError(f"A template must be one of {TEMPLATE_SUFFIXES}: {source}")
    return source


def save_target(path, overwrite):
    target = _absolute(path)
    if _suffix(target) not in SAVE_SUFFIXES:
        raise ValueError(f"A drawing must be saved as one of {SAVE_SUFFIXES}: {target}")
    return _write_target(target, overwrite)


@live_tool(mutates=True, start="if_missing")
def visio_new_drawing(template: str | None = None, units: Literal["us", "metric"] | None = None):
    """Create a new drawing, starting a visible Visio if none is running.

    template is a name from visio_list_templates or the absolute path of a template file.
    Omit it for a blank drawing. units picks US or metric paper and shapes.
    """
    if units not in MEASURES:
        raise ValueError("units must be 'us' or 'metric'.")
    source = template_source(template, units)
    document = current_session().app.Documents.AddEx(source, MEASURES[units], 0)
    return json.dumps(summarize(document), ensure_ascii=False)


@live_tool(mutates=True, start="if_missing")
def visio_open_drawing(path: str, read_only: bool = False):
    """Open a saved drawing, starting a visible Visio if none is running.

    If it's already open, returns that one with already_open=true, even if it was opened
    with different access. Check read_only in the result.
    """
    target = _absolute(path)
    if _suffix(target) not in OPEN_SUFFIXES:
        raise ValueError(f"A drawing must be one of {OPEN_SUFFIXES}: {target}")
    if not os.path.isfile(target):
        raise VisioError("file_not_found", f"File not found: {target}")
    app = current_session().app
    try:
        document = find_document(app, target)
    except VisioError as err:
        if err.code != "drawing_not_found":
            raise
    else:
        return json.dumps({**summarize(document), "already_open": True}, ensure_ascii=False)
    document = app.Documents.OpenEx(target, VIS_OPEN_RO if read_only else 0)
    return json.dumps(summarize(document), ensure_ascii=False)


@live_tool(mutates=True)
def visio_save(path: str | None = None, overwrite: bool = False, drawing: str | None = None):
    """Save a drawing, or save a copy to a new absolute path and continue in that copy.

    Refuses to replace another file unless overwrite is true. A read-only drawing can only
    be saved to a new path, and the result is writable.
    """
    document = find_document(current_session().app, drawing)
    if path is None:
        if not document.Path:
            raise VisioError(
                "save_needs_path",
                f"Drawing '{document.Name}' has never been saved",
                hint="Pass an absolute path.",
            )
        if document.ReadOnly:
            raise VisioError(
                "read_only",
                f"Drawing '{document.Name}' is open read-only",
                hint="Pass a new path to save a writable copy.",
            )
        document.Save()
    else:
        target = _absolute(path)
        if _same_path(document, target) and _suffix(target) in SAVE_SUFFIXES:
            if document.ReadOnly:
                raise VisioError(
                    "read_only",
                    f"Drawing '{document.Name}' is open read-only",
                    hint="Pass a new path to save a writable copy.",
                )
            document.Save()
        else:
            document.SaveAs(save_target(target, overwrite))
    return json.dumps(summarize(document), ensure_ascii=False)


def _export_format(target, format):
    suffix = _suffix(target)
    if format is None:
        for name, expected in EXPORT_SUFFIXES.items():
            if suffix == expected:
                return name
        raise ValueError(f"Can't tell the format from '{suffix}'. Use one of {list(EXPORT_SUFFIXES)}.")
    expected = EXPORT_SUFFIXES.get(format)
    if expected is None:
        raise ValueError(f"format must be one of {list(EXPORT_SUFFIXES)}.")
    if suffix != expected:
        raise ValueError(f"Format {format} needs a path ending in {expected}, because Visio goes by the extension.")
    return format


@live_tool(mutates=True)
def visio_export(
    path: str,
    format: Literal["png", "jpg", "svg", "pdf"] | None = None,
    overwrite: bool = False,
    drawing: str | None = None,
    page: int | str | None = None,
):
    """Export to an absolute path as PNG, JPG, SVG or PDF. The format comes from the extension.

    Images export one page: the named one, else the one on screen. PDF exports the named
    page, else every page. Refuses to replace a file unless overwrite is true.
    """
    target = _absolute(path)
    kind = _export_format(target, format)
    _write_target(target, overwrite)
    document = find_document(current_session().app, drawing)
    if kind != "pdf":
        exported = find_page(document, page)
        exported.Export(target)
        return json.dumps({"path": target, "format": kind, "page": exported.Name}, ensure_ascii=False)
    if page is None:
        document.ExportAsFixedFormat(VIS_FIXED_FORMAT_PDF, target, VIS_DOC_EX_INTENT_PRINT, VIS_PRINT_ALL)
        return json.dumps({"path": target, "format": kind, "page": None}, ensure_ascii=False)
    exported = find_page(document, page)
    index = exported.Index
    document.ExportAsFixedFormat(VIS_FIXED_FORMAT_PDF, target, VIS_DOC_EX_INTENT_PRINT, VIS_PRINT_FROM_TO, index, index)
    return json.dumps({"path": target, "format": kind, "page": exported.Name}, ensure_ascii=False)


@live_tool(mutates=True)
def visio_close_drawing(discard: bool = False, drawing: str | None = None):
    """Close a drawing. Visio itself stays open.

    Refuses if there are unsaved changes unless discard is true, which loses them and
    must name the drawing, so it can't take whichever drawing happens to be active.
    """
    if discard and not drawing:
        raise ValueError("discard=true loses unsaved changes, so it must name the drawing.")
    document = find_document(current_session().app, drawing)
    name = document.Name
    unsaved = not document.Saved
    if unsaved and not discard:
        raise VisioError(
            "unsaved_changes",
            f"Drawing '{name}' has unsaved changes",
            hint="Call visio_save first, or pass discard=true and the drawing's name to lose them.",
        )
    if unsaved:
        # Marks it clean so Close can't raise a prompt that blocks this call.
        document.Saved = True
    document.Close()
    return json.dumps({"closed": name, "discarded": unsaved}, ensure_ascii=False)
