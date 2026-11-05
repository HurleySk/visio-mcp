"""Find masters and templates from Visio's installed files, without Visio running."""
import dataclasses
import difflib
import logging
import os
import re
import threading
import xml.etree.ElementTree as ET
import zipfile

from visio_mcp.errors import VisioError

logger = logging.getLogger(__name__)

CONTENT_ENV = "VISIO_MCP_CONTENT"
APP_PATHS = r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\VISIO.EXE"
VISIO_NS = "{http://schemas.microsoft.com/office/visio/2012/main}"
TITLE_TAG = "{http://purl.org/dc/elements/1.1/}title"
DESCRIPTION_TAG = "{http://purl.org/dc/elements/1.1/}description"
KEYWORDS_TAG = "{http://schemas.openxmlformats.org/package/2006/metadata/core-properties}keywords"
FILE_NAME = re.compile(r"^(?P<stem>.+)_(?P<units>[UM])\.(?P<kind>VSSX|VSTX)$", re.IGNORECASE)
UNITS = {"U": "us", "M": "metric"}
# The stencils people mean most often rank first among otherwise equal matches.
COMMON = ("BASFLO", "BASIC", "BLOCK", "CONNEC")
MAX_LIMIT = 100

_lock = threading.Lock()
_stencils = None
_templates = None


@dataclasses.dataclass
class Library:
    stem: str
    paths: dict
    title: str = ""
    description: str = ""
    keywords: str = ""
    masters: list = dataclasses.field(default_factory=list)


@dataclasses.dataclass(frozen=True)
class Master:
    name: str
    local_name: str
    prompt: str


@dataclasses.dataclass(frozen=True)
class MasterRef:
    stem: str
    name: str
    path: str

    @property
    def label(self):
        return f"{self.stem}:{self.name}"


def reset():
    global _stencils, _templates
    with _lock:
        _stencils = None
        _templates = None


def _content_folders():
    override = os.environ.get(CONTENT_ENV)
    if override:
        return [override]
    import winreg

    try:
        exe = winreg.QueryValue(winreg.HKEY_LOCAL_MACHINE, APP_PATHS)
    except OSError as exc:
        raise VisioError(
            "visio_not_installed",
            "Could not find Visio's install folder",
            hint=f"Is Visio installed? To point at the shapes folder directly, set {CONTENT_ENV}.",
        ) from exc
    root = os.path.join(os.path.dirname(exe), "Visio Content")
    if not os.path.isdir(root):
        raise VisioError(
            "visio_content_not_found",
            f"Visio's shapes folder is missing: {root}",
            hint=f"To point at the shapes folder directly, set {CONTENT_ENV}.",
        )
    languages = sorted(name for name in os.listdir(root) if name.isdigit())
    languages.sort(key=lambda name: name != "1033")
    return [os.path.join(root, name) for name in languages]


def _scan(kind):
    found = {}
    for folder in _content_folders():
        try:
            names = os.listdir(folder)
        except OSError:
            continue
        for name in names:
            match = FILE_NAME.match(name)
            if not match or match["kind"].upper() != kind:
                continue
            stem = match["stem"].upper()
            library = found.setdefault(stem, Library(stem, {}))
            library.paths.setdefault(UNITS[match["units"].upper()], os.path.join(folder, name))
    return found


def _read_properties(archive, library):
    try:
        root = ET.fromstring(archive.read("docProps/core.xml"))
    except (KeyError, ET.ParseError):
        return
    library.title = root.findtext(TITLE_TAG) or ""
    library.description = root.findtext(DESCRIPTION_TAG) or ""
    library.keywords = root.findtext(KEYWORDS_TAG) or ""


def _read_masters(archive):
    try:
        root = ET.fromstring(archive.read("visio/masters/masters.xml"))
    except KeyError:
        return []
    masters = []
    for element in root.findall(f"{VISIO_NS}Master"):
        name = element.get("NameU")
        if name and element.get("Hidden", "0") == "0":
            masters.append(Master(name, element.get("Name", name), element.get("Prompt", "")))
    return masters


def _load(kind, with_masters):
    libraries = _scan(kind)
    for library in libraries.values():
        path = library.paths.get("us") or library.paths["metric"]
        try:
            with zipfile.ZipFile(path) as archive:
                _read_properties(archive, library)
                if with_masters:
                    library.masters = _read_masters(archive)
        except (OSError, zipfile.BadZipFile, ET.ParseError) as exc:
            logger.warning("skipped %s: %s", path, exc)
    return libraries


def stencils():
    global _stencils
    with _lock:
        if _stencils is None:
            _stencils = _load("VSSX", with_masters=True)
        return _stencils


def templates():
    global _templates
    with _lock:
        if _templates is None:
            _templates = _load("VSTX", with_masters=False)
        return _templates


def _tokens(text):
    return [token for token in re.split(r"[^\w]+", text.lower()) if token]


def _rank(stem):
    return COMMON.index(stem) if stem in COMMON else len(COMMON)


def _score(library, master, phrase, tokens):
    names = f"{master.name} {master.local_name}".lower()
    if phrase in (master.name.lower(), master.local_name.lower()):
        return 4
    if names.startswith(phrase):
        return 3
    if all(token in names for token in tokens):
        return 2
    context = f"{names} {library.stem} {library.title} {library.keywords}".lower()
    if all(token in context for token in tokens):
        # A master with any query word in its own name beats the rest of its stencil.
        return 1 + sum(token in names for token in tokens) / (len(tokens) + 1)
    return 0


def stencil_stem(reference):
    stem = os.path.basename(reference.strip()).upper()
    return re.sub(r"_[UM]$", "", re.sub(r"\.VSSX$", "", stem))


def find_masters(query="", stencil=None, limit=25):
    """Search masters by name, then by their stencil's title and keywords."""
    phrase = (query or "").strip().lower()
    if not phrase and not stencil:
        raise ValueError("Pass a query, a stencil, or both.")
    libraries = stencils()
    if stencil:
        libraries = {stencil_stem(stencil): _library(libraries, stencil)}
    tokens = _tokens(phrase)
    scored = []
    for library in libraries.values():
        for master in library.masters:
            score = _score(library, master, phrase, tokens) if phrase else 1
            if score:
                scored.append((-score, _rank(library.stem), library.stem, master.name, library, master))
    # Stencil words alone only count when no master's own name matches. Otherwise a stencil
    # with "process" among its keywords would list all of its masters for "process".
    if phrase and any(row[0] < -1 for row in scored):
        scored = [row for row in scored if row[0] < -1]
    scored.sort(key=lambda row: row[:4])
    limit = max(1, min(limit or 25, MAX_LIMIT))
    return {
        "total": len(scored),
        "results": [_describe(row[4], row[5]) for row in scored[:limit]],
    }


def _describe(library, master):
    described = {"master": f"{library.stem}:{master.name}", "stencil": library.title or library.stem}
    if master.prompt:
        described["prompt"] = master.prompt
    return described


def _library(libraries, reference):
    stem = stencil_stem(reference)
    library = libraries.get(stem)
    if library is None:
        close = difflib.get_close_matches(stem, list(libraries), n=5, cutoff=0.5)
        raise VisioError(
            "unknown_stencil",
            f"No installed stencil is called '{reference}'",
            hint="Search with visio_find_masters.",
            details=close or None,
        )
    return library


def _path(library, units):
    return library.paths.get(units) or next(iter(library.paths.values()))


def _matches(master, name):
    wanted = name.casefold()
    return wanted in (master.name.casefold(), master.local_name.casefold())


def resolve_master(reference, units="us", within=None):
    """Turn 'BASFLO:Decision', or a bare 'Decision', into a stencil file and a master."""
    if not isinstance(reference, str) or not reference.strip():
        raise ValueError("A master must be a non-empty string like 'BASFLO:Decision'.")
    libraries = stencils()
    stem, separator, name = reference.partition(":")
    if separator:
        library = _library(libraries, stem)
        name = name.strip()
        for master in library.masters:
            if _matches(master, name):
                return MasterRef(library.stem, master.name, _path(library, units))
        close = difflib.get_close_matches(name, [m.name for m in library.masters], n=5, cutoff=0.5)
        raise VisioError(
            "unknown_master",
            f"Stencil {library.stem} has no master called '{name}'",
            hint="Search with visio_find_masters.",
            details=[f"{library.stem}:{option}" for option in close] or None,
        )
    name = reference.strip()
    pool = [_library(libraries, item) for item in within] if within else libraries.values()
    hits = sorted(
        ((_rank(library.stem), library.stem, library, master) for library in pool for master in library.masters if _matches(master, name)),
        key=lambda row: row[:2],
    )
    if len(hits) == 1:
        _, _, library, master = hits[0]
        return MasterRef(library.stem, master.name, _path(library, units))
    if hits:
        raise VisioError(
            "ambiguous_master",
            f"{len(hits)} stencils have a master called '{name}'",
            hint="Use one of the listed names, like 'BASFLO:Decision'.",
            details=[f"{row[2].stem}:{row[3].name}" for row in hits[:10]],
        )
    raise VisioError(
        "unknown_master",
        f"No installed master is called '{name}'",
        hint="Search with visio_find_masters.",
        details=[row["master"] for row in find_masters(name, limit=5)["results"]] or None,
    )


def list_templates(query=None):
    tokens = _tokens(query or "")
    rows = []
    for library in templates().values():
        haystack = f"{library.stem} {library.title} {library.description} {library.keywords}".lower()
        if all(token in haystack for token in tokens):
            row = {"template": library.stem, "title": library.title}
            if library.description:
                row["description"] = library.description
            rows.append(row)
    rows.sort(key=lambda row: (row["title"] or row["template"]).lower())
    return rows


def resolve_template(name, units="us"):
    library = templates().get(stencil_stem(re.sub(r"\.VSTX$", "", name, flags=re.IGNORECASE)))
    if library is None:
        close = difflib.get_close_matches(stencil_stem(name), list(templates()), n=5, cutoff=0.5)
        raise VisioError(
            "unknown_template",
            f"No installed template is called '{name}'",
            hint="List them with visio_list_templates.",
            details=close or None,
        )
    return _path(library, units)
