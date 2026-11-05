import zipfile

import pytest

from tests.fakes import FakeConnector
from visio_mcp import com_runtime, events, stencil_index

NS = "http://schemas.microsoft.com/office/visio/2012/main"


@pytest.fixture(autouse=True)
def events_marker(tmp_path, monkeypatch):
    """Keeps tests away from the marker of a real server running on this PC."""
    marker = tmp_path / "visio-mcp-marker" / "events-off"
    monkeypatch.setattr(events, "_marker", lambda: str(marker))
    return marker


@pytest.fixture
def connector():
    fake = FakeConnector()
    com_runtime.configure(lambda: fake, sleep=lambda seconds: None)
    yield fake
    com_runtime.reset()


def _core(title, keywords=""):
    return (
        '﻿<?xml version="1.0" encoding="UTF-8"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/">'
        f"<dc:title>{title}</dc:title><cp:keywords>{keywords}</cp:keywords>"
        "<dc:description>About it.</dc:description></cp:coreProperties>"
    ).encode("utf-8")


def _masters(*names, hidden=()):
    rows = [f"<Master ID='{i}' NameU='{n}' Name='{n}' Prompt='Drag {n}.' Hidden='0'><PageSheet/></Master>" for i, n in enumerate(names)]
    rows += [f"<Master ID='9{i}' NameU='{n}' Name='{n}' Hidden='1'/>" for i, n in enumerate(hidden)]
    return f"<?xml version='1.0'?><Masters xmlns='{NS}'>{''.join(rows)}</Masters>".encode("utf-8")


def _stencil(folder, file_name, title, keywords, *names, hidden=()):
    with zipfile.ZipFile(folder / file_name, "w") as archive:
        archive.writestr("docProps/core.xml", _core(title, keywords))
        archive.writestr("visio/masters/masters.xml", _masters(*names, hidden=hidden))


@pytest.fixture
def content(tmp_path, monkeypatch):
    """Stencil and template files like Visio's own, built here because Microsoft's can't ship."""
    monkeypatch.setenv(stencil_index.CONTENT_ENV, str(tmp_path))
    stencil_index.reset()
    # Real stencils list their masters' words as keywords, so every master matches on them.
    _stencil(tmp_path, "BASFLO_U.VSSX", "Basic Flowchart Shapes", "flowchart,process,decision", "Process", "Decision", hidden=("Secret",))
    _stencil(tmp_path, "BASFLO_M.VSSX", "Basic Flowchart Shapes", "flowchart,process,decision", "Process", "Decision")
    _stencil(tmp_path, "XFUNC_U.VSSX", "Cross-functional", "swimlane", "Process", "Swimlane")
    _stencil(tmp_path, "ONLYM_M.VSSX", "Metric only", "", "Metric Thing")
    with zipfile.ZipFile(tmp_path / "BASFLO_U.VSTX", "w") as archive:
        archive.writestr("docProps/core.xml", _core("Basic Flowchart", "flow"))
    with zipfile.ZipFile(tmp_path / "BASFLO_M.VSTX", "w") as archive:
        archive.writestr("docProps/core.xml", _core("Basic Flowchart", "flow"))
    (tmp_path / "RIGHTLEFT_R2L.VSTX").write_bytes(b"ignored by the name filter")
    (tmp_path / "BROKEN_U.VSSX").write_bytes(b"not a zip")
    yield tmp_path
    stencil_index.reset()
