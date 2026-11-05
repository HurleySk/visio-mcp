"""Runtime behaviour against real Visio.

Undo scopes and the undo stack are app-wide, so these tests never touch a Visio
the user has open. They start their own instance and check that attach reached it.
"""
import asyncio
import json
import subprocess
import threading
import time

import pytest

from visio_mcp import com_runtime
from visio_mcp.com_runtime import NO_SCOPE, current_session, drawings_in, edit_doc
from visio_mcp.live_tool import live_tool

pytestmark = pytest.mark.visio

VIS_MS_US = 2
VIS_OPEN_RO = 2
VIS_CMD_PAGE_SETUP = 1076


def _pids():
    listing = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq VISIO.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
    ).stdout
    return {int(line.split('","')[1]) for line in listing.splitlines() if "VISIO.EXE" in line}


def _pid_of(app):
    import win32process

    return win32process.GetWindowThreadProcessId(app.WindowHandle32)[1]


def _launch_own():
    """Start a new instance and return it, once attach reaches that instance."""
    connector = com_runtime.Win32Connector()
    before = _pids()
    connector.launch()
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline:
        new = _pids() - before
        if new:
            try:
                app = connector.attach()
                if _pid_of(app) in new:
                    return app
            except Exception:
                pass
        time.sleep(0.5)
    raise AssertionError("attach did not reach the instance these tests started")


def _dialogs_of(pid):
    import win32gui
    import win32process

    found = []

    def visit(hwnd, _):
        if (
            win32process.GetWindowThreadProcessId(hwnd)[1] == pid
            and win32gui.IsWindowVisible(hwnd)
            and win32gui.GetClassName(hwnd) == "#32770"
        ):
            found.append(hwnd)
        return True

    win32gui.EnumWindows(visit, None)
    return found


async def _until(predicate, seconds=30):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if predicate():
            return
        await asyncio.sleep(0.1)
    raise AssertionError("condition not met in time")


@live_tool(mutates=False)
def visio_names():
    return json.dumps({"drawings": [d.Name for d in drawings_in(current_session().app)]})


@live_tool(mutates=True, start="if_missing")
def visio_starting():
    return json.dumps({"docs": current_session().app.Documents.Count})


@live_tool(mutates=True)
def visio_draw_two(drawing, fail=False):
    document = edit_doc(drawing)
    page = document.Pages(1)
    page.DrawRectangle(1, 1, 2, 2)
    page.DrawRectangle(3, 3, 4, 4)
    if fail:
        raise KeyError("stop")
    return json.dumps({"count": page.Shapes.Count})


@pytest.fixture(autouse=True)
def fresh_runtime():
    com_runtime.reset()
    yield
    com_runtime.reset()


async def test_read_tool_starts_nothing_with_no_visio():
    if _pids():
        pytest.skip("Visio is running")
    result = json.loads(await visio_names())
    assert result["code"] == "visio_not_running"
    assert not _pids()


async def test_if_missing_starts_a_visible_visio():
    import pythoncom

    if _pids():
        pytest.skip("Visio is running")
    pythoncom.CoInitialize()
    assert json.loads(await visio_starting()) == {"docs": 0}
    app = com_runtime.Win32Connector().attach()
    assert app.Visible
    app.Quit()
    del app
    await _until(lambda: not _pids())


@pytest.fixture(scope="module")
def own_visio():
    import pythoncom

    pythoncom.CoInitialize()
    app = _launch_own()
    yield app
    for index in range(app.Documents.Count, 0, -1):
        app.Documents(index).Saved = True
    app.Quit()
    del app
    pythoncom.CoUninitialize()


@pytest.fixture
def scratch(own_visio):
    document = own_visio.Documents.AddEx("", VIS_MS_US, 0)
    yield document
    document.Saved = True
    document.Close()


async def test_attach_to_user_instance(scratch):
    payload = json.loads(await visio_names())
    assert scratch.Name in payload["drawings"]


async def test_idle_visio_reports_no_scope(own_visio):
    assert own_visio.CurrentScope == NO_SCOPE


async def test_mutating_call_is_one_undo_step(scratch):
    assert json.loads(await visio_draw_two(scratch.Name)) == {"count": 2}
    assert scratch.Application.CurrentScope == NO_SCOPE
    scratch.Application.Undo()
    assert scratch.Pages(1).Shapes.Count == 0
    scratch.Application.Redo()
    assert scratch.Pages(1).Shapes.Count == 2


async def test_exception_rolls_back_for_real(scratch):
    result = json.loads(await visio_draw_two(scratch.Name, fail=True))
    assert result["code"] == "internal"
    assert result["rolled_back"] is True
    assert scratch.Pages(1).Shapes.Count == 0
    assert scratch.Application.CurrentScope == NO_SCOPE


async def test_timed_out_edit_rolls_back_for_real(scratch):
    drawn = threading.Event()
    release = threading.Event()

    @live_tool(mutates=True, timeout=2)
    def visio_slow_edit(drawing):
        document = edit_doc(drawing)
        document.Pages(1).DrawRectangle(1, 1, 2, 2)
        drawn.set()
        release.wait(30)
        return "{}"

    result = json.loads(await visio_slow_edit(scratch.Name))
    assert result["code"] == "timeout"
    assert drawn.is_set()
    assert scratch.Pages(1).Shapes.Count == 1

    refused = json.loads(await visio_draw_two(scratch.Name))
    assert refused["code"] == "edit_in_progress"
    assert scratch.Pages(1).Shapes.Count == 1

    release.set()
    await _until(lambda: scratch.Pages(1).Shapes.Count == 0)
    await _until(lambda: scratch.Application.CurrentScope == NO_SCOPE)
    await _until(lambda: all(future.done() for future in com_runtime._stuck_edits))
    assert json.loads(await visio_draw_two(scratch.Name)) == {"count": 2}


async def test_leftover_scope_is_committed_and_keeps_its_edits(scratch):
    app = scratch.Application
    leftover = app.BeginUndoScope("left behind")
    scratch.Pages(1).DrawRectangle(5, 5, 6, 6)
    try:
        result = json.loads(await visio_draw_two(scratch.Name))
    finally:
        if app.CurrentScope == leftover:
            app.EndUndoScope(leftover, True)
    assert result["count"] == 3
    assert [w["code"] for w in result["warnings"]] == ["recovered_scope"]
    assert app.CurrentScope == NO_SCOPE
    app.Undo()
    assert scratch.Pages(1).Shapes.Count == 1
    app.Undo()
    assert scratch.Pages(1).Shapes.Count == 0


async def test_read_only_drawing_is_refused_for_real(own_visio, tmp_path):
    path = str(tmp_path / "locked.vsdx")
    writable = own_visio.Documents.AddEx("", VIS_MS_US, 0)
    writable.SaveAs(path)
    writable.Close()
    locked = own_visio.Documents.OpenEx(path, VIS_OPEN_RO)
    try:
        result = json.loads(await visio_draw_two(locked.Name))
        assert result["code"] == "read_only"
        assert locked.Pages(1).Shapes.Count == 0
    finally:
        locked.Saved = True
        locked.Close()


async def test_open_dialog_is_reported_as_modal_dialog(scratch, own_visio):
    import pythoncom
    import win32con
    import win32gui

    pid = _pid_of(own_visio)

    def open_dialog():
        pythoncom.CoInitialize()
        try:
            com_runtime.Win32Connector().attach().DoCmd(VIS_CMD_PAGE_SETUP)
        except Exception:
            pass
        finally:
            pythoncom.CoUninitialize()

    opener = threading.Thread(target=open_dialog, daemon=True)
    opener.start()
    try:
        await _until(lambda: _dialogs_of(pid))
        result = json.loads(await visio_names())
        assert result["code"] == "modal_dialog"
        assert result["retryable"] is True
    finally:
        for hwnd in _dialogs_of(pid):
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        opener.join(10)
    assert not opener.is_alive()
    assert scratch.Name in json.loads(await visio_names())["drawings"]


# Keep this last. Once the newer instance closes, attach can stop reaching older ones.
async def test_saved_drawing_in_an_older_instance_is_reached(own_visio, tmp_path):
    path = str(tmp_path / "older.vsdx")
    document = own_visio.Documents.AddEx("", VIS_MS_US, 0)
    document.SaveAs(path)
    newer = _launch_own()
    try:
        assert json.loads(await visio_draw_two(path)) == {"count": 2}
        assert document.Pages(1).Shapes.Count == 2
        assert newer.Documents.Count == 0
        assert own_visio.CurrentScope == NO_SCOPE
        own_visio.Undo()
        assert document.Pages(1).Shapes.Count == 0
    finally:
        newer.Quit()
        document.Saved = True
        document.Close()
