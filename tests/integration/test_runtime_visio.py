import json
import threading

import pytest

from tests.integration.conftest import VIS_MS_US
from tests.integration.helpers import dialogs_of, launch_own, pid_of, pids, until
from visio_mcp import com_runtime
from visio_mcp.com_runtime import NO_SCOPE, current_session, drawings_in, edit_doc
from visio_mcp.live_tool import live_tool

pytestmark = pytest.mark.visio

VIS_OPEN_RO = 2
VIS_CMD_PAGE_SETUP = 1076


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


async def test_read_tool_starts_nothing_with_no_visio():
    if pids():
        pytest.skip("Visio is running")
    result = json.loads(await visio_names())
    assert result["code"] == "visio_not_running"
    assert not pids()


async def test_if_missing_starts_a_visible_visio():
    import pythoncom

    if pids():
        pytest.skip("Visio is running")
    pythoncom.CoInitialize()
    assert json.loads(await visio_starting()) == {"docs": 0}
    app = com_runtime.Win32Connector().attach()
    assert app.Visible
    app.Quit()
    del app
    await until(lambda: not pids())


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
    await until(lambda: scratch.Pages(1).Shapes.Count == 0)
    await until(lambda: scratch.Application.CurrentScope == NO_SCOPE)
    await until(lambda: all(future.done() for future in com_runtime._stuck_edits))
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

    pid = pid_of(own_visio)

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
        await until(lambda: dialogs_of(pid))
        result = json.loads(await visio_names())
        assert result["code"] == "modal_dialog"
        assert result["retryable"] is True
    finally:
        for hwnd in dialogs_of(pid):
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        opener.join(10)
    assert not opener.is_alive()
    assert scratch.Name in json.loads(await visio_names())["drawings"]


# Keep this last. Once the newer instance closes, attach can stop reaching older ones.
async def test_saved_drawing_in_an_older_instance_is_reached(own_visio, tmp_path):
    path = str(tmp_path / "older.vsdx")
    document = own_visio.Documents.AddEx("", VIS_MS_US, 0)
    document.SaveAs(path)
    newer = launch_own()
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
