"""Starting Visio from inside the kind of Job some MCP clients run servers in."""
import ctypes
import json
import os
import subprocess
import sys
import time

import pytest

from tests.integration.helpers import pid_of, pids
from visio_mcp import com_runtime, events
from visio_mcp.com_runtime import Win32Connector, current_session
from visio_mcp.live_tool import live_tool

pytestmark = pytest.mark.visio

CLIENT = """
import time
from visio_mcp.com_runtime import Win32Connector

connector = Win32Connector()
connector.thread_init()
before = set(connector.visio_pids())
connector.launch()
for _ in range(240):
    new = set(connector.visio_pids()) - before
    if new:
        try:
            app = connector.attach()
            if app.WindowHandle32:
                break
        except Exception:
            pass
    time.sleep(0.25)
document = app.Documents.AddEx("", 2, 0)
document.Pages(1).DrawRectangle(1, 1, 2, 2)
print(document.Name, new.pop(), flush=True)
"""


def _run_in_kill_on_close_job():
    """Run the client in a Job like the Python MCP SDK's, then release the Job as its exit would."""
    import win32api
    import win32con
    import win32job
    import win32process

    job = win32job.CreateJobObject(None, "")
    info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
    info["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)
    client = subprocess.Popen(
        [sys.executable, "-B", "-c", CLIENT],
        stdout=subprocess.PIPE,
        text=True,
        creationflags=win32process.CREATE_SUSPENDED,
    )
    process = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, False, client.pid)
    win32job.AssignProcessToJobObject(job, process)
    ctypes.WinDLL("ntdll").NtResumeProcess(int(process))
    output, _ = client.communicate(timeout=120)
    job.Close()
    return output.split()


def _windows_of(pid):
    import win32gui
    import win32process

    found = []

    def visit(hwnd, _):
        if win32process.GetWindowThreadProcessId(hwnd)[1] == pid and win32gui.IsWindowVisible(hwnd):
            found.append(win32gui.GetWindowText(hwnd))
        return True

    win32gui.EnumWindows(visit, None)
    return found


def test_visio_and_unsaved_work_outlive_a_client_that_kills_its_children():
    import pythoncom

    pythoncom.CoInitialize()
    name, pid = _run_in_kill_on_close_job()
    pid = int(pid)
    app = None
    try:
        time.sleep(3)
        assert pid in pids()
        assert any(name in title for title in _windows_of(pid))
        app = Win32Connector().attach()
        assert pid_of(app) == pid
        document = app.Documents(name)
        assert document.Saved == 0
        assert document.Pages(1).Shapes.Count == 1
    finally:
        if app is None and pid in pids():
            app = Win32Connector().attach()
        if app is not None:
            for index in range(app.Documents.Count, 0, -1):
                app.Documents(index).Saved = True
            app.Quit()
            del app
        document = None
        pythoncom.CoUninitialize()
        deadline = time.monotonic() + 30
        while pid in pids() and time.monotonic() < deadline:
            time.sleep(0.25)


def test_a_visio_we_started_has_our_rights(own_visio):
    connector = Win32Connector()
    assert connector.elevated(pid_of(own_visio)) == connector.elevated()


@live_tool(mutates=False)
def visio_count():
    return json.dumps({"docs": current_session().app.Documents.Count})


async def test_events_a_lost_server_left_off_are_restored(own_visio, events_marker):
    own_visio.EventsEnabled = events.EVENTS_OFF
    try:
        events_marker.parent.mkdir(parents=True, exist_ok=True)
        events_marker.write_text("")
        old = time.time() - events.MARKER_STALE_AFTER - 5
        os.utime(events_marker, (old, old))
        result = json.loads(await visio_count())
        assert [warning["code"] for warning in result["warnings"]] == ["events_restored"]
        assert own_visio.EventsEnabled == events.EVENTS_ON
        assert not events_marker.exists()
    finally:
        own_visio.EventsEnabled = events.EVENTS_ON
        com_runtime.reset()
