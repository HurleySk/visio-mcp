import asyncio
import subprocess
import time

from visio_mcp import com_runtime


def pids():
    listing = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq VISIO.EXE", "/FO", "CSV", "/NH"],
        capture_output=True,
        text=True,
    ).stdout
    return {int(line.split('","')[1]) for line in listing.splitlines() if "VISIO.EXE" in line}


def pid_of(app):
    import win32process

    return win32process.GetWindowThreadProcessId(app.WindowHandle32)[1]


def dialogs_of(pid):
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


def launch_own():
    """Start a new instance and return it, once attach reaches that instance."""
    connector = com_runtime.Win32Connector()
    before = pids()
    connector.launch()
    deadline = time.monotonic() + 60
    while time.monotonic() < deadline:
        new = pids() - before
        if new:
            try:
                app = connector.attach()
                if pid_of(app) in new:
                    return app
            except Exception:
                pass
        time.sleep(0.5)
    raise AssertionError("attach did not reach the instance these tests started")


async def until(predicate, seconds=30):
    deadline = time.monotonic() + seconds
    while time.monotonic() < deadline:
        if predicate():
            return
        await asyncio.sleep(0.1)
    raise AssertionError("condition not met in time")
