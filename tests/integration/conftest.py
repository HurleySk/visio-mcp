"""Undo scopes and the undo stack are app-wide, so these tests never touch a Visio the
user has open. Each module starts its own instance and checks that attach reached it."""
import time

import pytest

from tests.integration.helpers import launch_own, pid_of, pids
from visio_mcp import com_runtime

VIS_MS_US = 2


@pytest.fixture(autouse=True)
def fresh_runtime():
    com_runtime.reset()
    yield
    com_runtime.reset()


@pytest.fixture(scope="module")
def own_visio():
    import pythoncom

    pythoncom.CoInitialize()
    app = launch_own()
    pid = pid_of(app)
    yield app
    for index in range(app.Documents.Count, 0, -1):
        app.Documents(index).Saved = True
    app.Quit()
    del app
    pythoncom.CoUninitialize()
    # Later tests check for no Visio running, and Visio takes a moment to exit.
    deadline = time.monotonic() + 30
    while pid in pids() and time.monotonic() < deadline:
        time.sleep(0.25)


@pytest.fixture
def scratch(own_visio):
    document = own_visio.Documents.AddEx("", VIS_MS_US, 0)
    yield document
    document.Saved = True
    document.Close()
