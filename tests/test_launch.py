import json
import threading

import pytest

from tests.fakes import FakeClock, FakeConnector
from visio_mcp import com_runtime
from visio_mcp.com_runtime import Win32Connector, current_session
from visio_mcp.errors import MK_E_UNAVAILABLE, VisioError
from visio_mcp.live_tool import live_tool

pythoncom = pytest.importorskip("pythoncom")
REGDB_E_CLASSNOTREG = -2147221164


def test_a_missing_visio_is_reported(monkeypatch):
    def refuse(*args):
        raise pythoncom.com_error(REGDB_E_CLASSNOTREG, "Class not registered", None, None)

    monkeypatch.setattr(pythoncom, "CoCreateInstance", refuse)
    with pytest.raises(VisioError) as caught:
        Win32Connector().launch()
    assert caught.value.code == "visio_not_installed"


def test_a_slow_start_does_not_hold_up_the_caller(monkeypatch):
    # Visio's boot, or a question it asks at start-up, must not use up the tool's time limit.
    entered = threading.Event()
    release = threading.Event()

    def slow(*args):
        entered.set()
        release.wait(10)

    monkeypatch.setattr(pythoncom, "CoCreateInstance", slow)
    monkeypatch.setattr(com_runtime, "LAUNCH_ERROR_WAIT", 0.05)
    try:
        Win32Connector().launch()
        assert entered.is_set()
    finally:
        release.set()


def test_it_activates_the_local_server_by_class(monkeypatch):
    calls = []
    monkeypatch.setattr(pythoncom, "CoCreateInstance", lambda *args: calls.append(args))
    Win32Connector().launch()
    assert calls == [(com_runtime.VISIO_CLSID, None, pythoncom.CLSCTX_LOCAL_SERVER, pythoncom.IID_IDispatch)]


class Unreachable(Win32Connector):
    """Visio is running, and Windows won't hand over a reference to it."""

    def __init__(self, mine, theirs):
        self.mine = mine
        self.theirs = theirs

    def visio_pids(self):
        return [101]

    def other_documents(self):
        return []

    def elevated(self, pid=None):
        if pid is None:
            return self.mine
        if isinstance(self.theirs, Exception):
            raise self.theirs
        return self.theirs


@pytest.fixture
def no_active_object(monkeypatch):
    def missing(clsid):
        raise pythoncom.com_error(MK_E_UNAVAILABLE, "Operation unavailable", None, None)

    monkeypatch.setattr(pythoncom, "GetActiveObject", missing)


@pytest.mark.parametrize(
    "mine, theirs",
    [(True, False), (False, True), (False, PermissionError("denied"))],
    ids=["server admin", "visio admin", "visio admin, so it won't say"],
)
def test_different_rights_get_their_own_explanation(no_active_object, mine, theirs):
    with pytest.raises(VisioError) as caught:
        Unreachable(mine, theirs).attach()
    assert caught.value.code == "visio_unreachable"
    assert "administrator" in caught.value.message
    assert "administrator" in caught.value.hint
    assert caught.value.retryable is False


@pytest.mark.parametrize("rights", [True, False])
def test_the_same_rights_keep_the_general_explanation(no_active_object, rights):
    with pytest.raises(VisioError) as caught:
        Unreachable(rights, rights).attach()
    assert caught.value.code == "visio_unreachable"
    assert "administrator" not in caught.value.message
    assert caught.value.retryable is True


@live_tool(mutates=True, start="if_missing")
def create_tool():
    return json.dumps({"docs": current_session().app.Documents.Count})


@live_tool(mutates=False)
def read_tool():
    return json.dumps({"docs": current_session().app.Documents.Count})


@pytest.fixture
def unreachable():
    fake = FakeConnector(apps=[], unreachable=True)
    clock = FakeClock()
    com_runtime.configure(lambda: fake, clock=clock, sleep=clock.sleep)
    yield fake
    com_runtime.reset()


async def test_a_second_window_is_reported_because_their_drawings_are_not_in_it(unreachable):
    result = json.loads(await create_tool())
    assert unreachable.launch_count == 1
    assert [warning["code"] for warning in result["warnings"]] == ["second_visio"]
    assert "cannot be reached" in result["warnings"][0]["message"]


async def test_the_warning_is_only_on_the_call_that_started_it(unreachable):
    await create_tool()
    assert "warnings" not in json.loads(await read_tool())


async def test_starting_the_only_visio_warns_of_nothing():
    fake = FakeConnector(apps=[])
    clock = FakeClock()
    com_runtime.configure(lambda: fake, clock=clock, sleep=clock.sleep)
    try:
        assert json.loads(await create_tool()) == {"docs": 0}
    finally:
        com_runtime.reset()
