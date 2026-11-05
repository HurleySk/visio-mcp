import json
import os
import time

import pytest

from tests.fakes import FakeApp, FakeConnector
from visio_mcp import com_runtime, events
from visio_mcp.building import StencilPool
from visio_mcp.com_runtime import current_session
from visio_mcp.live_tool import live_tool
from visio_mcp.stencil_index import MasterRef


class Cell:
    def __init__(self, formula):
        self.FormulaU = formula


class Sheet:
    def __init__(self, handler="", children=None):
        self.handler = handler
        self.Type = 2 if children is not None else 3
        self.Shapes = Sheets(children or [])

    def CellExistsU(self, name, flag):
        return True

    def CellsU(self, name):
        return Cell(self.handler)


class Sheets:
    def __init__(self, rows):
        self._rows = rows
        self.reads = 0

    @property
    def Count(self):
        self.reads += 1
        return len(self._rows)

    def __call__(self, index):
        return self._rows[index - 1]


class Master:
    def __init__(self, *sheets):
        self.Shapes = Sheets(list(sheets))


class Masters:
    def __init__(self, masters):
        self._masters = masters

    def ItemU(self, name):
        return self._masters[name]


class Document:
    def __init__(self, masters):
        self.Masters = Masters(masters)


class Page:
    def __init__(self, app, error=None):
        self.app = app
        self.error = error
        self.seen = []

    def Drop(self, master, x, y):
        self.seen.append(self.app.EventsEnabled)
        if self.error:
            raise self.error
        return master


@pytest.fixture
def app():
    return FakeApp(docs=[])


def pool_with(app, **masters):
    return StencilPool(app, Document({name: Master(sheet) for name, sheet in masters.items()}))


def ref(name):
    return MasterRef("", name, None)


def set_up_app(app):
    app.EventsEnabled = events.EVENTS_ON
    return app


def test_a_master_that_opens_a_dialog_drops_with_events_off(app, events_marker):
    set_up_app(app)
    page = Page(app)
    pool = pool_with(app, Position=Sheet('RUNADDONWARGS("OrgC11","/cmd=ShapeDropped")'))
    pool.drop(page, ref("Position"), 1, 1)
    assert page.seen == [events.EVENTS_OFF]
    assert app.EventsEnabled == events.EVENTS_ON
    assert not events_marker.exists()
    assert "Position" in pool.skipped_warning()["message"]


@pytest.mark.parametrize("handler", ["", "0", "SETF(GetRef(HideText),FALSE)"])
def test_a_handler_that_only_sets_up_the_shape_keeps_running(app, handler):
    set_up_app(app)
    page = Page(app)
    pool = pool_with(app, Plain=Sheet(handler))
    pool.drop(page, ref("Plain"), 1, 1)
    assert page.seen == [events.EVENTS_ON]
    assert pool.skipped_warning() is None


def test_a_handler_deep_in_a_group_counts(app):
    set_up_app(app)
    page = Page(app)
    deep = Sheet(children=[Sheet(), Sheet(children=[Sheet("DOCMD(1312)")])])
    pool = pool_with(app, Group=deep)
    pool.drop(page, ref("Group"), 1, 1)
    assert page.seen == [events.EVENTS_OFF]


def test_events_come_back_when_the_drop_fails(app, events_marker):
    set_up_app(app)
    page = Page(app, error=RuntimeError("no"))
    pool = pool_with(app, Position=Sheet("RUNADDON(1)"))
    with pytest.raises(RuntimeError):
        pool.drop(page, ref("Position"), 1, 1)
    assert app.EventsEnabled == events.EVENTS_ON
    assert not events_marker.exists()


def test_each_master_is_checked_once_per_call(app):
    set_up_app(app)
    page = Page(app)
    master = Master(Sheet("RUNADDON(1)"))
    pool = StencilPool(app, Document({"Position": master}))
    pool.drop(page, ref("Position"), 1, 1)
    pool.drop(page, ref("Position"), 2, 2)
    assert master.Shapes.reads == 1
    assert pool.skipped_warning()["message"].count("Position") == 1


def test_events_the_user_turned_off_stay_off(app, events_marker):
    app.EventsEnabled = events.EVENTS_OFF
    with events.events_off(app):
        assert not events_marker.exists()
    assert app.EventsEnabled == events.EVENTS_OFF
    assert not events_marker.exists()


def test_the_marker_is_there_while_events_are_off(app, events_marker):
    set_up_app(app)
    with events.events_off(app):
        assert events_marker.exists()
    assert not events_marker.exists()


class RefusesToRestore(FakeApp):
    def __init__(self):
        self._events = events.EVENTS_ON
        self._gone = False
        super().__init__(docs=[])
        self._gone = True

    @property
    def EventsEnabled(self):
        return self._events

    @EventsEnabled.setter
    def EventsEnabled(self, value):
        if self._gone and value == events.EVENTS_ON:
            raise RuntimeError("Visio has gone")
        self._events = value


def test_the_marker_stays_when_events_could_not_be_restored(events_marker):
    with events.events_off(RefusesToRestore()):
        pass
    assert events_marker.exists()


def _leave_marker(marker, age):
    marker.parent.mkdir(parents=True, exist_ok=True)
    marker.write_text("")
    old = time.time() - age
    os.utime(marker, (old, old))


def test_a_stale_marker_puts_events_back(app, events_marker):
    app.EventsEnabled = events.EVENTS_OFF
    _leave_marker(events_marker, events.MARKER_STALE_AFTER + 5)
    assert events.restore_if_lost(app) is True
    assert app.EventsEnabled == events.EVENTS_ON
    assert not events_marker.exists()


def test_a_recent_marker_may_be_another_server_at_work(app, events_marker):
    app.EventsEnabled = events.EVENTS_OFF
    _leave_marker(events_marker, 1)
    assert events.restore_if_lost(app) is False
    assert app.EventsEnabled == events.EVENTS_OFF
    assert events_marker.exists()


def test_no_marker_leaves_the_setting_alone(app):
    app.EventsEnabled = events.EVENTS_OFF
    assert events.restore_if_lost(app) is False
    assert app.EventsEnabled == events.EVENTS_OFF


@live_tool(mutates=False)
def read_tool():
    return json.dumps({"docs": current_session().app.Documents.Count})


async def test_the_next_call_restores_events_and_says_so(connector, events_marker):
    app = connector.apps[0]
    app.EventsEnabled = events.EVENTS_OFF
    _leave_marker(events_marker, events.MARKER_STALE_AFTER + 5)
    result = json.loads(await read_tool())
    assert app.EventsEnabled == events.EVENTS_ON
    assert [warning["code"] for warning in result["warnings"]] == ["events_restored"]
    assert "warnings" not in json.loads(await read_tool())


async def test_failing_to_restore_events_does_not_fail_the_call(connector, events_marker):
    connector.apps[0] = RefusesToRestore()
    _leave_marker(events_marker, events.MARKER_STALE_AFTER + 5)
    assert json.loads(await read_tool()) == {"docs": 0}
    assert events_marker.exists()


async def test_a_call_with_no_visio_leaves_the_marker_for_later(events_marker):
    fake = FakeConnector(apps=[])
    com_runtime.configure(lambda: fake, sleep=lambda seconds: None)
    try:
        _leave_marker(events_marker, events.MARKER_STALE_AFTER + 5)
        assert json.loads(await read_tool())["code"] == "visio_not_running"
        assert events_marker.exists()
    finally:
        com_runtime.reset()
