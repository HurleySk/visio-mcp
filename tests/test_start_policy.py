import json

import pytest

from tests.fakes import FakeClock, FakeConnector
from visio_mcp import com_runtime
from visio_mcp.com_runtime import current_session
from visio_mcp.live_tool import live_tool


@pytest.fixture
def stopped():
    """No Visio running, with a clock the runtime's waits advance."""
    fake = FakeConnector(apps=[])
    clock = FakeClock()
    com_runtime.configure(lambda: fake, clock=clock, sleep=clock.sleep)
    yield fake, clock
    com_runtime.reset()


@live_tool(mutates=False)
def read_tool():
    return json.dumps({"docs": current_session().app.Documents.Count})


@live_tool(mutates=True, start="if_missing")
def create_tool():
    return json.dumps({"docs": current_session().app.Documents.Count})


@live_tool(mutates=False, start="report")
def report_tool():
    err = current_session().attach_error
    return json.dumps({"running": err is None, "reason": err.code if err else None})


async def test_read_tool_never_starts_visio(stopped):
    fake, _ = stopped
    result = json.loads(await read_tool())
    assert result["code"] == "visio_not_running"
    assert fake.launch_count == 0


async def test_read_tool_reports_an_unreachable_visio_and_starts_nothing():
    fake = FakeConnector(apps=[], unreachable=True)
    com_runtime.configure(lambda: fake, sleep=lambda seconds: None)
    try:
        result = json.loads(await read_tool())
        assert result["code"] == "visio_unreachable"
        assert result["retryable"] is True
        assert fake.launch_count == 0
    finally:
        com_runtime.reset()


async def test_report_runs_the_body_with_no_visio(stopped):
    fake, _ = stopped
    assert json.loads(await report_tool()) == {"running": False, "reason": "visio_not_running"}
    assert fake.launch_count == 0


async def test_report_runs_the_body_with_visio(connector):
    assert json.loads(await report_tool()) == {"running": True, "reason": None}


async def test_if_missing_does_not_launch_when_visio_is_running(connector):
    assert json.loads(await create_tool()) == {"docs": 1}
    assert connector.launch_count == 0


async def test_if_missing_launches_and_waits_for_the_boot(stopped):
    fake, _ = stopped
    fake.boot_attempts = 5
    assert json.loads(await create_tool()) == {"docs": 0}
    assert fake.launch_count == 1


async def test_no_second_launch_during_boot(stopped):
    fake, clock = stopped
    fake.boot_attempts = 10**6
    first = json.loads(await create_tool())
    assert first["code"] == "visio_start_timeout"
    assert first["retryable"] is True
    assert clock.now < com_runtime.LAUNCH_GRACE

    fake.boot_attempts = 2
    assert json.loads(await create_tool()) == {"docs": 0}
    assert fake.launch_count == 1


async def test_a_replaced_worker_does_not_launch_again(stopped):
    fake, _ = stopped
    fake.boot_attempts = 10**6
    await create_tool()
    com_runtime._get_worker().poison()
    fake.boot_attempts = 0
    await create_tool()
    assert fake.launch_count == 1


async def test_launches_again_once_the_last_launch_is_stale(stopped):
    fake, clock = stopped
    fake.boot_attempts = 10**6
    await create_tool()
    clock.now += com_runtime.LAUNCH_GRACE + 1
    await create_tool()
    assert fake.launch_count == 2


async def test_an_unreachable_visio_gets_a_new_instance_rather_than_a_wait():
    fake = FakeConnector(apps=[], unreachable=True)
    clock = FakeClock()
    com_runtime.configure(lambda: fake, clock=clock, sleep=clock.sleep)
    try:
        assert json.loads(await create_tool())["docs"] == 0
        assert fake.launch_count == 1
        assert clock.now < 1.0
    finally:
        com_runtime.reset()
