import asyncio
import json
import threading
import time

import pytest

from tests.fakes import FakeApp, FakeClock, FakeComError, FakeConnector, FakeDoc
from visio_mcp import com_runtime
from visio_mcp.com_runtime import classify_live, current_session, error_json, find_document, get_visio_app, run_com
from visio_mcp.errors import RPC_E_DISCONNECTED, VisioError


async def test_calls_run_on_the_visio_com_thread(connector):
    names = [await run_com(lambda s: threading.current_thread().name) for _ in range(3)]
    assert names == ["visio-com"] * 3
    assert connector.thread_names == ["visio-com"]
    assert connector.attach_count == 1


async def test_event_loop_stays_responsive(connector):
    ticks = 0

    async def ticker():
        nonlocal ticks
        while True:
            await asyncio.sleep(0.01)
            ticks += 1

    task = asyncio.create_task(ticker())
    await run_com(lambda s: time.sleep(0.3))
    task.cancel()
    assert ticks >= 10


async def test_timeout_then_fresh_worker(connector):
    release = threading.Event()
    with pytest.raises(VisioError) as caught:
        await run_com(lambda s: release.wait(5), timeout=0.2)
    assert caught.value.code == "timeout"
    assert caught.value.retryable is True
    assert await run_com(lambda s: "ok") == "ok"
    assert len(connector.thread_names) == 2
    release.set()


async def test_queued_call_fails_fast_when_worker_is_poisoned(connector):
    release = threading.Event()
    stuck = asyncio.create_task(run_com(lambda s: release.wait(5), timeout=0.2))
    await asyncio.sleep(0.05)
    started = time.perf_counter()
    with pytest.raises(VisioError) as caught:
        await run_com(lambda s: "never", timeout=5)
    assert caught.value.code == "timeout"
    assert time.perf_counter() - started < 1.0
    with pytest.raises(VisioError):
        await stuck
    release.set()


async def test_reattaches_after_disconnect(connector):
    await run_com(lambda s: s.app.Documents.Count)
    connector.apps[0].dead = True
    connector.apps.append(FakeApp())
    assert await run_com(lambda s: s.app.Documents.Count) == 1
    assert connector.attach_count == 2


async def test_visio_not_running():
    fake = FakeConnector(apps=[])
    com_runtime.configure(lambda: fake, sleep=lambda seconds: None)
    try:
        with pytest.raises(VisioError) as caught:
            await run_com(lambda s: s.app)
        assert caught.value.code == "visio_not_running"
    finally:
        com_runtime.reset()


async def test_waits_out_a_short_busy_spell(connector):
    connector.apps[0].busy_calls = 3
    assert await run_com(lambda s: s.app.Documents.Count) == 1


async def test_persistent_busy_is_reported_as_modal_dialog():
    fake = FakeConnector()
    fake.apps[0].busy_calls = 10**6
    clock = FakeClock()
    com_runtime.configure(lambda: fake, clock=clock, sleep=clock.sleep)
    try:
        with pytest.raises(VisioError) as caught:
            await run_com(lambda s: s.app.Documents.Count)
        assert caught.value.code == "modal_dialog"
        assert clock.now >= 10.0
    finally:
        com_runtime.reset()


async def test_busy_drops_the_proxy_so_a_typeless_one_is_never_kept(connector):
    connector.apps[0].busy_calls = 2
    await run_com(lambda s: s.app.Documents.Count)
    assert connector.attach_count == 3


async def test_rejected_name_lookup_during_a_dialog_is_busy_not_internal(connector):
    def body(session):
        session.app
        connector.apps[0].rejecting = True
        return classify_live(AttributeError("<unknown>.Documents")).code

    assert await run_com(body) == "busy"


async def test_a_real_attribute_error_stays_internal(connector):
    def body(session):
        session.app
        return classify_live(AttributeError("<unknown>.Typo")).code

    assert await run_com(body) == "internal"


async def test_get_visio_app_only_works_on_the_com_thread(connector):
    with pytest.raises(VisioError):
        get_visio_app()
    app = await run_com(lambda s: get_visio_app())
    assert app is connector.apps[0]


async def test_find_document(connector):
    name = await run_com(lambda s: find_document(s.app, "FLOW.vsdx").Name)
    assert name == "flow.vsdx"
    with pytest.raises(VisioError) as caught:
        await run_com(lambda s: find_document(s.app, "missing.vsdx"))
    assert caught.value.code == "drawing_not_found"
    connector.apps[0].docs.clear()
    with pytest.raises(VisioError) as caught:
        await run_com(lambda s: find_document(s.app, None))
    assert caught.value.code == "no_drawing"


async def test_find_document_ignores_stencils(connector):
    connector.apps[0].docs[:] = [FakeDoc("BASFLO_U.vssx", type_=2)]
    with pytest.raises(VisioError) as caught:
        await run_com(lambda s: find_document(s.app, None))
    assert caught.value.code == "no_drawing"
    with pytest.raises(VisioError) as caught:
        await run_com(lambda s: find_document(s.app, "BASFLO_U.vssx"))
    assert caught.value.code == "drawing_not_found"


async def test_find_document_prefers_a_drawing_when_a_stencil_is_active(connector):
    connector.apps[0].docs[:] = [FakeDoc("BASFLO_U.vssx", type_=2), FakeDoc("flow.vsdx")]
    assert await run_com(lambda s: find_document(s.app, None).Name) == "flow.vsdx"


async def test_find_document_reaches_a_saved_drawing_in_another_instance(connector):
    other = FakeDoc("other.vsdx", folder="C:\\elsewhere")
    connector.others.append(other)
    assert await run_com(lambda s: find_document(s.app, "other.vsdx")) is other
    assert await run_com(lambda s: find_document(s.app, "C:\\elsewhere\\other.vsdx")) is other


async def test_error_json_drops_the_app_on_disconnect(connector):
    def body(session):
        assert session.app is connector.apps[0]
        text = error_json(FakeComError(RPC_E_DISCONNECTED))
        return text, session._app, current_session() is session

    text, cached, same = await run_com(body)
    assert json.loads(text)["code"] == "disconnected"
    assert cached is None
    assert same is True


async def test_dead_worker_thread_is_replaced():
    class CrashingConnector(FakeConnector):
        def pump(self):
            raise RuntimeError("pump failed")

    crashing = CrashingConnector()
    healthy = FakeConnector()
    connectors = [crashing, healthy]
    com_runtime.configure(lambda: connectors.pop(0), sleep=lambda seconds: None)
    try:
        first = com_runtime._get_worker()
        first._thread.join(2)
        assert not first._thread.is_alive()
        assert first.poisoned
        result = await com_runtime.run_com(lambda session: session.app.Documents.Count, timeout=2)
        assert result == 1
        assert healthy.attach_count == 1
    finally:
        com_runtime.reset()


async def test_result_that_lands_at_the_timeout_is_kept(connector, monkeypatch):
    async def late_timeout(awaitable, timeout):
        await awaitable
        raise asyncio.TimeoutError

    monkeypatch.setattr(com_runtime.asyncio, "wait_for", late_timeout)
    assert await com_runtime.run_com(lambda session: 42, timeout=1) == 42
    assert not com_runtime._get_worker().poisoned
