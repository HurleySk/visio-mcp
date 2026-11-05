import asyncio
import json
import threading

import pytest

from tests.fakes import FakeApp, FakeClock, FakeComError, FakeConnector, FakeDoc
from visio_mcp import com_runtime
from visio_mcp.com_runtime import current_session, edit_doc
from visio_mcp.errors import RPC_E_CALL_REJECTED, VisioError
from visio_mcp.live_tool import live_tool


@live_tool(mutates=True)
def visio_rename_things(drawing=None, fail=None):
    document = edit_doc(drawing)
    edit_doc(drawing)
    if fail is not None:
        raise fail
    return json.dumps({"drawing": document.Name})


async def test_one_scope_per_mutating_call(connector):
    app = connector.apps[0]
    assert json.loads(await visio_rename_things()) == {"drawing": "flow.vsdx"}
    assert app.scopes_begun == ["rename things"]
    assert app.scopes_ended == [(101, True)]
    assert app.CurrentScope == -1

    await visio_rename_things()
    assert len(app.scopes_begun) == 2


async def test_exception_rolls_back(connector):
    app = connector.apps[0]
    result = json.loads(await visio_rename_things(fail=KeyError("boom")))
    assert result["code"] == "internal"
    assert result["rolled_back"] is True
    assert app.scopes_ended == [(101, False)]


async def test_validation_failure_opens_no_scope(connector):
    @live_tool(mutates=True)
    def strict():
        raise ValueError("bad spec")

    result = json.loads(await strict())
    assert result["code"] == "invalid_argument"
    assert "rolled_back" not in result
    assert connector.apps[0].scopes_begun == []


async def test_tool_without_edit_doc_opens_no_scope(connector):
    @live_tool(mutates=True)
    def save_like():
        return "{}"

    assert await save_like() == "{}"
    assert connector.apps[0].scopes_begun == []


async def test_read_only_drawing_is_refused_before_any_scope(connector):
    app = connector.apps[0]
    app.docs[0].ReadOnly = -1
    result = json.loads(await visio_rename_things())
    assert result["code"] == "read_only"
    assert "rolled_back" not in result
    assert app.scopes_begun == []


async def test_scope_opens_on_the_drawings_own_instance(connector):
    other_app = FakeApp()
    other = FakeDoc("other.vsdx", folder="C:\\elsewhere")
    other.Application = other_app
    connector.others.append(other)
    assert json.loads(await visio_rename_things("other.vsdx")) == {"drawing": "other.vsdx"}
    assert other_app.scopes_ended == [(101, True)]
    assert connector.apps[0].scopes_begun == []


async def test_busy_mid_edit_rolls_back_and_can_be_retried(connector):
    result = json.loads(await visio_rename_things(fail=FakeComError(RPC_E_CALL_REJECTED)))
    assert result["code"] == "busy"
    assert result["retryable"] is True
    assert result["rolled_back"] is True
    assert connector.apps[0].scopes_ended == [(101, False)]


async def test_failed_rollback_is_reported_and_not_retryable(connector):
    app = connector.apps[0]
    app.end_scope_error = FakeComError(RPC_E_CALL_REJECTED)
    result = json.loads(await visio_rename_things(fail=FakeComError(RPC_E_CALL_REJECTED)))
    assert result["code"] == "busy"
    assert result["retryable"] is False
    assert result["rolled_back"] is False


async def test_failed_commit_is_an_error_not_a_success(connector):
    connector.apps[0].end_scope_error = FakeComError(-2147352567, "cannot end")
    result = json.loads(await visio_rename_things())
    assert result["code"] == "com_error"


class _Stuck:
    """An edit that opens its scope, then hangs until released."""

    def __init__(self):
        self.in_scope = threading.Event()
        self.release = threading.Event()

        @live_tool(mutates=True, timeout=0.2)
        def stuck_edit():
            edit_doc()
            self.in_scope.set()
            self.release.wait(5)
            return json.dumps({"finished": True})

        self.tool = stuck_edit


async def _until(predicate):
    for _ in range(200):
        if predicate():
            return
        await asyncio.sleep(0.01)
    raise AssertionError("condition not met in time")


async def test_timed_out_edit_rolls_back(connector):
    app = connector.apps[0]
    stuck = _Stuck()
    result = json.loads(await stuck.tool())
    assert result["code"] == "timeout"
    assert result["retryable"] is False
    assert app.scopes_ended == []

    stuck.release.set()
    await _until(lambda: app.scopes_ended)
    assert app.scopes_ended == [(101, False)]


async def test_edits_are_refused_while_a_timed_out_edit_runs(connector):
    stuck = _Stuck()
    await stuck.tool()
    result = json.loads(await visio_rename_things())
    assert result["code"] == "edit_in_progress"
    assert result["retryable"] is True
    assert connector.apps[0].scopes_begun == ["stuck edit"]
    stuck.release.set()


async def test_reads_still_work_while_a_timed_out_edit_runs(connector):
    stuck = _Stuck()
    await stuck.tool()

    @live_tool(mutates=False)
    def read_tool():
        return json.dumps({"docs": current_session().app.Documents.Count})

    assert json.loads(await read_tool()) == {"docs": 1}
    stuck.release.set()


async def test_edits_resume_after_the_timed_out_edit_ends(connector):
    app = connector.apps[0]
    stuck = _Stuck()
    await stuck.tool()
    stuck.release.set()
    await _until(lambda: app.scopes_ended)
    await _until(lambda: not com_runtime._stuck_edits or all(f.done() for f in com_runtime._stuck_edits))
    assert json.loads(await visio_rename_things()) == {"drawing": "flow.vsdx"}
    assert app.scopes_ended == [(101, False), (102, True)]


@pytest.fixture
def leftover():
    fake = FakeConnector()
    clock = FakeClock()
    com_runtime.configure(lambda: fake, clock=clock, sleep=clock.sleep)
    yield fake.apps[0], clock
    com_runtime.reset()


async def test_leftover_scope_is_committed_and_reported(leftover):
    app, clock = leftover
    app.CurrentScope = 55
    result = json.loads(await visio_rename_things())
    assert result["drawing"] == "flow.vsdx"
    assert [w["code"] for w in result["warnings"]] == ["recovered_scope"]
    assert app.scopes_ended == [(55, True), (101, True)]
    assert clock.now >= com_runtime.SCOPE_WAIT


async def test_scope_that_closes_during_the_wait_is_left_alone(leftover):
    app, clock = leftover
    app.CurrentScope = 55
    real_sleep = com_runtime._sleep

    def sleep_then_close(seconds):
        real_sleep(seconds)
        app.CurrentScope = -1

    com_runtime._sleep = sleep_then_close
    result = json.loads(await visio_rename_things())
    assert "warnings" not in result
    assert app.scopes_ended == [(101, True)]
    assert clock.now < com_runtime.SCOPE_WAIT


async def test_warning_is_kept_on_a_failed_edit(leftover):
    app, _ = leftover
    app.CurrentScope = 55
    result = json.loads(await visio_rename_things(fail=KeyError("boom")))
    assert result["code"] == "internal"
    assert [w["code"] for w in result["warnings"]] == ["recovered_scope"]


async def test_warning_is_added_to_a_result_with_an_image(leftover):
    app, _ = leftover
    app.CurrentScope = 55

    @live_tool(mutates=True)
    def with_image():
        edit_doc()
        return [{"ok": True}, "image"]

    result = await with_image()
    assert result[0]["warnings"][0]["code"] == "recovered_scope"
    assert result[1] == "image"


async def test_edit_doc_only_works_on_the_com_thread():
    with pytest.raises(VisioError):
        edit_doc()
