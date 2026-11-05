import inspect
import json
import threading

from tests.fakes import FakeComError
from visio_mcp.com_runtime import error_json, get_visio_app
from visio_mcp.errors import RPC_E_CALL_REJECTED
from visio_mcp.live_tool import live_tool


async def test_body_runs_on_the_com_thread_and_keeps_its_identity(connector):
    @live_tool(mutates=False)
    def visio_example(drawing: str = None, count: int = 1) -> str:
        """Example docstring."""
        return json.dumps({"thread": threading.current_thread().name, "count": count, "docs": get_visio_app().Documents.Count})

    assert inspect.iscoroutinefunction(visio_example)
    assert visio_example.__name__ == "visio_example"
    assert visio_example.__doc__ == "Example docstring."
    assert list(inspect.signature(visio_example).parameters) == ["drawing", "count"]
    assert visio_example.mutates is False
    result = json.loads(await visio_example(None, 7))
    assert result == {"thread": "visio-com", "count": 7, "docs": 1}


async def test_escaped_exception_becomes_error_json(connector):
    @live_tool(mutates=True)
    def broken():
        raise KeyError("oops")

    result = json.loads(await broken())
    assert set(result) == {"error", "code", "retryable", "hint"}
    assert result["code"] == "internal"


async def test_timeout_becomes_error_json(connector):
    release = threading.Event()

    @live_tool(mutates=False, timeout=0.2)
    def slow():
        release.wait(5)
        return "{}"

    result = json.loads(await slow())
    release.set()
    assert result["code"] == "timeout"
    assert result["retryable"] is True


async def test_read_only_tool_is_retried_after_mid_body_busy(connector):
    calls = []

    @live_tool(mutates=False)
    def flaky():
        calls.append(1)
        if len(calls) < 3:
            raise FakeComError(RPC_E_CALL_REJECTED)
        return json.dumps({"success": True})

    assert json.loads(await flaky()) == {"success": True}
    assert len(calls) == 3


async def test_mutating_tool_is_never_retried(connector):
    calls = []

    @live_tool(mutates=True)
    def flaky():
        calls.append(1)
        raise FakeComError(RPC_E_CALL_REJECTED)

    result = json.loads(await flaky())
    assert len(calls) == 1
    assert result["code"] == "busy"
    assert result["retryable"] is False


async def test_mutating_timeout_is_not_retryable(connector):
    release = threading.Event()

    @live_tool(mutates=True, timeout=0.2)
    def slow():
        release.wait(5)
        return "{}"

    result = json.loads(await slow())
    release.set()
    assert result["code"] == "timeout"
    assert result["retryable"] is False


async def test_malformed_timeout_env_falls_back(connector, monkeypatch):
    monkeypatch.setenv("VISIO_MCP_TIMEOUT", "soon")

    @live_tool(mutates=False)
    def quick():
        return "{}"

    assert await quick() == "{}"
