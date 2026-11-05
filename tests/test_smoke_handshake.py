import inspect
import json
import pathlib
import sys

from tests.handshake import handshake, live_tools

SNAPSHOT = pathlib.Path(__file__).parent / "snapshots" / "live_tools.json"
COMMAND = [sys.executable, "-c", "from visio_mcp.server import run; run()"]
CONSOLE_SCRIPT = [str(pathlib.Path(sys.executable).with_name("visio-mcp.exe"))]
TOOL_COUNT = len(json.loads(SNAPSHOT.read_text(encoding="utf-8")))


def _normalised(tools):
    return [{**t, "description": inspect.cleandoc(t.get("description") or "")} for t in tools]


def test_tool_schemas_match_snapshot():
    expected = _normalised(json.loads(SNAPSHOT.read_text(encoding="utf-8")))
    actual = _normalised(live_tools(handshake(COMMAND)["tools"]))
    assert [t["name"] for t in actual] == [t["name"] for t in expected]
    assert actual == expected


def test_tool_count_is_25():
    assert TOOL_COUNT == 25


def test_console_script_starts_fast_and_exits_clean():
    handshake(CONSOLE_SCRIPT)
    result = handshake(CONSOLE_SCRIPT)
    assert len(result["tools"]) == TOOL_COUNT
    assert result["elapsed"] < 10.0
    assert result["exit_code"] == 0
    assert "FastMCP" not in result["stderr"]
