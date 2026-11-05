"""Write the tool snapshot that tests/test_smoke_handshake.py checks against.

Usage: python scripts/dump_tools.py tests/snapshots/live_tools.json [server command...]
The server command defaults to this checkout's own visio-mcp.
"""
import json
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from tests.handshake import handshake, live_tools

out = pathlib.Path(sys.argv[1])
command = sys.argv[2:] or [str(ROOT / ".venv" / "Scripts" / "visio-mcp.exe")]
result = handshake(command)
tools = live_tools(result["tools"])
out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(json.dumps(tools, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print(len(tools))
