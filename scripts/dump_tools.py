import json
import pathlib
import sys

sys.path.insert(0, str(pathlib.Path(__file__).resolve().parents[1]))

from tests.handshake import handshake, live_tools

out = pathlib.Path(sys.argv[1])
result = handshake(sys.argv[2:])
tools = live_tools(result["tools"])
out.parent.mkdir(parents=True, exist_ok=True)
out.write_text(json.dumps(tools, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
print(len(tools))
