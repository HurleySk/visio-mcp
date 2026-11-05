import importlib.util
import json
import pathlib

ROOT = pathlib.Path(__file__).resolve().parents[1]


def _render():
    spec = importlib.util.spec_from_file_location("tools_md", ROOT / "scripts" / "tools_md.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    tools = json.loads((ROOT / "tests" / "snapshots" / "live_tools.json").read_text(encoding="utf-8"))
    return module.render(tools)


def test_tools_md_matches_the_snapshot():
    """Regenerate with: uv run python scripts/tools_md.py tests/snapshots/live_tools.json TOOLS.md"""
    assert (ROOT / "TOOLS.md").read_text(encoding="utf-8") == _render()
