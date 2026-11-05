import pathlib

SOURCE = sorted(pathlib.Path("visio_mcp").rglob("*.py"))


def test_no_quit_or_invisible_app():
    for path in SOURCE:
        text = path.read_text(encoding="utf-8")
        assert ".Quit(" not in text, path
        assert "InvisibleApp" not in text, path


def test_nothing_can_start_visio_through_dispatch():
    for path in SOURCE:
        text = path.read_text(encoding="utf-8")
        assert 'Dispatch("Visio.' not in text, path
        assert "DispatchEx" not in text, path


def test_tool_bodies_stay_synchronous():
    for path in SOURCE:
        if path.parent.name == "tools":
            text = path.read_text(encoding="utf-8")
            assert "async def" not in text, path
            assert "await " not in text, path
