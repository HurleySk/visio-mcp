import json
import os

import pytest
from PIL import Image

from tests.integration.helpers import dialogs_of, pid_of
from visio_mcp import events, stencil_index
from visio_mcp.tools import files, masters

pytestmark = pytest.mark.visio


@pytest.fixture(autouse=True)
def visio(own_visio):
    """Every test here runs against this module's own instance."""
    return own_visio


async def _call(tool, *args, **kwargs):
    return json.loads(await tool(*args, **kwargs))


def _draw(app, name):
    app.Documents(name).Pages(1).DrawRectangle(1, 1, 3, 2)


async def test_new_save_close_open_round_trip(visio, tmp_path):
    path = str(tmp_path / "round.vsdx")
    made = await _call(files.visio_new_drawing, units="us")
    assert made["path"] is None
    _draw(visio, made["drawing"])

    dirty = await _call(files.visio_close_drawing, drawing=made["drawing"])
    assert dirty["code"] == "unsaved_changes"

    saved = await _call(files.visio_save, path, drawing=made["drawing"])
    assert saved["drawing"] == "round.vsdx"
    assert saved["path"] == path
    assert saved["saved"] is True
    assert os.path.getsize(path) > 0

    assert (await _call(files.visio_close_drawing, drawing="round.vsdx"))["discarded"] is False
    opened = await _call(files.visio_open_drawing, path)
    assert opened["read_only"] is False
    assert "already_open" not in opened
    assert visio.Documents("round.vsdx").Pages(1).Shapes.Count == 1

    again = await _call(files.visio_open_drawing, path)
    assert again["already_open"] is True
    assert (await _call(files.visio_close_drawing, drawing=path))["closed"] == "round.vsdx"


async def test_discard_closes_without_a_prompt(visio):
    made = await _call(files.visio_new_drawing)
    _draw(visio, made["drawing"])
    closed = await _call(files.visio_close_drawing, discard=True, drawing=made["drawing"])
    assert closed == {"closed": made["drawing"], "discarded": True}


async def test_read_only_save_refused(tmp_path):
    path = str(tmp_path / "ro.vsdx")
    made = await _call(files.visio_new_drawing)
    await _call(files.visio_save, path, drawing=made["drawing"])
    await _call(files.visio_close_drawing, drawing="ro.vsdx")

    opened = await _call(files.visio_open_drawing, path, read_only=True)
    assert opened["read_only"] is True
    assert (await _call(files.visio_save, drawing="ro.vsdx"))["code"] == "read_only"
    assert (await _call(files.visio_save, path, drawing="ro.vsdx"))["code"] == "read_only"

    copy = str(tmp_path / "ro-copy.vsdx")
    saved = await _call(files.visio_save, copy, drawing="ro.vsdx")
    assert saved["read_only"] is False
    assert saved["path"] == copy
    await _call(files.visio_close_drawing, drawing="ro-copy.vsdx")


async def test_save_refuses_to_replace_a_real_file(tmp_path):
    taken = tmp_path / "taken.vsdx"
    taken.write_bytes(b"keep me")
    made = await _call(files.visio_new_drawing)
    result = await _call(files.visio_save, str(taken), drawing=made["drawing"])
    assert result["code"] == "file_exists"
    assert taken.read_bytes() == b"keep me"
    await _call(files.visio_close_drawing, discard=True, drawing=made["drawing"])


@pytest.mark.parametrize("name", ["page.png", "page.jpg", "page.svg", "page.pdf"])
async def test_export_writes_each_format(visio, tmp_path, name):
    made = await _call(files.visio_new_drawing)
    _draw(visio, made["drawing"])
    path = tmp_path / name
    try:
        result = await _call(files.visio_export, str(path), drawing=made["drawing"])
        assert result["format"] == name.rsplit(".")[-1]
        assert path.stat().st_size > 0
        before = path.read_bytes()
        refused = await _call(files.visio_export, str(path), drawing=made["drawing"])
        assert refused["code"] == "file_exists"
        assert path.read_bytes() == before
        replaced = await _call(files.visio_export, str(path), overwrite=True, drawing=made["drawing"])
        assert "code" not in replaced
    finally:
        await _call(files.visio_close_drawing, discard=True, drawing=made["drawing"])


async def test_png_export_is_cropped_to_the_shapes_and_keeps_saved_state(visio, tmp_path):
    made = await _call(files.visio_new_drawing, units="us")
    document = visio.Documents(made["drawing"])
    _draw(visio, made["drawing"])
    document.Saved = True
    path = tmp_path / "crop.png"
    try:
        await _call(files.visio_export, str(path), drawing=made["drawing"])
        assert document.Saved
        with Image.open(path) as image:
            assert image.format == "PNG"
            assert image.size[0] / image.size[1] == pytest.approx(2.0, abs=0.1)
    finally:
        await _call(files.visio_close_drawing, discard=True, drawing=made["drawing"])


async def test_pdf_of_one_named_page(visio, tmp_path):
    made = await _call(files.visio_new_drawing)
    document = visio.Documents(made["drawing"])
    second = document.Pages.Add()
    second.Name = "Second"
    path = tmp_path / "one.pdf"
    try:
        result = await _call(files.visio_export, str(path), page="Second", drawing=made["drawing"])
        assert result["page"] == "Second"
        assert path.stat().st_size > 0
    finally:
        await _call(files.visio_close_drawing, discard=True, drawing=made["drawing"])


async def test_template_by_name_opens_its_stencils(visio):
    stencil_index.reset()
    made = await _call(files.visio_new_drawing, "BASFLO", "us")
    try:
        names = [visio.Documents(i).Name.upper() for i in range(1, visio.Documents.Count + 1)]
        assert "BASFLO_U.VSSX" in names
    finally:
        await _call(files.visio_close_drawing, discard=True, drawing=made["drawing"])


@pytest.mark.parametrize("template", ["GANTT", "XFUNC", "ORGWIZ"])
async def test_a_template_with_a_wizard_opens_without_it(visio, template):
    import win32con
    import win32gui

    stencil_index.reset()
    made = await _call(files.visio_new_drawing, template, "us")
    left = dialogs_of(pid_of(visio))
    for hwnd in left:
        win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    try:
        assert "error" not in made, made
        assert not left
        assert visio.EventsEnabled == events.EVENTS_ON
    finally:
        if "drawing" in made:
            await _call(files.visio_close_drawing, discard=True, drawing=made["drawing"])


async def test_installed_shapes_are_searchable():
    stencil_index.reset()
    found = json.loads(await masters.visio_find_masters("flowchart decision"))
    assert found["results"][0]["master"] == "BASFLO:Decision"
    assert stencil_index.resolve_master("BASFLO:Decision", "metric").path.upper().endswith("BASFLO_M.VSSX")
    templates = json.loads(await masters.visio_list_templates("flowchart"))["templates"]
    assert "BASFLO" in [row["template"] for row in templates]


async def test_corrupt_file_is_a_clean_error(tmp_path):
    path = tmp_path / "corrupt.vsdx"
    path.write_bytes(b"this is not a drawing")
    result = await _call(files.visio_open_drawing, str(path))
    assert result["code"] == "com_error"
