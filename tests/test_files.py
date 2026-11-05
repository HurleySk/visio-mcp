import json

import pytest

from tests.fakes import FakeConnector, FakeDoc, FakePage
from visio_mcp import com_runtime, stencil_index
from visio_mcp.tools import files


@pytest.fixture
def app(connector):
    return connector.apps[0]


async def _call(tool, *args, **kwargs):
    return json.loads(await tool(*args, **kwargs))


async def test_new_blank_drawing(app):
    result = await _call(files.visio_new_drawing, units="metric")
    assert result["drawing"] == "Drawing1"
    assert result["path"] is None
    assert app.Documents.added == [("", files.VIS_MS_METRIC, 0)]


async def test_new_blank_drawing_uses_visios_default_units(app):
    await _call(files.visio_new_drawing)
    assert app.Documents.added == [("", files.VIS_MS_DEFAULT, 0)]


async def test_new_drawing_resolves_a_template_name(app, monkeypatch):
    monkeypatch.setattr(stencil_index, "resolve_template", lambda name, units: f"C:\\content\\{name}_{units}.VSTX")
    await _call(files.visio_new_drawing, "BASFLO", "metric")
    assert app.Documents.added == [("C:\\content\\BASFLO_metric.VSTX", files.VIS_MS_METRIC, 0)]


async def test_a_template_opens_with_events_off_so_its_wizard_cannot_block(app, monkeypatch, events_marker):
    monkeypatch.setattr(stencil_index, "resolve_template", lambda name, units: f"C:\\content\\{name}_{units}.VSTX")
    await _call(files.visio_new_drawing, "GANTT", "us")
    assert app.Documents.events_while_adding == [0]
    assert app.EventsEnabled == -1
    assert not events_marker.exists()


async def test_a_blank_drawing_keeps_events_on(app):
    await _call(files.visio_new_drawing)
    assert app.Documents.events_while_adding == [-1]


async def test_new_drawing_rejects_a_missing_template_path(app, tmp_path):
    result = await _call(files.visio_new_drawing, str(tmp_path / "missing.vstx"))
    assert result["code"] == "file_not_found"
    assert app.Documents.added == []


async def test_new_drawing_rejects_bad_units(app):
    assert (await _call(files.visio_new_drawing, units="imperial"))["code"] == "invalid_argument"


async def test_new_drawing_starts_visio_when_none_is_running():
    fake = FakeConnector(apps=[])
    com_runtime.configure(lambda: fake, sleep=lambda seconds: None)
    try:
        assert (await _call(files.visio_new_drawing))["drawing"] == "Drawing1"
        assert fake.launch_count == 1
    finally:
        com_runtime.reset()


async def test_open_drawing(app, tmp_path):
    path = tmp_path / "plan.vsdx"
    path.write_bytes(b"x")
    result = await _call(files.visio_open_drawing, str(path), read_only=True)
    assert result["drawing"] == "plan.vsdx"
    assert app.Documents.opened == [(str(path), files.VIS_OPEN_RO)]


async def test_open_drawing_that_is_already_open_returns_it(app, tmp_path):
    path = tmp_path / "flow.vsdx"
    path.write_bytes(b"x")
    app.docs[0].FullName = str(path)
    result = await _call(files.visio_open_drawing, str(path))
    assert result["already_open"] is True
    assert app.Documents.opened == []


@pytest.mark.parametrize(
    "path, code",
    [
        ("relative.vsdx", "invalid_argument"),
        ("C:\\no\\such\\file.vsdx", "file_not_found"),
        ("C:\\shapes.vssx", "invalid_argument"),
    ],
)
async def test_open_drawing_rejects_bad_paths(app, path, code):
    assert (await _call(files.visio_open_drawing, path))["code"] == code
    assert app.Documents.opened == []


async def test_save_in_place(app):
    await _call(files.visio_save)
    assert app.docs[0].calls == [("Save",)]


async def test_save_never_saved_drawing_needs_a_path(app):
    app.docs[0].Path = ""
    assert (await _call(files.visio_save))["code"] == "save_needs_path"
    assert app.docs[0].calls == []


async def test_save_refuses_existing_file(app, tmp_path):
    path = tmp_path / "taken.vsdx"
    path.write_bytes(b"keep me")
    result = await _call(files.visio_save, str(path))
    assert result["code"] == "file_exists"
    assert app.docs[0].calls == []
    assert path.read_bytes() == b"keep me"


async def test_save_replaces_with_overwrite(app, tmp_path):
    path = tmp_path / "taken.vsdx"
    path.write_bytes(b"old")
    await _call(files.visio_save, str(path), overwrite=True)
    assert app.docs[0].calls == [("SaveAs", str(path))]


async def test_save_to_its_own_path_is_a_plain_save(app, tmp_path):
    path = tmp_path / "flow.vsdx"
    path.write_bytes(b"x")
    document = app.docs[0]
    document.FullName = str(path)
    document.Path = f"{tmp_path}\\"
    await _call(files.visio_save, str(path))
    assert document.calls == [("Save",)]


async def test_read_only_save_refused_without_a_new_path(app):
    app.docs[0].ReadOnly = -1
    assert (await _call(files.visio_save))["code"] == "read_only"
    assert app.docs[0].calls == []


async def test_read_only_can_be_saved_to_a_new_path(app, tmp_path):
    app.docs[0].ReadOnly = -1
    path = tmp_path / "copy.vsdx"
    await _call(files.visio_save, str(path))
    assert app.docs[0].calls == [("SaveAs", str(path))]


async def test_save_checks_the_folder_and_extension(app, tmp_path):
    assert (await _call(files.visio_save, str(tmp_path / "no" / "x.vsdx")))["code"] == "folder_not_found"
    assert (await _call(files.visio_save, str(tmp_path / "x.png")))["code"] == "invalid_argument"
    assert app.docs[0].calls == []


async def test_export_refuses_existing_file(app, tmp_path):
    path = tmp_path / "page.png"
    path.write_bytes(b"keep me")
    result = await _call(files.visio_export, str(path))
    assert result["code"] == "file_exists"
    assert app.docs[0].Pages(1).exports == []
    assert path.read_bytes() == b"keep me"


async def test_export_image_replaces_with_overwrite(app, tmp_path):
    path = tmp_path / "page.png"
    path.write_bytes(b"old")
    result = await _call(files.visio_export, str(path), overwrite=True)
    assert result == {"path": str(path), "format": "png", "page": "Page-1"}
    assert app.docs[0].Pages(1).exports == [str(path)]


async def test_export_checks_the_format_before_touching_visio(app, tmp_path):
    assert (await _call(files.visio_export, str(tmp_path / "x.bmp")))["code"] == "invalid_argument"
    assert (await _call(files.visio_export, str(tmp_path / "x.png"), format="pdf"))["code"] == "invalid_argument"
    assert (await _call(files.visio_export, str(tmp_path / "no" / "x.png")))["code"] == "folder_not_found"
    assert app.docs[0].Pages(1).exports == []


async def test_export_pdf_of_every_page(app, tmp_path):
    path = str(tmp_path / "all.pdf")
    result = await _call(files.visio_export, path)
    assert result["page"] is None
    assert app.docs[0].calls == [("ExportAsFixedFormat", 1, path, 1, files.VIS_PRINT_ALL)]


async def test_export_pdf_of_one_page(app, tmp_path):
    document = FakeDoc(pages=[FakePage("One", 1), FakePage("Two", 2)])
    document.Application = app
    app.docs[:] = [document]
    path = str(tmp_path / "one.pdf")
    result = await _call(files.visio_export, path, page="Two")
    assert result["page"] == "Two"
    assert document.calls == [("ExportAsFixedFormat", 1, path, 1, files.VIS_PRINT_FROM_TO, 2, 2)]


async def test_close_refuses_dirty(app):
    app.docs[0].Saved = False
    result = await _call(files.visio_close_drawing)
    assert result["code"] == "unsaved_changes"
    assert app.docs[0].calls == []


async def test_close_saved_drawing(app):
    document = app.docs[0]
    assert await _call(files.visio_close_drawing) == {"closed": "flow.vsdx", "discarded": False}
    assert document.calls == [("Close", True)]


async def test_close_with_discard_marks_it_clean_first(app):
    document = app.docs[0]
    document.Saved = False
    result = await _call(files.visio_close_drawing, discard=True, drawing="flow.vsdx")
    assert result == {"closed": "flow.vsdx", "discarded": True}
    assert document.calls == [("Close", True)]


async def test_discard_must_name_the_drawing(app):
    app.docs[0].Saved = False
    assert (await _call(files.visio_close_drawing, discard=True))["code"] == "invalid_argument"
    assert app.docs[0].calls == []


@pytest.mark.parametrize("tool", [files.visio_save, files.visio_export, files.visio_close_drawing])
def test_file_tools_never_start_visio(tool):
    assert tool.start == "never"


def test_only_create_and_open_can_start_visio():
    assert files.visio_new_drawing.start == "if_missing"
    assert files.visio_open_drawing.start == "if_missing"
