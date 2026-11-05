import json

import pytest

from tests.fakes import FakeConnector, FakeDoc, FakePage, FakeShape
from visio_mcp import com_runtime
from visio_mcp.tools import read


def _use_page(app, page):
    document = FakeDoc(pages=[page])
    document.Application = app
    app.docs[:] = [document]


async def _call(tool, *args, **kwargs):
    return json.loads(await tool(*args, **kwargs))


async def test_get_page_pages_results(connector):
    _use_page(connector.apps[0], FakePage(shapes=[FakeShape(i, f"Sheet.{i}") for i in range(1, 6)]))

    first = await _call(read.visio_get_page, 0, 2)
    assert first["total"] == 5
    assert first["next_offset"] == 2
    assert [row["id"] for row in first["shapes"]] == [1, 2]

    second = await _call(read.visio_get_page, first["next_offset"], 2)
    assert [row["id"] for row in second["shapes"]] == [3, 4]

    last = await _call(read.visio_get_page, second["next_offset"], 2)
    assert [row["id"] for row in last["shapes"]] == [5]
    assert last["next_offset"] is None


async def test_get_page_past_the_end_is_empty(connector):
    _use_page(connector.apps[0], FakePage(shapes=[FakeShape(1, "a")]))
    result = await _call(read.visio_get_page, 10, 5)
    assert result["shapes"] == []
    assert result["next_offset"] is None


async def test_get_page_rejects_bad_paging(connector):
    assert (await _call(read.visio_get_page, -1))["code"] == "invalid_argument"
    assert (await _call(read.visio_get_page, 0, 0))["code"] == "invalid_argument"


async def test_read_does_not_mint_guids(connector):
    shape = FakeShape(1, "Sheet.1")
    _use_page(connector.apps[0], FakePage(shapes=[shape]))
    result = await _call(read.visio_get_shape, 1)
    assert "guid" not in result
    assert shape.guid_flags == [read.VIS_GET_GUID]


async def test_find_shapes_needs_a_filter(connector):
    assert (await _call(read.visio_find_shapes))["code"] == "invalid_argument"


async def test_find_shapes_matches_part_of_the_text_ignoring_case(connector):
    _use_page(connector.apps[0], FakePage(shapes=[FakeShape(1, "a", "Approve Request"), FakeShape(2, "b", "Reject")]))
    result = await _call(read.visio_find_shapes, text="approve")
    assert result["total"] == 1
    assert result["shapes"][0]["id"] == 1


@pytest.mark.parametrize("wanted", ["Process", "process", "BASFLO:Process"])
async def test_find_shapes_by_master(connector, wanted):
    _use_page(connector.apps[0], FakePage(shapes=[FakeShape(1, "a", master="Process"), FakeShape(2, "b")]))
    result = await _call(read.visio_find_shapes, master=wanted)
    assert [row["id"] for row in result["shapes"]] == [1]


async def test_list_open_with_no_visio_starts_nothing():
    fake = FakeConnector(apps=[])
    com_runtime.configure(lambda: fake, sleep=lambda seconds: None)
    try:
        assert await _call(read.visio_list_open) == {"running": False, "reachable": False, "drawings": []}
        assert fake.launch_count == 0
    finally:
        com_runtime.reset()


async def test_list_open_says_when_visio_cant_be_reached():
    fake = FakeConnector(apps=[], unreachable=True)
    com_runtime.configure(lambda: fake, sleep=lambda seconds: None)
    try:
        result = await _call(read.visio_list_open)
        assert result["running"] is True
        assert result["reachable"] is False
        assert fake.launch_count == 0
    finally:
        com_runtime.reset()


async def test_list_open_marks_the_active_drawing_and_skips_stencils(connector):
    app = connector.apps[0]
    other = FakeDoc("other.vsdx")
    stencil = FakeDoc("BASFLO_U.vssx", type_=2)
    for document in (other, stencil):
        document.Application = app
        app.docs.append(document)
    result = await _call(read.visio_list_open)
    assert [(row["drawing"], row["active"]) for row in result["drawings"]] == [("flow.vsdx", True), ("other.vsdx", False)]


async def test_list_open_includes_another_instance_once(connector):
    app = connector.apps[0]
    far = FakeDoc("far.vsdx", folder="C:\\elsewhere")
    connector.others.extend([far, app.docs[0]])
    result = await _call(read.visio_list_open)
    assert [row["drawing"] for row in result["drawings"]] == ["flow.vsdx", "far.vsdx"]
