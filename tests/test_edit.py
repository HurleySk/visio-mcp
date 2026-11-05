import json

import pytest

from tests.fakes import FakeClock, FakeConnector, FakeDoc, FakePage, FakePages, FakeShape
from visio_mcp import com_runtime
from visio_mcp.models import EdgeSpec, LayoutSpec, NodeSpec, ShapeUpdate
from visio_mcp.tools import edit


@pytest.fixture
def page(connector):
    target = FakePage(shapes=[FakeShape(1, "Process", "Pay", master="Process"), FakeShape(2, "Decision", "OK?", master="Decision")])
    connector.apps[0].docs[0].Pages = FakePages([target])
    return target


def _changed_nothing(connector):
    assert connector.apps[0].scopes_begun == []


async def test_update_echoes_identity(connector, page):
    result = json.loads(await edit.visio_update_shapes([ShapeUpdate(shape="Pay", text="Pay supplier")]))
    assert result["updated"] == [{"id": 1, "name": "Process", "text": "Pay supplier", "master": "Process"}]
    assert connector.apps[0].scopes_begun == ["update shapes"]
    assert [commit for _, commit in connector.apps[0].scopes_ended] == [True]


async def test_two_updates_to_one_shape_echo_it_once(connector, page):
    result = json.loads(
        await edit.visio_update_shapes([ShapeUpdate(shape=1, text="a"), ShapeUpdate(shape="Process", text="b")])
    )
    assert [row["text"] for row in result["updated"]] == ["b"]


async def test_an_unknown_shape_in_any_update_changes_nothing(connector, page):
    result = json.loads(
        await edit.visio_update_shapes([ShapeUpdate(shape="Pay", text="changed"), ShapeUpdate(shape=99, text="x")])
    )
    assert result["code"] == "shape_not_found"
    assert page.Shapes(1).Text == "Pay"
    _changed_nothing(connector)


@pytest.mark.parametrize(
    "update",
    [{"shape": 1}, {"shape": 1, "fill": "red"}, {"shape": 1, "w": 0}, {"shape": 1, "data": {"": "x"}}],
)
async def test_a_bad_update_is_refused(connector, page, update):
    result = json.loads(await edit.visio_update_shapes([ShapeUpdate.model_validate(update)]))
    assert result["code"] == "invalid_argument"
    _changed_nothing(connector)


async def test_delete_page_refuses_last_page(connector, page):
    result = json.loads(await edit.visio_delete_page(1))
    assert result["code"] == "last_page"
    _changed_nothing(connector)


async def test_a_background_page_is_not_a_second_page(connector, page):
    background = FakePage("Background-1", index=2)
    background.Background = -1
    connector.apps[0].docs[0].Pages = FakePages([page, background])
    result = json.loads(await edit.visio_delete_page("Page-1"))
    assert result["code"] == "last_page"


async def test_delete_page_needs_a_page_that_exists(connector, page):
    result = json.loads(await edit.visio_delete_page("Nope"))
    assert result["code"] == "page_not_found"
    assert result["details"] == ["Page-1"]


def _undo_removes_a_shape(app, page):
    def undo():
        page.Shapes._shapes.pop()

    app.Undo = undo


async def test_undo_reports_shape_counts(connector, page):
    app = connector.apps[0]
    other = FakeDoc("other.vsdx", pages=[FakePage(shapes=[FakeShape(5, "Box")])])
    other.Application = app
    app.docs.append(other)
    _undo_removes_a_shape(app, page)
    result = json.loads(await edit.visio_undo())
    assert result == {
        "steps": 1,
        "drawings": [
            {"drawing": "flow.vsdx", "shapes_before": 2, "shapes_after": 1, "pages_before": 1, "pages_after": 1},
            {"drawing": "other.vsdx", "shapes_before": 1, "shapes_after": 1, "pages_before": 1, "pages_after": 1},
        ],
    }
    _changed_nothing(connector)


@pytest.mark.parametrize("steps", [0, 100, True])
async def test_undo_steps_are_checked(connector, page, steps):
    called = []
    connector.apps[0].Undo = lambda: called.append(1)
    result = json.loads(await edit.visio_undo(steps))
    assert result["code"] == "invalid_argument"
    assert called == []


async def test_undo_first_closes_a_scope_left_open():
    fake = FakeConnector()
    clock = FakeClock()
    com_runtime.configure(lambda: fake, clock=clock, sleep=clock.sleep)
    try:
        app = fake.apps[0]
        app.CurrentScope = 55
        order = []
        app.Undo = lambda: order.append(("undo", app.CurrentScope))
        result = json.loads(await edit.visio_undo())
        assert app.scopes_ended == [(55, True)]
        assert order == [("undo", -1)]
        assert [warning["code"] for warning in result["warnings"]] == ["recovered_scope"]
    finally:
        com_runtime.reset()


async def test_add_elements_needs_something_to_add(connector, page):
    result = json.loads(await edit.visio_add_elements())
    assert result["code"] == "invalid_argument"
    _changed_nothing(connector)


async def test_an_edge_to_a_missing_shape_changes_nothing(connector, page):
    result = json.loads(await edit.visio_add_elements(edges=[EdgeSpec.model_validate({"from": "Pay", "to": "Nowhere"})]))
    assert result["code"] == "shape_not_found"
    _changed_nothing(connector)


async def test_a_mistyped_key_is_suggested(connector, page):
    result = json.loads(
        await edit.visio_add_elements(
            nodes=[NodeSpec(key="audit", master="BASFLO:Process")],
            edges=[EdgeSpec.model_validate({"from": "Pay", "to": "audt"})],
        )
    )
    assert result["code"] == "unknown_node"
    assert result["details"] == ["audit"]
    _changed_nothing(connector)


async def test_positions_and_a_layout_still_conflict(connector, page):
    result = json.loads(
        await edit.visio_add_elements(
            nodes=[NodeSpec(key="a", master="BASFLO:Process", x=1, y=1)],
            layout=LayoutSpec(style="left_right"),
        )
    )
    assert result["code"] == "invalid_argument"
    _changed_nothing(connector)


async def test_add_page_refuses_a_taken_name(connector, page):
    result = json.loads(await edit.visio_add_page(" page-1 "))
    assert result["code"] == "page_exists"
    _changed_nothing(connector)


@pytest.mark.parametrize(
    "fields",
    [{}, {"fit": False}, {"fit": True, "width": 5}, {"height": 0}, {"theme": ""}, {"theme": -1}, {"name": "  "}],
)
async def test_a_bad_page_update_is_refused(connector, page, fields):
    result = json.loads(await edit.visio_update_page(**fields))
    assert result["code"] == "invalid_argument"
    _changed_nothing(connector)


async def test_a_shape_inside_a_group_is_refused(connector, page):
    group = FakeShape(7, "Group")
    group.Type = 2
    page.Shapes(1).ContainingShape = group
    result = json.loads(await edit.visio_delete_shapes([1]))
    assert result["code"] == "inside_group"
    _changed_nothing(connector)


async def test_a_container_needs_shapes(connector, page):
    result = json.loads(await edit.visio_add_container([], "Team"))
    assert result["code"] == "invalid_argument"
    _changed_nothing(connector)
