import json

import pytest
from fastmcp import Client

from tests.integration.helpers import FLOW
from visio_mcp import building, server
from visio_mcp.models import DiagramSpec, EdgeSpec, NodeSpec, ShapeUpdate
from visio_mcp.tools import diagnose, edit, generate, read

pytestmark = pytest.mark.visio


async def _flow(app):
    result = json.loads(await generate.visio_create_diagram(DiagramSpec.model_validate({**FLOW, "return_image": False})))
    assert "error" not in result, result
    return result, app.Documents(result["drawing"]).Pages(1)


def _edge(start, end):
    return EdgeSpec.model_validate({"from": start, "to": end})


def _ends(page, connector_id):
    connector = page.Shapes.ItemFromID(connector_id)
    connects = connector.Connects
    return {connects(i).FromCell.Name: connects(i).ToSheet.ID for i in range(1, connects.Count + 1)}


def _positions(page):
    shapes = page.Shapes
    return {
        shapes(i).ID: (round(shapes(i).CellsU("PinX").ResultIU, 4), round(shapes(i).CellsU("PinY").ResultIU, 4))
        for i in range(1, shapes.Count + 1)
        if not shapes(i).OneD
    }


async def test_failed_edit_rolls_back(made):
    result, page = await _flow(made)
    width = page.PageSheet.CellsU("PageWidth").ResultIU
    reply = json.loads(
        await edit.visio_update_page(name="Renamed", width=20, theme="No such theme", drawing=result["drawing"])
    )
    assert reply["code"] == "unknown_theme"
    assert reply["rolled_back"] is True
    assert page.Name == "Page-1"
    assert page.PageSheet.CellsU("PageWidth").ResultIU == width
    assert made.CurrentScope == -1


async def test_a_failed_add_leaves_the_page_as_it_was(made, monkeypatch):
    result, page = await _flow(made)
    before = _positions(page)

    def fail(*args, **kwargs):
        raise RuntimeError("layout went wrong")

    monkeypatch.setattr(building, "lay_out", fail)
    reply = json.loads(
        await edit.visio_add_elements(
            nodes=[NodeSpec(key="audit", master="BASFLO:Process", text="Audit")],
            edges=[_edge("Pay", "audit")],
            drawing=result["drawing"],
        )
    )
    assert reply["rolled_back"] is True
    assert page.Shapes.Count == 10
    assert _positions(page) == before


async def test_add_elements_joins_existing_shapes(made):
    result, page = await _flow(made)
    reply = json.loads(
        await edit.visio_add_elements(
            nodes=[NodeSpec(key="audit", master="BASFLO:Process", text="Audit")],
            edges=[_edge("Pay", "audit"), _edge("audit", "End")],
            drawing=result["drawing"],
        )
    )
    assert "error" not in reply, reply
    audit = reply["ids"]["audit"]
    first, second = reply["connectors"]
    assert _ends(page, first["id"]) == {"BeginX": result["ids"]["pay"], "EndX": audit}
    assert _ends(page, second["id"]) == {"BeginX": audit, "EndX": result["ids"]["end"]}
    assert reply["issues"] == []
    assert page.Shapes.Count == 13
    made.Undo()
    assert page.Shapes.Count == 10


async def test_nodes_with_positions_move_nothing_else(made):
    result, page = await _flow(made)
    before = _positions(page)
    reply = json.loads(
        await edit.visio_add_elements(
            nodes=[NodeSpec(key="note", master="BASIC:Rectangle", text="Note", x=0.75, y=0.75)],
            drawing=result["drawing"],
        )
    )
    assert "error" not in reply, reply
    after = _positions(page)
    assert after.pop(reply["ids"]["note"]) == (0.75, 0.75)
    assert after == before


async def test_edges_alone_move_nothing(made):
    result, page = await _flow(made)
    before = _positions(page)
    reply = json.loads(await edit.visio_add_elements(edges=[_edge("Reject", "Start")], drawing=result["drawing"]))
    assert "error" not in reply, reply
    assert _ends(page, reply["connectors"][0]["id"]) == {"BeginX": result["ids"]["reject"], "EndX": result["ids"]["start"]}
    assert _positions(page) == before


async def test_update_shapes_changes_what_it_names(made):
    result, page = await _flow(made)
    pay = result["ids"]["pay"]
    reply = json.loads(
        await edit.visio_update_shapes(
            [ShapeUpdate(shape=pay, text="Pay now", x=1, y=1, w=2, fill="#FF0000", data={"Owner": "Ann"})],
            drawing=result["drawing"],
        )
    )
    assert reply["updated"] == [{"id": pay, "name": "Process", "text": "Pay now", "master": "Process"}]
    shape = page.Shapes.ItemFromID(pay)
    assert (shape.CellsU("PinX").ResultIU, shape.CellsU("PinY").ResultIU, shape.CellsU("Width").ResultIU) == (1, 1, 2)
    assert shape.CellsU("FillForegnd").FormulaU == "THEMEGUARD(RGB(255,0,0))"
    assert {label: value for label, _, value in read.shape_data(shape)}["Owner"] == "Ann"
    made.Undo()
    assert page.Shapes.ItemFromID(pay).Text == "Pay"


async def test_delete_shapes_takes_glued_connectors(made):
    result, page = await _flow(made)
    reply = json.loads(await edit.visio_delete_shapes(["Reject"], drawing=result["drawing"]))
    assert [row["text"] for row in reply["deleted"]] == ["Reject"]
    assert {row["text"] for row in reply["connectors_deleted"]} == {"no", ""}
    assert page.Shapes.Count == 7
    issues = json.loads(await diagnose.visio_diagnose_page(drawing=result["drawing"]))["issues"]
    assert [issue for issue in issues if issue["kind"] == "unglued_connector"] == []
    made.Undo()
    assert page.Shapes.Count == 10


async def test_a_group_member_is_refused_and_the_group_is_not(scratch):
    page = scratch.Pages(1)
    first = page.DrawRectangle(1, 1, 2, 2)
    second = page.DrawRectangle(3, 1, 4, 2)
    selection = page.CreateSelection(0)
    selection.Select(first, 2)
    selection.Select(second, 2)
    group = selection.Group()
    member = first.ID
    reply = json.loads(await edit.visio_delete_shapes([member], drawing=scratch.Name))
    assert reply["code"] == "inside_group"
    assert "rolled_back" not in reply
    reply = json.loads(await edit.visio_delete_shapes([group.ID], drawing=scratch.Name))
    assert "error" not in reply, reply
    assert page.Shapes.Count == 0


async def test_add_container_holds_the_shapes(made):
    result, page = await _flow(made)
    reply = json.loads(await edit.visio_add_container(["Pay", "Reject"], "Outcomes", drawing=result["drawing"]))
    assert "error" not in reply, reply
    assert {result["ids"]["pay"], result["ids"]["reject"]} <= set(reply["contains"])
    container = reply["container"]
    assert (container["text"], container["master"]) == ("Outcomes", "Plain")
    listing = json.loads(await read.visio_get_page(drawing=result["drawing"]))
    row = next(row for row in listing["shapes"] if row["id"] == container["id"])
    assert row["kind"] == "container"
    assert sorted(row["contains"]) == sorted(reply["contains"])
    made.Undo()
    assert page.Shapes.Count == 10


async def test_page_tools_round_trip(made):
    result, _ = await _flow(made)
    drawing = result["drawing"]
    document = made.Documents(drawing)
    added = json.loads(await edit.visio_add_page("Second", drawing=drawing))
    assert added["pages"] == ["Page-1", "Second"]
    changed = json.loads(
        await edit.visio_update_page(page="Second", name="Third", width=11, height=8.5, theme="linear", drawing=drawing)
    )
    assert "error" not in changed, changed
    assert (changed["page"]["name"], changed["page"]["width"], changed["page"]["height"]) == ("Third", 11, 8.5)
    assert changed["page"]["theme"] > 0
    plain = json.loads(await edit.visio_update_page(page="Third", theme="None", drawing=drawing))
    assert plain["page"]["theme"] == 0
    taken = json.loads(await edit.visio_update_page(page="Third", name="page-1", drawing=drawing))
    assert taken["code"] == "page_exists"
    gone = json.loads(await edit.visio_delete_page("Third", drawing=drawing))
    assert gone["pages"] == ["Page-1"]
    last = json.loads(await edit.visio_delete_page(1, drawing=drawing))
    assert last["code"] == "last_page"
    made.Undo()
    assert [document.Pages(i).Name for i in range(1, document.Pages.Count + 1)] == ["Page-1", "Third"]


async def test_undo_and_redo_report_counts(made):
    result, _ = await _flow(made)
    drawing = result["drawing"]
    added = json.loads(
        await edit.visio_add_elements(nodes=[NodeSpec(key="x", master="BASIC:Rectangle", x=1, y=1)], drawing=drawing)
    )
    assert "error" not in added, added

    def row(reply):
        found = next(row for row in reply["drawings"] if row["drawing"] == drawing)
        return found["shapes_before"], found["shapes_after"]

    assert row(json.loads(await edit.visio_undo())) == (11, 10)
    assert row(json.loads(await edit.visio_redo())) == (10, 11)


async def test_edit_tools_take_plain_json(made):
    result, page = await _flow(made)
    server.register_tools()
    async with Client(server.mcp) as client:
        reply = await client.call_tool(
            "visio_add_elements",
            {
                "nodes": [{"key": "audit", "master": "BASFLO:Process", "text": "Audit", "x": 1, "y": 1}],
                "edges": [{"from": "Pay", "to": "audit"}],
                "drawing": result["drawing"],
            },
        )
        added = json.loads(reply.content[0].text)
        assert "error" not in added, added
        reply = await client.call_tool(
            "visio_update_shapes",
            {"updates": [{"shape": added["ids"]["audit"], "text": "Audit trail"}], "drawing": result["drawing"]},
        )
    assert json.loads(reply.content[0].text)["updated"][0]["text"] == "Audit trail"
    assert page.Shapes.Count == 12
