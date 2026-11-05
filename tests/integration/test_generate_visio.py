import json
import os

import pytest
from fastmcp import Client

from visio_mcp import building, server, stencil_index
from visio_mcp.models import DiagramSpec
from visio_mcp.tools import diagnose, generate

pytestmark = pytest.mark.visio

VIS_MS_US = 2
VIS_OPEN_RO_HIDDEN = 2 + 64
VIS_ADD_STENCIL = 512
VIS_TYPE_STENCIL = 2

FLOW = {
    "nodes": [
        {"key": "start", "master": "BASFLO:Start/End", "text": "Start"},
        {"key": "check", "master": "BASFLO:Decision", "text": "Approved?"},
        {"key": "pay", "master": "BASFLO:Process", "text": "Pay"},
        {"key": "reject", "master": "BASFLO:Process", "text": "Reject"},
        {"key": "end", "master": "BASFLO:Start/End", "text": "End"},
    ],
    "edges": [
        {"from": "start", "to": "check"},
        {"from": "check", "to": "pay", "text": "yes"},
        {"from": "check", "to": "reject", "text": "no"},
        {"from": "pay", "to": "end"},
        {"from": "reject", "to": "end"},
    ],
}


@pytest.fixture
def made(own_visio):
    """Closes whatever drawings a test makes, and reports any left behind hidden."""
    before = {own_visio.Documents(i).Name for i in range(1, own_visio.Documents.Count + 1)}
    yield own_visio
    for index in range(own_visio.Documents.Count, 0, -1):
        document = own_visio.Documents(index)
        if document.Name not in before and document.Type != VIS_TYPE_STENCIL:
            document.Saved = True
            document.Close()


async def _create(**fields):
    fields.setdefault("return_image", False)
    return json.loads(await generate.visio_create_diagram(DiagramSpec.model_validate(fields)))


def _documents(app):
    return {app.Documents(i).Name for i in range(1, app.Documents.Count + 1)}


async def test_create_diagram_glues_connectors(made):
    result = await _create(**FLOW)
    assert "error" not in result, result
    page = made.Documents(result["drawing"]).Pages(result["page"])
    ids = result["ids"]
    assert set(ids) == {"start", "check", "pay", "reject", "end"}
    assert len(result["connectors"]) == 5
    for row in result["connectors"]:
        connector = page.Shapes.ItemFromID(row["id"])
        joined = {connector.Connects(i).FromCell.Name: connector.Connects(i).ToSheet.ID for i in (1, 2)}
        assert joined == {"BeginX": ids[row["from"]], "EndX": ids[row["to"]]}
        assert connector.CellsU("EndArrow").FormulaU == "13"
    assert page.Shapes.ItemFromID(ids["check"]).Text == "Approved?"


async def test_moving_a_shape_drags_its_connector(made):
    result = await _create(**FLOW)
    page = made.Documents(result["drawing"]).Pages(result["page"])
    connector = page.Shapes.ItemFromID(result["connectors"][0]["id"])
    before = connector.BoundingBox(1)
    shape = page.Shapes.ItemFromID(result["ids"]["start"])
    shape.CellsU("PinX").ResultIU += 3
    assert connector.BoundingBox(1) != before


async def test_create_is_one_undo_step(made):
    result = await _create(**FLOW)
    page = made.Documents(result["drawing"]).Pages(result["page"])
    assert page.Shapes.Count == 10
    made.Undo()
    assert page.Shapes.Count == 0
    made.Redo()
    assert page.Shapes.Count == 10


async def test_the_flow_lays_out_with_no_issues(made):
    result = await _create(**FLOW)
    assert result["issues"] == []


async def test_a_diagram_comes_back_with_a_picture(made):
    server.register_tools()
    async with Client(server.mcp) as client:
        reply = await client.call_tool("visio_create_diagram", {"spec": FLOW})
    assert [item.type for item in reply.content] == ["text", "image"]
    assert json.loads(reply.content[0].text)["issues"] == []
    keep = os.environ.get("VISIO_MCP_KEEP_IMAGE")
    if keep:
        import base64

        with open(keep, "wb") as kept:
            kept.write(base64.b64decode(reply.content[1].data))


async def test_hidden_stencils_are_closed_afterwards(made):
    before = _documents(made)
    result = await _create(**FLOW)
    assert _documents(made) - before == {result["drawing"]}


async def test_a_template_keeps_its_own_stencils_open(made):
    result = await _create(template="BASFLO", units="us", nodes=[{"key": "a", "master": "Process"}])
    assert "error" not in result, result
    stencils = {name for name in _documents(made) if name.upper() == "BASFLO_U.VSSX"}
    assert stencils


async def test_an_open_stencil_is_reused_and_left_open(made):
    stencil = made.Documents.OpenEx(stencil_index.resolve_master("BASFLO:Process").path, VIS_OPEN_RO_HIDDEN)
    try:
        await _create(nodes=[{"key": "a", "master": "BASFLO:Process"}])
        assert stencil.Name in _documents(made)
    finally:
        stencil.Close()


async def test_a_custom_open_stencil_works(made, tmp_path):
    stencil = made.Documents.AddEx("", VIS_MS_US, VIS_ADD_STENCIL)
    source = made.Documents.AddEx("", VIS_MS_US, 0)
    stencil.Masters.Drop(source.Pages(1).DrawRectangle(1, 1, 2, 2), 0, 0).Name = "Widget"
    source.Saved = True
    source.Close()
    path = str(tmp_path / "Mine.vssx")
    stencil.SaveAs(path)
    try:
        result = await _create(nodes=[{"key": "a", "master": "Mine:Widget"}])
        assert "error" not in result, result
        page = made.Documents(result["drawing"]).Pages(result["page"])
        assert page.Shapes.ItemFromID(result["ids"]["a"]).Master.NameU == "Widget"
        assert "Mine.vssx" in _documents(made)
    finally:
        stencil.Close()


async def test_data_goes_into_rows_old_and_new(made):
    result = await _create(
        nodes=[
            {
                "key": "a",
                "master": "BASFLO:Process",
                "data": {"Owner": 'She said "hi"', "Cost Centre": "Finance", "Cost_Centre": "Other", "2024 budget": "12"},
            },
            {"key": "b", "master": "BASIC:Rectangle", "data": {"Kind": "plain"}},
        ]
    )
    assert "error" not in result, result
    page = made.Documents(result["drawing"]).Pages(result["page"])
    from visio_mcp.tools.read import shape_data

    rows = {label: value for label, _, value in shape_data(page.Shapes.ItemFromID(result["ids"]["a"]))}
    assert rows["Owner"] == 'She said "hi"'
    assert rows["Cost Centre"] == "Finance"
    assert rows["Cost_Centre"] == "Other"
    assert rows["2024 budget"] == "12"
    plain = {label: value for label, _, value in shape_data(page.Shapes.ItemFromID(result["ids"]["b"]))}
    assert plain == {"Kind": "plain"}


async def test_size_and_fill(made):
    result = await _create(nodes=[{"key": "a", "master": "BASFLO:Process", "w": 3, "h": 0.5, "fill": "#C8E6FF"}])
    shape = made.Documents(result["drawing"]).Pages(1).Shapes.ItemFromID(result["ids"]["a"])
    assert shape.CellsU("Width").ResultIU == pytest.approx(3)
    assert shape.CellsU("Height").ResultIU == pytest.approx(0.5)
    assert shape.CellsU("FillForegnd").FormulaU == "THEMEGUARD(RGB(200,230,255))"


async def test_positions_are_kept_and_the_page_is_left_alone(made):
    result = await _create(
        nodes=[
            {"key": "a", "master": "BASFLO:Process", "x": 2, "y": 3},
            {"key": "b", "master": "BASFLO:Process", "x": 5, "y": 7},
        ],
        edges=[{"from": "a", "to": "b"}],
    )
    page = made.Documents(result["drawing"]).Pages(1)
    shape = page.Shapes.ItemFromID(result["ids"]["b"])
    assert (shape.CellsU("PinX").ResultIU, shape.CellsU("PinY").ResultIU) == (5, 7)
    assert page.PageSheet.CellsU("PageWidth").ResultIU == 8.5


async def test_an_unconnected_node_is_placed_clear_of_the_rest(made):
    result = await _create(
        nodes=[
            {"key": "a", "master": "BASIC:Rectangle", "text": "a"},
            {"key": "b", "master": "BASIC:Rectangle", "text": "b"},
            {"key": "c", "master": "BASIC:Rectangle", "text": "c"},
            {"key": "alone", "master": "BASIC:Rectangle", "text": "alone"},
        ],
        edges=[{"from": "a", "to": "b"}, {"from": "b", "to": "c"}],
    )
    assert result["issues"] == []


async def test_siblings_keep_the_order_given(made):
    nodes = [{"key": key, "master": "BASFLO:Process", "text": key} for key in ("root", "one", "two", "three")]
    edges = [{"from": "root", "to": key} for key in ("one", "two", "three")]

    down = await _create(nodes=nodes, edges=edges, layout={"style": "top_down"})
    page = made.Documents(down["drawing"]).Pages(1)
    xs = [page.Shapes.ItemFromID(down["ids"][key]).CellsU("PinX").ResultIU for key in ("one", "two", "three")]
    assert xs == sorted(xs)

    across = await _create(nodes=nodes, edges=edges, layout={"style": "left_right"})
    page = made.Documents(across["drawing"]).Pages(1)
    ys = [page.Shapes.ItemFromID(across["ids"][key]).CellsU("PinY").ResultIU for key in ("one", "two", "three")]
    assert ys == sorted(ys, reverse=True)


async def test_spacing_spreads_shapes_out(made):
    async def width(spacing):
        result = await _create(**FLOW, layout={"style": "top_down", "spacing": spacing})
        return made.Documents(result["drawing"]).Pages(1).PageSheet.CellsU("PageHeight").ResultIU

    assert await width(1.5) > await width(0.25)


async def test_save_as_saves_the_finished_drawing(made, tmp_path):
    path = tmp_path / "flow.vsdx"
    result = await _create(**FLOW, save_as=str(path))
    assert path.exists()
    assert result["drawing"] == "flow.vsdx"
    assert result["path"] == str(path)
    document = made.Documents("flow.vsdx")
    assert document.Saved
    assert document.Pages(1).Shapes.Count == 10


async def test_a_failed_build_leaves_nothing_behind(made, monkeypatch):
    before = _documents(made)

    def fail(*args, **kwargs):
        raise RuntimeError("layout went wrong")

    monkeypatch.setattr(generate, "lay_out", fail)
    result = await _create(**FLOW)
    assert result["rolled_back"] is True
    assert "layout went wrong" in result["error"]
    assert [warning["code"] for warning in result["warnings"]] == ["new_drawing_closed"]
    assert _documents(made) == before
    assert made.CurrentScope == -1


async def test_a_bad_bare_name_with_a_template_closes_the_new_drawing(made):
    before = _documents(made)
    result = await _create(template="BASFLO", units="us", nodes=[{"key": "a", "master": "No Such Master Anywhere"}])
    assert result["code"] == "unknown_master"
    assert [warning["code"] for warning in result["warnings"]] == ["new_drawing_closed"]
    assert not [name for name in _documents(made) - before if not name.upper().endswith(".VSSX")]


async def _flow_page(app):
    result = await _create(**FLOW)
    return result, app.Documents(result["drawing"]).Pages(1)


def _position(page, shape_id):
    shape = page.Shapes.ItemFromID(shape_id)
    return shape.CellsU("PinX").ResultIU, shape.CellsU("PinY").ResultIU


async def test_layout_page_lays_out_again_and_undoes_as_one(made):
    result, page = await _flow_page(made)
    ids = list(result["ids"].values())
    before = [_position(page, shape_id) for shape_id in ids]
    reply = json.loads(await generate.visio_layout_page("left_right", drawing=result["drawing"]))
    assert "error" not in reply, reply
    assert reply["issues"] == []
    assert reply["page"]["shapes"] == 10
    after = [_position(page, shape_id) for shape_id in ids]
    assert after != before
    made.Undo()
    assert [_position(page, shape_id) for shape_id in ids] == before


async def test_layout_page_on_some_shapes_leaves_the_rest(made):
    result, page = await _flow_page(made)
    keep = result["ids"]["end"]
    moving = [result["ids"]["pay"], result["ids"]["reject"]]
    before = _position(page, keep)
    width = page.PageSheet.CellsU("PageWidth").ResultIU
    reply = json.loads(await generate.visio_layout_page("left_right", shapes=moving, drawing=result["drawing"]))
    assert "error" not in reply, reply
    assert _position(page, keep) == before
    assert page.PageSheet.CellsU("PageWidth").ResultIU == width


async def test_layout_page_reports_an_unknown_shape_without_a_scope(made):
    result, _ = await _flow_page(made)
    reply = json.loads(await generate.visio_layout_page(shapes=[9999], drawing=result["drawing"]))
    assert reply["code"] == "shape_not_found"
    assert "rolled_back" not in reply


async def test_diagnose_finds_real_problems(scratch, own_visio):
    stencil = own_visio.Documents.OpenEx(stencil_index.resolve_master("BASFLO:Process").path, VIS_OPEN_RO_HIDDEN)
    page = scratch.Pages(1)
    master = stencil.Masters.ItemU("Process")
    first = page.Drop(master, 2, 8)
    second = page.Drop(master, 2.2, 8)
    outside = page.Drop(master, 20, 8)
    loose = page.Drop(own_visio.ConnectorToolDataObject, 4, 4)
    loose.CellsU("BeginX").GlueTo(first.CellsU("PinX"))
    drawn = page.DrawLine(5, 1, 6, 1)
    stencil.Close()
    reply = json.loads(await diagnose.visio_diagnose_page(drawing=scratch.Name))
    kinds = {issue["kind"] for issue in reply["issues"]}
    assert kinds == {"overlap", "off_page", "unglued_connector"}
    overlap = next(issue for issue in reply["issues"] if issue["kind"] == "overlap")
    assert sorted(row["id"] for row in overlap["shapes"]) == sorted([first.ID, second.ID])
    assert next(issue for issue in reply["issues"] if issue["kind"] == "off_page")["id"] == outside.ID
    unglued = next(issue for issue in reply["issues"] if issue["kind"] == "unglued_connector")
    assert (unglued["id"], unglued["from"], unglued["to"]) == (loose.ID, first.ID, None)
    assert drawn.ID not in [issue.get("id") for issue in reply["issues"]]


async def test_diagnose_ignores_a_container_and_its_members(scratch, own_visio):
    flow = own_visio.Documents.OpenEx(stencil_index.resolve_master("BASFLO:Process").path, VIS_OPEN_RO_HIDDEN)
    containers = own_visio.Documents.OpenEx(stencil_index.resolve_master("SDCONT:Plain").path, VIS_OPEN_RO_HIDDEN)
    page = scratch.Pages(1)
    inside = page.Drop(flow.Masters.ItemU("Process"), 3, 6)
    selection = page.CreateSelection(0)
    selection.Select(inside, 2)
    page.DropContainer(containers.Masters.ItemU("Plain"), selection)
    flow.Close()
    containers.Close()
    reply = json.loads(await diagnose.visio_diagnose_page(drawing=scratch.Name))
    assert reply["issues"] == []
