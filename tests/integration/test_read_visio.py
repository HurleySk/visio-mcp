import io
import json
import os

import pytest
from fastmcp import Client
from PIL import Image

from visio_mcp import server
from visio_mcp.tools import image, read

pytestmark = pytest.mark.visio

VIS_MS_US = 2
VIS_OPEN_RO_HIDDEN = 2 + 64
VIS_SELECT = 2


@pytest.fixture
def flow(own_visio):
    """A small flow made directly through COM: three masters, two glued connectors, one drawn box."""
    from visio_mcp import stencil_index

    document = own_visio.Documents.AddEx("", VIS_MS_US, 0)
    stencil = own_visio.Documents.OpenEx(stencil_index.resolve_master("BASFLO:Process").path, VIS_OPEN_RO_HIDDEN)
    page = document.Pages(1)
    start = page.Drop(stencil.Masters.ItemU("Process"), 2, 8)
    start.Text = "Start"
    decide = page.Drop(stencil.Masters.ItemU("Decision"), 2, 6)
    decide.Text = "Approve?"
    done = page.Drop(stencil.Masters.ItemU("Process"), 5, 6)
    done.Text = "Done"
    stencil.Close()
    start.CellsU("Prop.Owner").FormulaU = '"Finance"'
    connectors = []
    for begin, end, label in ((start, decide, ""), (decide, done, "yes")):
        connector = page.Drop(own_visio.ConnectorToolDataObject, 0, 0)
        connector.CellsU("BeginX").GlueTo(begin.CellsU("PinX"))
        connector.CellsU("EndX").GlueTo(end.CellsU("PinX"))
        connector.Text = label
        connectors.append(connector)
    drawn = page.DrawRectangle(6, 1, 7, 2)
    ids = {
        "start": start.ID,
        "decide": decide.ID,
        "done": done.ID,
        "first": connectors[0].ID,
        "second": connectors[1].ID,
        "drawn": drawn.ID,
    }
    document.Saved = True
    yield document, ids
    document.Saved = True
    document.Close()


async def _call(tool, *args, **kwargs):
    return json.loads(await tool(*args, **kwargs))


async def test_list_open_finds_the_active_drawing(flow):
    document, _ = flow
    result = await _call(read.visio_list_open)
    row = next(row for row in result["drawings"] if row["drawing"] == document.Name)
    assert row["active"] is True
    assert row["path"] is None
    assert not [r for r in result["drawings"] if r["drawing"].upper().endswith(".VSSX")]


async def test_get_page_describes_shapes_and_connectors(flow):
    document, ids = flow
    result = await _call(read.visio_get_page, drawing=document.Name)
    assert result["page"]["width"] == 8.5
    assert result["total"] == 6
    rows = {row["id"]: row for row in result["shapes"]}

    start = rows[ids["start"]]
    assert (start["kind"], start["text"], start["master"]) == ("shape", "Start", "Process")
    assert (start["x"], start["y"]) == (2, 8)
    assert start["w"] > 0 and start["h"] > 0

    second = rows[ids["second"]]
    assert (second["kind"], second["text"], second["master"]) == ("connector", "yes", "Dynamic connector")
    assert (second["from"], second["to"]) == (ids["decide"], ids["done"])

    assert rows[ids["drawn"]]["master"] is None


async def test_get_page_does_not_change_the_drawing(flow):
    document, _ = flow
    await _call(read.visio_get_page, drawing=document.Name)
    assert document.Saved


async def test_get_shape_reads_data_and_connections(flow):
    document, ids = flow
    start = await _call(read.visio_get_shape, ids["start"], ["PinX", "NoSuchCell"], document.Name)
    assert start["data"] == {"Owner": "Finance"}
    assert "Cost" in start["unset_data"]
    assert start["outgoing"] == [ids["decide"]]
    assert start["incoming"] == []
    assert start["cells"]["PinX"]["value"] == 2
    assert start["cells"]["NoSuchCell"] is None

    decide = await _call(read.visio_get_shape, "Approve?", drawing=document.Name)
    assert decide["id"] == ids["decide"]
    assert decide["incoming"] == [ids["start"]]
    assert decide["outgoing"] == [ids["done"]]


async def test_read_does_not_mint_a_guid(flow):
    document, ids = flow
    before = document.Pages(1).Shapes.ItemFromID(ids["start"]).UniqueID(0)
    result = await _call(read.visio_get_shape, ids["start"], drawing=document.Name)
    assert before == ""
    assert "guid" not in result
    assert document.Pages(1).Shapes.ItemFromID(ids["start"]).UniqueID(0) == ""
    assert document.Saved


async def test_get_shape_ambiguous_and_missing(flow):
    document, ids = flow
    document.Pages(1).Shapes.ItemFromID(ids["done"]).Text = "Start"
    ambiguous = await _call(read.visio_get_shape, "Start", drawing=document.Name)
    assert ambiguous["code"] == "ambiguous_shape"
    assert sorted(row["id"] for row in ambiguous["details"]) == sorted([ids["start"], ids["done"]])
    assert (await _call(read.visio_get_shape, 9999, drawing=document.Name))["code"] == "shape_not_found"


async def test_find_shapes_by_each_filter(flow):
    document, ids = flow
    by_text = await _call(read.visio_find_shapes, text="approve", drawing=document.Name)
    assert [row["id"] for row in by_text["shapes"]] == [ids["decide"]]
    by_master = await _call(read.visio_find_shapes, master="BASFLO:Process", drawing=document.Name)
    assert sorted(row["id"] for row in by_master["shapes"]) == sorted([ids["start"], ids["done"]])
    by_data = await _call(read.visio_find_shapes, data={"owner": "finance"}, drawing=document.Name)
    assert [row["id"] for row in by_data["shapes"]] == [ids["start"]]
    assert (await _call(read.visio_find_shapes, data={"owner": "nobody"}, drawing=document.Name))["total"] == 0


async def test_selection_is_read_from_the_window(flow, own_visio):
    document, ids = flow
    window = own_visio.ActiveWindow
    window.DeselectAll()
    assert (await _call(read.visio_get_selection))["shapes"] == []
    window.Select(document.Pages(1).Shapes.ItemFromID(ids["done"]), VIS_SELECT)
    result = await _call(read.visio_get_selection)
    assert result["drawing"] == document.Name
    assert [row["id"] for row in result["shapes"]] == [ids["done"]]


async def test_page_image_is_png(flow, own_visio):
    document, _ = flow
    settings = own_visio.Settings
    before = (settings.GetRasterExportResolution(), settings.GetRasterExportSize())
    metadata, data = (await image.visio_get_page_image(drawing=document.Name))
    metadata_image = data.to_image_content()
    assert metadata["drawing"] == document.Name
    assert metadata_image.mime_type == "image/png"
    assert (settings.GetRasterExportResolution(), settings.GetRasterExportSize()) == before
    assert document.Saved


async def test_page_image_through_the_server_has_text_and_image(flow):
    document, _ = flow
    server.register_tools()
    async with Client(server.mcp) as client:
        result = await client.call_tool("visio_get_page_image", {"drawing": document.Name})
    kinds = [item.type for item in result.content]
    assert kinds == ["text", "image"]
    assert json.loads(result.content[0].text)["drawing"] == document.Name


async def test_max_px_keeps_the_aspect_ratio(flow):
    document, _ = flow
    page = document.Pages(1)
    left, bottom, right, top = page.BoundingBox(image.VIS_BBOX_UPRIGHT_WH)
    metadata, data = await _render(document, 500)
    with Image.open(io.BytesIO(data)) as picture:
        assert max(picture.size) == pytest.approx(500, abs=6)
        assert picture.size[0] / picture.size[1] == pytest.approx((right - left) / (top - bottom), abs=0.02)
    assert metadata["width_px"] == picture.size[0]


async def _render(document, max_px, annotate_ids=False):
    from visio_mcp.com_runtime import run_com

    # COM references belong to their thread, so the worker looks the drawing up itself.
    name = document.Name

    def render(session):
        found = session.app.Documents(name)
        return image.render_page(found, found.Pages(1), max_px, annotate_ids)

    return await run_com(render)


async def test_annotated_image_marks_the_shapes(flow, tmp_path):
    document, _ = flow
    _, plain = await _render(document, 800)
    metadata, marked = await _render(document, 800, annotate_ids=True)

    def badge_pixels(picture):
        colours = picture.convert("RGB").getcolors(maxcolors=1 << 24)
        return sum(count for count, colour in colours if colour == image.BADGE_FILL)

    with Image.open(io.BytesIO(plain)) as a, Image.open(io.BytesIO(marked)) as b:
        margin = 2 * image.BADGE_MARGIN
        assert b.size == (a.size[0] + margin, a.size[1] + margin)
        assert (metadata["width_px"], metadata["height_px"]) == b.size
        assert badge_pixels(a) == 0
        assert badge_pixels(b) > 100
    out = os.environ.get("VISIO_MCP_KEEP_IMAGE")
    if out:
        with open(out, "wb") as kept:
            kept.write(marked)


async def test_blank_page_image(own_visio):
    document = own_visio.Documents.AddEx("", VIS_MS_US, 0)
    try:
        metadata, _ = await image.visio_get_page_image(drawing=document.Name)
        assert metadata["shapes"] == 0
    finally:
        document.Saved = True
        document.Close()
