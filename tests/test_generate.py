import json

import pytest
from fastmcp import Client
from pydantic import ValidationError

from tests.fakes import FakeDoc
from visio_mcp import server
from visio_mcp.building import MasterResolver
from visio_mcp.models import DiagramSpec, EdgeSpec, LayoutSpec, NodeSpec, check_graph, resolve_layout
from visio_mcp.tools import generate


def _spec(**fields):
    fields.setdefault(
        "nodes",
        [{"key": "a", "master": "BASFLO:Process"}, {"key": "b", "master": "BASFLO:Decision"}],
    )
    return DiagramSpec.model_validate(fields)


async def _create(spec):
    return json.loads(await generate.visio_create_diagram(spec))


def _nothing_happened(connector):
    app = connector.apps[0]
    assert app.Documents.added == []
    assert app.scopes_begun == []


async def test_bad_edge_key_fails_before_mutation(connector, content):
    result = await _create(_spec(edges=[{"from": "a", "to": "bb"}]))
    assert result["code"] == "unknown_node"
    assert result["details"] == ["b"]
    _nothing_happened(connector)


async def test_unknown_master_fails_before_mutation(connector, content):
    result = await _create(_spec(nodes=[{"key": "a", "master": "BASFLO:Procss"}]))
    assert result["code"] == "unknown_master"
    assert result["details"] == ["BASFLO:Process"]
    _nothing_happened(connector)


async def test_ambiguous_master_fails_before_mutation(connector, content):
    result = await _create(_spec(nodes=[{"key": "a", "master": "Process"}]))
    assert result["code"] == "ambiguous_master"
    assert sorted(result["details"]) == ["BASFLO:Process", "XFUNC:Process"]
    _nothing_happened(connector)


async def test_stencils_settle_an_ambiguous_master(connector, content):
    # It gets as far as building, which the fake can't do, so the new drawing is created.
    await _create(_spec(nodes=[{"key": "a", "master": "Process"}], stencils=["XFUNC"], return_image=False))
    assert len(connector.apps[0].Documents.added) == 1


async def test_duplicate_keys_are_refused(connector, content):
    result = await _create(_spec(nodes=[{"key": "a", "master": "BASFLO:Process"}, {"key": "a", "master": "BASFLO:Process"}]))
    assert result["code"] == "invalid_argument"
    _nothing_happened(connector)


@pytest.mark.parametrize("fill", ["red", "#12345", "#GGGGGG", "123456"])
async def test_a_bad_fill_is_refused(connector, content, fill):
    result = await _create(_spec(nodes=[{"key": "a", "master": "BASFLO:Process", "fill": fill}]))
    assert result["code"] == "invalid_argument"
    _nothing_happened(connector)


async def test_an_existing_save_target_is_refused_before_mutation(connector, content):
    existing = content / "there.vsdx"
    existing.write_bytes(b"precious")
    result = await _create(_spec(save_as=str(existing)))
    assert result["code"] == "file_exists"
    assert existing.read_bytes() == b"precious"
    _nothing_happened(connector)


async def test_a_missing_save_folder_is_refused_before_mutation(connector, content):
    result = await _create(_spec(save_as=str(content / "nowhere" / "new.vsdx")))
    assert result["code"] == "folder_not_found"
    _nothing_happened(connector)


async def test_a_template_defers_bare_names_until_it_is_open(connector, content):
    await _create(_spec(nodes=[{"key": "a", "master": "Process"}], template="BASFLO"))
    assert len(connector.apps[0].Documents.added) == 1


async def test_a_template_still_checks_names_that_say_their_stencil(connector, content):
    result = await _create(_spec(nodes=[{"key": "a", "master": "BASFLO:Nope"}], template="BASFLO"))
    assert result["code"] == "unknown_master"
    _nothing_happened(connector)


async def test_a_failed_build_rolls_back_and_closes_the_new_drawing(connector, content):
    app = connector.apps[0]
    before = list(app.docs)
    # The fake can't drop shapes, which is the failure part-way through a build.
    result = await _create(_spec())
    assert result["rolled_back"] is True
    assert [row[1] for row in app.scopes_ended] == [False]
    assert app.docs == before
    assert [warning["code"] for warning in result["warnings"]] == ["new_drawing_closed"]


def test_unknown_fields_are_refused():
    with pytest.raises(ValidationError):
        NodeSpec.model_validate({"key": "a", "master": "m", "label": "typo"})


def test_numbers_are_accepted_as_keys_and_data():
    node = NodeSpec.model_validate({"key": 1, "master": "m", "data": {"Cost": 5}})
    assert node.key == "1"
    assert node.data == {"Cost": "5"}


def test_from_is_the_edges_field_name():
    edge = EdgeSpec.model_validate({"from": "a", "to": "b"})
    assert edge.from_ == "a"
    assert edge.arrow == "end"


def _nodes(*rows):
    return [NodeSpec(key=str(index), master="m", **row) for index, row in enumerate(rows)]


def test_layout_defaults_to_top_down_without_positions():
    assert resolve_layout(None, _nodes({}, {})).style == "top_down"


def test_layout_defaults_to_none_with_positions():
    assert resolve_layout(None, _nodes({"x": 1, "y": 2})).style == "none"


def test_positions_and_a_layout_conflict():
    with pytest.raises(ValueError, match="would move"):
        resolve_layout(LayoutSpec(style="left_right"), _nodes({"x": 1, "y": 2}))


def test_style_none_needs_every_position():
    with pytest.raises(ValueError, match="missing"):
        resolve_layout(LayoutSpec(style="none"), _nodes({"x": 1, "y": 2}, {}))
    with pytest.raises(ValueError, match="missing"):
        resolve_layout(None, _nodes({"x": 1, "y": 2}, {}))


def test_x_without_y_is_refused():
    with pytest.raises(ValueError, match="go together"):
        resolve_layout(None, _nodes({"x": 1}))


def test_style_none_refuses_route_and_spacing():
    with pytest.raises(ValueError, match="only apply"):
        resolve_layout(LayoutSpec(style="none", route="tree"), _nodes({"x": 1, "y": 2}))


def test_spacing_must_be_positive():
    with pytest.raises(ValueError, match="spacing"):
        resolve_layout(LayoutSpec(spacing=0), _nodes({}))


def test_a_self_loop_is_allowed():
    check_graph(_nodes({}), [EdgeSpec.model_validate({"from": "0", "to": "0"})])


class _Master:
    def __init__(self, name):
        self.NameU = name
        self.Name = name


class _Masters:
    def __init__(self, *names):
        self._rows = [_Master(name) for name in names]

    @property
    def Count(self):
        return len(self._rows)

    def __call__(self, index):
        return self._rows[index - 1]


def _open_stencil(app, name, *masters):
    stencil = FakeDoc(name, folder="C:\\mine", type_=2)
    stencil.Masters = _Masters(*masters)
    stencil.Application = app
    app.docs.append(stencil)
    return stencil


def test_resolver_finds_a_custom_open_stencil(connector, content):
    app = connector.apps[0]
    _open_stencil(app, "My Shapes.vssx", "Widget")
    ref = MasterResolver(app).resolve("My Shapes:widget")
    assert (ref.stem, ref.name, ref.path) == ("MY SHAPES", "Widget", "C:\\mine\\My Shapes.vssx")


def test_resolver_says_what_a_custom_stencil_lacks(connector, content):
    app = connector.apps[0]
    _open_stencil(app, "Mine.vssx", "Widget")
    with pytest.raises(Exception, match="no master called 'Wodget'") as caught:
        MasterResolver(app).resolve("Mine:Wodget")
    assert caught.value.details == ["MINE:Widget"]


def test_resolver_prefers_a_bare_name_in_an_open_stencil(connector, content):
    app = connector.apps[0]
    _open_stencil(app, "BASFLO_U.vssx", "Process", "Decision")
    ref = MasterResolver(app).resolve("Process")
    assert ref.label == "BASFLO:Process"
    assert ref.path.upper().endswith("BASFLO_U.VSSX")


def test_resolver_uses_the_drawings_units_for_an_open_stencil(connector, content):
    app = connector.apps[0]
    _open_stencil(app, "BASFLO_U.vssx", "Process")
    ref = MasterResolver(app, units="metric").resolve("Process")
    assert ref.path.upper().endswith("BASFLO_M.VSSX")


def test_resolver_prefers_a_master_already_in_the_drawing(connector, content):
    app = connector.apps[0]
    document = app.docs[0]
    document.Masters = _Masters("Process")
    ref = MasterResolver(app, document).resolve("Process")
    assert ref.path is None
    assert ref.name == "Process"


def test_named_stencils_outrank_the_drawings_masters(connector, content):
    app = connector.apps[0]
    document = app.docs[0]
    document.Masters = _Masters("Process")
    ref = MasterResolver(app, document, within=["XFUNC"]).resolve("Process")
    assert ref.label == "XFUNC:Process"


def test_two_open_stencils_with_the_name_are_ambiguous(connector, content):
    app = connector.apps[0]
    _open_stencil(app, "One.vssx", "Thing")
    _open_stencil(app, "Two.vssx", "Thing")
    with pytest.raises(Exception, match="2 open stencils") as caught:
        MasterResolver(app).resolve("Thing")
    assert caught.value.code == "ambiguous_master"


async def test_the_spec_schema_reaches_clients_with_from_and_no_extras():
    server.register_tools()
    async with Client(server.mcp) as client:
        tools = {tool.name: tool for tool in await client.list_tools()}
    schema = tools["visio_create_diagram"].input_schema
    assert schema["required"] == ["spec"]
    edge = schema["properties"]["spec"]["properties"]["edges"]["items"]
    assert "from" in edge["properties"] and "from_" not in edge["properties"]
    assert edge["additionalProperties"] is False


async def test_a_spec_sent_through_the_server_is_validated(connector, content):
    server.register_tools()
    async with Client(server.mcp) as client:
        result = await client.call_tool(
            "visio_create_diagram",
            {"spec": {"nodes": [{"key": "a", "master": "BASFLO:Process"}], "edges": [{"from": "a", "to": "z"}]}},
        )
    assert json.loads(result.content[0].text)["code"] == "unknown_node"
