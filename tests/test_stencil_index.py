import json
import zipfile

import pytest

from visio_mcp import stencil_index
from visio_mcp.errors import VisioError
from visio_mcp.tools import masters

NS = "http://schemas.microsoft.com/office/visio/2012/main"


def _core(title, keywords=""):
    return (
        '﻿<?xml version="1.0" encoding="UTF-8"?>'
        '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/">'
        f"<dc:title>{title}</dc:title><cp:keywords>{keywords}</cp:keywords>"
        "<dc:description>About it.</dc:description></cp:coreProperties>"
    ).encode("utf-8")


def _masters(*names, hidden=()):
    rows = [f"<Master ID='{i}' NameU='{n}' Name='{n}' Prompt='Drag {n}.' Hidden='0'><PageSheet/></Master>" for i, n in enumerate(names)]
    rows += [f"<Master ID='9{i}' NameU='{n}' Name='{n}' Hidden='1'/>" for i, n in enumerate(hidden)]
    return f"<?xml version='1.0'?><Masters xmlns='{NS}'>{''.join(rows)}</Masters>".encode("utf-8")


def _stencil(folder, file_name, title, keywords, *names, hidden=()):
    with zipfile.ZipFile(folder / file_name, "w") as archive:
        archive.writestr("docProps/core.xml", _core(title, keywords))
        archive.writestr("visio/masters/masters.xml", _masters(*names, hidden=hidden))


@pytest.fixture
def content(tmp_path, monkeypatch):
    monkeypatch.setenv(stencil_index.CONTENT_ENV, str(tmp_path))
    stencil_index.reset()
    # Real stencils list their masters' words as keywords, so every master matches on them.
    _stencil(tmp_path, "BASFLO_U.VSSX", "Basic Flowchart Shapes", "flowchart,process,decision", "Process", "Decision", hidden=("Secret",))
    _stencil(tmp_path, "BASFLO_M.VSSX", "Basic Flowchart Shapes", "flowchart,process,decision", "Process", "Decision")
    _stencil(tmp_path, "XFUNC_U.VSSX", "Cross-functional", "swimlane", "Process", "Swimlane")
    _stencil(tmp_path, "ONLYM_M.VSSX", "Metric only", "", "Metric Thing")
    with zipfile.ZipFile(tmp_path / "BASFLO_U.VSTX", "w") as archive:
        archive.writestr("docProps/core.xml", _core("Basic Flowchart", "flow"))
    with zipfile.ZipFile(tmp_path / "BASFLO_M.VSTX", "w") as archive:
        archive.writestr("docProps/core.xml", _core("Basic Flowchart", "flow"))
    (tmp_path / "RIGHTLEFT_R2L.VSTX").write_bytes(b"ignored by the name filter")
    (tmp_path / "BROKEN_U.VSSX").write_bytes(b"not a zip")
    yield tmp_path
    stencil_index.reset()


def test_parses_fixture_vssx(content):
    library = stencil_index.stencils()["BASFLO"]
    assert library.title == "Basic Flowchart Shapes"
    assert [m.name for m in library.masters] == ["Process", "Decision"]
    assert library.masters[0].prompt == "Drag Process."


def test_a_broken_stencil_file_is_skipped(content):
    assert stencil_index.stencils()["BROKEN"].masters == []


def test_ambiguous_master_lists_candidates(content):
    with pytest.raises(VisioError) as caught:
        stencil_index.resolve_master("Process")
    assert caught.value.code == "ambiguous_master"
    assert caught.value.details == ["BASFLO:Process", "XFUNC:Process"]


def test_within_settles_an_ambiguous_master(content):
    assert stencil_index.resolve_master("Process", within=["XFUNC"]).label == "XFUNC:Process"


def test_a_unique_bare_name_resolves(content):
    assert stencil_index.resolve_master("decision").label == "BASFLO:Decision"


def test_units_pick_u_or_m(content):
    assert stencil_index.resolve_master("BASFLO:Decision", "us").path.endswith("BASFLO_U.VSSX")
    assert stencil_index.resolve_master("BASFLO:Decision", "metric").path.endswith("BASFLO_M.VSSX")


def test_units_fall_back_when_a_stencil_has_only_one_kind(content):
    assert stencil_index.resolve_master("XFUNC:Swimlane", "metric").path.endswith("XFUNC_U.VSSX")
    assert stencil_index.resolve_master("ONLYM:Metric Thing", "us").path.endswith("ONLYM_M.VSSX")


def test_a_file_suffix_in_the_stencil_name_is_forgiven(content):
    assert stencil_index.resolve_master("basflo_u:Decision", "metric").path.endswith("BASFLO_M.VSSX")


def test_unknown_stencil_suggests_near_names(content):
    with pytest.raises(VisioError) as caught:
        stencil_index.resolve_master("BASFLOW:Decision")
    assert caught.value.code == "unknown_stencil"
    assert "BASFLO" in caught.value.details


def test_unknown_master_suggests_near_names(content):
    with pytest.raises(VisioError) as caught:
        stencil_index.resolve_master("BASFLO:Decisoin")
    assert caught.value.code == "unknown_master"
    assert caught.value.details == ["BASFLO:Decision"]


def test_hidden_masters_are_left_out(content):
    with pytest.raises(VisioError):
        stencil_index.resolve_master("BASFLO:Secret")


def test_search_finds_names_before_stencil_words(content):
    found = stencil_index.find_masters("flowchart decision")
    assert [row["master"] for row in found["results"]] == ["BASFLO:Decision"]
    assert found["results"][0]["stencil"] == "Basic Flowchart Shapes"


def test_search_ranks_the_common_stencil_first(content):
    found = stencil_index.find_masters("process")
    assert [row["master"] for row in found["results"]] == ["BASFLO:Process", "XFUNC:Process"]


def test_search_limits_results_but_reports_the_total(content):
    found = stencil_index.find_masters("process", limit=1)
    assert found["total"] == 2
    assert len(found["results"]) == 1


def test_search_can_list_one_stencil(content):
    found = stencil_index.find_masters(stencil="XFUNC")
    assert [row["master"] for row in found["results"]] == ["XFUNC:Process", "XFUNC:Swimlane"]


def test_templates_are_listed_and_resolved(content):
    assert stencil_index.list_templates("flow") == [{"template": "BASFLO", "title": "Basic Flowchart", "description": "About it."}]
    assert stencil_index.list_templates("nothing like this") == []
    assert stencil_index.resolve_template("BASFLO", "metric").endswith("BASFLO_M.VSTX")
    assert stencil_index.resolve_template("basflo_u.vstx", "us").endswith("BASFLO_U.VSTX")


def test_unknown_template_suggests_near_names(content):
    with pytest.raises(VisioError) as caught:
        stencil_index.resolve_template("BASFLOW")
    assert caught.value.code == "unknown_template"
    assert caught.value.details == ["BASFLO"]


async def test_tools_answer_without_visio(content):
    found = json.loads(await masters.visio_find_masters("decision"))
    assert found["results"][0]["master"] == "BASFLO:Decision"
    assert json.loads(await masters.visio_list_templates())["templates"][0]["template"] == "BASFLO"


async def test_tools_return_errors_as_json(content):
    assert json.loads(await masters.visio_find_masters(""))["code"] == "invalid_argument"
    assert json.loads(await masters.visio_find_masters("x", stencil="NOPE"))["code"] == "unknown_stencil"
