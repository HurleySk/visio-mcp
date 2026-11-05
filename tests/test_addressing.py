import pytest

from tests.fakes import FakeApp, FakeDoc, FakePage, FakeShape
from visio_mcp.addressing import ShapeFinder, find_page, identity, summarize
from visio_mcp.errors import VisioError


def _page(*shapes):
    return FakePage(shapes=shapes)


def test_id_beats_name():
    named_seven = FakeShape(3, "7")
    with_id_seven = FakeShape(7, "Other")
    finder = ShapeFinder(_page(named_seven, with_id_seven))
    assert finder.find(7) is with_id_seven
    assert finder.find("7") is with_id_seven


def test_digit_string_falls_back_to_name_when_no_such_id():
    shape = FakeShape(3, "42")
    assert ShapeFinder(_page(shape)).find("42") is shape


def test_integer_never_falls_back_to_name():
    with pytest.raises(VisioError) as caught:
        ShapeFinder(_page(FakeShape(3, "42"))).find(42)
    assert caught.value.code == "shape_not_found"


def test_name_beats_text():
    by_name = FakeShape(1, "Start")
    by_text = FakeShape(2, "Sheet.2", "Start")
    assert ShapeFinder(_page(by_text, by_name)).find("Start") is by_name


def test_finds_by_exact_text():
    shape = FakeShape(2, "Sheet.2", "Approve?")
    assert ShapeFinder(_page(shape)).find("Approve?") is shape


def test_ambiguous_text_errors():
    finder = ShapeFinder(_page(FakeShape(1, "Sheet.1", "Same", "Process"), FakeShape(2, "Sheet.2", "Same")))
    with pytest.raises(VisioError) as caught:
        finder.find("Same")
    assert caught.value.code == "ambiguous_shape"
    assert caught.value.details == [
        {"id": 1, "name": "Sheet.1", "text": "Same", "master": "Process"},
        {"id": 2, "name": "Sheet.2", "text": "Same", "master": None},
    ]


def test_not_found_suggests_near_matches():
    finder = ShapeFinder(_page(FakeShape(1, "Sheet.1", "Approve request")))
    with pytest.raises(VisioError) as caught:
        finder.find("Approve requests")
    assert caught.value.code == "shape_not_found"
    assert caught.value.details[0]["id"] == 1


@pytest.mark.parametrize("bad", [True, None, "", 1.5])
def test_bad_shape_references_are_invalid(bad):
    with pytest.raises(ValueError):
        ShapeFinder(_page()).find(bad)


def test_identity_reports_no_master_for_a_drawn_shape():
    assert identity(FakeShape(9, "Sheet.9", "x"))["master"] is None


def _document(*pages):
    document = FakeDoc(pages=list(pages))
    document.Application = FakeApp(docs=[document])
    return document


def test_find_page_by_index_and_name():
    first, second = FakePage("One", 1), FakePage("Two", 2)
    document = _document(first, second)
    assert find_page(document, 2) is second
    assert find_page(document, "two") is second
    assert find_page(document) is first


def test_find_page_uses_the_page_on_screen():
    first, second = FakePage("One", 1), FakePage("Two", 2)
    document = _document(first, second)

    class Window:
        Document = document
        Page = second

    document.Application.ActiveWindow = Window()
    assert find_page(document) is second


def test_find_page_ignores_a_window_on_another_drawing():
    first, second = FakePage("One", 1), FakePage("Two", 2)
    document = _document(first, second)

    class Window:
        Document = FakeDoc("other.vsdx")
        Page = second

    document.Application.ActiveWindow = Window()
    assert find_page(document) is first


@pytest.mark.parametrize("missing", [3, 0, "Nope"])
def test_missing_page_lists_the_pages(missing):
    document = _document(FakePage("One", 1), FakePage("Two", 2))
    with pytest.raises(VisioError) as caught:
        find_page(document, missing)
    assert caught.value.code == "page_not_found"
    assert caught.value.details == ["One", "Two"]


def test_summary_reports_no_path_for_an_unsaved_drawing():
    document = FakeDoc("Drawing1", folder="")
    assert summarize(document) == {
        "drawing": "Drawing1",
        "path": None,
        "read_only": False,
        "saved": True,
        "pages": ["Page-1"],
    }
