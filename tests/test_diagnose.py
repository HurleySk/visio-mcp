import pytest

from tests.fakes import FakeCell, FakePage, FakeShape
from visio_mcp.tools import diagnose


class BoxShape(FakeShape):
    def __init__(self, id_, box, master=None, containers=()):
        super().__init__(id_, f"Sheet.{id_}", master=master)
        self.box = box
        self.MemberOfContainers = containers

    def BoundingBox(self, flags):
        assert flags == diagnose.VIS_BBOX_UPRIGHT_WH
        return self.box


class Connect:
    def __init__(self, cell, target):
        self.FromCell = type("Cell", (), {"Name": cell})
        self.ToSheet = type("Sheet", (), {"ID": target})


class Connects:
    def __init__(self, rows):
        self._rows = rows
        self.Count = len(rows)

    def __call__(self, index):
        return self._rows[index - 1]


class Line(FakeShape):
    def __init__(self, id_, master, ends=()):
        super().__init__(id_, f"Line.{id_}", master=master)
        self.OneD = 1
        self.Connects = Connects([Connect(cell, target) for cell, target in ends])


class Sheet:
    def CellsU(self, name):
        return FakeCell({"PageWidth": 8.5, "PageHeight": 11.0}[name])


def _page(*shapes):
    page = FakePage(shapes=shapes)
    page.PageSheet = Sheet()
    return page


def _kinds(page):
    return [issue["kind"] for issue in diagnose.page_issues(page)]


def test_a_clean_page_has_no_issues():
    assert _kinds(_page(BoxShape(1, (1, 1, 2, 2)), BoxShape(2, (3, 3, 4, 4)))) == []


def test_overlap_names_both_shapes():
    issues = diagnose.page_issues(_page(BoxShape(1, (1, 1, 3, 3)), BoxShape(2, (2, 2, 4, 4))))
    assert [issue["kind"] for issue in issues] == ["overlap"]
    assert sorted(row["id"] for row in issues[0]["shapes"]) == [1, 2]


@pytest.mark.parametrize(
    "second",
    [(2, 1, 3, 2), (1, 2, 2, 3), (2.005, 1, 3, 2)],
    ids=["touching side", "touching top", "within tolerance"],
)
def test_touching_shapes_do_not_overlap(second):
    assert _kinds(_page(BoxShape(1, (1, 1, 2, 2)), BoxShape(2, second))) == []


def test_side_by_side_at_different_heights_do_not_overlap():
    assert _kinds(_page(BoxShape(1, (1, 1, 3, 2)), BoxShape(2, (2, 5, 4, 6)))) == []


def test_overlap_is_found_past_a_shape_that_does_not_overlap():
    page = _page(BoxShape(1, (0, 0, 5, 1)), BoxShape(2, (1, 3, 2, 4)), BoxShape(3, (4, 0.5, 6, 2)))
    issues = diagnose.page_issues(page)
    assert [sorted(row["id"] for row in issue["shapes"]) for issue in issues] == [[1, 3]]


def test_a_container_does_not_overlap_its_members():
    container = BoxShape(1, (0, 0, 5, 5))
    member = BoxShape(2, (1, 1, 2, 2), containers=(1,))
    assert _kinds(_page(container, member)) == []


def test_two_members_of_a_container_can_overlap():
    page = _page(BoxShape(1, (0, 0, 5, 5)), BoxShape(2, (1, 1, 3, 3), containers=(1,)), BoxShape(3, (2, 2, 4, 4), containers=(1,)))
    issues = diagnose.page_issues(page)
    assert [sorted(row["id"] for row in issue["shapes"]) for issue in issues] == [[2, 3]]


def test_a_shape_on_a_container_it_is_not_in_overlaps():
    assert _kinds(_page(BoxShape(1, (0, 0, 5, 5)), BoxShape(2, (1, 1, 2, 2)))) == ["overlap"]


@pytest.mark.parametrize(
    "box",
    [(-1, 1, 1, 2), (1, -1, 2, 1), (8, 1, 9, 2), (1, 10, 2, 12)],
    ids=["left", "bottom", "right", "top"],
)
def test_off_page_on_each_side(box):
    issues = diagnose.page_issues(_page(BoxShape(1, box)))
    assert [issue["kind"] for issue in issues] == ["off_page"]
    assert issues[0]["id"] == 1


def test_exactly_filling_the_page_is_on_it():
    assert _kinds(_page(BoxShape(1, (0, 0, 8.5, 11)))) == []


def test_unglued_connector_says_which_end():
    issues = diagnose.page_issues(_page(Line(1, "Dynamic connector", [("BeginX", 7)])))
    assert issues == [{"kind": "unglued_connector", "id": 1, "text": "", "from": 7, "to": None}]


def test_a_glued_connector_is_fine():
    assert _kinds(_page(Line(1, "Dynamic connector", [("BeginX", 7), ("EndX", 8)]))) == []


def test_a_line_someone_drew_is_not_a_connector():
    assert _kinds(_page(Line(1, None))) == []


def test_the_report_is_capped_but_says_how_many():
    shapes = [BoxShape(index, (-1, 1, 1, 2)) for index in range(1, 61)]
    # Stacked up, they overlap too: 60 off the page and every pair among them.
    payload = diagnose.report(_page(*shapes))
    assert len(payload["issues"]) == diagnose.MAX_ISSUES
    assert payload["issue_count"] == 60 + 60 * 59 // 2


def test_the_report_leaves_the_count_out_when_all_are_listed():
    assert diagnose.report(_page(BoxShape(1, (1, 1, 2, 2)))) == {"issues": []}
