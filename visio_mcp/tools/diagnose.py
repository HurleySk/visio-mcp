import json

from visio_mcp.addressing import find_page, master_name
from visio_mcp.com_runtime import current_session, find_document
from visio_mcp.live_tool import live_tool
from visio_mcp.tools.read import _inches, connector_ends

VIS_BBOX_UPRIGHT_WH = 1
# Shapes touching, or a hair over the edge, are how diagrams normally look.
TOLERANCE = 0.01
MAX_ISSUES = 50


def _brief(shape):
    return {"id": shape.ID, "text": shape.Text}


def _unglued(shape):
    ends = connector_ends(shape)
    if ends["from"] is not None and ends["to"] is not None:
        return None
    return {"kind": "unglued_connector", **_brief(shape), "from": ends["from"], "to": ends["to"]}


def _off_page(shape, box, width, height):
    left, bottom, right, top = box
    if left >= -TOLERANCE and bottom >= -TOLERANCE and right <= width + TOLERANCE and top <= height + TOLERANCE:
        return None
    return {
        "kind": "off_page",
        **_brief(shape),
        "box": {"left": round(left, 3), "bottom": round(bottom, 3), "right": round(right, 3), "top": round(top, 3)},
    }


def _overlaps(rows):
    """Pairs whose boxes overlap, leaving out a container and its own members."""
    rows = sorted(rows, key=lambda row: row[1][0])
    for position, (first, box, containers) in enumerate(rows):
        for second, other, other_containers in rows[position + 1:]:
            if other[0] >= box[2] - TOLERANCE:
                break
            if other[1] >= box[3] - TOLERANCE or box[1] >= other[3] - TOLERANCE:
                continue
            if first.ID in other_containers or second.ID in containers:
                continue
            yield {"kind": "overlap", "shapes": [_brief(first), _brief(second)]}


def page_issues(page):
    """Find shapes on top of each other, shapes off the page, and connectors not glued at both ends."""
    sheet = page.PageSheet
    width, height = _inches(sheet, "PageWidth"), _inches(sheet, "PageHeight")
    issues = []
    rows = []
    shapes = page.Shapes
    for index in range(1, shapes.Count + 1):
        shape = shapes(index)
        if shape.OneD:
            # A line someone drew has no master, and isn't meant to join anything.
            if master_name(shape) is not None:
                issue = _unglued(shape)
                if issue:
                    issues.append(issue)
            continue
        box = shape.BoundingBox(VIS_BBOX_UPRIGHT_WH)
        issue = _off_page(shape, box, width, height)
        if issue:
            issues.append(issue)
        rows.append((shape, box, tuple(shape.MemberOfContainers or ())))
    issues.extend(_overlaps(rows))
    return issues


def report(page):
    """Issues for a tool result, capped so a badly tangled page can't flood it."""
    issues = page_issues(page)
    payload = {"issues": issues[:MAX_ISSUES]}
    if len(issues) > MAX_ISSUES:
        payload["issue_count"] = len(issues)
    return payload


@live_tool(mutates=False)
def visio_diagnose_page(drawing: str | None = None, page: int | str | None = None):
    """Check a page for overlapping shapes, shapes off the page, and connectors not glued at both ends.

    Overlaps compare shapes' boxes, so a container and its own members don't count.
    Connectors crossing shapes, and text too big for its shape, aren't checked, so look at
    visio_get_page_image as well. Lists at most 50 issues, with issue_count saying if there are more.
    """
    document = find_document(current_session().app, drawing)
    target = find_page(document, page)
    return json.dumps({"drawing": document.Name, "page": target.Name, **report(target)}, ensure_ascii=False)
