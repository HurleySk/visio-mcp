"""What a diagram request looks like, and the checks that need no Visio."""
import difflib
import re
from typing import Literal

from pydantic import BaseModel, ConfigDict, Field

from visio_mcp.errors import VisioError

COLOUR = re.compile(r"^#[0-9a-fA-F]{6}$")


class _Spec(BaseModel):
    # A misspelt field must fail, or a label would silently go missing from the drawing.
    model_config = ConfigDict(extra="forbid", coerce_numbers_to_str=True, populate_by_name=True)


class NodeSpec(_Spec):
    key: str = Field(description="Your name for this node, which edges and the returned ids use.")
    master: str = Field(description="A master like 'BASFLO:Decision'. Find one with visio_find_masters.")
    text: str | None = None
    data: dict[str, str] | None = Field(default=None, description="Shape Data fields, by label.")
    x: float | None = Field(default=None, description="Centre, in inches from the page's left. Only with layout style 'none'.")
    y: float | None = Field(default=None, description="Centre, in inches from the page's bottom. Only with layout style 'none'.")
    w: float | None = Field(default=None, description="Width in inches. Leave out to use the master's.")
    h: float | None = Field(default=None, description="Height in inches. Setting it stops the shape growing to fit its text.")
    fill: str | None = Field(default=None, description="Fill colour as '#RRGGBB'.")


class EdgeSpec(_Spec):
    from_: str = Field(alias="from", description="The key of the node it starts at.")
    to: str = Field(description="The key of the node it ends at.")
    text: str | None = None
    arrow: Literal["end", "both", "none"] = "end"


class LayoutSpec(_Spec):
    style: Literal["top_down", "left_right", "radial", "circular", "none"] = Field(
        default="top_down",
        description="top_down and left_right give the cleanest results. radial and circular are weaker. none keeps each node's x and y.",
    )
    route: Literal["flowchart", "tree", "right_angle", "straight"] | None = Field(
        default=None,
        description="Connector style. Defaults to flowchart, or straight for radial and circular.",
    )
    spacing: float | None = Field(default=None, description="Gap between shapes in inches. Visio's default is 0.5.")
    fit_page: bool | None = Field(
        default=None,
        description="Resize the page to fit its shapes. Defaults to true, or false for style 'none', since it moves shapes.",
    )


class DiagramSpec(_Spec):
    nodes: list[NodeSpec]
    edges: list[EdgeSpec] = Field(default_factory=list)
    layout: LayoutSpec | None = Field(
        default=None,
        description="Defaults to top_down, or to none if the nodes have x and y.",
    )
    template: str | None = Field(default=None, description="A name from visio_list_templates, or a template file's absolute path.")
    units: Literal["us", "metric"] | None = None
    stencils: list[str] | None = Field(default=None, description="Stencils to search for masters named without one, like 'BASFLO'.")
    save_as: str | None = Field(default=None, description="An absolute path to save to.")
    overwrite: bool = Field(default=False, description="Replace the file at save_as if there is one.")
    return_image: bool = True


def check_graph(nodes, edges):
    """Everything about the nodes and edges that can be wrong without looking at Visio."""
    keys = set()
    for node in nodes:
        if not node.key:
            raise ValueError("A node's key can't be empty.")
        if node.key in keys:
            raise ValueError(f"Two nodes use the key '{node.key}'. Keys must be unique.")
        keys.add(node.key)
        if node.fill is not None and not COLOUR.match(node.fill):
            raise ValueError(f"Node '{node.key}': fill must look like '#RRGGBB', not '{node.fill}'.")
        for size in (node.w, node.h):
            if size is not None and size <= 0:
                raise ValueError(f"Node '{node.key}': w and h must be greater than 0.")
        if node.data and not all(node.data):
            raise ValueError(f"Node '{node.key}': a Shape Data label can't be empty.")
    for edge in edges:
        for end in (edge.from_, edge.to):
            if end not in keys:
                raise VisioError(
                    "unknown_node",
                    f"An edge refers to '{end}', which isn't a node's key",
                    details=difflib.get_close_matches(end, list(keys), n=5, cutoff=0.5) or None,
                )


def resolve_layout(layout, nodes):
    """Settle how nodes get placed. x and y and a layout would both place them, so one has to win."""
    positioned = [node for node in nodes if node.x is not None or node.y is not None]
    half = [node.key for node in positioned if node.x is None or node.y is None]
    if half:
        raise ValueError(f"x and y go together. Only one of them is set on: {half}")
    if layout is None:
        layout = LayoutSpec(style="none" if positioned else "top_down")
    if layout.style == "none":
        missing = [node.key for node in nodes if node.x is None]
        if missing:
            raise ValueError(f"Layout style 'none' needs x and y on every node. They're missing on: {missing}")
        if layout.route is not None or layout.spacing is not None:
            raise ValueError("route and spacing only apply when a layout runs, so they can't go with style 'none'.")
    elif positioned:
        raise ValueError(
            f"Layout style '{layout.style}' would move nodes with x and y. "
            "Use style 'none' to keep positions, or leave them out."
        )
    check_spacing(layout.spacing)
    return layout


def check_spacing(spacing):
    if spacing is not None and spacing <= 0:
        raise ValueError("spacing must be greater than 0.")
