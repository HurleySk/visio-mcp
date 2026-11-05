"""Write TOOLS.md from the tool snapshot, so the reference says exactly what clients are sent.

Usage: python scripts/tools_md.py tests/snapshots/live_tools.json TOOLS.md
"""
import inspect
import json
import pathlib
import sys

GROUPS = [
    ("Generate", ["visio_create_diagram", "visio_find_masters", "visio_list_templates", "visio_layout_page"]),
    (
        "Edit",
        [
            "visio_add_elements",
            "visio_update_shapes",
            "visio_delete_shapes",
            "visio_add_container",
            "visio_add_page",
            "visio_update_page",
            "visio_delete_page",
            "visio_undo",
            "visio_redo",
        ],
    ),
    (
        "Read",
        [
            "visio_list_open",
            "visio_get_page",
            "visio_get_shape",
            "visio_find_shapes",
            "visio_get_selection",
            "visio_get_page_image",
            "visio_diagnose_page",
        ],
    ),
    ("Files", ["visio_new_drawing", "visio_open_drawing", "visio_save", "visio_export", "visio_close_drawing"]),
]
NAMES = {"integer": "int", "number": "number", "string": "str", "boolean": "bool"}


def _type(schema):
    if "anyOf" in schema:
        return " or ".join(_type(option) for option in schema["anyOf"] if option.get("type") != "null")
    if "enum" in schema:
        return " or ".join(json.dumps(value) for value in schema["enum"])
    kind = schema.get("type")
    if kind == "array":
        return f"list of {_type(schema.get('items', {}))}"
    if kind == "object":
        values = schema.get("additionalProperties")
        if isinstance(values, dict):
            return f"map of str to {_type(values)}"
        return "object"
    return NAMES.get(kind, kind or "any")


def _fields(schema):
    """The object schema whose fields to list under a parameter, if it has one."""
    if schema.get("properties"):
        return schema
    for option in schema.get("anyOf", []):
        found = _fields(option)
        if found:
            return found
    if schema.get("type") == "array":
        return _fields(schema.get("items", {}))
    return None


def _parameters(schema, depth=0):
    lines = []
    required = set(schema.get("required", []))
    for name, field in schema.get("properties", {}).items():
        notes = [_type(field)]
        if name in required:
            notes.append("required")
        elif field.get("default") not in (None, [], {}):
            notes.append(f"default {json.dumps(field['default'])}")
        line = f"{'  ' * depth}- `{name}` ({', '.join(notes)})"
        if field.get("description"):
            line += f": {field['description']}"
        lines.append(line)
        inner = _fields(field)
        if inner:
            lines.extend(_parameters(inner, depth + 1))
    return lines


def _section(tool):
    annotations = tool.get("annotations", {})
    if annotations.get("readOnlyHint"):
        effect = "Reads only."
    elif annotations.get("destructiveHint"):
        effect = "May delete or overwrite something."
    else:
        effect = "Changes the drawing."
    parts = [f"### {tool['name']}", "", f"**{annotations.get('title', tool['name'])}.** {effect}", ""]
    parts.append(inspect.cleandoc(tool.get("description") or ""))
    parameters = _parameters(tool["inputSchema"])
    parts += ["", "Parameters:", "", *parameters] if parameters else ["", "No parameters."]
    return "\n".join(parts)


def render(tools):
    by_name = {tool["name"]: tool for tool in tools}
    grouped = [name for _, names in GROUPS for name in names]
    missing = sorted(set(by_name) - set(grouped))
    if missing or len(grouped) != len(by_name):
        raise SystemExit(f"Put every tool in one group in scripts/tools_md.py. Not grouped: {missing}")
    lines = [
        "# Tools",
        "",
        f"visio-mcp has {len(tools)} tools. This page is generated from `tests/snapshots/live_tools.json`",
        "by `scripts/tools_md.py`, so it matches what an MCP client is sent.",
        "",
        "Positions and sizes are in inches, with x and y the shape's centre measured from the page's",
        "bottom-left corner. A shape is named by its ID, else its exact name, else its exact text.",
        "Every tool that changes a drawing is one Ctrl+Z step, and rolls back if it fails part-way.",
        "",
        "| Group | Tools |",
        "|---|---|",
    ]
    for group, names in GROUPS:
        lines.append(f"| {group} | {', '.join(f'`{name}`' for name in names)} |")
    for group, names in GROUPS:
        lines += ["", f"## {group}"]
        for name in names:
            lines += ["", _section(by_name[name])]
    return "\n".join(lines) + "\n"


if __name__ == "__main__":
    snapshot, out = pathlib.Path(sys.argv[1]), pathlib.Path(sys.argv[2])
    out.write_text(render(json.loads(snapshot.read_text(encoding="utf-8"))), encoding="utf-8", newline="\n")
    print(out)
