# visio-mcp

MCP server for Microsoft Visio on Windows. It drives Visio over COM to generate diagrams with real stencil shapes, glued connectors and automatic layout, to edit drawings open in Visio, and to read drawings as context.

One `visio_create_diagram` call builds a whole diagram: it drops the shapes, glues the connectors, lays out the page and returns a PNG, so the model can check the result in the same call.

## Requirements

- Windows with the Visio desktop app installed. Tested with Visio Professional 16.0, click-to-run.
- Python 3.11 or later, and [uv](https://docs.astral.sh/uv/).

## Install

The PyPI name `visio-mcp` belongs to another project, so install from GitHub:

```sh
uv tool install --compile-bytecode git+https://github.com/HurleySk/visio-mcp
```

Then register it with your MCP client. For Claude Code:

```sh
claude mcp add --scope user visio-mcp -- visio-mcp
```

For a client configured with JSON:

```json
{
  "mcpServers": {
    "visio-mcp": { "command": "visio-mcp" }
  }
}
```

Restart the client to see the tools. To update, run the install command again with `--force`.

## Tools

25 tools in four groups. [TOOLS.md](TOOLS.md) lists every parameter and is generated from what clients are sent.

| Group | Tools |
|---|---|
| Generate | `visio_create_diagram`, `visio_find_masters`, `visio_list_templates`, `visio_layout_page` |
| Edit | `visio_add_elements`, `visio_update_shapes`, `visio_delete_shapes`, `visio_add_container`, `visio_add_page`, `visio_update_page`, `visio_delete_page`, `visio_undo`, `visio_redo` |
| Read | `visio_list_open`, `visio_get_page`, `visio_get_shape`, `visio_find_shapes`, `visio_get_selection`, `visio_get_page_image`, `visio_diagnose_page` |
| Files | `visio_new_drawing`, `visio_open_drawing`, `visio_save`, `visio_export`, `visio_close_drawing` |

## How it behaves

- **It uses the Visio you have open.** If Visio isn't running, only the tools that create or open a drawing start it, in a visible window. Read tools never start it. The server never quits Visio.
- **Each change is one Ctrl+Z step.** Every tool that changes a drawing does so in one undo scope, and rolls back if it fails part-way. A bad request is refused before anything changes.
- **Masters are found by search.** `visio_find_masters` searches the shapes installed with Visio, and a master is named like `BASFLO:Decision`. The server picks US or metric shapes to match the drawing.
- **Shapes are named by ID, else exact name, else exact text.** If a name matches several shapes, the error lists them. Edits echo each shape's ID, text and master, because Visio can reuse an ID after a delete.
- **Positions are in inches,** with x and y the shape's centre measured from the page's bottom-left corner.
- **Files are never replaced by accident.** Saving or exporting over an existing file needs `overwrite`, and closing a drawing with unsaved changes needs `discard`.

## Settings

| Variable | Default | Effect |
|---|---|---|
| `VISIO_MCP_TIMEOUT` | `60` | Seconds a tool waits for Visio. Generating a diagram and adding elements allow 120. |
| `VISIO_MCP_CONTENT` | Visio's own | The folder to search for stencils and templates, if Visio's install folder can't be found. |

## Troubleshooting

- **"Visio is running with different administrator rights from this server".** Windows blocks COM between a process running as administrator and one that isn't. Start Visio and the app that runs this server the same way: both as administrator, or neither.
- **"Visio is running but cannot be reached".** With two Visio windows open, only the newer one can be reached, and the older one may stay out of reach after the newer one closes. Click the Visio window and retry, or close and reopen Visio.
- **`modal_dialog` or `busy` errors.** Visio is showing a dialog or is in the middle of editing text. Close the dialog or click off the text, then retry.
- **`second_visio` warning.** A Visio this server can't reach was already running, so a new window was opened. Drawings in the other window aren't available until the rights match.
- **`drop_handler_skipped` warning.** Some masters, such as org chart shapes and off-page references, run an add-on when dropped, and its dialog would block the call. Those masters are dropped with Visio's automation events off for that one drop, so the shape arrives without what the add-on would have added.
- **`events_restored` warning.** Visio saves its automation events setting when it exits. If a server stopped while it had events off for a drop, the next call turns them back on. To check by hand, with Visio closed: `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Visio\Application\EventsEnabled` set to `0` means events are off.

## Development

```sh
uv sync
uv run pytest            # unit tests, no Visio needed
uv run pytest -m visio   # against real Visio: starts its own instances, so save your work first
```

The real-Visio tests start and quit their own Visio instances and never touch a drawing they didn't make. [docs/spike-results.md](docs/spike-results.md) records what testing against real Visio established, including the traps.

After changing a tool's signature or docstring, regenerate the snapshot and TOOLS.md:

```sh
uv run python scripts/dump_tools.py tests/snapshots/live_tools.json
uv run python scripts/tools_md.py tests/snapshots/live_tools.json TOOLS.md
```

## Licence

MIT. See [LICENSE](LICENSE). The COM runtime derives from [word-mcp](https://github.com/HurleySk/word-mcp).
