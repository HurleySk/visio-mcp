# visio-mcp

MCP server for Microsoft Visio on Windows. It drives Visio over COM to generate diagrams with real stencil shapes, glued connectors and automatic layout, to edit drawings open in Visio, and to read drawings as context.

**Status: under construction.** See [docs/spike-results.md](docs/spike-results.md) for what testing against real Visio established.

## Troubleshooting

- **"Visio is running with different administrator rights from this server".** Windows blocks COM between a process running as administrator and one that isn't. Start Visio and the app that runs this server the same way: both as administrator, or neither.
- **`second_visio` warning.** A Visio this server can't reach was already running, so a new window was opened. Drawings in the other window aren't available until the rights match.
- **`drop_handler_skipped` warning.** Some masters, such as org chart shapes and off-page references, run an add-on when dropped, and its dialog would block the call. Those masters are dropped with Visio's automation events off for that one drop, so the shape arrives without what the add-on would have added.
- **`events_restored` warning.** Visio saves its automation events setting when it exits. If a server stopped while it had events off for a drop, the next call turns them back on. To check by hand, with Visio closed: `HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Visio\Application\EventsEnabled` set to `0` means events are off.

## Licence

MIT. See [LICENSE](LICENSE). The COM runtime derives from [word-mcp](https://github.com/HurleySk/word-mcp).
