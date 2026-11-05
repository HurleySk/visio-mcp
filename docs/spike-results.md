# Spike results

Throwaway scripts ran against real Visio before any tool code was written, to settle what the documentation leaves open.

**Environment:** Visio Professional 16.0.20430 (click-to-run, en-US), Windows 11, Python 3.13, pywin32 with late binding through `win32com.client.dynamic`.

## What the design relies on

| # | Question | Result | What the code does |
|---|---|---|---|
| S1 | Does late binding handle Visio's cell API? | Yes. `ResultIU`, `FormulaU`, `Result("in")`, `Result(65)`, and `SetResult` all work. No makepy needed. | Uses dynamic dispatch only. |
| S1 | Reading a missing cell? | Raises `0x80020009`. | Checks `CellExistsU` first. |
| S2 | Do array methods work? | `DropMany` and `GetResults` both take plain lists and return tuples. `GetResults` needs a non-empty units list. | Uses a `Drop` loop; 200 drops take 0.4 to 0.9 s, so batching isn't needed. |
| S3 | Does `GetActiveObject` find click-to-run Visio? | Yes, 2.5 s after launch. Backstage does not block COM. With no Visio it returns `MK_E_UNAVAILABLE` and starts nothing. | Attaches by CLSID, never `Dispatch`, which could start Visio. |
| S3 | How many instances? | Every launch is a new process: `VISIO.EXE`, `VISIO.EXE <file>`, and the shell's open action (what a double-click runs). | Several instances are treated as normal. |
| S3 | Which instance does attach reach? | The newest one. | |
| S3 | After the newest closes? | **Intermittent.** Attach failed with `MK_E_UNAVAILABLE` in some runs and worked in others, while older instances kept running. It stayed broken for over a minute in one run. A focus request preceded recovery once, but the cause is not established. These runs came from an elevated session and some instances were started by the shell, so S12 may explain some of them; that wasn't separated out. | "Can't attach" is not "not running", so it reports `visio_unreachable`. |
| S3 | Can anything else reach an older instance? | Saved drawings: a ROT file moniker binds to the document in any instance, and edits through it work. `OBJID_NATIVEOM` fails on Visio windows (`0x80004005`), so there is no window route. | Falls back to file monikers, filtered to Visio extensions. |
| S4 | One Ctrl+Z per scope? | Yes. Rollback (`EndUndoScope(id, False)`) restores shapes and removes masters pulled in during the scope. | The decorator owns one scope per mutating call. |
| S4 | Nesting? | Behaves as expected. Inner cancel keeps outer work; outer cancel removes everything, including a committed inner scope. | |
| S4 | Is undo per document? | **No.** `Application.Undo()` walks one app-wide stack, across documents, whichever is active. | `visio_undo` reports per-drawing shape counts. |
| S4 | Are scopes tied to a client? | **No.** A scope belongs to the Visio app. An edit from any other client while it's open joins it and is rolled back with it. | New edits are refused while a timed-out edit is still running. |
| S4 | What survives a crash? | **The scope.** It stays open after its thread ends, and after its process hard-exits. Anyone can close it by ID. | Mutating calls detect an unowned scope and commit it, which keeps any user edits caught inside. The result carries a warning. |
| S5 | COM during a modal dialog (Page Setup)? | Calls are rejected with `RPC_E_CALL_REJECTED`, **except** that pywin32 reports a rejected name lookup as `AttributeError: <unknown>.X`. A proxy created during the dialog has no type info. | `AttributeError` triggers a raw `GetIDsOfNames` probe; a rejection there is `busy`, and the proxy is dropped. |
| S5 | Dirty close? | `Close()` shows a prompt and blocks only the caller. Visio still answers other clients. | Close refuses unsaved changes unless `discard`, which sets `Saved = True` first, so the prompt never appears. |
| S5 | Bad inputs? | Missing file, bad template and corrupt file all raise errors, with no dialog. | `AlertResponse` isn't used. Setting it app-wide could answer the user's own prompts. |
| S5 | Text-edit mode? | **Not established.** Automation could not enter it reliably, so behaviour while a user is typing in a shape is unknown. | Errors pass through the normal classifier. To check by hand. |
| S6 | Are IDs stable? | Yes, across delete+undo, create+undo+create, and save+reopen. Not reused in these runs, though Microsoft documents that they can be. | Edits echo `{id, name, text, master}`. |
| S6 | Can COM references be kept? | **No.** After an undo a held shape reference reads `ID == 0`. | Shapes are looked up afresh on every call. |
| S7 | Is `Page.Layout` good enough? | Yes for `top_down` and `left_right` with flowchart routing: clean results on a 12-node branching flow with a loop. Radial and tree routing are weaker. Circular is poor. All connectors stayed glued. | Default is `top_down` with flowchart routing. |
| S8 | How to use extra stencils? | Open hidden, drop, close: the master stays in the document, and undo and redo still work. A hidden stencil left open outlives its drawing. Microsoft documents that docked needs an active drawing window; that limit wasn't tested here. | Hidden and closed straight away, unless it was already open. |
| S9 | `Page.Export`? | Crops to the shapes' bounding box at 96 dpi: exact `BoundingBox(1)` size. No focus change, `Saved` unchanged. A blank page still exports. | Maps pin positions to pixels directly for ID badges. |
| S9 | Image size control? | Custom resolution scales correctly and restores exactly. Custom size distorts the aspect ratio. | `max_px` picks a resolution; settings restored in `finally`. |
| S9 | Overwrites? | `Page.Export`, `SaveAs` and `ExportAsFixedFormat` all **replace files silently**. | Every write checks `overwrite` in Python. |
| S9 | Export errors? | A missing folder reports "requires a fully qualified file name". An unknown extension reports "Identify the file type". | Checks folder and format first, with clear messages. |
| S10 | Read-only drawings? | Edits are allowed and undoable, but `Save` fails ("presently disabled"). `SaveAs` works and the result is writable. Opening a file already open returns it. Opening it writable while open read-only is an error. | Edits and in-place save are refused; save-as is allowed. |
| S11 | Does a Visio the server started outlive the client? | **Not if it's a child of the server.** The Python MCP SDK runs servers in a kill-on-close Job on Windows. Visio started with a shell start joined that Job and died as soon as the client process exited, losing unsaved work, with nothing in the event log. With a plain subprocess client and no Job, it survived. | Starts Visio through COM instead. |
| S11 | Does COM activation escape the Job? | **Yes.** `CoCreateInstance(CLSID, None, CLSCTX_LOCAL_SERVER, IID_IDispatch)` from inside a kill-on-close Job: Visio up in 1.5 to 2.4 s, visible, same rights as the caller. An unsaved drawing made inside the Job was still there and still unsaved 10 s after the client exited and the Job was released, and Visio stayed attachable. Releasing the only reference doesn't close Visio. | `launch()` uses it, on a thread of its own so a slow boot can't use up the tool's time limit. |
| S12 | Do administrator rights matter? | **Yes.** An elevated process can't see a non-elevated Visio: `GetActiveObject` returns `MK_E_UNAVAILABLE` even though the window is on screen. Visio started by COM activation keeps the caller's rights. | `visio_unreachable` says when the rights differ. `if_missing` warns (`second_visio`) if it has to start another window. |
| S13 | Do masters run code when dropped? | **Often.** Of 9,267 US masters in 363 stencils, 566 have an active `EventDrop`: `RUNADDON` 406, `DOCMD` 223, `OPENTEXTWIN` 86. Another 1,767 use `SETF(GetRef(HideText),FALSE)`, which only finishes setting the shape up. So events can't simply be turned off for every drop. | Checks each master, including group members, once per call. Masters with a handler that can open something drop with events off. |
| S13 | Which ones block? | `ORGCH:Executive` (the first org shape in a drawing), `BASFLO:Off-page reference` and `XFUNC:Separator (vertical)` next to a horizontal lane each opened a modal dialog and blocked `Drop`. `XFUNC:Swimlane`, `PERIPH:Legend` and `MATRIXDOT:Dot 1` didn't. | Treats the whole list as interactive, since which ones block depends on what's already on the page. |
| S13 | Does events off clear them? | **Yes,** at 0.1 to 0.2 s per drop, with the master's own text and data in place. `app.EventsEnabled` reads back as set. | `events_off(app)` wraps that one `Drop`, restoring in `finally`, and the result carries `drop_handler_skipped`. |
| S14 | Is `EventsEnabled` saved? | **Yes.** Set it to 0 and quit: Visio writes `HKCU\Software\Microsoft\Office\16.0\Visio\Application\EventsEnabled` = `0` on exit, and the next start reads 0. The value isn't written until exit. Microsoft's reference says the property is True whenever an instance starts, which isn't what happened here. | A server lost inside the guard would disable events for all later sessions, so the guard writes a marker file first. The next call finding a stale marker restores the setting and reports `events_restored`. A guard does nothing if the user already had events off. |

## Generating a diagram

- **Units:** `PageWidth.Units` is 73 (`visInchFrac`) on a blank US drawing and 65 on a US template. Blank metric and metric templates are 70. Metric is 69, 70, 71 or 72.
- **Inside an open undo scope,** `Export` and `SaveAs` both work, `Saved` stays True after the commit, and one Undo then empties the page.
- **A new drawing can be closed inside an open scope and the scope rolled back,** or the other way round. Both work.
- **Templates open their stencils as Type 2 documents,** which close with the drawing. `BASFLO_U.VSTX` opens `BASFLO_U.vssx` and `XFUNC_U.vssx`. Templates don't reliably list their stencils in `visio/windows.xml` (`BLOCK_U` lists none), so a bare master name can only be checked after the template opens.
- **`Page.Layout` leaves an unconnected shape where it is** unless `ObjType` is set explicitly to `1`, even though it already reads 1 by inheritance.
- **Layout orders siblings by where they start.** Dropping on a diagonal (x = 1 + 0.25i, y = 10 − 0.25i) keeps the caller's order, for `top_down` and for `left_right`.
- **Setting `Height` through `ResultIU` replaces the `User.ResizeTxtHeight` formula,** so that shape stops growing to fit its text.
- **Routes, by eye on a contact sheet:** flowchart is clean, tree is slanted, right_angle is uglier, radial looks grid-like because route has no effect, and circular is poor. `AvenueSizeX` and `AvenueSizeY` control spacing.
- **Connector labels aren't checked.** Real diagrams showed a label colliding with a shape's text, which the `issues` list doesn't report but the PNG shows. Several connectors can also share one line, which is Visio's router.

## Tried and dropped

- **Starting Visio with `explorer.exe <exe>`,** to escape the Job. It works only for a non-elevated server. From an elevated one it produces a Visio the server can't reach (S12).
- **A false positive on the way there.** One run looked as if a shell-started Visio survived an SDK client. It didn't show that: the server couldn't attach, started a second instance, and that one died with the client. The survivor counted was the shell's own instance.
- **`IsProcessInJob(handle, None)` to tell who's in a Job.** It's true for everything here, so it can't discriminate. Watching the process when the Job handle closes can: the kill fires then, not when the client's last call ends.
- **`WM_ACTIVATEAPP`, `WM_ACTIVATE` and a real focus change,** to make a non-elevated Visio reachable from an elevated process. None helped, because the block is between rights levels.
- **Closing the add-on's dialog with `WM_CLOSE` from a watcher.** It unblocks the call, but what happens next depends on the add-on: an org shape stayed, and a vertical separator failed with "Inappropriate target". It also flashes a dialog, so events off is cleaner.
- **Force-killing Visio to clean up after tests.** A kill can make the next start show "Visio couldn't start last time" and offer safe mode, which blocks everything until it's answered. It came from a script calling `Quit` while still holding a reference, which keeps Visio alive, then killing it. Drop the reference first, then wait.

## Other traps

- **Deleting the last page wipes it.** `Pages(1).Delete(0)` on a one-page drawing succeeds, deletes all content and creates a blank page. Undo restores it if it was in a scope. The tool refuses.
- **`AddNamedRow` throws on a row the master already has.** Flowchart masters ship with `Prop.Owner` and others. Check `CellExistsU` first. Row names only allow `a-zA-Z0-9_`.
- **Unset Shape Data reads as `0.0000`.** `FormulaU == ""` means unset.
- **`Application.ProcessID` is not the Windows PID.** Use `GetWindowThreadProcessId(app.WindowHandle32)`.
- **Closing a drawing can close its stencils.** Iterating `Documents` while closing fails with "Invalid document identifier". Snapshot names and re-check.
- **`GluedShapes` needs three arguments** under late binding: `(flags, "", None)`.
- **`_U` and `_M` are US and metric masters.** Process in `BASFLO_U` is 1 in wide; in `BASFLO_M`, 25 mm.
- **"Windows fatal exception: code 0x80010001" under pytest is not a crash.** COM raises call rejections as Windows exceptions, and pytest's faulthandler prints a dump even when Python catches them. The tests pass with it on. The project turns it off.
- **Visio takes a moment to exit after `Quit`.** A process check straight afterwards can still find it.
- **`MATRIXDOT:Dot 2` faults whether events are on or off.** `Drop` raises `RPC_E_SERVERFAULT` after about 7 s and the shape lands anyway. It looks like a defect in that master, not its handler. The tool reports the error and rolls back.

## Not tested

- A saved drawing that runs an add-on when opened. `visio_open_drawing` leaves events on, because that's the user's own file, so a dialog there ends in a `timeout` or `modal_dialog` error.
- A server that isn't elevated. All of this ran from an elevated session, so the non-elevated path needs its own run.

- A drawing on OneDrive. Word has a known problem there where `GetActiveObject` returns an empty proxy. No test files were created in a synced folder to find out.
- Non-English locale folders.
- Pages with a drawing scale.
- Behaviour while a user is typing in a shape (see S5).
