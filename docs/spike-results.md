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
| S3 | After the newest closes? | **Intermittent.** Attach failed with `MK_E_UNAVAILABLE` in some runs and worked in others, while older instances kept running. It stayed broken for over a minute in one run. A focus request preceded recovery once, but the cause is not established. | "Can't attach" is not "not running", so it reports `visio_unreachable`. |
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

## Other traps

- **Deleting the last page wipes it.** `Pages(1).Delete(0)` on a one-page drawing succeeds, deletes all content and creates a blank page. Undo restores it if it was in a scope. The tool refuses.
- **`AddNamedRow` throws on a row the master already has.** Flowchart masters ship with `Prop.Owner` and others. Check `CellExistsU` first. Row names only allow `a-zA-Z0-9_`.
- **Unset Shape Data reads as `0.0000`.** `FormulaU == ""` means unset.
- **`Application.ProcessID` is not the Windows PID.** Use `GetWindowThreadProcessId(app.WindowHandle32)`.
- **Closing a drawing can close its stencils.** Iterating `Documents` while closing fails with "Invalid document identifier". Snapshot names and re-check.
- **`GluedShapes` needs three arguments** under late binding: `(flags, "", None)`.
- **`_U` and `_M` are US and metric masters.** Process in `BASFLO_U` is 1 in wide; in `BASFLO_M`, 25 mm.

## Not tested

- A drawing on OneDrive. Word has a known problem there where `GetActiveObject` returns an empty proxy. No test files were created in a synced folder to find out.
- Non-English locale folders.
- Pages with a drawing scale.
- Behaviour while a user is typing in a shape (see S5).
