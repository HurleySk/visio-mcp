import asyncio
import concurrent.futures
import json
import logging
import os
import queue
import subprocess
import threading
import time
import unicodedata

from visio_mcp.errors import MK_E_UNAVAILABLE, BUSY_HRESULTS, VisioError, busy_error, classify

logger = logging.getLogger(__name__)
_local = threading.local()
_lock = threading.Lock()
_worker = None
_connector_factory = None
_clock = time.monotonic
_sleep = time.sleep

VISIO_CLSID = "{00021A20-0000-0000-C000-000000000046}"
VISIO_EXTENSIONS = (".vsdx", ".vsdm", ".vsd", ".vdx", ".vstx", ".vstm", ".vst")
VIS_TYPE_STENCIL = 2
NO_SCOPE = -1

NOT_ATTACHED = ("visio_not_running", "visio_unreachable")
LAUNCH_GRACE = 90.0
START_WAIT = 45.0
SCOPE_WAIT = 2.0
LEFTOVER_SCOPE_LIMIT = 16

_launch_lock = threading.Lock()
_last_launch = None
_stuck_edits = []


class Win32Connector:
    def thread_init(self):
        import pythoncom

        pythoncom.CoInitialize()

    def thread_exit(self):
        import pythoncom

        pythoncom.CoUninitialize()

    def pump(self):
        import pythoncom

        pythoncom.PumpWaitingMessages()

    def attach(self):
        """Reach a running Visio. Never starts one, which is why this avoids Dispatch."""
        import pythoncom
        from win32com.client import dynamic

        try:
            unknown = pythoncom.GetActiveObject(VISIO_CLSID)
        except pythoncom.com_error as exc:
            if exc.hresult != MK_E_UNAVAILABLE:
                raise
        else:
            return dynamic.Dispatch(unknown.QueryInterface(pythoncom.IID_IDispatch))
        for document in self.other_documents():
            return document.Application
        if self.process_running():
            raise VisioError(
                "visio_unreachable",
                "Visio is running but cannot be reached",
                retryable=True,
                hint="This can happen after a newer Visio window closes. Click the Visio window and retry, or open a saved drawing.",
            )
        raise VisioError(
            "visio_not_running",
            "Microsoft Visio is not running",
            hint="Open Visio, or call visio_new_drawing or visio_open_drawing, which start it.",
        )

    def other_documents(self):
        """Saved drawings registered as file monikers, in any Visio instance."""
        import pythoncom
        from win32com.client import dynamic

        table = pythoncom.GetRunningObjectTable(0)
        monikers = table.EnumRunning()
        found = []
        while True:
            batch = monikers.Next(1)
            if not batch:
                return found
            try:
                name = batch[0].GetDisplayName(pythoncom.CreateBindCtx(0), None)
                if not name.lower().endswith(VISIO_EXTENSIONS):
                    continue
                bound = table.GetObject(batch[0])
                document = dynamic.Dispatch(bound.QueryInterface(pythoncom.IID_IDispatch))
                document.Type
                found.append(document)
            except Exception:
                continue

    def rejecting(self, app):
        """pywin32 reports a rejected name lookup as AttributeError, so ask directly."""
        try:
            app._oleobj_.GetIDsOfNames("Documents")
        except Exception as exc:
            return getattr(exc, "hresult", None) in BUSY_HRESULTS
        return False

    def process_running(self):
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq VISIO.EXE", "/NH"],
            capture_output=True,
            text=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return "VISIO.EXE" in result.stdout.upper()

    def launch(self):
        try:
            os.startfile("visio.exe")
        except OSError as exc:
            raise VisioError(
                "visio_not_installed",
                f"Could not start Microsoft Visio: {exc}",
                hint="Install Visio, or open it by hand and retry.",
            ) from exc


class VisioSession:
    def __init__(self, connector):
        self._connector = connector
        self._app = None
        self._scope = None
        self._label = None
        self.last_error = None
        self.attach_error = None
        self.abandoned = False
        self.warnings = []

    @property
    def app(self):
        if self._app is None:
            self._app = self._connector.attach()
        return self._app

    @property
    def connector(self):
        return self._connector

    def drop(self):
        self._app = None

    def prepare(self, start):
        """Apply a tool's start policy: never, if_missing or report."""
        self.attach_error = None
        if start == "if_missing":
            self._start_if_missing()
        elif start == "report":
            try:
                self.wait_until_ready()
            except VisioError as err:
                if err.code not in NOT_ATTACHED:
                    raise
                self.attach_error = err
        else:
            self.wait_until_ready()

    def _ready_if_attached(self):
        try:
            self.wait_until_ready()
        except VisioError as err:
            if err.code not in NOT_ATTACHED:
                raise
            return False
        return True

    def _start_if_missing(self):
        global _last_launch
        if self._ready_if_attached():
            return
        # Only our own launch is worth waiting for. An unreachable instance may
        # stay that way, and every launch makes a new process, so start another.
        with _launch_lock:
            now = _clock()
            if _last_launch is None or now - _last_launch > LAUNCH_GRACE:
                self._connector.launch()
                _last_launch = now
        deadline = _clock() + START_WAIT
        while not self._ready_if_attached():
            if _clock() >= deadline:
                raise VisioError(
                    "visio_start_timeout",
                    f"Visio did not become reachable within {START_WAIT:g} seconds of starting",
                    retryable=True,
                    hint="Visio may still be starting. Retry shortly.",
                )
            _sleep(0.25)

    def begin_call(self, label):
        self.last_error = None
        self.warnings = []
        self._scope = None
        self._label = label

    def open_scope(self, app):
        """Open this call's undo scope. Later calls in the same tool call share it."""
        if self._scope is not None:
            return
        self._close_leftover_scope(app)
        self._scope = (app, app.BeginUndoScope(self._label))

    def _close_leftover_scope(self, app):
        # Scopes outlive the client that opened them, and any edit made while one
        # is open joins it. Commit, because a rollback would take the user's edits too.
        deadline = _clock() + SCOPE_WAIT
        closed = 0
        while app.CurrentScope != NO_SCOPE and closed < LEFTOVER_SCOPE_LIMIT:
            if _clock() < deadline:
                _sleep(0.1)
                continue
            app.EndUndoScope(app.CurrentScope, True)
            closed += 1
        if closed:
            self.warnings.append(
                {
                    "code": "recovered_scope",
                    "message": "Closed an undo scope left open by an earlier edit. Its changes were kept and may share one undo step with edits made while it was open.",
                }
            )

    def commit_scope(self):
        if self._scope is None:
            return
        app, scope = self._scope
        self._scope = None
        app.EndUndoScope(scope, True)

    def rollback_scope(self):
        """Returns None when there was no scope, else whether the rollback worked."""
        if self._scope is None:
            return None
        app, scope = self._scope
        self._scope = None
        try:
            app.EndUndoScope(scope, False)
        except Exception:
            logger.exception("undo scope rollback failed")
            return False
        return True

    def pause(self, seconds):
        _sleep(seconds)

    def rejecting(self):
        return self._app is not None and self._connector.rejecting(self._app)

    def wait_until_ready(self, window=10.0):
        deadline = _clock() + window
        delay = 0.1
        while True:
            try:
                self.app.Documents.Count
                return
            except Exception as exc:
                err = classify_live(exc)
            if err.code in ("disconnected", "busy"):
                # A proxy built during a modal dialog has no type info, so never keep one.
                self.drop()
            else:
                raise err
            if _clock() >= deadline:
                if err.code == "busy":
                    raise VisioError(
                        "modal_dialog",
                        f"Visio did not accept calls for {window:g} seconds",
                        retryable=True,
                        hint="Visio may have a dialog open. Close it and retry.",
                    )
                raise err
            _sleep(delay)
            delay = min(delay * 2, 1.0)


class ComWorker:
    def __init__(self, connector):
        self._connector = connector
        self._queue = queue.SimpleQueue()
        self._session = None
        self.poisoned = False
        self._thread = threading.Thread(target=self._run, name="visio-com", daemon=True)
        self._thread.start()

    def submit(self, fn, start="never"):
        future = concurrent.futures.Future()
        self._queue.put((fn, start, future))
        if self.poisoned:
            self.poison()
        return future

    def alive(self):
        return not self.poisoned and self._thread.is_alive()

    def poison(self):
        self.poisoned = True
        session = self._session
        if session is not None:
            # The caller has been told this call failed, so it must not commit.
            session.abandoned = True
        while True:
            try:
                _, _, future = self._queue.get_nowait()
            except queue.Empty:
                return
            if future.set_running_or_notify_cancel():
                future.set_exception(_timeout_error("Visio stopped responding"))

    def _run(self):
        session = None
        try:
            self._connector.thread_init()
            session = VisioSession(self._connector)
            session.abandoned = self.poisoned
            self._session = session
            _local.session = session
            while not self.poisoned:
                try:
                    fn, start, future = self._queue.get(timeout=0.1)
                except queue.Empty:
                    self._connector.pump()
                    continue
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    session.prepare(start)
                    future.set_result(fn(session))
                except BaseException as exc:
                    future.set_exception(exc)
        except Exception:
            logger.exception("visio-com worker stopped")
        finally:
            self.poison()
            _local.session = None
            if session is not None:
                session.drop()
            try:
                self._connector.thread_exit()
            except Exception:
                pass


def _timeout_error(message, *, applied=False):
    if applied:
        return VisioError(
            "timeout",
            message,
            retryable=False,
            hint="The edit is rolled back if Visio responds later. Check the drawing before retrying.",
        )
    return VisioError(
        "timeout",
        message,
        retryable=True,
        hint="Visio may be showing a dialog or processing a large drawing.",
    )


def configure(connector_factory, *, clock=time.monotonic, sleep=time.sleep):
    global _connector_factory, _clock, _sleep
    reset()
    _connector_factory = connector_factory
    _clock = clock
    _sleep = sleep


def reset():
    global _worker, _connector_factory, _clock, _sleep, _last_launch
    with _lock:
        if _worker is not None:
            _worker.poison()
        _worker = None
        _stuck_edits.clear()
    _last_launch = None
    _connector_factory = None
    _clock = time.monotonic
    _sleep = time.sleep


def _get_worker():
    global _worker
    with _lock:
        if _worker is None or not _worker.alive():
            factory = _connector_factory or Win32Connector
            _worker = ComWorker(factory())
        return _worker


def _refuse_while_an_edit_is_stuck():
    # An open scope sweeps in any edit made meanwhile, and the timed-out edit's
    # rollback would then undo that too.
    with _lock:
        _stuck_edits[:] = [future for future in _stuck_edits if not future.done()]
        if _stuck_edits:
            raise VisioError(
                "edit_in_progress",
                "An earlier edit that timed out is still running in Visio",
                retryable=True,
                hint="Close any dialog open in Visio. The earlier edit rolls back when it ends, then this can be retried.",
            )


async def run_com(fn, *, timeout=60.0, mutates=False, start="never"):
    if mutates:
        _refuse_while_an_edit_is_stuck()
    worker = _get_worker()
    future = worker.submit(fn, start)
    try:
        return await asyncio.wait_for(asyncio.wrap_future(future), timeout)
    except asyncio.TimeoutError:
        if future.cancel():
            raise _timeout_error(f"Visio was still busy with an earlier call after {timeout:g} seconds") from None
        if future.done():
            return future.result()
        worker.poison()
        if mutates:
            with _lock:
                _stuck_edits.append(future)
        raise _timeout_error(f"Visio did not respond within {timeout:g} seconds", applied=mutates) from None


def current_session():
    return getattr(_local, "session", None)


def get_visio_app():
    session = current_session()
    if session is None:
        raise VisioError("internal", "get_visio_app was called off the COM thread")
    return session.app


def classify_live(exc):
    if isinstance(exc, AttributeError):
        session = current_session()
        if session is not None and session.rejecting():
            return busy_error()
    return classify(exc)


def error_json(exc):
    err = classify_live(exc)
    session = current_session()
    if session is not None:
        session.last_error = err
        if err.code in ("disconnected", "busy"):
            session.drop()
    return json.dumps(err.to_dict(), ensure_ascii=False)


def _normal(text):
    return unicodedata.normalize("NFC", text).lower()


def _matches(document, target):
    name = _normal(os.path.basename(target))
    if _normal(document.Name) == name:
        return True
    return os.path.isabs(target) and _normal(os.path.normpath(document.FullName)) == _normal(os.path.normpath(target))


def drawings_in(app):
    """Open drawings only. Documents also lists stencils."""
    return [
        app.Documents(i)
        for i in range(1, app.Documents.Count + 1)
        if app.Documents(i).Type != VIS_TYPE_STENCIL
    ]


def find_document(app, drawing=None):
    """Find an open drawing by name or full path, or the active one when none is named."""
    drawings = drawings_in(app)
    session = current_session()
    if not drawing:
        if not drawings:
            raise VisioError(
                "no_drawing",
                "No drawings are open in Visio",
                hint="Call visio_new_drawing or visio_open_drawing.",
            )
        try:
            active = app.ActiveDocument
        except Exception:
            return drawings[0]
        return active if active.Type != VIS_TYPE_STENCIL else drawings[0]
    for document in drawings:
        if _matches(document, drawing):
            return document
    if session is not None:
        for document in session.connector.other_documents():
            if _matches(document, drawing):
                return document
    raise VisioError(
        "drawing_not_found",
        f"Drawing '{drawing}' is not open in Visio. Open drawings: {[d.Name for d in drawings]}",
        hint="Pass one of the open drawing names, or omit drawing to use the active one.",
    )


def edit_doc(drawing=None):
    """Find a drawing for editing and open this call's undo scope on its Visio.

    Call it after validating, so a bad request opens no scope. The scope belongs to
    the drawing's own instance, which need not be the one this session attached to.
    """
    session = current_session()
    if session is None:
        raise VisioError("internal", "edit_doc was called off the COM thread")
    document = find_document(session.app, drawing)
    if document.ReadOnly:
        raise VisioError(
            "read_only",
            f"Drawing '{document.Name}' is open read-only",
            hint="Save a copy with visio_save(path=...), which is writable, then edit that.",
        )
    session.open_scope(document.Application)
    return document
