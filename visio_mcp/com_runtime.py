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
        self.last_error = None

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
        self.poisoned = False
        self._thread = threading.Thread(target=self._run, name="visio-com", daemon=True)
        self._thread.start()

    def submit(self, fn):
        future = concurrent.futures.Future()
        self._queue.put((fn, future))
        if self.poisoned:
            self.poison()
        return future

    def alive(self):
        return not self.poisoned and self._thread.is_alive()

    def poison(self):
        self.poisoned = True
        while True:
            try:
                _, future = self._queue.get_nowait()
            except queue.Empty:
                return
            if future.set_running_or_notify_cancel():
                future.set_exception(_timeout_error("Visio stopped responding"))

    def _run(self):
        session = None
        try:
            self._connector.thread_init()
            session = VisioSession(self._connector)
            _local.session = session
            while not self.poisoned:
                try:
                    fn, future = self._queue.get(timeout=0.1)
                except queue.Empty:
                    self._connector.pump()
                    continue
                if not future.set_running_or_notify_cancel():
                    continue
                try:
                    session.wait_until_ready()
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
    global _worker, _connector_factory, _clock, _sleep
    with _lock:
        if _worker is not None:
            _worker.poison()
        _worker = None
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


async def run_com(fn, *, timeout=60.0, mutates=False):
    worker = _get_worker()
    future = worker.submit(fn)
    try:
        return await asyncio.wait_for(asyncio.wrap_future(future), timeout)
    except asyncio.TimeoutError:
        if future.cancel():
            raise _timeout_error(f"Visio was still busy with an earlier call after {timeout:g} seconds") from None
        if future.done():
            return future.result()
        worker.poison()
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
