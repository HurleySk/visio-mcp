import threading

from visio_mcp.errors import RPC_E_CALL_REJECTED, RPC_E_DISCONNECTED, VisioError


class FakeComError(Exception):
    def __init__(self, hresult, strerror="fake com error"):
        super().__init__(strerror)
        self.hresult = hresult
        self.strerror = strerror
        self.excepinfo = None


class FakeDoc:
    def __init__(self, name="flow.vsdx", folder="C:\\docs", type_=1):
        self.Name = name
        self.FullName = f"{folder}\\{name}"
        self.Saved = True
        self.Type = type_


class FakeDocuments:
    def __init__(self, app):
        self._app = app

    @property
    def Count(self):
        if self._app.dead:
            raise FakeComError(RPC_E_DISCONNECTED)
        if self._app.busy_calls > 0:
            self._app.busy_calls -= 1
            raise FakeComError(RPC_E_CALL_REJECTED)
        return len(self._app.docs)

    def __call__(self, index):
        return self._app.docs[index - 1]


class FakeApp:
    def __init__(self, docs=None):
        self.docs = [FakeDoc()] if docs is None else docs
        self.busy_calls = 0
        self.dead = False
        self.rejecting = False
        self.Documents = FakeDocuments(self)

    @property
    def ActiveDocument(self):
        return self.docs[0]


class FakeConnector:
    def __init__(self, apps=None):
        self.apps = [FakeApp()] if apps is None else apps
        self.attach_count = 0
        self.thread_names = []
        self.others = []

    def thread_init(self):
        self.thread_names.append(threading.current_thread().name)

    def thread_exit(self):
        pass

    def pump(self):
        pass

    def attach(self):
        self.attach_count += 1
        for app in self.apps:
            if not app.dead:
                return app
        raise VisioError(
            "visio_not_running",
            "Microsoft Visio is not running",
            hint="Open Visio, or call visio_new_drawing or visio_open_drawing, which start it.",
        )

    def other_documents(self):
        return list(self.others)

    def rejecting(self, app):
        return app.rejecting


class FakeClock:
    def __init__(self):
        self.now = 0.0

    def __call__(self):
        return self.now

    def sleep(self, seconds):
        self.now += seconds
