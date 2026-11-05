import threading

from visio_mcp.errors import RPC_E_CALL_REJECTED, RPC_E_DISCONNECTED, VisioError


class FakeComError(Exception):
    def __init__(self, hresult, strerror="fake com error"):
        super().__init__(strerror)
        self.hresult = hresult
        self.strerror = strerror
        self.excepinfo = None


class FakeMaster:
    def __init__(self, name):
        self.NameU = name


class FakeCell:
    Units = 65

    def __init__(self, value=0.0):
        self.ResultIU = value


class FakeShape:
    def __init__(self, id_, name, text="", master=None):
        self.ID = id_
        self.Name = name
        self.NameU = name
        self.Text = text
        self.Master = FakeMaster(master) if master else None
        self.OneD = 0
        self.Type = 3
        self.MemberOfContainers = ()
        self.guid_flags = []

    def CellsU(self, name):
        return FakeCell(float(self.ID))

    def UniqueID(self, flag):
        self.guid_flags.append(flag)
        return ""

    def SectionExists(self, section, flag):
        return False

    def ConnectedShapes(self, flag, category):
        return ()


class FakeShapes:
    def __init__(self, shapes):
        self._shapes = shapes

    @property
    def Count(self):
        return len(self._shapes)

    def __call__(self, index):
        return self._shapes[index - 1]

    def ItemFromID(self, shape_id):
        for shape in self._shapes:
            if shape.ID == shape_id:
                return shape
        raise FakeComError(-2147352567, "Invalid sheet identifier.")


class FakePage:
    def __init__(self, name="Page-1", index=1, shapes=()):
        self.Name = name
        self.NameU = name
        self.Index = index
        self.Shapes = FakeShapes(list(shapes))
        self.PageSheet = FakeShape(0, "ThePage")
        self.exports = []

    def Export(self, path):
        self.exports.append(path)


class FakePages:
    def __init__(self, pages):
        self._pages = pages

    @property
    def Count(self):
        return len(self._pages)

    def __call__(self, index):
        return self._pages[index - 1]


class FakeDoc:
    def __init__(self, name="flow.vsdx", folder="C:\\docs", type_=1, read_only=0, pages=None):
        self.Name = name
        self.FullName = f"{folder}\\{name}" if folder else name
        self.Path = f"{folder}\\" if folder else ""
        self.Saved = True
        self.Type = type_
        self.ReadOnly = read_only
        self.Application = None
        self.ID = id(self)
        self.Pages = FakePages(pages if pages is not None else [FakePage()])
        self.calls = []

    def Save(self):
        self.calls.append(("Save",))
        self.Saved = True

    def SaveAs(self, path):
        self.calls.append(("SaveAs", path))
        self.Saved = True

    def ExportAsFixedFormat(self, *args):
        self.calls.append(("ExportAsFixedFormat", *args))

    def Close(self):
        self.calls.append(("Close", self.Saved))
        self.Application.docs.remove(self)


class FakeDocuments:
    def __init__(self, app):
        self._app = app
        self.added = []
        self.events_while_adding = []
        self.opened = []

    def AddEx(self, path, measure, flags):
        self.added.append((path, measure, flags))
        self.events_while_adding.append(self._app.EventsEnabled)
        document = FakeDoc("Drawing1", folder="")
        document.Application = self._app
        self._app.docs.append(document)
        return document

    def OpenEx(self, path, flags):
        self.opened.append((path, flags))
        document = FakeDoc(path.rsplit("\\", 1)[-1], folder=path.rsplit("\\", 1)[0])
        document.Application = self._app
        self._app.docs.append(document)
        return document

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
        for doc in self.docs:
            doc.Application = self
        self.busy_calls = 0
        self.dead = False
        self.rejecting = False
        self.EventsEnabled = -1
        self.ActiveWindow = None
        self.Documents = FakeDocuments(self)
        self.CurrentScope = -1
        self.scopes_begun = []
        self.scopes_ended = []
        self.end_scope_error = None
        self._next_scope = 100

    @property
    def ActiveDocument(self):
        return self.docs[0]

    def BeginUndoScope(self, label):
        self._next_scope += 1
        self.scopes_begun.append(label)
        self.CurrentScope = self._next_scope
        return self._next_scope

    def EndUndoScope(self, scope, commit):
        if self.end_scope_error is not None:
            raise self.end_scope_error
        self.scopes_ended.append((scope, commit))
        self.CurrentScope = -1


class FakeConnector:
    def __init__(self, apps=None, unreachable=False, boot_attempts=0):
        self.apps = [FakeApp()] if apps is None else apps
        self.unreachable = unreachable
        self.boot_attempts = boot_attempts
        self.attach_count = 0
        self.launch_count = 0
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
        if self.launch_count and not self.apps:
            if self.boot_attempts > 0:
                self.boot_attempts -= 1
            else:
                self.apps.append(FakeApp(docs=[]))
        for app in self.apps:
            if not app.dead:
                return app
        if self.unreachable:
            raise VisioError("visio_unreachable", "Visio is running but cannot be reached", retryable=True)
        raise VisioError(
            "visio_not_running",
            "Microsoft Visio is not running",
            hint="Open Visio, or call visio_new_drawing or visio_open_drawing, which start it.",
        )

    def launch(self):
        self.launch_count += 1

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
