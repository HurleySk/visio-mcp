RPC_E_CALL_REJECTED = -2147418111
RPC_E_SERVERCALL_RETRYLATER = -2147417846
RPC_E_DISCONNECTED = -2147417848
RPC_S_SERVER_UNAVAILABLE = -2147023174
RPC_S_CALL_FAILED = -2147023170
CO_E_OBJNOTCONNECTED = -2147220995
MK_E_UNAVAILABLE = -2147221021

BUSY_HRESULTS = frozenset({RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER})
DISCONNECT_HRESULTS = frozenset(
    {RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, RPC_S_CALL_FAILED, CO_E_OBJNOTCONNECTED}
)


class VisioError(Exception):
    def __init__(self, code, message, *, retryable=False, hint=None, details=None):
        super().__init__(message)
        self.code = code
        self.message = message
        self.retryable = retryable
        self.hint = hint
        self.details = details
        self.rolled_back = None

    def to_dict(self):
        result = {
            "error": self.message,
            "code": self.code,
            "retryable": self.retryable,
            "hint": self.hint,
        }
        if self.rolled_back is not None:
            result["rolled_back"] = self.rolled_back
        if self.details is not None:
            result["details"] = self.details
        return result


def busy_error():
    return VisioError(
        "busy",
        "Visio rejected the call because it is busy",
        retryable=True,
        hint="Wait for Visio to finish, close any open dialog, then retry.",
    )


def _com_message(exc, hresult):
    info = getattr(exc, "excepinfo", None)
    detail = None
    if info and len(info) > 2:
        detail = info[2]
    detail = (detail or getattr(exc, "strerror", None) or "COM error").strip()
    return f"{detail} (HRESULT 0x{hresult & 0xFFFFFFFF:08X})"


def classify(exc):
    if isinstance(exc, VisioError):
        return exc
    hresult = getattr(exc, "hresult", None)
    if isinstance(hresult, int):
        if hresult in BUSY_HRESULTS:
            return busy_error()
        if hresult in DISCONNECT_HRESULTS:
            return VisioError(
                "disconnected",
                "Lost the connection to Visio",
                retryable=True,
                hint="Visio was closed or restarted. Retry to reconnect.",
            )
        return VisioError("com_error", _com_message(exc, hresult))
    if isinstance(exc, ValueError):
        return VisioError("invalid_argument", str(exc))
    return VisioError("internal", f"{type(exc).__name__}: {exc}")
