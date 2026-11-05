from visio_mcp.errors import (
    CO_E_OBJNOTCONNECTED,
    RPC_E_CALL_REJECTED,
    RPC_E_DISCONNECTED,
    RPC_E_SERVERCALL_RETRYLATER,
    RPC_S_SERVER_UNAVAILABLE,
    VisioError,
    classify,
)


class ComError(Exception):
    def __init__(self, hresult, strerror="failed", excepinfo=None):
        super().__init__(strerror)
        self.hresult = hresult
        self.strerror = strerror
        self.excepinfo = excepinfo


def test_to_dict_has_the_four_keys():
    err = VisioError("busy", "Visio is busy", retryable=True, hint="wait")
    assert err.to_dict() == {"error": "Visio is busy", "code": "busy", "retryable": True, "hint": "wait"}


def test_to_dict_adds_rolled_back_and_details_only_when_set():
    err = VisioError("ambiguous_shape", "two matches", details=[{"id": 1}])
    err.rolled_back = True
    result = err.to_dict()
    assert result["rolled_back"] is True
    assert result["details"] == [{"id": 1}]


def test_visio_error_passes_through():
    err = VisioError("timeout", "slow")
    assert classify(err) is err


def test_busy_hresults():
    for hresult in (RPC_E_CALL_REJECTED, RPC_E_SERVERCALL_RETRYLATER):
        err = classify(ComError(hresult))
        assert err.code == "busy"
        assert err.retryable is True


def test_disconnect_hresults():
    for hresult in (RPC_E_DISCONNECTED, RPC_S_SERVER_UNAVAILABLE, CO_E_OBJNOTCONNECTED):
        err = classify(ComError(hresult))
        assert err.code == "disconnected"
        assert err.retryable is True


def test_other_com_error_uses_visios_description_without_its_blank_lines():
    exc = ComError(-2147352567, "Exception occurred.", (0, "Visio", "\n\nFile not found.", None, 0, -2032465466))
    err = classify(exc)
    assert err.code == "com_error"
    assert err.retryable is False
    assert err.message.startswith("File not found.")
    assert "0x80020009" in err.message


def test_value_error_is_invalid_argument():
    assert classify(ValueError("bad position")).code == "invalid_argument"


def test_anything_else_is_internal():
    err = classify(KeyError("x"))
    assert err.code == "internal"
    assert "KeyError" in err.message
