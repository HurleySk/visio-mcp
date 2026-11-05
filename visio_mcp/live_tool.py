import asyncio
import functools
import json
import logging
import os

from visio_mcp.com_runtime import classify_live, error_json, run_com
from visio_mcp.errors import VisioError

logger = logging.getLogger(__name__)
READ_ATTEMPTS = 3


def _default_timeout():
    try:
        value = float(os.environ.get("VISIO_MCP_TIMEOUT", "60"))
    except ValueError:
        return 60.0
    return value if value > 0 else 60.0


def _label(body):
    return body.__name__.removeprefix("visio_").replace("_", " ")


def _with_warnings(result, warnings):
    if not warnings:
        return result
    if isinstance(result, str):
        try:
            payload = json.loads(result)
        except ValueError:
            return result
        if isinstance(payload, dict):
            payload["warnings"] = warnings
            return json.dumps(payload, ensure_ascii=False)
        return result
    if isinstance(result, list) and result and isinstance(result[0], dict):
        result[0]["warnings"] = warnings
    return result


def _call(session, body, args, kwargs):
    session.begin_call(_label(body))
    try:
        result = body(*args, **kwargs)
    except Exception as exc:
        err = classify_live(exc)
        err.rolled_back = session.rollback_scope()
        return _with_warnings(error_json(err), session.warnings)
    if session.abandoned:
        # The caller was already told this edit timed out.
        if session.rollback_scope() is False:
            logger.error("a timed-out edit could not be rolled back")
        return result
    try:
        session.commit_scope()
    except Exception as exc:
        return _with_warnings(error_json(exc), session.warnings)
    return _with_warnings(result, session.warnings)


def _mutating_busy(rolled_back):
    if rolled_back:
        return VisioError(
            "busy",
            "Visio became busy, so the edit was rolled back",
            retryable=True,
            hint="Close any dialog open in Visio, then retry.",
        )
    return VisioError(
        "busy",
        "Visio became busy part-way through the edit",
        retryable=False,
        hint="Check the drawing, or call visio_undo, before retrying.",
    )


def _invoke(session, body, args, kwargs, mutates):
    attempts = 1 if mutates else READ_ATTEMPTS
    for attempt in range(attempts):
        if attempt:
            session.pause(0.5 * attempt)
            session.wait_until_ready()
        result = _call(session, body, args, kwargs)
        err = session.last_error
        if err is None or err.code != "busy":
            return result
    if mutates:
        replacement = _mutating_busy(err.rolled_back)
        replacement.rolled_back = err.rolled_back
        return json.dumps(replacement.to_dict(), ensure_ascii=False)
    return result


def offline_tool(body):
    """Run a tool body that doesn't need Visio on a thread pool, away from the COM thread."""

    @functools.wraps(body)
    async def tool(*args, **kwargs):
        try:
            return await asyncio.to_thread(body, *args, **kwargs)
        except Exception as exc:
            return error_json(exc)

    tool.sync = body
    tool.mutates = False
    tool.start = None
    return tool


def live_tool(*, mutates, start="never", timeout=None):
    """Run a tool body on the COM thread.

    start: "never" fails if Visio isn't running. "if_missing" starts it. "report" runs
    the body anyway, with session.attach_error saying why there is no Visio.

    A mutating body gets one undo step by calling edit_doc after validating: committed
    when the body returns, rolled back when it raises or has timed out.
    """

    def decorate(body):
        @functools.wraps(body)
        async def tool(*args, **kwargs):
            limit = _default_timeout() if timeout is None else timeout
            try:
                return await run_com(
                    lambda session: _invoke(session, body, args, kwargs, mutates),
                    timeout=limit,
                    mutates=mutates,
                    start=start,
                )
            except Exception as exc:
                return error_json(exc)

        tool.sync = body
        tool.mutates = mutates
        tool.start = start
        return tool

    return decorate
