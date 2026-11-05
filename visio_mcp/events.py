"""Turn Visio's automation events off around one call, for handlers that would open a dialog.

Visio saves this setting when it exits and reads it back at its next start. A server lost while
events are off would leave the user's Visio without them from then on, with nothing to show for
it. So a marker file records that a guard is open, and the next call puts the setting back.
"""
import contextlib
import logging
import os
import tempfile
import time

logger = logging.getLogger(__name__)

EVENTS_OFF = 0
EVENTS_ON = -1
# A younger marker may belong to another server, part-way through its own guard.
MARKER_STALE_AFTER = 120.0


def _marker():
    folder = os.environ.get("LOCALAPPDATA") or tempfile.gettempdir()
    return os.path.join(folder, "visio-mcp", "events-off")


@contextlib.contextmanager
def events_off(app):
    previous = app.EventsEnabled
    if previous == EVENTS_OFF:
        # This user's own setting, so there is nothing to put back.
        yield
        return
    marker = _marker()
    os.makedirs(os.path.dirname(marker), exist_ok=True)
    with open(marker, "w", encoding="utf-8"):
        pass
    app.EventsEnabled = EVENTS_OFF
    try:
        yield
    finally:
        try:
            app.EventsEnabled = previous
            os.remove(marker)
        except Exception:
            # The marker stays, so a later call tries again. Raising would hide the real result.
            logger.warning("could not turn Visio's events back on yet", exc_info=True)


def restore_if_lost(app):
    """Put events back on if a server was lost inside a guard. Returns whether it did."""
    marker = _marker()
    try:
        age = time.time() - os.path.getmtime(marker)
    except OSError:
        return False
    if age < MARKER_STALE_AFTER:
        return False
    app.EventsEnabled = EVENTS_ON
    os.remove(marker)
    return True
