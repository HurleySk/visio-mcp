import json
import subprocess
import threading
import time

INIT = {
    "jsonrpc": "2.0",
    "id": 1,
    "method": "initialize",
    "params": {
        "protocolVersion": "2025-06-18",
        "capabilities": {},
        "clientInfo": {"name": "handshake", "version": "0"},
    },
}
INITIALIZED = {"jsonrpc": "2.0", "method": "notifications/initialized"}
LIST = {"jsonrpc": "2.0", "id": 2, "method": "tools/list"}


def live_tools(tools):
    return [t for t in tools if t["name"].startswith("visio_")]


def handshake(command, timeout=60.0):
    started = time.perf_counter()
    proc = subprocess.Popen(
        command,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    responses = {}
    done = threading.Event()

    def read():
        try:
            for raw in proc.stdout:
                raw = raw.strip()
                if not raw:
                    continue
                message = json.loads(raw)
                if "id" in message:
                    responses[message["id"]] = message
                if 2 in responses:
                    return
        finally:
            done.set()

    threading.Thread(target=read, daemon=True).start()
    try:
        for message in (INIT, INITIALIZED, LIST):
            proc.stdin.write((json.dumps(message) + "\n").encode())
            proc.stdin.flush()
    except OSError:
        pass
    done.wait(timeout)
    answered = 2 in responses
    elapsed = time.perf_counter() - started
    if not answered:
        proc.kill()
    try:
        proc.stdin.close()
    except OSError:
        pass
    try:
        exit_code = proc.wait(15)
    except subprocess.TimeoutExpired:
        proc.kill()
        exit_code = None
    stderr = proc.stderr.read().decode(errors="replace")
    if not answered:
        raise RuntimeError(f"no tools/list response within {timeout}s\n{stderr}")
    return {
        "tools": responses[2]["result"]["tools"],
        "elapsed": elapsed,
        "exit_code": exit_code,
        "stderr": stderr,
    }
