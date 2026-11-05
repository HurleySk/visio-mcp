import subprocess
import sys


def test_importing_the_server_does_not_import_pywin32_or_pillow():
    code = (
        "import sys, visio_mcp.server as s; s.register_tools(); "
        "bad = [m for m in ('win32com', 'pythoncom', 'win32gui', 'win32ui', 'PIL') if m in sys.modules]; "
        "print(bad); raise SystemExit(1 if bad else 0)"
    )
    proc = subprocess.run([sys.executable, "-c", code], capture_output=True, text=True)
    assert proc.returncode == 0, proc.stdout + proc.stderr
