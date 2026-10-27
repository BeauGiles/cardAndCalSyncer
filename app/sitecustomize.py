"""Imported by Python for every process in the image, including ``vdirsyncer``.

vdirsyncer has no plugin hook, so the storage types have to be registered
before the CLI builds its storage index. This file is copied into
site-packages by the image build. Failures are swallowed: a Graph import
problem must not take down the healthcheck or iCloud/Google sync.
"""
import sys
from pathlib import Path

for candidate in (Path("/app"), Path(__file__).resolve().parent):
    if (candidate / "graphsetup.py").exists():
        entry = str(candidate)
        if entry not in sys.path:
            sys.path.insert(0, entry)
        break

try:
    import graphsetup

    graphsetup.install()
except Exception:
    pass
