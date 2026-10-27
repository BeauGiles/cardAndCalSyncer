"""Put the flat app modules on sys.path and keep test data out of /data."""
import os
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "app"))
os.environ.setdefault("CACS_DATA", tempfile.mkdtemp(prefix="cacs-test-"))
