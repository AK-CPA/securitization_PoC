"""Centralized path configuration."""
import os

# Project root = parent of app/
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Directories (static, read-only)
TEMPLATES_DIR = os.path.join(BASE_DIR, "app", "templates")
STATIC_DIR = os.path.join(BASE_DIR, "app", "static")

# Writable directories — use /tmp on cloud platforms where project dir may be read-only
_writable_base = os.environ.get("RENDER", None)
if _writable_base is not None:
    _storage = "/tmp/odrt"
else:
    _storage = BASE_DIR

DATA_DIR = os.path.join(_storage, "data")
UPLOAD_DIR = os.path.join(_storage, "uploads")
OUTPUT_DIR = os.path.join(_storage, "outputs")

# Ensure runtime directories exist
for d in [DATA_DIR, UPLOAD_DIR, OUTPUT_DIR]:
    os.makedirs(d, exist_ok=True)
