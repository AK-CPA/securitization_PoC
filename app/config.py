"""Centralized path configuration."""
import os

# Project root = parent of app/
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# Directories
TEMPLATES_DIR = os.path.join(BASE_DIR, "app", "templates")
STATIC_DIR = os.path.join(BASE_DIR, "app", "static")
DATA_DIR = os.path.join(BASE_DIR, "data")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

# Ensure runtime directories exist
for d in [DATA_DIR, UPLOAD_DIR, OUTPUT_DIR]:
    os.makedirs(d, exist_ok=True)
