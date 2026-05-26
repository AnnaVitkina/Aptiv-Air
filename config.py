"""Centralized folder paths — works both locally and in Google Colab via exec(open(...))."""

from pathlib import Path

# ── Colab paths (edit these when running in Colab) ─────────────────────────────
HARDCODED_INPUT_DIR = "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Air/input"
HARDCODED_PROCESSING_DIR = "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Air/processing"
HARDCODED_OUTPUT_DIR = "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team /Documents/AI Adoption RMT/RMT_APTIV_VERSIGENT/RMT_Air/output"

# ── Runtime detection ──────────────────────────────────────────────────────────
_running_in_colab = False
try:
    import google.colab  # type: ignore[import-untyped]
    _running_in_colab = True
except ImportError:
    pass


def _local_project_root() -> Path:
    return Path(__file__).resolve().parent


def get_input_dir() -> Path:
    if _running_in_colab:
        return Path(HARDCODED_INPUT_DIR)
    return _local_project_root() / "input"


def get_processing_dir() -> Path:
    if _running_in_colab:
        return Path(HARDCODED_PROCESSING_DIR)
    return _local_project_root() / "processing"


def get_output_dir() -> Path:
    if _running_in_colab:
        return Path(HARDCODED_OUTPUT_DIR)
    return _local_project_root() / "output"


INPUT_DIR = get_input_dir()
PROCESSING_DIR = get_processing_dir()
OUTPUT_DIR = get_output_dir()
