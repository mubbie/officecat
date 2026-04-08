"""Shared error reporting for readers."""

from __future__ import annotations

import sys
from pathlib import Path


def _is_cloud_only(path: Path) -> bool:
    """Check if a file is a OneDrive cloud-only placeholder (Windows)."""
    if sys.platform != "win32":
        return False
    try:
        import ctypes

        attrs = ctypes.windll.kernel32.GetFileAttributesW(str(path))
        if attrs == -1:
            return False
        # FILE_ATTRIBUTE_RECALL_ON_DATA_ACCESS or RECALL_ON_OPEN
        return bool(attrs & (0x00400000 | 0x00040000))
    except Exception:
        return False


def report_open_error(path: Path) -> None:
    """Print a helpful error depending on why the file can't be opened."""
    import zipfile

    # Check cloud-only first
    if _is_cloud_only(path):
        print(
            f"Error: '{path.name}' is a cloud-only file on OneDrive "
            f"and has not been downloaded yet.\n"
            f"Right-click the file and choose "
            f"'Always keep on this device', then try again.",
            file=sys.stderr,
        )
        return

    # Try reading raw bytes to distinguish locked vs corrupt
    try:
        with open(path, "rb") as f:
            f.read(4)
    except PermissionError:
        print(
            f"Error: '{path.name}' is locked by another process.\n"
            f"Close the file in any other application "
            f"(e.g. Word, Excel) and try again.",
            file=sys.stderr,
        )
        return
    except OSError as e:
        print(
            f"Error: '{path.name}' could not be read: {e}",
            file=sys.stderr,
        )
        return

    # File is readable but not a valid ZIP/OOXML
    try:
        with zipfile.ZipFile(str(path)) as zf:
            zf.namelist()
    except zipfile.BadZipFile:
        on_onedrive = "onedrive" in str(path).lower()
        if on_onedrive:
            print(
                f"Error: '{path.name}' could not be opened.\n"
                f"This file is on OneDrive and may not be fully "
                f"synced. Try opening it in its application first.",
                file=sys.stderr,
            )
        else:
            print(
                f"Error: '{path.name}' appears to be corrupt "
                f"or is not a valid {path.suffix} file.",
                file=sys.stderr,
            )
        return
    except Exception:
        pass

    # Fallback
    print(
        f"Error: '{path.name}' appears to be corrupt or invalid.",
        file=sys.stderr,
    )
