"""M365 Email Hermes Plugin."""

import sys
from pathlib import Path

_plugin_root = str(Path(__file__).resolve().parent)
if _plugin_root not in sys.path:
    sys.path.insert(0, _plugin_root)

from adapter import register  # noqa: E402

__version__ = "0.1.0"
__all__ = ["register"]
