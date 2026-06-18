"""M365 Email Hermes Plugin."""

import sys
from pathlib import Path

# adapter.py lives at project root for Hermes plugin discovery
_plugin_root = Path(__file__).resolve().parents[1]
if str(_plugin_root) not in sys.path:
    sys.path.insert(0, str(_plugin_root))

from adapter import register  # noqa: E402

__version__ = "0.1.0"
__all__ = ["register"]
