"""M365 Email Hermes Plugin."""

import sys
from pathlib import Path

_plugin_root = str(Path(__file__).resolve().parent)
if _plugin_root not in sys.path:
    sys.path.insert(0, _plugin_root)

__version__ = "0.1.0"


def register(ctx):
    """Lazy proxy to adapter.register -- only loads adapter when called."""
    from adapter import register as _register

    return _register(ctx)


__all__ = ["register"]
