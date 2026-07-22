"""Baseline scaffold test to verify the package imports cleanly."""


def test_package_version():
    """Package exposes a non-empty, dotted __version__ string."""
    import __init__ as pkg_init
    version = pkg_init.__version__
    assert isinstance(version, str)
    assert version.count(".") >= 1
    assert all(part for part in version.split("."))


def test_core_imports():
    """Core dependencies are available at runtime."""
    import httpx  # noqa: F401
    import dotenv  # noqa: F401
    from bs4 import BeautifulSoup  # noqa: F401
