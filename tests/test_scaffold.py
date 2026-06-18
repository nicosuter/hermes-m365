"""Baseline scaffold test to verify the package imports cleanly."""


def test_package_version():
    """Package __version__ is set."""
    import __init__ as pkg_init
    assert pkg_init.__version__ == "0.1.0"


def test_core_imports():
    """Core dependencies are available at runtime."""
    import httpx  # noqa: F401
    import dotenv  # noqa: F401
    from bs4 import BeautifulSoup  # noqa: F401
