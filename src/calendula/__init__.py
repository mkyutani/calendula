"""calendula: an MCP server that reads your Outlook (classic) calendar from WSL."""

from importlib.metadata import PackageNotFoundError, version

try:
    __version__ = version("calendula")
except PackageNotFoundError:  # running from a source tree that is not installed
    __version__ = "0.0.0"
