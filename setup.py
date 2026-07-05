import sys

from setuptools import setup
from setuptools.dist import Distribution

if sys.platform != "win32":
    sys.exit(
        "outlook-mcp-com only supports Windows: it drives classic Outlook "
        "desktop via the Win32 COM API (pywin32), which is unavailable on "
        "this platform."
    )


class BinaryDistribution(Distribution):
    """Report as platform-specific so wheels get a win_amd64 tag instead of
    py3-none-any, preventing pip from resolving them on other platforms."""

    def has_ext_modules(self):
        return True


setup(distclass=BinaryDistribution)
