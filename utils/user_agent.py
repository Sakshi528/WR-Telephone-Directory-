"""Minimal User-Agent string parsing for audit/request-metadata capture.

Werkzeug 3.x removed its bundled User-Agent parser entirely --
request.user_agent.browser/.platform return None unconditionally unless a
custom parser is configured (there is no built-in fallback anymore). This
is a small, good-enough-for-an-audit-trail regex parser rather than adding
a new third-party dependency for something that only needs to be roughly
readable ("Chrome on Windows"), not exact.
"""

import re

_BROWSER_PATTERNS = [
    ("Edge", re.compile(r"Edg(?:e|A|iOS)?/([\d.]+)")),
    ("Opera", re.compile(r"(?:OPR|Opera)/([\d.]+)")),
    ("Chrome", re.compile(r"Chrome/([\d.]+)")),
    ("Firefox", re.compile(r"Firefox/([\d.]+)")),
    ("Safari", re.compile(r"Version/([\d.]+).*Safari")),
    ("Internet Explorer", re.compile(r"MSIE ([\d.]+)")),
    ("Internet Explorer", re.compile(r"rv:([\d.]+)\).*like Gecko$")),
]

_OS_PATTERNS = [
    ("Windows 11/10", re.compile(r"Windows NT 10\.0")),
    ("Windows 8.1", re.compile(r"Windows NT 6\.3")),
    ("Windows 8", re.compile(r"Windows NT 6\.2")),
    ("Windows 7", re.compile(r"Windows NT 6\.1")),
    ("Windows", re.compile(r"Windows NT")),
    ("Android", re.compile(r"Android ([\d.]+)")),
    ("iOS", re.compile(r"i(?:Phone|Pad|Pod).*OS (\d+[_\d]*)")),
    ("macOS", re.compile(r"Mac OS X ([\d_.]+)")),
    ("Linux", re.compile(r"Linux")),
]


def parse_user_agent(ua_string):
    """Returns (browser, operating_system) as short display strings, or
    (None, None) if ua_string is empty. Best-effort only."""
    if not ua_string:
        return None, None

    browser = None
    for name, pattern in _BROWSER_PATTERNS:
        m = pattern.search(ua_string)
        if m:
            version = m.group(1) if m.groups() else ""
            browser = f"{name} {version}".strip()
            break

    operating_system = None
    for name, pattern in _OS_PATTERNS:
        m = pattern.search(ua_string)
        if m:
            operating_system = name
            break

    return browser, operating_system
