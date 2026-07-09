"""
Shared logging setup for import/validation scripts.

Usage:
    from utils.logger import get_logger
    logger = get_logger(__name__)
    logger.info("...")
"""

import logging
import os
import sys

LOG_DIR = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "logs")


def get_logger(name):
    """Return a logger that writes to both the console and a log file
    under logs/<script name>.log. Safe to call repeatedly with the same
    name (won't duplicate handlers).

    When called with __name__ from a script run directly (not imported),
    name is the literal string "__main__" for every such script, which
    would otherwise make them all collide into logs/__main__.log. In that
    case the log filename falls back to the invoked script's own filename
    (sys.argv[0]) so each script still gets its own log file.
    """
    logger = logging.getLogger(name)
    if logger.handlers:
        return logger

    logger.setLevel(logging.INFO)

    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    os.makedirs(LOG_DIR, exist_ok=True)
    if name == "__main__":
        base = os.path.splitext(os.path.basename(sys.argv[0]))[0] or "main"
    else:
        base = name.split(".")[0]
    log_file = os.path.join(LOG_DIR, f"{base}.log")
    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    logger.propagate = False
    return logger
