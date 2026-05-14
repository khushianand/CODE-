"""Central logging utilities for console and UI log streaming."""

import logging
from datetime import datetime
from typing import Callable, Optional


class UILogHandler(logging.Handler):
    def __init__(self, callback: Callable[[str], None]):
        """Explain workflow and purpose of `__init__` in this module."""
        super().__init__()
        self.callback = callback

    def emit(self, record: logging.LogRecord) -> None:
        """Explain workflow and purpose of `emit` in this module."""
        msg = self.format(record)
        self.callback(msg)


def get_logger(name: str = "vuln_automation", ui_callback: Optional[Callable[[str], None]] = None) -> logging.Logger:
    """Explain workflow and purpose of `get_logger` in this module."""
    logger = logging.getLogger(name)
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%Y-%m-%d %H:%M:%S")
        sh = logging.StreamHandler()
        sh.setFormatter(fmt)
        logger.addHandler(sh)
    if ui_callback:
        uh = UILogHandler(ui_callback)
        uh.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", "%H:%M:%S"))
        logger.addHandler(uh)
    logger.info("Logger initialized at %s", datetime.utcnow().isoformat())
    return logger
