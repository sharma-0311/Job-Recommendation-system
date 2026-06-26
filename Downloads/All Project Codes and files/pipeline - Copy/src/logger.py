import logging
import sys

# Globally configure standard streams to UTF-8 on Windows to prevent UnicodeEncodeError
for stream in (sys.stdout, sys.stderr):
    if stream is not None and hasattr(stream, "reconfigure"):
        try:
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
        except Exception:
            pass

import time
import uuid
from typing import Dict, Any, Optional

class ProductionLogger:
    """
    Enterprise-grade logging utility that dynamically injects correlation IDs,
    batch IDs, and timing metrics. Supports structured console/file output.
    """
    _correlation_id: Optional[str] = None
    _batch_id: Optional[str] = None
    _registration_id: Optional[str] = None

    def __init__(self, name: str = "GST-Enterprise"):
        self.logger = logging.getLogger(name)
        self.logger.setLevel(logging.INFO)
        
        # Ensure we don't duplicate handlers if class is re-instantiated
        if not self.logger.handlers:
            formatter = logging.Formatter(
                "%(asctime)s [%(levelname)s] [%(correlation_id)s] [%(batch_id)s] [%(registration_id)s] %(message)s",
                datefmt="%H:%M:%S"
            )
            
            # Stream Handler (console)
            sh = logging.StreamHandler(sys.stdout)
            sh.setFormatter(formatter)
            self.logger.addHandler(sh)
            
            # File Handler
            fh = logging.FileHandler("gst_amendment.log", encoding="utf-8")
            fh.setFormatter(formatter)
            self.logger.addHandler(fh)

    @classmethod
    def set_context(cls, correlation_id: Optional[str] = None, batch_id: Optional[str] = None, registration_id: Optional[str] = None):
        if correlation_id:
            cls._correlation_id = correlation_id
        if batch_id:
            cls._batch_id = batch_id
        if registration_id:
            cls._registration_id = registration_id

    @classmethod
    def clear_context(cls):
        cls._correlation_id = None
        cls._batch_id = None
        cls._registration_id = None

    def _get_extra(self) -> Dict[str, str]:
        return {
            "correlation_id": self._correlation_id or "GLOBAL",
            "batch_id": self._batch_id or "SYSTEM",
            "registration_id": self._registration_id or "SYS"
        }

    def info(self, msg: str, *args, **kwargs):
        self.logger.info(msg, *args, extra=self._get_extra(), **kwargs)

    def warning(self, msg: str, *args, **kwargs):
        self.logger.warning(msg, *args, extra=self._get_extra(), **kwargs)

    def error(self, msg: str, *args, **kwargs):
        self.logger.error(msg, *args, extra=self._get_extra(), **kwargs)

    def exception(self, msg: str, *args, **kwargs):
        self.logger.exception(msg, *args, extra=self._get_extra(), **kwargs)

    def debug(self, msg: str, *args, **kwargs):
        self.logger.debug(msg, *args, extra=self._get_extra(), **kwargs)

# Global Logger instance
log = ProductionLogger()

def track_execution_time(action_name: str):
    """
    Decorator that calculates execution timing metrics for any function and logs it.
    """
    def decorator(func):
        def wrapper(*args, **kwargs):
            corr_id = str(uuid.uuid4())[:8]
            ProductionLogger.set_context(correlation_id=corr_id)
            t0 = time.perf_counter()
            log.info(f"[PERF] Starting {action_name}...")
            try:
                result = func(*args, **kwargs)
                duration = time.perf_counter() - t0
                log.info(f"[PERF] Completed {action_name} in {duration:.3f} seconds.")
                return result
            except Exception as e:
                duration = time.perf_counter() - t0
                log.error(f"[PERF] Failed {action_name} after {duration:.3f} seconds: {e}")
                raise
        return wrapper
    return decorator
