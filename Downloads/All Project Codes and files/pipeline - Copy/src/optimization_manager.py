import os
import time
import logging
import threading
from typing import Dict, Any, Optional, List, Tuple

log = logging.getLogger("GST-Amendment.Optimization-Manager")

# ── UPLOAD TRACKER SERVICE ───────────────────────────────────────────────────

class UploadTrackerService:
    """
    Tracks dynamic document uploads by injecting monkeypatches into the page context
    to intercept upload API responses, completely bypassing blind sleeps.
    """
    JS_INTERCEPTOR = """
    (function() {
        if (window._xhrIntercepted) return;
        window._xhrIntercepted = true;
        window._lastUploadResponse = null;
        
        // Intercept XMLHttpRequest
        var open = XMLHttpRequest.prototype.open;
        XMLHttpRequest.prototype.open = function(method, url) {
            this._url = url;
            return open.apply(this, arguments);
        };
        var send = XMLHttpRequest.prototype.send;
        XMLHttpRequest.prototype.send = function() {
            var self = this;
            var onreadystatechange = this.onreadystatechange;
            this.onreadystatechange = function() {
                if (self.readyState === 4) {
                    if (self._url && self._url.indexOf('/documenthb/upload') !== -1) {
                        try {
                            window._lastUploadResponse = JSON.parse(self.responseText);
                        } catch(e) {}
                    }
                }
                if (onreadystatechange) {
                    return onreadystatechange.apply(self, arguments);
                }
            };
            return send.apply(this, arguments);
        };

        // Intercept Fetch
        var originalFetch = window.fetch;
        window.fetch = function(input, init) {
            var url = typeof input === 'string' ? input : (input ? input.url : '');
            return originalFetch.apply(this, arguments).then(function(response) {
                if (url && url.indexOf('/documenthb/upload') !== -1) {
                    response.clone().json().then(function(data) {
                        window._lastUploadResponse = data;
                    }).catch(function() {});
                }
                return response;
            });
        };
    })();
    """

    @classmethod
    def inject_interceptor(cls, driver) -> None:
        """Injects network interception monkeypatches in the page context."""
        try:
            if hasattr(driver, "execute_script"):
                driver.execute_script(cls.JS_INTERCEPTOR)
                log.info("[UPLOAD-TRACKER] Injected network interceptor into browser session.")
            elif hasattr(driver, "execute_js"):
                driver.execute_js(cls.JS_INTERCEPTOR)
                log.info("[UPLOAD-TRACKER] Injected network interceptor into browser session.")
        except Exception as e:
            log.warning(f"[UPLOAD-TRACKER] Interceptor injection failed: {e}")

    @classmethod
    def wait_for_upload(cls, driver, timeout: int = 30) -> Tuple[bool, Optional[Dict[str, Any]]]:
        """Polls intercepted network response and DOM indicators to detect upload completion."""
        log.info("[UPLOAD-TRACKER] Tracking upload status event-driven...")
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                # Query intercepted response
                res = None
                if hasattr(driver, "execute_script"):
                    res = driver.execute_script("return window._lastUploadResponse;")
                elif hasattr(driver, "execute_js"):
                    res = driver.execute_js("return window._lastUploadResponse;")
                
                if res:
                    status = res.get("status")
                    upload_id = res.get("id") or res.get("uploadId")
                    log.info(f"[UPLOAD-TRACKER] API Response Intercepted: Status={status}, UploadID={upload_id}")
                    if status == "SUCCESS" or status == "COMPLETE" or not status: # Empty status or standard success
                        log.info(f"[UPLOAD-TRACKER] Upload completed successfully (elapsed: {time.time()-start_time:.1f}s)")
                        return True, res
                    elif status == "ERROR" or status == "FAILED":
                        log.error(f"[UPLOAD-TRACKER] Upload reported failure by API: {res}")
                        return False, res
                
                # Fallback DOM check for the trash / delete icon indicating a completed document upload
                trash_exists = False
                delete_xpath = "//a[contains(@class,'fa-trash') or contains(@ng-click,'delete')] | //button[contains(.,'Delete') or contains(@class,'btn-danger') or contains(@ng-click,'delete')]"
                if hasattr(driver, "find_elements"):
                    trash_exists = len(driver.find_elements("xpath", delete_xpath)) > 0
                elif hasattr(driver, "find_xpath_elements"):
                    trash_exists = len(driver.find_xpath_elements(delete_xpath)) > 0

                if trash_exists:
                    log.info(f"[UPLOAD-TRACKER] Trash/Delete element found in DOM. Confirming completion.")
                    return True, res

            except Exception as e:
                log.warning(f"[UPLOAD-TRACKER] Status polling error: {e}")

            time.sleep(0.5)

        log.warning("[UPLOAD-TRACKER] Upload tracking timed out.")
        return False, None

# ── CACHING & REDUCER LAYERS ──────────────────────────────────────────────────

class RuntimeCacheManager:
    """Thread-safe runtime memory-only cache mapping processing states."""
    _lock = threading.Lock()
    _cache: Dict[str, Any] = {}

    @classmethod
    def get(cls, key: str, default: Any = None) -> Any:
        with cls._lock:
            return cls._cache.get(key, default)

    @classmethod
    def set(cls, key: str, value: Any) -> None:
        with cls._lock:
            cls._cache[key] = value

    @classmethod
    def delete(cls, key: str) -> None:
        with cls._lock:
            cls._cache.pop(key, None)

    @classmethod
    def clear(cls) -> None:
        with cls._lock:
            cls._cache.clear()
            log.info("[CACHE] Runtime cache cleared.")

class SessionCacheManager:
    """Thread-safe persistent session cache for metadata parameters."""
    _lock = threading.Lock()
    _session_cache: Dict[str, Any] = {}

    @classmethod
    def get(cls, key: str, default: Any = None) -> Any:
        with cls._lock:
            return cls._session_cache.get(key, default)

    @classmethod
    def set(cls, key: str, value: Any) -> None:
        with cls._lock:
            cls._session_cache[key] = value

    @classmethod
    def clear(cls) -> None:
        with cls._lock:
            cls._session_cache.clear()
            log.info("[CACHE] Session cache cleared.")

class NavigationCache:
    """Caches browser routes and skips redundant clicks when already loaded."""
    @classmethod
    def verify_and_skip(cls, driver, target_url_fragment: str) -> bool:
        try:
            current_url = driver.current_url.lower()
            if target_url_fragment.lower() in current_url:
                log.info(f"[NAV-CACHE] Route '{target_url_fragment}' matches current URL. Skipping redundant navigation.")
                return True
        except Exception:
            pass
        return False

class DOMReadReducer:
    """Caches DOM lookup results to minimize expensive driver DOM searches."""
    _lock = threading.Lock()
    _reads_cache: Dict[str, Tuple[Any, float]] = {}
    _TTL = 3.0 # Cache DOM values for max 3 seconds

    @classmethod
    def get_cached_element(cls, key: str) -> Optional[Any]:
        with cls._lock:
            val, timestamp = cls._reads_cache.get(key, (None, 0.0))
            if time.time() - timestamp < cls._TTL:
                return val
            return None

    @classmethod
    def cache_element(cls, key: str, value: Any) -> None:
        with cls._lock:
            cls._reads_cache[key] = (value, time.time())

# ── OCR SINGLETON WITH THREAD-SAFE DOUBLE-CHECKED LOCK ───────────────────────

class OCRSingleton:
    """Thread-safe singleton engine to ensure TrOCR loads exactly once and is shared."""
    _instance = None
    _lock = threading.Lock()

    @classmethod
    def get_ocr_engine(cls, model_name: Optional[str] = None):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    # Delay import until needed
                    from transformers import TrOCRProcessor, VisionEncoderDecoderModel
                    import torch
                    
                    device = "cuda" if torch.cuda.is_available() else "cpu"
                    log.info(f"[OCR-SINGLETON] Loading TrOCR engine on {device.upper()} (thread-safe)...")
                    name = model_name or "microsoft/trocr-base-printed"
                    
                    try:
                        processor = TrOCRProcessor.from_pretrained(name)
                        model = VisionEncoderDecoderModel.from_pretrained(name, low_cpu_mem_usage=False)
                        model.to(device)
                        model.eval()
                        
                        cls._instance = {
                            "processor": processor,
                            "model": model,
                            "device": device,
                            "model_name": name
                        }
                        log.info("[OCR-SINGLETON] TrOCR engine loaded successfully.")
                    except Exception as e:
                        log.error(f"[OCR-SINGLETON] Failed to load TrOCR engine: {e}")
                        raise
        return cls._instance

# ── CONNECTION POOL MANAGER ──────────────────────────────────────────────────

class ConnectionPoolManager:
    """Manages global connection pools to keep HTTP sessions alive."""
    _lock = threading.Lock()
    _sessions: Dict[str, Any] = {}

    @classmethod
    def get_session(cls, name: str = "default") -> Any:
        with cls._lock:
            if name not in cls._sessions:
                import requests
                from requests.adapters import HTTPAdapter
                from urllib3.util import Retry
                
                session = requests.Session()
                retry = Retry(total=3, backoff_factor=0.5, status_forcelist=[500, 502, 503, 504])
                adapter = HTTPAdapter(pool_connections=25, pool_maxsize=25, max_retries=retry)
                session.mount("http://", adapter)
                session.mount("https://", adapter)
                
                cls._sessions[name] = session
                log.info(f"[CONNECTION-POOL] Initialized global connection pool for '{name}' (Size: 25)")
            return cls._sessions[name]
