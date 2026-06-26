import logging
import json
from typing import Dict, Any, Optional

log = logging.getLogger("GST-Amendment.API-Reader")

def _safe_parse_json(val: Any) -> Any:
    """Recursively parses JSON strings if detected, ensuring robust dictionary yields."""
    if not val:
        return val
    if isinstance(val, dict):
        return val
    if isinstance(val, str):
        val_str = val.strip()
        if not val_str:
            return val
        # 1. Try standard JSON load
        try:
            parsed = json.loads(val_str)
            if isinstance(parsed, dict):
                return parsed
            if isinstance(parsed, str):
                parsed2 = json.loads(parsed)
                if isinstance(parsed2, dict):
                    return parsed2
        except Exception:
            pass
            
        # 2. Try ast.literal_eval for Python dictionary string representation fallback
        try:
            import ast
            parsed = ast.literal_eval(val_str)
            if isinstance(parsed, dict):
                return parsed
        except Exception:
            pass
            
        # 3. Clean Angular security prefix if present
        if val_str.startswith(")]}',"):
            try:
                cleaned = val_str[5:].strip()
                parsed = json.loads(cleaned)
                if isinstance(parsed, dict):
                    return parsed
            except Exception:
                pass
    return val

class ContactReaderService:
    @staticmethod
    def fetch_contacts(driver) -> Optional[Dict[str, Any]]:
        """
        Executes an authorized profile contact details request using Angular $http
        or fetch() with explicit XSRF-TOKEN cookie header mapping.
        """
        log.info("[API-READ] Fetching contacts via profile contact API (handshake)...")
        js_code = """
            var callback = arguments[arguments.length - 1];

            function getCookie(name) {
                var value = "; " + document.cookie;
                var parts = value.split("; " + name + "=");
                if (parts.length == 2) return decodeURIComponent(parts.pop().split(";").shift());
                return '';
            }

            // Tier 1: Try the active page's AngularJS $http client to reuse configured CSRF interceptors
            var el = document.querySelector('.ng-scope') || document.querySelector('[ng-app]') || document.body;
            if (window.angular && el) {
                try {
                    var $injector = angular.element(el).injector();
                    if ($injector) {
                        var $http = $injector.get('$http');
                        $http({
                            method: 'POST',
                            url: '/services/auth/profile/contacts',
                            headers: { 'Accept': 'application/json' },
                            data: {}
                        }).then(function(res) {
                            callback(JSON.stringify({success: true, data: res.data}));
                        }).catch(function(err) {
                            runFetch();
                        });
                        return;
                    }
                } catch(e) {
                    // Fall through to manual fetch
                }
            }

            runFetch();

            // Tier 2: Fallback to manual window.fetch() with manual XSRF-TOKEN cookie extraction
            function runFetch() {
                var xsrfToken = getCookie('XSRF-TOKEN');
                fetch('/services/auth/profile/contacts', {
                    method: 'POST',
                    headers: {
                        'Content-Type': 'application/json',
                        'Accept': 'application/json',
                        'X-XSRF-TOKEN': xsrfToken
                    },
                    body: '{}'
                })
                .then(function(res) {
                    if (!res.ok) throw new Error('HTTP Status ' + res.status);
                    return res.json();
                })
                .then(function(data) {
                    callback(JSON.stringify({success: true, data: data}));
                })
                .catch(function(err) {
                    callback(JSON.stringify({success: false, error: err.toString()}));
                });
            }
        """
        try:
            if hasattr(driver, "execute_async_script"):
                res = driver.execute_async_script(js_code)
            elif hasattr(driver, "execute_async_js"):
                res = driver.execute_async_js(js_code)
            else:
                raise ValueError("Unsupported driver type")

            res = _safe_parse_json(res)
            if res:
                # If the response directly contains our targeted keys, return it safely
                if isinstance(res, dict) and ("authorisedSignatoryDetl" in res or "authorizedSignatory" in res):
                    log.info("[API-READ] Contact details fetched successfully.")
                    return res
                elif isinstance(res, str) and ("authorisedSignatoryDetl" in res or "authorizedSignatory" in res):
                    try:
                        parsed = json.loads(res)
                        if isinstance(parsed, dict):
                            log.info("[API-READ] Contact details fetched successfully (parsed).")
                            return parsed
                    except Exception:
                        pass
                
                # Fallback check if it was packed in a success/data structure by interceptors
                if isinstance(res, dict) and res.get("success") and res.get("data"):
                    log.info("[API-READ] Contact details fetched successfully via data wrapper.")
                    return _safe_parse_json(res.get("data"))
            else:
                log.warning(f"[API-READ] Profile contact API failed: {res.get('error') if isinstance(res, dict) else 'No response'}")
        except Exception as e:
            log.error(f"[API-READ] Exception during contact API fetch: {e}")
        return None

class DraftReaderService:
    @staticmethod
    def fetch_draft_core(driver) -> Optional[Dict[str, Any]]:
        """
        Retrieves active core registration drafts using direct page context request.
        """
        log.info("[API-READ] Fetching core registration draft via API (handshake)...")
        js_code = """
            var callback = arguments[arguments.length - 1];

            function getCookie(name) {
                var value = "; " + document.cookie;
                var parts = value.split("; " + name + "=");
                if (parts.length == 2) return decodeURIComponent(parts.pop().split(";").shift());
                return '';
            }

            var el = document.querySelector('.ng-scope') || document.querySelector('[ng-app]') || document.body;
            if (window.angular && el) {
                try {
                    var $injector = angular.element(el).injector();
                    if ($injector) {
                        var $http = $injector.get('$http');
                        $http({
                            method: 'GET',
                            url: '/registration/auth/api/getRegDraftAmndCore',
                            headers: { 'Accept': 'application/json' }
                        }).then(function(res) {
                            callback({success: true, data: res.data});
                        }).catch(function(err) {
                            runFetch();
                        });
                        return;
                    }
                } catch(e) {
                    // Fall through
                }
            }

            runFetch();

            function runFetch() {
                var xsrfToken = getCookie('XSRF-TOKEN');
                fetch('/registration/auth/api/getRegDraftAmndCore', {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json',
                        'X-XSRF-TOKEN': xsrfToken
                    }
                })
                .then(function(res) {
                    if (!res.ok) throw new Error('HTTP Status ' + res.status);
                    return res.json();
                })
                .then(function(data) {
                    callback({success: true, data: data});
                })
                .catch(function(err) {
                    callback({success: false, error: err.toString()});
                });
            }
        """
        try:
            if hasattr(driver, "execute_async_script"):
                res = driver.execute_async_script(js_code)
            elif hasattr(driver, "execute_async_js"):
                res = driver.execute_async_js(js_code)
            else:
                raise ValueError("Unsupported driver type")

            res = _safe_parse_json(res)
            if res and res.get("success"):
                log.info("[API-READ] Core draft data retrieved successfully.")
                return _safe_parse_json(res.get("data"))
            else:
                log.warning(f"[API-READ] Core draft API failed: {res.get('error') if res else 'No response'}")
        except Exception as e:
            log.error(f"[API-READ] Exception during core draft API fetch: {e}")
        return None

class APOBReaderService:
    @staticmethod
    def fetch_apob_details(driver) -> Optional[Dict[str, Any]]:
        """
        Retrieves active draft APoB amendment details.
        """
        log.info("[API-READ] Fetching APOB details via API (handshake)...")
        js_code = """
            var callback = arguments[arguments.length - 1];

            function getCookie(name) {
                var value = "; " + document.cookie;
                var parts = value.split("; " + name + "=");
                if (parts.length == 2) return decodeURIComponent(parts.pop().split(";").shift());
                return '';
            }

            var el = document.querySelector('.ng-scope') || document.querySelector('[ng-app]') || document.body;
            if (window.angular && el) {
                try {
                    var $injector = angular.element(el).injector();
                    if ($injector) {
                        var $http = $injector.get('$http');
                        $http({
                            method: 'GET',
                            url: '/registration/auth/api/apob/apobamndcoredetails',
                            headers: { 'Accept': 'application/json' }
                        }).then(function(res) {
                            callback({success: true, data: res.data});
                        }).catch(function(err) {
                            runFetch();
                        });
                        return;
                    }
                } catch(e) {
                    // Fall through
                }
            }

            runFetch();

            function runFetch() {
                var xsrfToken = getCookie('XSRF-TOKEN');
                fetch('/registration/auth/api/apob/apobamndcoredetails', {
                    method: 'GET',
                    headers: {
                        'Accept': 'application/json',
                        'X-XSRF-TOKEN': xsrfToken
                    }
                })
                .then(function(res) {
                    if (!res.ok) throw new Error('HTTP Status ' + res.status);
                    return res.json();
                })
                .then(function(data) {
                    callback({success: true, data: data});
                })
                .catch(function(err) {
                    callback({success: false, error: err.toString()});
                });
            }
        """
        try:
            if hasattr(driver, "execute_async_script"):
                res = driver.execute_async_script(js_code)
            elif hasattr(driver, "execute_async_js"):
                res = driver.execute_async_js(js_code)
            else:
                raise ValueError("Unsupported driver type")

            res = _safe_parse_json(res)
            if res and res.get("success"):
                log.info("[API-READ] APOB details retrieved successfully.")
                return _safe_parse_json(res.get("data"))
            else:
                log.warning(f"[API-READ] APOB details API failed: {res.get('error') if res else 'No response'}")
        except Exception as e:
            log.error(f"[API-READ] Exception during APOB details API fetch: {e}")
        return None

class DraftIdCache:
    """In-memory cache mapping GSTIN to draft ID."""
    _cache: Dict[str, str] = {}

    @classmethod
    def get(cls, gstin: str) -> Optional[str]:
        return cls._cache.get(str(gstin).strip().upper())

    @classmethod
    def set(cls, gstin: str, draft_id: str) -> None:
        cls._cache[str(gstin).strip().upper()] = str(draft_id).strip()

    @classmethod
    def clear(cls) -> None:
        cls._cache.clear()
