import os
import json
import time
from typing import Optional
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from src.logger import log

class GSTSessionManager:
    """
    Manages session cookie persistence, validation, and auto recovery
    to avoid redundant logins and solve less CAPTCHAs.
    """
    def __init__(self, session_dir: str = "session_vault"):
        self.session_dir = session_dir
        os.makedirs(self.session_dir, exist_ok=True)

    def _get_cookie_path(self, username: str) -> str:
        safe_username = "".join(c for c in username if c.isalnum() or c in ("_", "-"))
        return os.path.join(self.session_dir, f"cookies_{safe_username}.json")

    def save_session(self, driver: webdriver.Chrome, username: str) -> bool:
        """
        Saves the current driver cookies to local file storage.
        """
        try:
            cookies = driver.get_cookies()
            cookie_path = self._get_cookie_path(username)
            with open(cookie_path, "w", encoding="utf-8") as f:
                json.dump(cookies, f, indent=4)
            log.info(f"[SESSION] Successfully persisted session cookies for '{username}' to '{cookie_path}'.")
            return True
        except Exception as e:
            log.error(f"[SESSION] Failed to save cookies for '{username}': {e}")
            return False

    def load_session(self, driver: webdriver.Chrome, username: str) -> bool:
        """
        Loads local saved session cookies into the webdriver.
        """
        cookie_path = self._get_cookie_path(username)
        if not os.path.exists(cookie_path):
            log.info(f"[SESSION] No saved cookies found for '{username}'.")
            return False
        
        try:
            # Navigate to the domain first to set context, otherwise selenium rejects cookies
            driver.get("https://www.gst.gov.in/")
            time.sleep(1.5)
            
            with open(cookie_path, "r", encoding="utf-8") as f:
                cookies = json.load(f)
            
            for cookie in cookies:
                # Handle expiry discrepancies
                if "expiry" in cookie:
                    cookie["expiry"] = int(cookie["expiry"])
                try:
                    driver.add_cookie(cookie)
                except Exception as ex:
                    log.debug(f"[SESSION] Ignored invalid cookie entry: {ex}")
            
            log.info(f"[SESSION] Loaded {len(cookies)} cookies into WebDriver context for '{username}'.")
            return True
        except Exception as e:
            log.error(f"[SESSION] Failed to load session cookies for '{username}': {e}")
            return False

    def validate_session(self, driver: webdriver.Chrome) -> bool:
        """
        Verifies if loaded session is still active on the portal
        by redirecting to dashboard and checking for presence of authorized menus.
        """
        try:
            log.info("[SESSION] Validating active session health via Dashboard redirect...")
            driver.get("https://services.gst.gov.in/services/auth/dashboard")
            time.sleep(3.0)  # Allow a bit more time for slow network redirects
            
            # Check current URL and DOM markers
            current_url = driver.current_url.lower()
            if "login" in current_url or "/auth/" not in current_url:
                log.info(f"[SESSION] Stale session: redirected to non-auth URL '{driver.current_url}'.")
                return False
                
            # If log-out elements exist, session is alive (Logout does not exist publicly)
            logout_el = driver.find_elements(By.XPATH, "//a[contains(text(),'Logout') or contains(@href,'logout')]")
            if logout_el:
                log.info("[SESSION] Session validation SUCCESS via Logout indicator. Reusing active dashboard context.")
                return True
                
            # Fallback welcome display check
            welcome_el = driver.find_elements(By.XPATH, "//*[contains(text(),'Welcome') or contains(@class,'usrName')]")
            if welcome_el:
                log.info("[SESSION] Session validation SUCCESS via Welcome display. Reusing active dashboard context.")
                return True
                
            # Additional fallback check for username field
            un_el = driver.find_elements(By.ID, "username")
            if un_el and un_el[0].is_displayed():
                log.info("[SESSION] Stale session: login form is visible.")
                return False

            log.info("[SESSION] Session validation failed (no authenticated logout/welcome indicators found).")
            return False
        except Exception as e:
            log.warning(f"[SESSION] Validation exception: {e}")
            return False

# Maintainability class alias
SessionManager = GSTSessionManager

