import time
import threading
from typing import Dict, List, Optional
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from src.logger import log

class DriverPoolManager:
    """
    Enterprise WebDriver Pool Manager.
    Features:
    - Driver reuse & session affinity
    - Automatic health checks (recreates crashed drivers)
    - Prevents dangling Chrome instances
    """
    _instance = None
    _lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(DriverPoolManager, cls).__new__(cls)
                cls._instance._initialized = False
            return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self.drivers: Dict[str, webdriver.Chrome] = {} # username -> driver
        self.lock = threading.Lock()
        self._initialized = True

    def check_driver_health(self, driver: webdriver.Chrome) -> bool:
        """
        Validates if the driver is still active, alive, and responsive.
        """
        try:
            # Execute a simple lightweight script to test browser responsiveness
            driver.execute_script("return 1;")
            return True
        except Exception:
            return False

    def get_driver(self, username: str, headless: bool = False, build_fn = None) -> webdriver.Chrome:
        """
        Retrieves an active healthy webdriver from the pool with session affinity.
        Creates a fresh driver if none exists or if the existing one is unhealthy.
        """
        with self.lock:
            driver = self.drivers.get(username)
            if driver:
                log.info(f"[POOL] Found existing driver instance in pool for user '{username}'. Testing health...")
                if self.check_driver_health(driver):
                    log.info(f"[POOL] Driver for '{username}' is healthy. Reusing instance.")
                    return driver
                else:
                    log.warning(f"[POOL] Driver for '{username}' is unresponsive. Terminating and recreating...")
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    self.drivers.pop(username, None)
            
            # Construct a new Chrome webdriver
            log.info(f"[POOL] Launching a fresh Chrome WebDriver instance for user '{username}'...")
            if build_fn:
                new_driver = build_fn(headless=headless)
            else:
                # Fallback simple build if no build function is supplied
                options = Options()
                if headless:
                    options.add_argument("--headless=new")
                options.add_argument("--no-sandbox")
                options.add_argument("--disable-dev-shm-usage")
                options.add_argument("--window-size=1920,1080")
                service = Service(ChromeDriverManager().install())
                new_driver = webdriver.Chrome(service=service, options=options)
                new_driver.maximize_window()
                
            new_driver.current_username = username
            self.drivers[username] = new_driver
            return new_driver

    def release_driver(self, username: str) -> None:
        """
        Placeholder to cleanly release driver back.
        Since we keep session affinity, we keep it inside self.drivers.
        """
        pass

    def close_all(self) -> None:
        """
        Forcefully terminates all active WebDrivers in the pool during lifecycle shutdown.
        """
        with self.lock:
            for username, driver in list(self.drivers.items()):
                log.info(f"[POOL] Cleaning up active pool driver for '{username}'...")
                try:
                    driver.quit()
                except Exception as e:
                    log.debug(f"[POOL] Error closing driver for '{username}': {e}")
            self.drivers.clear()
