import logging
from abc import ABC, abstractmethod
from typing import Any, List, Optional

log = logging.getLogger("GST-Amendment.Automation-Adapter")

# ── BASE ABSTRACT CLASS ───────────────────────────────────────────────────────

class PortalAutomationAdapter(ABC):
    """
    Unified automation engine interface for operating GST Portal activities.
    Supports seamless hot-swapping between Selenium and Playwright backends.
    """
    @abstractmethod
    def navigate(self, url: str) -> None:
        pass

    @abstractmethod
    def click(self, by: str, value: str, timeout: int = 15) -> Any:
        pass

    @abstractmethod
    def send_keys(self, by: str, value: str, text: str, timeout: int = 15) -> Any:
        pass

    @abstractmethod
    def execute_script(self, script: str, *args) -> Any:
        pass

    @abstractmethod
    def execute_async_script(self, script: str, *args) -> Any:
        pass

    @abstractmethod
    def get_url(self) -> str:
        pass

    @abstractmethod
    def get_title(self) -> str:
        pass

    @abstractmethod
    def find_element(self, by: str, value: str):
        pass

    @abstractmethod
    def find_elements(self, by: str, value: str) -> List[Any]:
        pass

    @abstractmethod
    def quit(self) -> None:
        pass

    @property
    @abstractmethod
    def raw_driver(self) -> Any:
        pass

# ── SELENIUM IMPLEMENTATION ───────────────────────────────────────────────────

class SeleniumAdapter(PortalAutomationAdapter):
    """Selenium Chrome WebDriver implementation of PortalAutomationAdapter."""
    def __init__(self, driver):
        self.driver = driver

    def navigate(self, url: str) -> None:
        self.driver.get(url)

    def click(self, by: str, value: str, timeout: int = 15) -> Any:
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.by import By
        
        by_map = {"id": By.ID, "xpath": By.XPATH, "css": By.CSS_SELECTOR, "class": By.CLASS_NAME}
        sel_by = by_map.get(by.lower(), By.XPATH)
        
        elem = WebDriverWait(self.driver, timeout).until(EC.element_to_be_clickable((sel_by, value)))
        self.driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", elem)
        try:
            elem.click()
        except Exception:
            self.driver.execute_script("arguments[0].click();", elem)
        return elem

    def send_keys(self, by: str, value: str, text: str, timeout: int = 15) -> Any:
        from selenium.webdriver.support.ui import WebDriverWait
        from selenium.webdriver.support import expected_conditions as EC
        from selenium.webdriver.common.by import By
        
        by_map = {"id": By.ID, "xpath": By.XPATH, "css": By.CSS_SELECTOR, "class": By.CLASS_NAME}
        sel_by = by_map.get(by.lower(), By.XPATH)
        
        elem = WebDriverWait(self.driver, timeout).until(EC.presence_of_element_located((sel_by, value)))
        self.driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        self.driver.execute_script(
            "arguments[0].value = '';"
            "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
            elem,
        )
        elem.send_keys(text)
        return elem

    def execute_script(self, script: str, *args) -> Any:
        return self.driver.execute_script(script, *args)

    def execute_async_script(self, script: str, *args) -> Any:
        return self.driver.execute_async_script(script, *args)

    def get_url(self) -> str:
        return self.driver.current_url

    def get_title(self) -> str:
        return self.driver.title

    def find_element(self, by: str, value: str):
        from selenium.webdriver.common.by import By
        by_map = {"id": By.ID, "xpath": By.XPATH, "css": By.CSS_SELECTOR}
        return self.driver.find_element(by_map.get(by.lower(), By.XPATH), value)

    def find_elements(self, by: str, value: str) -> List[Any]:
        from selenium.webdriver.common.by import By
        by_map = {"id": By.ID, "xpath": By.XPATH, "css": By.CSS_SELECTOR}
        return self.driver.find_elements(by_map.get(by.lower(), By.XPATH), value)

    def quit(self) -> None:
        self.driver.quit()

    @property
    def raw_driver(self) -> Any:
        return self.driver

# ── PLAYWRIGHT IMPLEMENTATION ─────────────────────────────────────────────────

class PlaywrightAdapter(PortalAutomationAdapter):
    """Playwright synchronous implementation of PortalAutomationAdapter."""
    def __init__(self, headless: bool = False):
        try:
            from playwright.sync_api import sync_playwright
        except ImportError:
            log.error("[PLAYWRIGHT-ADAPTER] Playwright package is missing. Fall back to Selenium.")
            raise RuntimeError("Playwright package is not installed. Please install 'playwright' using pip.")
            
        self.playwright = sync_playwright().start()
        self.browser = self.playwright.chromium.launch(
            headless=headless,
            args=[
                "--no-sandbox",
                "--disable-dev-shm-usage",
                "--disable-gpu",
                "--window-size=1920,1080"
            ]
        )
        self.context = self.browser.new_context(viewport={"width": 1920, "height": 1080})
        self.page = self.context.new_page()

    def navigate(self, url: str) -> None:
        self.page.goto(url)

    def click(self, by: str, value: str, timeout: int = 15) -> Any:
        selector = self._get_selector(by, value)
        locator = self.page.locator(selector).first
        locator.wait_for(state="visible", timeout=timeout * 1000)
        locator.scroll_into_view_if_needed()
        locator.click(timeout=timeout * 1000)
        return locator

    def send_keys(self, by: str, value: str, text: str, timeout: int = 15) -> Any:
        selector = self._get_selector(by, value)
        locator = self.page.locator(selector).first
        locator.wait_for(state="visible", timeout=timeout * 1000)
        locator.scroll_into_view_if_needed()
        locator.fill(text, timeout=timeout * 1000)
        return locator

    def execute_script(self, script: str, *args) -> Any:
        # standard execute
        js_args = list(args)
        return self.page.evaluate(script, js_args)

    def execute_async_script(self, script: str, *args) -> Any:
        # Convert callback pattern into promise eval
        wrapped_script = f"""
            new Promise((resolve) => {{
                var callback = resolve;
                {script}
            }})
        """
        js_args = list(args)
        return self.page.evaluate(wrapped_script, js_args)

    def get_url(self) -> str:
        return self.page.url

    def get_title(self) -> str:
        return self.page.title

    def find_element(self, by: str, value: str):
        selector = self._get_selector(by, value)
        return self.page.locator(selector).first

    def find_elements(self, by: str, value: str) -> List[Any]:
        selector = self._get_selector(by, value)
        locators = self.page.locator(selector)
        return [locators.nth(i) for i in range(locators.count())]

    def quit(self) -> None:
        try:
            self.page.close()
            self.browser.close()
            self.playwright.stop()
        except Exception:
            pass

    @property
    def raw_driver(self) -> Any:
        return self.page

    def _get_selector(self, by: str, value: str) -> str:
        if by.lower() == "xpath":
            if not value.startswith("xpath="):
                return f"xpath={value}"
            return value
        elif by.lower() == "id":
            return f"id={value}"
        elif by.lower() == "css":
            return value
        return value
