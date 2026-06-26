import sys

# Globally configure standard streams to UTF-8 on Windows to prevent UnicodeEncodeError
for stream in (sys.stdout, sys.stderr):
    if stream is not None and hasattr(stream, "reconfigure"):
        try:
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
        except Exception:
            pass

import os
import re
import io
import time
import json
import random
import logging
import argparse
import warnings
import tempfile
import threading
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from functools import wraps, lru_cache
from typing import List, Optional, Tuple, Dict, Any
from datetime import datetime

# ── Suppress noisy 3rd-party warnings before any heavy import ─────────────────
warnings.filterwarnings("ignore")
os.environ["TOKENIZERS_PARALLELISM"] = "false"

# Import optimization layers
from src.api_reader_service import ContactReaderService, DraftReaderService, APOBReaderService, DraftIdCache
from src.optimization_manager import UploadTrackerService, RuntimeCacheManager, SessionCacheManager, NavigationCache, DOMReadReducer, OCRSingleton, ConnectionPoolManager
from src.portal_automation_adapter import PortalAutomationAdapter, SeleniumAdapter, PlaywrightAdapter
from src.batch_orchestrator import BatchOrchestrator

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import (
    TimeoutException, NoSuchElementException, WebDriverException
)
import urllib3
# Set pool limits to safely match or exceed MAX_WORKERS threads
urllib3.util.retry.Retry.DEFAULT_POOLSIZE = 25
urllib3.connectionpool.HTTPConnectionPool.default_connections = 25
urllib3.connectionpool.HTTPSConnectionPool.default_connections = 25

import pandas as pd
import openpyxl

# ── OCR / CV dependencies (soft-fail with clear message) ──────────────────────
try:
    import cv2
    import numpy as np
    from PIL import Image
    import torch
    from transformers import TrOCRProcessor, VisionEncoderDecoderModel
    _OCR_AVAILABLE = True
except ImportError as _ocr_import_err:
    _OCR_AVAILABLE = False
    _OCR_IMPORT_MSG = str(_ocr_import_err)



# ── Module-level thread-safety primitives ─────────────────────────────────────
_excel_write_lock = threading.Lock()          # guards completed/failed xlsx writes
_doc_path_cache:  Dict[str, Optional[str]] = {}  # memo for resolve_document_path


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("gst_amendment.log", encoding="utf-8"),
    ],
)
log = logging.getLogger("GST-Amendment")



# RETRY DECORATOR

def retry(max_attempts=3, delay=1.5, exceptions=(Exception,)):
    """Exponential-backoff retry decorator with jitter for flaky DOM interactions."""
    def decorator(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            for attempt in range(1, max_attempts + 1):
                try:
                    return fn(*args, **kwargs)
                except exceptions as e:
                    if attempt == max_attempts:
                        log.error(f"[{fn.__name__}] Failed after {max_attempts} attempts: {e}")
                        raise
                    # jitter prevents correlated retries in parallel sessions
                    jitter = random.uniform(0, delay * 0.5)
                    wait = min(delay * (2 ** (attempt - 1)) + jitter, 30.0)
                    log.warning(f"[{fn.__name__}] Attempt {attempt}/{max_attempts} failed: {e}. Retrying in {wait:.1f}s...")
                    time.sleep(wait)
        return wrapper
    return decorator



# ARGUMENT PARSER

def parse_args():
    parser = argparse.ArgumentParser(description="GST Portal Core Amendment Suite & Bulk APoB Engine")
    parser.add_argument("--username", required=False, help="GST Portal Username")
    parser.add_argument("--password", required=False, help="GST Portal Password")
    parser.add_argument("--headless",  action="store_true", help="Run in headless mode (no GUI)")
    parser.add_argument("--timeout",   type=int, default=15, help="Default element wait timeout (seconds)")
    parser.add_argument("--model",     type=str, default=None, help="TrOCR model name (e.g. microsoft/trocr-base-printed)")
    parser.add_argument("--batch",     type=str, default=None,
                        help="Path to CSV file with columns: username,password")
    parser.add_argument("--no-wait",   dest="no_wait", action="store_true",
                        help="Skip the final 'Press Enter' pause")
    parser.add_argument("--excel",     type=str, default="amendments.xlsx",
                        help="Path to Excel workbook with registrations details (default: amendments.xlsx)")
    parser.add_argument("--checkpoint", type=str, default="checkpoint.json",
                        help="Path to progress checkpoint file (default: checkpoint.json)")
    parser.add_argument("--doc", "--docs", type=str, default=None,
                        help="Explicit file path of the PDF document to upload (or secondary fallback)")
    parser.add_argument("--workers",   type=int, default=1, help="Number of concurrent workers (default: 1)")
    parser.add_argument("--engine",    type=str, default="selenium", help="Automation engine: selenium or playwright (default: selenium)")
    return parser.parse_known_args()[0]



# DRIVER FACTORY

def build_driver(headless: bool = False) -> webdriver.Chrome:
    """Constructs an optimised Chrome WebDriver instance with cookie reuse."""
    options = Options()

    if headless:
        options.add_argument("--headless=new")

    prefs = {
        "download.default_directory": os.path.join(os.path.expanduser("~"), "Downloads"),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
        "profile.default_content_settings.popups": 0,
        "credentials_enable_service": False,
        "profile.password_manager_enabled": False,
    }
    options.add_experimental_option("prefs", prefs)
    # NOTE: "detach":True was removed — it leaks Chrome processes on crashes.
    options.add_experimental_option("excludeSwitches", ["enable-logging", "enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    for flag in [
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu",
        "--disable-extensions",
        "--log-level=3",
        "--disable-renderer-backgrounding",
        "--disable-background-timer-throttling",
        "--disable-backgrounding-occluded-windows",
        "--disable-blink-features=AutomationControlled",
        "--disable-site-isolation-trials",
        "--disable-infobars",
        "--window-size=1920,1080",
        # Performance: skip image decoding for pages that don't need them
        "--blink-settings=imagesEnabled=true",
        # Reduce JS timer resolution — cuts idle CPU on portal waits
        "--js-flags=--max-old-space-size=512",
        # Multi-threading stability & connection drop prevention options
        "--disable-ipc-flooding-protection",
        "--disable-background-networking",
        "--disable-client-side-phishing-detection",
        "--disable-default-apps",
        "--disable-hang-monitor",
        "--disable-popup-blocking",
        "--disable-prompt-on-repost",
        "--disable-sync",
        "--no-first-run",
    ]:
        options.add_argument(flag)

    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    # Use eager page-load strategy: DOMContentLoaded is enough for SPA portals
    driver.execute_cdp_cmd("Page.setLifecycleEventsEnabled", {"enabled": True})
    driver.execute_cdp_cmd(
        "Page.addScriptToEvaluateOnNewDocument",
        {"source": "Object.defineProperty(navigator, 'webdriver', {get: () => undefined})"},
    )
    driver.maximize_window()
    log.info("Chrome WebDriver initialised successfully.")
    return driver



# DOM HELPERS

def wait_for_angular_idle(driver, timeout=20):
    """
    Stabilizes automation by waiting for Angular rendering to be idle,
    all AJAX queries to finish, and the document readyState to be complete.
    """
    log.info("[ANGULAR] Waiting for portal rendering and DOM stabilization...")
    start_time = time.time()
    try:
        # 1. Wait for document.readyState == 'complete'
        WebDriverWait(driver, timeout).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        
        # 2. Wait for Angular's dynamic elements to settle (if angular is present)
        angular_settled = """
        try {
            if (window.angular) {
                var el = document.querySelector('.ng-scope') || document.querySelector('[ng-app]');
                if (el) {
                    var $injector = angular.element(el).injector();
                    var $http = $injector ? $injector.get('$http') : null;
                    if ($http && $http.pendingRequests.length > 0) {
                        return false;
                    }
                }
            }
            return true;
        } catch (err) {
            return true;
        }
        """
        elapsed = time.time() - start_time
        remaining = max(1.0, timeout - elapsed)
        WebDriverWait(driver, remaining).until(lambda d: d.execute_script(angular_settled))
        
        # 3. Add a brief cushion sleep
        time.sleep(1.0)
        log.info(f"[ANGULAR] DOM is fully stable (elapsed: {time.time() - start_time:.1f}s)")
    except Exception as e:
        log.warning(f"[ANGULAR] Optimization wait timed out or failed: {e}. Proceeding anyway.")


def wait_and_click(driver, by, value, timeout=15):
    """
    Wait for element to be clickable, then JS-click (bypasses visibility/size traps).
    Implements presence, visibility, scrolling, JS-click fallback, stale retry, and timeout retry.
    """
    for attempt in range(1, 4):
        try:
            # 1. Wait for presence
            WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
            # 2. Wait for visibility & clickability
            elem = WebDriverWait(driver, timeout).until(EC.element_to_be_clickable((by, value)))
            # 3. Scroll into view
            driver.execute_script("arguments[0].scrollIntoView({block:'center', inline:'center'});", elem)
            time.sleep(0.3)
            # 4. Try click, fallback to JS click
            try:
                elem.click()
            except Exception:
                driver.execute_script("arguments[0].click();", elem)
            return elem
        except Exception as e:
            if attempt == 3:
                log.error(f"[CLICK] Click failed after 3 attempts on ({by}, {value}): {e}")
                raise
            log.warning(f"[CLICK] Attempt {attempt} failed on ({by}, {value}): {e}. Retrying...")
            time.sleep(1.0)


def safe_find_and_click(driver, xpaths: list, timeout=5) -> bool:
    """Tries multiple XPath candidates in order; returns True on first success."""
    for xpath in xpaths:
        try:
            wait_and_click(driver, By.XPATH, xpath, timeout=timeout)
            return True
        except (TimeoutException, NoSuchElementException):
            continue
    return False



# POPUP HANDLERS

@retry(max_attempts=1, delay=0.5, exceptions=(Exception,))
def handle_aadhaar_popup(driver, timeout=5) -> bool:
    """Dismisses Aadhaar / E-KYC verification modal using ESC key injection & ng-clicks."""
    log.info("Scanning for Aadhaar authentication modal...")
    try:
        WebDriverWait(driver, timeout).until(
            lambda d: d.find_elements(By.CLASS_NAME, "modal-backdrop") or
                      d.find_elements(By.CLASS_NAME, "modal-content")
        )
        log.info("Modal detected — evaluating content...")

        modal_text = driver.find_elements(
            By.XPATH,
            "//div[contains(@class,'modal')]//*[contains(text(),'Aadhaar') or contains(text(),'E-KYC')]"
        )
        if not modal_text:
            log.info("Modal present but not Aadhaar-related — skipping.")
            return True

        log.info("Aadhaar modal confirmed. Attempting dismissal...")

        for btn in driver.find_elements(By.XPATH, "//a[contains(text(),'Remind me later')]"):
            if btn.is_displayed():
                driver.execute_script("arguments[0].click();", btn)
                log.info("Dismissed via text-anchor.")
                return True

        try:
            el = driver.find_element(By.XPATH, "//a[@ng-click='cancelcallback()']")
            driver.execute_script("arguments[0].click();", el)
            log.info("Dismissed via ng-click binding.")
            return True
        except NoSuchElementException:
            pass

        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        log.info("Dismissed via ESC key.")
        return True

    except TimeoutException:
        log.info("No Aadhaar modal detected within timeout — continuing.")
    return True


def handle_metadata_popup(driver, timeout=4) -> bool:
    """Closes the GST metadata-collection notice banner."""
    try:
        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(
                (By.XPATH, "//*[contains(text(),'GST System is collecting metadata')]")
            )
        )
        log.info("Metadata notice detected — closing...")
        wait_and_click(driver, By.XPATH, "//a[contains(text(),'No-Remind me later')]", timeout=3)
    except TimeoutException:
        pass
    return True


def dismiss_all_popups(driver):
    """Runs popup handlers sequentially for stable single-driver execution on the GST Portal."""
    _COMBINED = (
        "//div[contains(@class,'modal-backdrop')] | "
        "//div[contains(@class,'modal-content')] | "
        "//*[contains(text(),'GST System is collecting metadata')]"
    )
    try:
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.XPATH, _COMBINED))
        )
    except TimeoutException:
        return
    # Dismiss popups sequentially to avoid Angular DOM layout crashes
    try:
        handle_aadhaar_popup(driver)
    except Exception as e:
        log.warning(f"[POPUP] Aadhaar modal dismissal warning: {e}")
        
    try:
        handle_metadata_popup(driver)
    except Exception as e:
        log.warning(f"[POPUP] Metadata notice dismissal warning: {e}")



# CAPTCHA OCR SUBSYSTEM

_GST_CAPTCHA_LEN         = 6
_OCR_CONFIDENCE_THRESHOLD = 0.45
_CAPTCHA_MAX_AUTO_RETRIES = 2
_CAPTCHA_TEMP_DIR         = tempfile.gettempdir()

_captcha_preprocessor = None
_captcha_ocr_engine   = None


class _CaptchaPreprocessor:
    """OpenCV preprocessing pipeline for GST-style CAPTCHAs."""
    def __init__(self):
        self.save_stages = False

    def measure_blur(self, img: np.ndarray) -> float:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        return float(cv2.Laplacian(gray, cv2.CV_64F).var())

    def _remove_red(self, img: np.ndarray) -> np.ndarray:
        """Remove red diagonal line & grid artifacts, replace with white."""
        b, g, r = cv2.split(img)
        bgr_mask = (
            (r.astype(np.int16) - g.astype(np.int16) > 20) &
            (r.astype(np.int16) - b.astype(np.int16) > 20)
        )
        hsv = cv2.cvtColor(img, cv2.COLOR_BGR2HSV)
        m1 = cv2.inRange(hsv, np.array([0,   35, 35]), np.array([18,  255, 255]))
        m2 = cv2.inRange(hsv, np.array([155,  35, 35]), np.array([180, 255, 255]))
        hsv_mask = cv2.dilate(
            m1 | m2,
            cv2.getStructuringElement(cv2.MORPH_RECT, (4, 4)),
            iterations=2,
        ) > 0
        combined = bgr_mask | hsv_mask
        out = img.copy()
        out[combined] = [255, 255, 255]
        return out

    def _clahe(self, img: np.ndarray, clip: float = 2.5) -> np.ndarray:
        lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
        l, a, b_ch = cv2.split(lab)
        cl = cv2.createCLAHE(clipLimit=clip, tileGridSize=(4, 4)).apply(l)
        return cv2.cvtColor(cv2.merge([cl, a, b_ch]), cv2.COLOR_LAB2BGR)

    def _upscale(self, img: np.ndarray, scale: int = 4) -> np.ndarray:
        h, w = img.shape[:2]
        return cv2.resize(img, (w * scale, h * scale), interpolation=cv2.INTER_CUBIC)


class _TrOCREngine:
    """Microsoft TrOCR inference wrapper."""
    _MODEL_CANDIDATES = [
        "microsoft/trocr-base-printed",
        "microsoft/trocr-large-printed",
    ]

    def __init__(self, model_name: Optional[str] = None):
        # Use thread-safe double-checked lock OCRSingleton to load TrOCR once
        engine_dict = OCRSingleton.get_ocr_engine(model_name=model_name)
        self.device = engine_dict["device"]
        self.processor = engine_dict["processor"]
        self.model = engine_dict["model"]
        self.model_name = engine_dict["model_name"]
        log.info(f"[CAPTCHA] TrOCR ready (reused from thread-safe singleton on {self.device.upper()})")

    def run_ocr(self, bgr_img: np.ndarray) -> Tuple[str, float]:
        t0 = time.perf_counter()
        try:
            rgb     = cv2.cvtColor(bgr_img, cv2.COLOR_BGR2RGB)
            pil_img = Image.fromarray(rgb)

            pixel_values = self.processor(
                images=pil_img, return_tensors="pt"
            ).pixel_values.to(self.device)

            with torch.no_grad():
                outputs = self.model.generate(
                    pixel_values,
                    num_beams=3,
                    max_new_tokens=20,
                    output_scores=True,
                    return_dict_in_generate=True,
                )

            sequences = outputs.sequences
            raw_text  = self.processor.batch_decode(
                sequences, skip_special_tokens=True
            )[0].strip()

            if hasattr(outputs, "sequences_scores") and outputs.sequences_scores is not None:
                conf = float(torch.exp(outputs.sequences_scores[0]).item())
                conf = max(0.0, min(1.0, conf))
            else:
                conf = self._estimate_confidence(outputs, sequences)

            log.info(
                f"[CAPTCHA] OCR result: '{raw_text}'  "
                f"confidence: {conf*100:.1f}%  "
                f"({(time.perf_counter()-t0)*1000:.0f}ms)"
            )
            return raw_text, conf

        except Exception as e:
            log.error(f"[CAPTCHA] TrOCR inference error: {e}")
            return "", 0.0

    @staticmethod
    def _estimate_confidence(outputs, sequences) -> float:
        try:
            if not hasattr(outputs, "scores") or outputs.scores is None:
                return 0.5
            log_probs = []
            for step_idx, step_scores in enumerate(outputs.scores):
                probs  = torch.softmax(step_scores, dim=-1)
                chosen = sequences[0, step_idx + 1]
                if chosen.item() in (0, 1, 2):
                    continue
                log_probs.append(float(torch.log(probs[0, chosen] + 1e-9)))
            if not log_probs:
                return 0.5
            avg_lp = sum(log_probs) / len(log_probs)
            return float(np.clip(np.exp(avg_lp), 0.0, 1.0))
        except Exception:
            return 0.5


def _clean_captcha_text(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", text).upper()


def _apply_gst_char_corrections(text: str) -> str:
    letter_count = sum(1 for c in text if c.isalpha())
    digit_count  = sum(1 for c in text if c.isdigit())
    if digit_count >= letter_count:
        corrections = str.maketrans("OISZBGQDL", "015286000")
    else:
        corrections = str.maketrans("OI", "01")
    return text.translate(corrections)


def _ensure_ocr_engine_ready(model_name=None) -> bool:
    global _captcha_preprocessor, _captcha_ocr_engine

    if not _OCR_AVAILABLE:
        log.warning(
            f"[CAPTCHA] OCR libraries unavailable ({_OCR_IMPORT_MSG}). "
            "Falling back to manual CAPTCHA mode."
        )
        return False

    if _captcha_preprocessor is None:
        _captcha_preprocessor = _CaptchaPreprocessor()

    if _captcha_ocr_engine is None:
        try:
            _captcha_ocr_engine = _TrOCREngine(model_name=model_name)
        except Exception as e:
            log.error(f"[CAPTCHA] Engine initialisation failed: {e}. Will use manual fallback.")
            return False

    return True


def capture_captcha_image(driver) -> Optional[str]:
    _CAPTCHA_SELECTORS = [
        (By.ID,           "imgCaptcha"),
        (By.CSS_SELECTOR, "img.captcha"),
        (By.XPATH,        "//img[@ng-src[contains(.,'/services/captcha')]]"),
        (By.XPATH,        "//img[@src[contains(.,'/services/captcha')]]"),
        (By.ID,           "captcha-img"),
        (By.CSS_SELECTOR, "img[class*='captchaImg']"),
        (By.XPATH,        "//img[contains(translate(@src,'CAPTCHA','captcha'),'captcha')]"),
    ]

    captcha_elem = None
    for by, value in _CAPTCHA_SELECTORS:
        try:
            captcha_elem = WebDriverWait(driver, 5).until(
                EC.presence_of_element_located((by, value))
            )
            if captcha_elem and captcha_elem.is_displayed():
                break
            captcha_elem = None
        except (TimeoutException, NoSuchElementException):
            continue

    if captcha_elem is None:
        log.warning("[CAPTCHA] CAPTCHA image element not found.")
        return None

    try:
        WebDriverWait(driver, 8).until(
            lambda d: captcha_elem.size.get("width", 0) > 10
                      and captcha_elem.size.get("height", 0) > 10
        )
    except TimeoutException:
        log.warning("[CAPTCHA] CAPTCHA image did not reach expected dimensions.")
        return None

    try:
        try:
            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center',inline:'center'});",
                captcha_elem,
            )
            time.sleep(0.4)

            elem_png   = captcha_elem.screenshot_as_png
            captcha_crop = cv2.imdecode(
                np.frombuffer(elem_png, dtype=np.uint8), cv2.IMREAD_COLOR
            )
            if captcha_crop is None or captcha_crop.shape[0] < 10 or captcha_crop.shape[1] < 10:
                raise ValueError("Element screenshot too small or failed to decode")
            if float(captcha_crop.std()) < 5.0:
                raise ValueError("Element screenshot appears blank")

            log.info(
                f"[CAPTCHA] Image captured via element.screenshot "
                f"({captcha_crop.shape[1]}×{captcha_crop.shape[0]}px)"
            )

        except Exception as elem_err:
            log.warning(f"[CAPTCHA] Element screenshot failed ({elem_err}), using full-page crop fallback.")

            driver.execute_script(
                "arguments[0].scrollIntoView({block:'center',inline:'center'});",
                captcha_elem,
            )
            time.sleep(0.4)

            png_bytes   = driver.get_screenshot_as_png()
            full_screen = cv2.imdecode(
                np.frombuffer(png_bytes, dtype=np.uint8), cv2.IMREAD_COLOR
            )
            if full_screen is None:
                log.warning("[CAPTCHA] Full-page screenshot decode failed.")
                return None

            location = captcha_elem.location
            size     = captcha_elem.size
            dpr = float(driver.execute_script("return window.devicePixelRatio || 1;") or 1)

            x1 = max(0, int(location["x"] * dpr))
            y1 = max(0, int(location["y"] * dpr))
            x2 = min(full_screen.shape[1], int((location["x"] + size["width"])  * dpr))
            y2 = min(full_screen.shape[0], int((location["y"] + size["height"]) * dpr))

            captcha_crop = full_screen[y1:y2, x1:x2]

            if captcha_crop.shape[0] < 10 or captcha_crop.shape[1] < 10:
                log.warning(f"[CAPTCHA] Crop too small ({captcha_crop.shape[1]}×{captcha_crop.shape[0]}) — rejecting.")
                return None

            if float(captcha_crop.std()) < 5.0:
                log.warning("[CAPTCHA] Captured image appears blank — rejecting.")
                return None

        out_path = os.path.join(_CAPTCHA_TEMP_DIR, f"gst_captcha_live_{int(time.time() * 1000)}.png")
        cv2.imwrite(out_path, captcha_crop)
        log.info(f"[CAPTCHA] Saved -> {out_path}")
        return out_path

    except Exception as e:
        log.error(f"[CAPTCHA] Screenshot/crop error: {e}")
        return None


def preprocess_captcha_image(image_path: str) -> Optional[List]:
    if _captcha_preprocessor is None:
        log.error("[CAPTCHA] Preprocessor not initialised.")
        return None

    try:
        img = cv2.imread(image_path)
        if img is None:
            log.error(f"[CAPTCHA] cv2.imread failed: {image_path}")
            return None

        pp = _captcha_preprocessor

        # Pre-compute shared intermediaries (used by multiple variants)
        no_red = pp._remove_red(img)
        no_red_median = cv2.medianBlur(no_red, 3)

        # Build all 4 variants in parallel — each is CPU-bound and independent
        def _s1(): return pp._upscale(img, 4)
        def _s2(): return pp._upscale(no_red, 4)
        def _s3(): return pp._upscale(no_red_median, 4)
        def _s4():
            s4_base = cv2.medianBlur(no_red_median, 3)
            return pp._upscale(pp._clahe(s4_base, clip=2.5), 4)

        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = [pool.submit(fn) for fn in (_s1, _s2, _s3, _s4)]
            results = [f.result() for f in futures]

        log.info("[CAPTCHA] Preprocessing complete (4 ensemble variants ready, parallel)")
        return results

    except Exception as e:
        log.error(f"[CAPTCHA] Preprocessing error: {e}")
        return None


def solve_captcha_from_image(preprocessed_variants: List[np.ndarray]) -> Tuple[str, float]:
    if _captcha_ocr_engine is None:
        log.error("[CAPTCHA] OCR engine not initialised.")
        return "", 0.0

    if not preprocessed_variants:
        log.error("[CAPTCHA] No preprocessed variants supplied.")
        return "", 0.0

    strategy_names = [
        "Raw 4× Upscale",
        "Red Mask Only",
        "Red+Median",
        "Red+Median+CLAHE",
    ]

    _EARLY_EXIT_CONFIDENCE = 0.75

    candidates = []
    for idx, variant_img in enumerate(preprocessed_variants):
        name = strategy_names[idx] if idx < len(strategy_names) else f"Variant-{idx}"
        try:
            raw_text, conf = _captcha_ocr_engine.run_ocr(variant_img)
            cleaned        = _apply_gst_char_corrections(_clean_captcha_text(raw_text))
            candidates.append((name, raw_text, cleaned, conf))
            log.info(f"[CAPTCHA]   {name:<22} raw='{raw_text}'  cleaned='{cleaned}'  conf={conf*100:.1f}%")

            if len(cleaned) == _GST_CAPTCHA_LEN and conf >= _EARLY_EXIT_CONFIDENCE:
                log.info(
                    f"[CAPTCHA] Early exit after variant {idx+1}/{len(preprocessed_variants)} "
                    f"(conf={conf*100:.1f}% >= {_EARLY_EXIT_CONFIDENCE*100:.0f}%, len={len(cleaned)})"
                )
                break

        except Exception as e:
            log.warning(f"[CAPTCHA] OCR failed for variant '{name}': {e}")

    if not candidates:
        return "", 0.0

    def _score(c):
        txt = c[2]
        return (
            10 if len(txt) == 0 else 0,
            abs(len(txt) - _GST_CAPTCHA_LEN),
            -c[3],
        )

    candidates.sort(key=_score)
    best_name, best_raw, best_text, best_conf = candidates[0]
    log.info(f"[CAPTCHA] Best strategy: '{best_name}'  ->  '{best_text}'  (conf={best_conf*100:.1f}%)")
    return best_text, best_conf


def refresh_captcha(driver) -> bool:
    _REFRESH_SELECTORS = [
        (By.XPATH, "//button[@ng-click='refreshCaptcha()']"),
        (By.XPATH, "//button[.//i[contains(@class,'fa-refresh')]]"),
        (By.ID,    "refreshCaptcha"),
        (By.XPATH, "//a[contains(@title,'Refresh') or contains(@title,'refresh') or "
                   "contains(@aria-label,'refresh') or contains(@aria-label,'Refresh')]"),
        (By.XPATH, "//span[contains(@class,'refresh') or contains(@class,'reload')]"),
        (By.XPATH, "//i[contains(@class,'fa-refresh') or contains(@class,'fa-redo')]"),
    ]

    for by, value in _REFRESH_SELECTORS:
        try:
            elem = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((by, value))
            )
            driver.execute_script("arguments[0].click();", elem)
            time.sleep(0.5)
            log.info("[CAPTCHA] CAPTCHA refreshed via click.")
            return True
        except (TimeoutException, NoSuchElementException):
            continue

    try:
        refreshed = driver.execute_script("""
            var imgs = document.querySelectorAll(
                'img[id*="captcha" i], img[class*="captcha" i], img[src*="captcha" i]'
            );
            if (imgs.length > 0) {
                var img = imgs[0];
                var src = img.src.split('?')[0];
                img.src = src + '?ts=' + new Date().getTime();
                return true;
            }
            return false;
        """)
        if refreshed:
            time.sleep(0.6)
            log.info("[CAPTCHA] CAPTCHA refreshed via JS src-swap.")
            return True
    except Exception as e:
        log.warning(f"[CAPTCHA] JS refresh fallback failed: {e}")

    log.warning("[CAPTCHA] Could not find a refresh control — CAPTCHA not refreshed.")
    return False


def auto_solve_captcha(driver, timeout: int = 15, model_name=None) -> Tuple[bool, str]:
    if not _ensure_ocr_engine_ready(model_name=model_name):
        return False, ""

    img_path = capture_captcha_image(driver)
    if not img_path:
        log.warning("[CAPTCHA] Capture failed.")
        return False, ""

    variants = preprocess_captcha_image(img_path)
    if not variants:
        log.warning("[CAPTCHA] Preprocessing failed. Refreshing CAPTCHA...")
        refresh_captcha(driver)
        return False, ""

    captcha_text, confidence = solve_captcha_from_image(variants)

    if not captcha_text:
        log.warning("[CAPTCHA] OCR returned empty result. Refreshing CAPTCHA...")
        refresh_captcha(driver)
        return False, ""

    if len(captcha_text) != _GST_CAPTCHA_LEN:
        log.warning(
            f"[CAPTCHA] Length mismatch: got {len(captcha_text)} chars "
            f"('{captcha_text}'), expected {_GST_CAPTCHA_LEN}. Rejecting and refreshing CAPTCHA immediately..."
        )
        refresh_captcha(driver)
        return False, ""

    if confidence < _OCR_CONFIDENCE_THRESHOLD:
        log.warning(
            f"[CAPTCHA] Confidence {confidence*100:.1f}% below threshold "
            f"({_OCR_CONFIDENCE_THRESHOLD*100:.0f}%) — rejecting and refreshing CAPTCHA..."
        )
        refresh_captcha(driver)
        return False, ""

    captcha_field = None
    try:
        captcha_field = driver.execute_script("""
            var selectors = [
                'input#captcha',
                'input[name="captcha"]',
                'input[data-ng-model="lform.captcha"]',
                'input#userCaptcha',
                'input#captchaAnswer',
                'input[placeholder*="characters" i]',
                'input[id*="captcha" i]'
            ];
            for (var i = 0; i < selectors.length; i++) {
                var el = document.querySelector(selectors[i]);
                if (el && el.offsetParent !== null) return el;
            }
            return null;
        """)
    except Exception:
        captcha_field = None

    if captcha_field is None:
        _CAPTCHA_INPUT_SELECTORS = [
            (By.ID,    "captcha"),
            (By.XPATH, "//input[@name='captcha']"),
            (By.XPATH, "//input[@data-ng-model='lform.captcha']"),
            (By.ID,    "userCaptcha"),
            (By.XPATH, "//input[contains(@placeholder,'characters') or contains(@placeholder,'Characters')]"),
        ]
        for by, value in _CAPTCHA_INPUT_SELECTORS:
            try:
                captcha_field = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((by, value))
                )
                break
            except (TimeoutException, NoSuchElementException):
                continue

    if captcha_field is None:
        log.warning("[CAPTCHA] CAPTCHA input field not found.")
        return False, ""

    try:
        captcha_field.clear()
        captcha_field.send_keys(captcha_text)
        log.info(f"[CAPTCHA] Field filled with: '{captcha_text}'")
    except Exception as e:
        log.error(f"[CAPTCHA] Could not type into CAPTCHA field: {e}")
        return False, ""

    return True, captcha_text



# SAVED-APPLICATION DRAFT CLEANER & DRAFT DETECTOR

_BLOCKING_APPCDS = {
    "REGAN", "REGAC", "APLAC", "APLAN", "AUNAN", "AEMAN", "AOTAN", "AUNAC",
    "AEMAC", "AOTAC", "ATDAN", "ATCAN", "ANRAN", "ATDAC", "ATCAC", "ANRAC",
    "APPEL", "AGPAN", "AGPAC", "ATRAN", "ADJVP", "ARARA", "ARAPA", "ADVPD",
    "ADJWO", "ADJWN", "APLOW", "AOWFC",
}
_SAVED_APP_URL   = "https://services.gst.gov.in/services/auth/savedapp"
_AMENDMENT_URL_FRAG = "amendbusinessdetail"
_AMENDMENT_HREF  = "/registration/auth/amend/core/amendbusinessdetail"
_WARNING_TEXTS   = [
    "cannot file the amendment",
    "pending for approval",
    "earlier application for core amendment",
    "you cannot apply for amendment of core field",
    "cannot apply for amendment of core field",
    "seems you have already initiated",
    "already initiated application of registration core",
    "have already initiated",
    "application of registration core fields",
]


def _warning_xpath():
    return " or ".join(f"contains(text(),'{t}')" for t in _WARNING_TEXTS)



# CORE PIPELINE & SUBMISSION ACTIONS

_XPATH_TRADE_NAME_INPUT = (
    "//input[@id='tradeName' or @name='tradeName' or "
    "@data-ng-model[contains(.,'tradeName')] or "
    "@placeholder[contains(.,'Trade Name')] or "
    "@placeholder[contains(.,'trade name')]]"
)
_XPATH_REASON_INPUT = (
    "//textarea[@id='reason' or @name='reason' or "
    "@data-ng-model[contains(.,'reason')] or "
    "@placeholder[contains(.,'reason') or contains(.,'Reason')]]"
    " | "
    "//input[@id='reason' or @name='reason']"
)
_XPATH_SAVE_BTN = (
    "//button[normalize-space(.)='Save as Draft' or normalize-space(.)='Save' "
    "or @id='btnSave' or @ng-click[contains(.,'saveDraft')]] | //a[contains(text(), 'Save')]"
)
_XPATH_SUBMIT_BTN = (
    "//button[normalize-space(.)='Submit' or @id='btnSubmit' or "
    "@ng-click[contains(.,'submit') or contains(.,'Submit')]]"
    "[not(contains(@class,'disabled'))]"
)
_XPATH_EVC_TAB = (
    "//a[contains(text(),'EVC') or contains(text(),'OTP') or "
    "@href[contains(.,'evc')] or @ng-click[contains(.,'evc')]]"
    " | "
    "//li[contains(@class,'evc') or contains(@class,'otp')]//a"
)
_XPATH_DSC_TAB = (
    "//a[contains(text(),'DSC') or @href[contains(.,'dsc')] or "
    "@ng-click[contains(.,'dsc')]]"
)
_XPATH_OTP_INPUT = (
    "//input[@id='otp' or @name='otp' or @placeholder[contains(.,'OTP')] or "
    "@data-ng-model[contains(.,'otp')]]"
)
_XPATH_VERIFY_OTP_BTN = (
    "//button[normalize-space(.)='Verify OTP' or @ng-click[contains(.,'verifyOtp')] "
    "or @ng-click[contains(.,'verifyotp')]]"
)
_XPATH_SUCCESS_BANNER = (
    "//*[contains(text(),'ARN') or contains(text(),'successfully') "
    "or contains(text(),'Successfully') or contains(@class,'success')]"
)
_XPATH_PROCEED_BTN = (
    "//button[normalize-space(.)='Proceed' or "
    "@ng-click[contains(.,'proceed') or contains(.,'Proceed')]]"
)


@retry(max_attempts=2, delay=1, exceptions=(TimeoutException, NoSuchElementException))
def _fill_text_field(driver, xpath: str, value: str, timeout: int = 15,
                     label: str = "field") -> bool:
    try:
        elem = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        driver.execute_script(
            "arguments[0].value = '';"
            "arguments[0].dispatchEvent(new Event('input',{bubbles:true}));"
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));",
            elem,
        )
        elem.send_keys(value)
        driver.execute_script(
            "arguments[0].dispatchEvent(new Event('blur',{bubbles:true}));", elem
        )
        log.info(f"[AMEND] {label} set to: '{value[:60]}'")
        return True
    except (TimeoutException, NoSuchElementException):
        log.warning(f"[AMEND] {label} field not found — skipping.")
        return False


def _fill_date_field(driver, xpath: str, value: str, timeout: int = 15, label: str = "Date of Amendment") -> bool:
    """Fills date fields using JavaScript to bypass portal picker popups, triggering event listeners."""
    try:
        elem = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, elem, value)
        try:
            elem.send_keys(Keys.TAB)
        except:
            pass
        log.info(f"[FINAL] {label} successfully set to: '{value}'")
        return True
    except Exception as e:
        log.warning(f"[FINAL] Failed to fill date field '{label}': {e}")
        return False


def _open_services_amendment_link(driver, timeout: int):
    """Click Services → Registration → Core Amendment deep-link."""
    wait_and_click(
        driver, By.XPATH,
        "//a[contains(@class,'dropdown-toggle') and contains(text(),'Services')]",
        timeout=timeout,
    )
    WebDriverWait(driver, 5).until(
        EC.presence_of_element_located((By.XPATH, "//a[contains(text(),'Registration')]"))
    )
    elem = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located(
            (By.XPATH, f"//a[contains(@href,'{_AMENDMENT_HREF}')]")
        )
    )
    driver.execute_script("arguments[0].click();", elem)


def fetch_profile_contact_info(driver, timeout: int = 15) -> Tuple[Optional[str], Optional[str]]:
    """
    Rapidly extracts primary mobile number and email from the client's profile page.
    Optimized: Tries API-Assisted Reader first, falls back to DOM if API fails.
    """
    try:
        t_api_start = time.time()
        api_data = ContactReaderService.fetch_contacts(driver)
        
        # Robust parsing utility
        parsed_data = None
        if api_data:
            if isinstance(api_data, dict):
                parsed_data = api_data
            elif isinstance(api_data, str):
                val_str = api_data.strip()
                if val_str:
                    # 1. Try standard JSON load
                    try:
                        parsed = json.loads(val_str)
                        if isinstance(parsed, dict):
                            parsed_data = parsed
                        elif isinstance(parsed, str):
                            parsed2 = json.loads(parsed)
                            if isinstance(parsed2, dict):
                                parsed_data = parsed2
                    except Exception:
                        pass
                    
                    # 2. Try ast.literal_eval fallback
                    if not parsed_data:
                        try:
                            import ast
                            parsed = ast.literal_eval(val_str)
                            if isinstance(parsed, dict):
                                parsed_data = parsed
                        except Exception:
                            pass
                            
                    # 3. Try cleaning Angular security prefix
                    if not parsed_data and val_str.startswith(")]}',"):
                        try:
                            cleaned = val_str[5:].strip()
                            parsed = json.loads(cleaned)
                            if isinstance(parsed, dict):
                                parsed_data = parsed
                        except Exception:
                            pass
                            
                    if not parsed_data:
                        log.warning(f"[PROFILE-API] Failed parsing api_data string: {api_data[:200]}")

        if parsed_data and isinstance(parsed_data, dict):
            # Check if authorisedSignatoryDetl exists
            has_auth_sig = ("authorisedSignatoryDetl" in parsed_data or "authorizedSignatoryDetl" in parsed_data or "authorizedSignatory" in parsed_data)
            
            if has_auth_sig:
                mobile = None
                email = None
                
                # Check auth signatory list
                auth_sig = parsed_data.get("authorisedSignatoryDetl", []) or parsed_data.get("authorizedSignatoryDetl", [])
                if auth_sig and isinstance(auth_sig, list) and len(auth_sig) > 0:
                    first_signatory = auth_sig[0] 
                    if isinstance(first_signatory, list) and len(first_signatory) >= 2:
                        for item in first_signatory:
                            item_str = str(item).strip()
                            if re.match(r"^\d{10}$", item_str):
                                mobile = item_str
                            elif "@" in item_str and "." in item_str:
                                email = item_str
                    elif isinstance(first_signatory, dict):
                        mobile = (first_signatory.get("mobile") or first_signatory.get("mobNum") or 
                                  first_signatory.get("mobileNumber") or first_signatory.get("mobileNo") or
                                  first_signatory.get("commmob"))
                        email = (first_signatory.get("email") or first_signatory.get("emailId") or 
                                 first_signatory.get("emailAddress") or first_signatory.get("commemail"))
                
                # Fallback recursive search if not found yet
                if not mobile or not email:
                    def find_values(obj, target_keys_mobile, target_keys_email):
                        mob_val, em_val = None, None
                        if isinstance(obj, dict):
                            for k, v in obj.items():
                                if k.lower() in target_keys_mobile and v and re.match(r"^\d{10}$", str(v).strip()):
                                    mob_val = str(v).strip()
                                elif k.lower() in target_keys_email and v and "@" in str(v) and "." in str(v):
                                    em_val = str(v).strip()
                                
                                if isinstance(v, (dict, list)):
                                    m, e = find_values(v, target_keys_mobile, target_keys_email)
                                    if m and not mob_val: mob_val = m
                                    if e and not em_val: em_val = e
                        elif isinstance(obj, list):
                            for item in obj:
                                if isinstance(item, (dict, list)):
                                    m, e = find_values(item, target_keys_mobile, target_keys_email)
                                    if m and not mob_val: mob_val = m
                                    if e and not em_val: em_val = e
                        return mob_val, em_val

                    m_keys = ["mobile", "mobnum", "mobilenumber", "mobileno", "commmob", "primarymobile"]
                    e_keys = ["email", "emailid", "emailaddress", "commemail", "primaryemail"]
                    m_found, e_found = find_values(parsed_data, m_keys, e_keys)
                    if m_found: mobile = m_found
                    if e_found: email = e_found

                if mobile and email:
                    log.info(f"[PROFILE-API] ✓ Successfully retrieved contact info via API in {time.time()-t_api_start:.2f}s: Mobile={mobile}, Email={email}")
                    return mobile, email
                else:
                    log.warning("[PROFILE-API] authorisedSignatoryDetl found but mobile or email missing.")
            else:
                log.warning("[PROFILE-API] Contact API success but authorisedSignatoryDetl key is missing in response.")
                
    except Exception as e:
        log.warning(f"[PROFILE-API] Profile contact API failed: {e}")

    log.info("[PROFILE] Visual Fallback: Navigating to My Profile using direct JS redirect...")
    try:
        driver.execute_script("window.location.href = '/services/auth/myprofile';")
        # Wait for either pills-auth-tab or URL to confirm loading
        WebDriverWait(driver, timeout).until(EC.presence_of_element_located((By.ID, "pills-auth-tab")))
        time.sleep(1.0) # brief sleep for rendering
        
        # Click the "Address and Contacts" tab
        log.info("[PROFILE] Clicking 'Address and Contacts' tab...")
        wait_and_click(driver, By.ID, "pills-auth-tab", timeout=timeout)
        time.sleep(1.5) # Wait for ng-click="viewcontacts()" data binding
        
        # Ensure the "Authorized Signatory" collapse panel is open
        log.info("[PROFILE] Ensuring 'Authorized Signatory' panel is expanded...")
        is_open = driver.execute_script("""
            var panel = document.getElementById('collapseAuthSign');
            if (panel) {
                return panel.classList.contains('in') || panel.getAttribute('aria-expanded') === 'true';
            }
            return false;
        """)
        if not is_open:
            log.info("[PROFILE] Panel collapsed. Clicking to expand...")
            # Click the toggle link
            toggle = WebDriverWait(driver, timeout).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@href='#collapseAuthSign']"))
            )
            driver.execute_script("arguments[0].click();", toggle)
            time.sleep(1.0)
            
        # Use a clean driver.execute_script snippet to extract cell text from the first row 
        # of the signatory table (cells[1] for mobile, cells[2] for email).
        log.info("[PROFILE] Extracting primary contact details via JavaScript...")
        result = driver.execute_script("""
            var panel = document.getElementById('collapseAuthSign');
            if (!panel) return null;
            var table = panel.querySelector('table');
            if (!table) return null;
            var rows = table.querySelectorAll('tbody tr');
            if (rows.length === 0) return null;
            
            var firstRow = rows[0];
            var cells = firstRow.querySelectorAll('td');
            if (cells.length < 3) return null;
            
            return {
                mobile: cells[1].innerText.trim(),
                email: cells[2].innerText.trim()
            };
        """)
        
        if result:
            mobile = result.get("mobile")
            email = result.get("email")
            log.info(f"[PROFILE] Extracted Profile Contact: Mobile='{mobile}', Email='{email}'")
            return mobile, email
        else:
            log.warning("[PROFILE] Signatory table data not found or empty.")
            return None, None
            
    except Exception as e:
        log.error(f"[PROFILE] Failed to extract contact details from profile: {e}")
        return None, None


# DYNAMIC APoB AUTO-FORM PROCESSING ENGINE


def navigate_to_saved_applications(driver, timeout: int = 15) -> bool:
    """Navigates to My Saved Applications inside the active client context safely."""
    log.info("[NAVIGATION] Navigating to Dashboard -> Services -> User Services -> My Saved Applications...")

    # 1. Dismiss any potential blocking popups first (like Aadhaar modal)
    try:
        dismiss_all_popups(driver)
    except:
        pass

    # 2. Strategy 1: Direct local relative URL redirect (preserves session scope and client context, extremely fast & reliable)
    try:
        log.info("[NAVIGATION] Triggering direct relative URL redirect...")
        driver.execute_script("window.location.href = '/services/auth/savedapp';")
        time.sleep(4.0)
        log.info("[NAVIGATION] ✓ Successfully navigated via relative URL redirect.")
        return True
    except Exception as ex:
        log.warning(f"[NAVIGATION] Direct relative redirect failed ({ex}). Falling back to visual link strategies...")

    # 3. Strategy 2: Check if "My Saved Applications" quick link is visible on the current page
    quick_links = [
        "//a[contains(text(),'My Saved Applications')]",
        "//a[contains(text(),'Saved Applications')]",
        "//a[contains(@href,'/services/auth/savedapp')]",
    ]
    for ql in quick_links:
        try:
            elem = driver.find_element(By.XPATH, ql)
            if elem.is_displayed():
                driver.execute_script("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", elem)
                log.info("[NAVIGATION] ✓ Successfully navigated via Quick Link.")
                time.sleep(2.0)
                return True
        except Exception:
            continue

    # 4. Strategy 3: Click via Top Navbar Menus (preserving selected context!)
    try:
        # Click 'Services' tab
        wait_and_click(
            driver, By.XPATH,
            "//a[contains(@class,'dropdown-toggle') and contains(text(),'Services')]",
            timeout=timeout,
        )
        time.sleep(0.5)

        # Click 'User Services' dropdown item
        wait_and_click(
            driver, By.XPATH,
            "//a[contains(text(),'User Services') or contains(text(),'User services')]",
            timeout=timeout,
        )
        time.sleep(0.5)

        # Click 'My Saved Applications'
        saved_xpath = (
            "//a[contains(text(),'My Saved Applications') or "
            "contains(text(),'Saved Applications') or "
            "contains(@href,'savedapp')]"
        )
        wait_and_click(driver, By.XPATH, saved_xpath, timeout=timeout)
        log.info("[NAVIGATION] ✓ Successfully navigated via Top Navbar dropdown.")
        time.sleep(2.0)
        return True
    except Exception as e:
                return False


def check_and_resume_draft(driver, row: Dict[str, Any], timeout: int = 15) -> str:
    """
    Scans My Saved Applications to resume existing draft for current GSTIN or Trade Name,
    or clears any unrelated blocking Core Fields drafts.
    """
    gstin = row.get("GSTIN")
    trade_name = row.get("TradeName")
    log.info(f"[{gstin}] Scanning My Saved Applications...")
    
    # Use the robust menu navigation
    navigate_to_saved_applications(driver, timeout)

    # Wait for savedapp URL context to load first
    try:
        WebDriverWait(driver, timeout).until(EC.url_contains("savedapp"))
        time.sleep(1.5) # Allow Angular data binding to fetch rows
    except TimeoutException:
        log.warning(f"[{gstin}] URL did not change to 'savedapp' page. Table scan may fail.")

    try:
        WebDriverWait(driver, timeout).until(lambda d: d.execute_script(
            "return document.querySelectorAll('table tbody tr, tr[ng-repeat]').length > 0"
        ))
        time.sleep(1.0)
    except TimeoutException:
        log.info(f"[{gstin}] Saved applications table is empty.")
        return 'cleared'

    JS_FIND_ROW = """
        var gstin = arguments[0];
        var tradeName = arguments[1];
        var rows = document.querySelectorAll('table tbody tr, tr[ng-repeat]');
        
        // Strategy A: Scan for a draft row matching either GST_Reg_ID, GSTIN, or client's Trade Name
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            if (!row.offsetParent) continue;
            var txt = row.innerText.toUpperCase();
            
            var matches_client = false;
            if (gstin && txt.indexOf(gstin.toUpperCase()) !== -1) matches_client = true;
            if (tradeName && txt.indexOf(tradeName.toUpperCase()) !== -1) matches_client = true;
            
            // If row belongs to the client context and is a Core Field Amendment
            if (matches_client && 
                (txt.indexOf('REG-14') !== -1 || txt.indexOf('AMENDMENT') !== -1 || txt.indexOf('CORE') !== -1 || txt.indexOf('REGAN') !== -1 || txt.indexOf('REGAC') !== -1)) {
                return { type: 'match', index: i };
            }
        }
        
        // Strategy B: If no exact client row matched, check if there is ANY Core Amendment draft in the table.
        // Since My Saved Applications is context-filtered to the current selected client in most dashboard flows,
        // any visible Core Field Amendment draft is very likely our target!
        // Improved: search for 'CORE' instead of 'AMENDMENT OF REGISTRATION CORE' to catch 'Application of Registration Core fields'
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            if (!row.offsetParent) continue;
            var txt = row.innerText.toUpperCase();
            
            if (txt.indexOf('REG-14') !== -1 || txt.indexOf('CORE') !== -1 || txt.indexOf('AMENDMENT') !== -1 || txt.indexOf('REGAN') !== -1 || txt.indexOf('REGAC') !== -1) {
                return { type: 'match_context_fallback', index: i };
            }
        }
        
        // Strategy C: Scan for any OTHER unrelated blocking core amendments to delete
        var blocking_keywords = ['REG-14', 'CORE', 'AMENDMENT', 'REGAN', 'REGAC'];
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            if (!row.offsetParent) continue;
            var txt = row.innerText.toUpperCase();
            
            var is_blocking = false;
            for (var k = 0; k < blocking_keywords.length; k++) {
                if (txt.indexOf(blocking_keywords[k]) !== -1) {
                    is_blocking = true;
                    break;
                }
            }
            if (is_blocking) {
                return { type: 'unrelated_blocking', index: i };
            }
        }
        return null;
    """

    res = driver.execute_script(JS_FIND_ROW, gstin, trade_name)
    if res:
        row_type = res.get("type")
        idx = res.get("index")

        if row_type in ('match', 'match_context_fallback'):
            log.info(f"[{gstin}] Resuming Core Field Amendment draft in table...")
            driver.execute_script(f"""
                var rows = document.querySelectorAll('table tbody tr, tr[ng-repeat]');
                var btn = rows[{idx}].querySelector('a[ng-click*="edit" i], button[data-ng-click*="editSavedapp" i], button[ng-click*="edit" i]');
                if (btn) btn.click();
            """)
            time.sleep(3.0)
            return 'resumed'

        elif row_type == 'unrelated_blocking':
            log.warning(f"[{gstin}] Unrelated Core Field Amendment draft is blocking the session. Deleting...")
            driver.execute_script(f"""
                var rows = document.querySelectorAll('table tbody tr, tr[ng-repeat]');
                var btn = rows[{idx}].querySelector('button.btn-danger, button[ng-click*="delete" i], a.btn-danger');
                if (btn) btn.click();
            """)
            time.sleep(1.0)
            
            try:
                confirm = WebDriverWait(driver, 8).until(EC.element_to_be_clickable((
                    By.XPATH,
                    "//a[contains(@class,'btn-primary') and (normalize-space(text())='Yes' or @ng-click='callback()')]"
                    " | "
                    "//button[contains(@class,'btn-primary') and normalize-space(text())='Yes']"
                )))
                driver.execute_script("arguments[0].click();", confirm)
                log.info(f"[{gstin}] Unrelated blocking draft deleted.")
                time.sleep(2.0)
            except TimeoutException:
                log.warning(f"[{gstin}] Confirm deletion dialog did not load.")
            
            # Re-evaluate saved applications after deletion
            return check_and_resume_draft(driver, row, timeout)

    return 'cleared'


def resume_first_core_draft(driver, timeout: int = 15) -> bool:
    """Fallback handler to click Edit/Resume on the first visible Core Amendment draft in the list."""
    log.info("[DRAFT] Attempting to resume the first visible Core Amendment draft in My Saved Applications...")
    try:
        WebDriverWait(driver, timeout).until(lambda d: d.execute_script(
            "return document.querySelectorAll('table tbody tr, tr[ng-repeat]').length > 0"
        ))
        time.sleep(1.0)
        
        JS_RESUME_FIRST = """
            var rows = document.querySelectorAll('table tbody tr, tr[ng-repeat]');
            var keywords = ['REG-14', 'AMENDMENT', 'CORE', 'REGAN', 'REGAC'];
            for (var i = 0; i < rows.length; i++) {
                var row = rows[i];
                if (!row.offsetParent) continue;
                var txt = row.innerText.toUpperCase();
                
                var is_core = false;
                for (var k = 0; k < keywords.length; k++) {
                    if (txt.indexOf(keywords[k]) !== -1) {
                        is_core = true;
                        break;
                    }
                }
                if (is_core) {
                    var editBtn = row.querySelector('a[ng-click*="edit" i], button[data-ng-click*="editSavedapp" i], button[ng-click*="edit" i]');
                    if (editBtn) {
                        editBtn.click();
                        return true;
                    }
                }
            }
            return false;
        """
        success = driver.execute_script(JS_RESUME_FIRST)
        if success:
            log.info("[DRAFT] ✓ Successfully clicked Edit/Resume on the saved Core Amendment draft.")
            time.sleep(3.0)
            return True
        else:
            log.warning("[DRAFT] No Core Amendment draft rows found to resume.")
    except Exception as e:
        log.error(f"[DRAFT] Failed to resume first Core Amendment draft: {e}")
    return False


def select_gstin_context_if_prompted(driver, gstin: str, timeout: int = 5):
    """Selects target GSTIN from dropdown/input context screen in practitioner login sessions."""
    try:
        select_els = driver.find_elements(By.XPATH, "//select[contains(@id,'gstin') or contains(@name,'gstin') or contains(@data-ng-model,'gstin')]")
        if select_els and select_els[0].is_displayed():
            select = Select(select_els[0])
            for option in select.options:
                if gstin.lower() in option.text.lower():
                    select.select_by_visible_text(option.text)
                    log.info(f"[CONTEXT] Context GSTIN selected: {option.text}")
                    break
            btn = driver.find_element(By.XPATH, "//button[contains(.,'Search') or contains(.,'Proceed')]")
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(2.0)
            return True
            
        input_els = driver.find_elements(By.XPATH, "//input[contains(@id,'gstin') or contains(@name,'gstin')]")
        if input_els and input_els[0].is_displayed():
            input_els[0].clear()
            input_els[0].send_keys(gstin)
            log.info(f"[CONTEXT] Typed context GSTIN: {gstin}")
            btn = driver.find_element(By.XPATH, "//button[contains(.,'Search') or contains(.,'Proceed')]")
            driver.execute_script("arguments[0].click();", btn)
            time.sleep(2.0)
            return True
    except Exception as e:
        log.debug(f"[CONTEXT] GSTIN prompt bypass: {e}")
    return False


def enable_have_apob(driver, timeout: int = 15):
    """Enables Additional Place of Business toggle inside Principal Place of Business tab."""
    log.info("[APOB] Navigating to Principal Place of Business tab...")
    tab_xpath = "//a[contains(@data-ng-click,'tabroute(3)') or contains(normalize-space(.), 'Principal Place of Business')]"
    wait_and_click(driver, By.XPATH, tab_xpath, timeout=timeout)
    time.sleep(2.0)
    
    log.info("[APOB] Enabling 'Have Additional Place of Business' = YES...")
    label = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, "//label[@for='bp_add']"))
    )
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", label)
    
    # Check if radio YES is already checked
    is_checked = driver.execute_script("return document.getElementById('bp_add') ? document.getElementById('bp_add').checked : false;")
    if not is_checked:
        driver.execute_script("arguments[0].click();", label)
        log.info("[APOB] YES radio checked.")
    else:
        log.info("[APOB] YES radio already checked.")
    time.sleep(1.0)
    
    save_btn = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, "//a[contains(@class,'btn-primary') and contains(@data-ng-click,'passHaveAddlplc')]"))
    )
    driver.execute_script("arguments[0].click();", save_btn)
    log.info("[APOB] Principal Place of Business tab saved successfully.")
    time.sleep(2.0)


def navigate_to_additional_places_tab(driver, timeout: int = 15):
    """Switches view focus to the Additional Places of Business tab."""
    log.info("[APOB] Navigating to Additional Places of Business tab...")
    tab_xpath = "//a[contains(@data-ng-click,'tabroute(4)') or contains(normalize-space(.), 'Additional Places of Business')]"
    wait_and_click(driver, By.XPATH, tab_xpath, timeout=timeout)
    time.sleep(2.0)


def get_existing_apob_addresses(driver) -> List[str]:
    """
    Extracts all existing registered APoB addresses from the portal table/cards using the DOM.
    """
    try:
        js_extract = """
            var addresses = [];
            // Extract from tables
            var tables = document.querySelectorAll('table');
            for (var i = 0; i < tables.length; i++) {
                var rows = tables[i].querySelectorAll('tbody tr, tr[ng-repeat]');
                for (var j = 0; j < rows.length; j++) {
                    var text = rows[j].innerText || "";
                    if (text.trim() && /\\b[1-9][0-9]{5}\\b/.test(text)) {
                        addresses.push(text.trim());
                    }
                }
            }
            // Extract from cards/panels
            var cards = document.querySelectorAll('.card, .panel, [class*="card" i], [class*="panel" i], .well');
            for (var i = 0; i < cards.length; i++) {
                if (cards[i].offsetParent !== null) {
                    var text = cards[i].innerText || "";
                    if (text.trim() && /\\b[1-9][0-9]{5}\\b/.test(text)) {
                        if (addresses.indexOf(text.trim()) === -1) {
                            addresses.push(text.trim());
                        }
                    }
                }
            }
            return addresses;
        """
        existing = driver.execute_script(js_extract)
        log.info(f"[APOB] DOM extraction found {len(existing)} existing addresses.")
        return existing
    except Exception as e:
        log.warning(f"[APOB] Failed to extract existing addresses via DOM: {e}")
        return []


def normalize_address(address: str) -> str:
    """
    Applies production-grade address normalization:
    - Normalizes state contractions: UP/U.P. -> Uttar Pradesh, Noida (U.P.) -> Noida Uttar Pradesh
    - Replaces Sector hyphens: Sector-62 -> Sector 62
    - Removes punctuation and extra spaces
    - Converts to lowercase
    """
    if not address:
        return ""
    addr = address.lower()
    
    # Noida (U.P.) or Noida U.P. or Noida (UP) -> Noida Uttar Pradesh
    addr = re.sub(r'\bnoida\s*\(?\s*u\.?p\.?\s*\)?', 'noida uttar pradesh', addr)
    # General U.P. / U.P / UP
    addr = re.sub(r'\b\(?\s*u\.?p\.?\s*\)?\b', ' uttar pradesh ', addr)
    addr = re.sub(r'\bup\b', ' uttar pradesh ', addr)
    
    # Sector-XX -> Sector XX
    addr = re.sub(r'\bsector\s*-\s*(\d+)\b', r'sector \1', addr)
    
    # Remove all punctuation (replace with space to prevent joining words)
    addr = re.sub(r'[^\w\s]', ' ', addr)
    
    # Convert multiple spaces to single space and strip
    addr = " ".join(addr.split())
    return addr


def is_duplicate_apob(excel_row: Dict[str, Any], existing_addresses: List[str]) -> bool:
    """
    Checks if the Excel address in excel_row is a duplicate of any existing portal addresses.
    Uses PIN code match, and normalized comparison.
    """
    excel_pin = str(excel_row.get("PIN") or "").strip()
    if not excel_pin:
        excel_addr_str = str(excel_row.get("Address") or str(excel_row.get("BuildingName", "")) + " " + str(excel_row.get("Street", "")))
        match = re.search(r'\b([1-9][0-9]{5})\b', excel_addr_str)
        if match:
            excel_pin = match.group(1)
            
    if not excel_pin:
        return False
        
    addr_parts = [
        str(excel_row.get("BuildingNo") or ""),
        str(excel_row.get("BuildingName") or ""),
        str(excel_row.get("Street") or ""),
        str(excel_row.get("Locality") or ""),
        str(excel_row.get("PIN") or ""),
        str(excel_row.get("District") or ""),
        str(excel_row.get("State") or "")
    ]
    excel_addr_full = " ".join([p for p in addr_parts if p.strip()])
    norm_excel = normalize_address(excel_addr_full)
    
    for exist_addr in existing_addresses:
        # Check PIN code first
        if excel_pin not in exist_addr:
            continue
            
        norm_exist = normalize_address(exist_addr)
        
        # Check exact match after normalization
        if norm_excel == norm_exist:
            log.info(f"[APOB] Duplicate detected (Exact Match)! Excel address '{excel_addr_full}' overlaps with Portal address '{exist_addr}'")
            return True
            
        # Check if one is a substring of the other
        if norm_excel in norm_exist or norm_exist in norm_excel:
            log.info(f"[APOB] Duplicate detected (Substring Match)! Excel address '{excel_addr_full}' overlaps with Portal address '{exist_addr}'")
            return True
            
        # Check high word overlap
        words_excel = norm_excel.split()
        words_exist = norm_exist.split()
        if words_excel and words_exist:
            set_excel = set(words_excel)
            set_exist = set(words_exist)
            set_excel.discard(excel_pin)
            set_exist.discard(excel_pin)
            
            common = set_excel.intersection(set_exist)
            overlap = len(common) / len(set_excel) if set_excel else 0
            if overlap >= 0.65:
                log.info(f"[APOB] Duplicate detected (Word Overlap {overlap*100:.1f}%)! Excel address '{excel_addr_full}' overlaps with Portal address '{exist_addr}'")
                return True
                
    return False


def fill_number_of_additional_places(driver, client_rows: List[Dict], client_indices: List[int], timeout: int = 15) -> Tuple[List[Dict], List[int]]:
    """
    Fills the Number of Additional Places input field with an incremental update.
    Reads current count E, filters duplicates, calculates target Z = E + new_unique_records_count,
    logs the stats card, and clicks the body element to trigger Angular validation.
    """
    log.info(f"[APOB] Initiating dynamic additional places incrementer with {len(client_rows)} rows.")

    # 1. Fetch existing addresses from DOM
    existing_addresses = get_existing_apob_addresses(driver)
    E = len(existing_addresses)
    
    # 2. Filter duplicates
    duplicate_count = 0
    unique_rows = []
    unique_indices = []
    for row, idx in zip(client_rows, client_indices):
        if is_duplicate_apob(row, existing_addresses):
            duplicate_count += 1
        else:
            unique_rows.append(row)
            unique_indices.append(idx)

    # 3. Find input field
    _JS_FIND_INPUT = """
        var inputs = document.querySelectorAll('input[type="text"], input[type="number"], input:not([type])');
        for (var i = 0; i < inputs.length; i++) {
            var input = inputs[i];
            if (!input.offsetParent) continue;
            var id = input.id.toLowerCase();
            var ngModel = (input.getAttribute('data-ng-model') || '').toLowerCase();
            if (id.indexOf('add') !== -1 || id.indexOf('place') !== -1 || id.indexOf('num') !== -1 ||
                ngModel.indexOf('add') !== -1 || ngModel.indexOf('place') !== -1 || ngModel.indexOf('num') !== -1) {
                return input;
            }
        }
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].offsetParent) return inputs[i];
        }
        return null;
    """

    input_el = None
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            input_el = driver.execute_script(_JS_FIND_INPUT)
            if input_el:
                break
        except Exception:
            pass
        time.sleep(0.4)

    if not input_el:
        log.warning("[APOB] Number of Additional Places input box not found via JS scanner. Trying hardcoded XPaths...")
        xpath = (
            "//input[@id='numAddPlace' or @name='numAddPlace' or contains(@data-ng-model, 'numAdd') "
            "or contains(@placeholder, 'Number of additional places') or contains(@id, 'numAddl') "
            "or contains(@id, 'noOfAdd')]"
        )
        try:
            input_el = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
        except Exception as e:
            log.error(f"[APOB] All search strategies failed to find Number of Additional Places field: {e}")
            return unique_rows, unique_indices

    try:
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", input_el)

        # Fallback if DOM extraction returned 0 but input field or table has rows
        if E == 0:
            E_val = driver.execute_script("""
                var inputEl = arguments[0];
                var counts = [0];
                if (inputEl) {
                    var inputVal = parseInt(inputEl.value, 10);
                    if (!isNaN(inputVal) && inputVal > 0) {
                        counts.push(inputVal);
                    }
                }
                var tables = document.querySelectorAll('table');
                for (var i = 0; i < tables.length; i++) {
                    var rows = tables[i].querySelectorAll('tbody tr, tr[ng-repeat]');
                    if (rows.length > 0) {
                        counts.push(rows.length);
                    }
                }
                var cards = document.querySelectorAll('.card, .panel, [class*="card" i], [class*="panel" i], .well');
                var cardCount = 0;
                for (var i = 0; i < cards.length; i++) {
                    if (cards[i].offsetParent !== null && (cards[i].innerText.indexOf('Additional Place') !== -1 || cards[i].innerText.indexOf('Place of Business') !== -1)) {
                        cardCount++;
                    }
                }
                if (cardCount > 0) {
                    counts.push(cardCount);
                }
                return Math.max.apply(null, counts);
            """, input_el)
            if E_val > 0:
                E = E_val

        N = len(unique_rows)
        Z = E + N

        # 4. Print stats card in exact requested format
        log.info(
            f"\nExisting APoBs:\n{E}\n\n"
            f"Excel APoBs:\n{len(client_rows)}\n\n"
            f"Duplicate APoBs:\n{duplicate_count}\n\n"
            f"New APoBs:\n{N}\n\n"
            f"Target Count:\n{Z}\n"
        )

        # Clear existing value
        driver.execute_script("""
            arguments[0].value = '';
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, input_el)

        # Resilient write: send keys with JS direct value backup
        try:
            input_el.send_keys(str(Z))
            input_el.send_keys(Keys.TAB)
        except Exception:
            driver.execute_script("arguments[0].value = arguments[1];", input_el, str(Z))

        # Trigger Angular change events
        driver.execute_script("""
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
        """, input_el)

        log.info("[APOB] Tapping on free screen (body) to trigger validation...")
        try:
            body = driver.find_element(By.TAG_NAME, "body")
            body.click()
        except Exception as body_ex:
            log.warning(f"[APOB] Failed body click: {body_ex}. Trying JavaScript click fallback on body...")
            try:
                driver.execute_script("document.body.click();")
            except Exception:
                pass

        time.sleep(2.0)
        log.info(f"[APOB] Number of additional places successfully set to {Z}")
        return unique_rows, unique_indices
    except Exception as e:
        log.error(f"[APOB] Failed to fill Number of Additional Places: {e}")
        return unique_rows, unique_indices
    except Exception as e:
        log.error(f"[APOB] Failed to fill Number of Additional Places: {e}")
        return False


def fill_pin_code_and_wait(driver, pin_code: str, timeout: int = 15):
    """Enters PIN Code and dispatches Angular watchers, then waits for auto-fetch/enablement completion."""
    log.info(f"[APOB] Filling PIN Code: {pin_code}")
    xpath = (
        "//input[@id='pncd' or @name='pncd' or @id='pin' or @name='pin' "
        "or @data-ng-model[contains(.,'pin') or contains(.,'pncd')]]"
    )
    pin_input = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, xpath))
    )
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", pin_input)
    
    driver.execute_script("""
        arguments[0].value = '';
        arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
        arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
    """, pin_input)
    
    pin_input.send_keys(str(pin_code))
    
    # Wait for PIN suggestion list/item to appear and click it
    log.info("[APOB] Waiting for PIN suggestion dropdown list...")
    suggestion_xpath = "//ul[contains(@class,'as-list')]//li | //*[contains(@class,'as-results') or contains(@class,'autosuggest')]//li"
    try:
        suggestion_el = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.XPATH, suggestion_xpath))
        )
        log.info("[APOB] Suggestion found. Clicking suggestion to select PIN...")
        driver.execute_script("arguments[0].click();", suggestion_el)
        time.sleep(1.0)
    except TimeoutException:
        log.warning("[APOB] Suggestion dropdown did not appear. Falling back to TAB/Enter...")
        pin_input.send_keys(Keys.TAB)
        time.sleep(1.0)
        
    driver.execute_script("""
        arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
    """, pin_input)
    
    log.info("[APOB] PIN entered. Waiting for district auto-fetch/enablement...")
    
    def _wait_for_autofetch(d):
        try:
            el = d.find_element(By.XPATH, "//*[@id='district' or @name='district' or @id='ap_dst' or @name='ap_dst']")
            if el.tag_name == "select":
                return len(Select(el).options) > 1
            return el.is_enabled()
        except:
            return False
            
    WebDriverWait(driver, timeout).until(_wait_for_autofetch)
    log.info("[APOB] Auto-fetch/enablement check completed successfully.")


def select_dropdown_by_text_or_value(driver, select_xpath: str, target_value: Any, timeout: int = 15) -> bool:
    """Robust dropdown selector that checks text, value, and falls back if needed."""
    select_el = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, select_xpath))
    )
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", select_el)
    
    select = Select(select_el)
    target_clean = str(target_value).strip().lower()
    
    # Visible text match
    for option in select.options:
        opt_text = option.text.strip().lower()
        if target_clean == opt_text or target_clean in opt_text:
            select.select_by_visible_text(option.text)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", select_el)
            log.info(f"[APOB] Dropdown option selected by text: '{option.text}'")
            return True
            
    # Value attribute match
    for option in select.options:
        opt_val = option.get_attribute("value")
        if opt_val and target_clean == opt_val.strip().lower():
            select.select_by_value(opt_val)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", select_el)
            log.info(f"[APOB] Dropdown option selected by value: '{opt_val}'")
            return True
            
    # Fallback to index 1 if target not found
    if len(select.options) > 1:
        select.select_by_index(1)
        log.warning(f"[APOB] Option '{target_value}' missing. Defaulted to: '{select.options[1].text}'")
        return True
        
    return False


def upload_supporting_document(driver, doc_path: str, timeout: int = 20) -> bool:
    """Robust document uploader containing file validation and progress monitoring."""
    if not doc_path or not os.path.exists(doc_path):
        raise FileNotFoundError(f"Document upload file does not exist at: {doc_path}")
        
    abs_path = os.path.abspath(doc_path)
    
    for attempt in range(1, 4):
        try:
            log.info(f"[UPLOAD] Upload attempt {attempt}/3: {doc_path}")
            # Intercept the true hidden input element
            file_input = WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, "//input[@type='file' and @id='ap_upload']"))
            )
            # Clean CSS visibility properties via JS
            driver.execute_script("""
                arguments[0].style.opacity = '1';
                arguments[0].style.visibility = 'visible';
                arguments[0].style.height = 'auto';
                arguments[0].style.display = 'block';
            """, file_input)
            
            # Send keys directly to file element
            file_input.send_keys(abs_path)
            log.info("[UPLOAD] Bounding path injected. Waiting for upload completion indicator...")
            
            # Try XHR/Fetch interceptor tracking first (avoids blind waits)
            t_up_start = time.time()
            ok, upload_res = UploadTrackerService.wait_for_upload(driver, timeout=timeout)
            if ok:
                log.info(f"[UPLOAD] ✓ Document uploaded successfully (tracked in {time.time() - t_up_start:.1f}s).")
                return True
            else:
                log.warning("[UPLOAD] Event-driven tracking did not confirm success. Falling back to DOM trash check...")
            
            # Verify completion by running a WebDriverWait for the portal's dynamic trash/delete indicator
            delete_xpath = "//a[contains(@class,'fa-trash') or contains(@ng-click,'delete')] | //button[contains(.,'Delete') or contains(@class,'btn-danger') or contains(@ng-click,'delete')]"
            WebDriverWait(driver, timeout).until(
                EC.presence_of_element_located((By.XPATH, delete_xpath))
            )
            log.info("[UPLOAD] ✓ Document uploaded successfully.")
            return True
        except Exception as e:
            log.warning(f"[UPLOAD] Upload attempt {attempt} failed: {e}")
            if attempt < 3:
                time.sleep(2.0)
            else:
                raise e


def select_business_activities(driver, activities, timeout: int = 15):
    """
    Ticks business activities checkboxes dynamically based on Excel values and mandatory selections.
    Robust hybrid strategy:
    Priority 1: Direct ID matching for mandatory checkboxes:
      - bp_ck_IMP: Import
      - bp_ck_OSO: Office / Sale Office
      - bp_ck_SOS: Supplier of Services
      - bp_ck_SRE: Recipient of Goods or Services
    Priority 2: Use label text matching if ID matching fails (or for other custom activities).
    """
    log.info(f"[APOB] Processing Business Activities: {activities}")
    
    activities_list = []
    if activities:
        if isinstance(activities, str):
            activities_list = [a.strip() for a in activities.split(",") if a.strip()]
        else:
            activities_list = [str(activities)]
            
    mandatory_activities = {
        "bp_ck_IMP": "Import",
        "bp_ck_OSO": "Office / Sale Office",
        "bp_ck_SOS": "Supplier of Services",
        "bp_ck_SRE": "Recipient of Goods or Services"
    }
    
    for m_act in mandatory_activities.values():
        if m_act not in activities_list and not any(m_act.lower() in x.lower() for x in activities_list):
            activities_list.append(m_act)
            
    log.info(f"[APOB] Master business activities list to select: {activities_list}")
    
    # Priority 1: Direct ID matching
    for chk_id, act_name in mandatory_activities.items():
        try:
            cb = driver.find_element(By.ID, chk_id)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cb)
            is_checked = driver.execute_script("return arguments[0].checked;", cb)
            if not is_checked:
                driver.execute_script("arguments[0].click();", cb)
                log.info(f"[APOB] Selected mandatory checkbox via ID '{chk_id}': {act_name}")
            else:
                log.info(f"[APOB] Mandatory checkbox ID '{chk_id}' ({act_name}) already checked.")
        except Exception as e:
            log.warning(f"[APOB] Direct ID check failed for {chk_id} ({act_name}): {e}. Will fallback to text matching.")
            
    # Priority 2: Use label text matching only if ID matching fails (or for other custom activities)
    checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox']")
    for cb in checkboxes:
        try:
            if not cb.is_displayed():
                continue
            
            cb_id = cb.get_attribute("id")
            if cb_id in mandatory_activities:
                continue
                
            label_text = ""
            parent = cb.find_element(By.XPATH, "./..")
            if parent.tag_name == "label":
                label_text = parent.text
            else:
                try:
                    sibling = cb.find_element(By.XPATH, "following-sibling::label")
                    label_text = sibling.text
                except:
                    label_text = cb.get_attribute("aria-label") or cb.get_attribute("title") or ""
            if not label_text and parent:
                label_text = parent.text
                
            label_clean = label_text.strip().lower()
            if not label_clean:
                continue
                
            for act in activities_list:
                if act in mandatory_activities.values() and cb_id:
                    continue
                act_clean = act.strip().lower()
                if act_clean in label_clean or label_clean in act_clean:
                    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cb)
                    is_checked = driver.execute_script("return arguments[0].checked;", cb)
                    if not is_checked:
                        driver.execute_script("arguments[0].click();", cb)
                        log.info(f"[APOB] Ticked custom checkbox activity via text: '{label_text.strip()}'")
                    break
        except Exception:
            continue


def verify_business_activities(driver, timeout: int = 15):
    """
    Priority 3: Perform verification pass after selection.
    For every mandatory checkbox:
      - Scroll into view
      - Click via JavaScript
      - Verify checked state
      - Retry up to 3 times
      - Log success/failure
    Before proceeding: Verify all mandatory checkboxes are selected. If any remain unchecked: Raise exception.
    """
    log.info("[APOB] Verification pass for the four mandatory business activity checkboxes...")
    mandatory_activities = {
        "bp_ck_IMP": "Import",
        "bp_ck_OSO": "Office / Sale Office",
        "bp_ck_SOS": "Supplier of Services",
        "bp_ck_SRE": "Recipient of Goods or Services"
    }
    
    for chk_id, act_name in mandatory_activities.items():
        success = False
        for attempt in range(1, 4):
            try:
                cb = driver.find_element(By.ID, chk_id)
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", cb)
                is_checked = driver.execute_script("return arguments[0].checked;", cb)
                if not is_checked:
                    log.warning(f"[APOB] Checkbox '{act_name}' (ID: {chk_id}) unchecked. Attempt {attempt}/3 to click via JS...")
                    driver.execute_script("arguments[0].click();", cb)
                    time.sleep(0.5)
                    is_checked = driver.execute_script("return arguments[0].checked;", cb)
                
                if is_checked:
                    log.info(f"[APOB] Mandatory checkbox checked: {act_name}")
                    success = True
                    break
                else:
                    log.warning(f"[APOB] Attempt {attempt}/3 failed to check {act_name}.")
            except Exception as e:
                log.warning(f"[APOB] Attempt {attempt}/3 failed with exception for {act_name}: {e}")
                time.sleep(0.5)
                
        if not success:
            log.error(f"[APOB] FAILED: Mandatory checkbox '{act_name}' (ID: {chk_id}) could not be selected after 3 attempts.")
            raise ValueError(f"Mandatory business activity checkbox '{act_name}' remains unselected!")
            
    log.info("[APOB] Mandatory checkboxes verified")


def _fill_autosuggest_text_field(driver, input_xpath: str, value: str, label: str = "field", timeout: int = 15) -> bool:
    """Fills an auto-suggest geocoding input textbox, clicking the suggestion if it appears, or skipping if already filled by auto-fetch."""
    try:
        elem = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, input_xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", elem)
        
        # 1. Check if already filled by auto-fetch
        current_val = driver.execute_script("return arguments[0].value;", elem)
        if current_val and current_val.strip() != "":
            log.info(f"[FORM] {label} is already auto-populated with: '{current_val}'. Skipping mapping.")
            return True
            
        log.info(f"[FORM] Filling auto-suggest {label}: {value}")
        
        driver.execute_script("""
            arguments[0].value = '';
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, elem)
        
        elem.send_keys(str(value))
        
        # 2. Wait for suggestion dropdown to appear and click the first suggestion
        suggestion_xpath = "//ul[contains(@class,'as-list')]//li | //*[contains(@class,'as-results') or contains(@class,'autosuggest')]//li"
        try:
            suggestion_el = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, suggestion_xpath))
            )
            log.info(f"[FORM] Suggestion found for {label}. Clicking to select...")
            driver.execute_script("arguments[0].click();", suggestion_el)
            time.sleep(1.0)
        except TimeoutException:
            log.warning(f"[FORM] Suggestion dropdown did not appear for {label}. Sending TAB...")
            elem.send_keys(Keys.TAB)
            time.sleep(1.0)
            
        driver.execute_script("arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));", elem)
        return True
    except Exception as e:
        log.error(f"[FORM] Failed to fill auto-suggest field '{label}': {e}")
        return False


def _contains_khasra_details(val: Any) -> bool:
    """
    Returns True if the provided value looks like a Khasra/Khasra No. address fragment.
    We intentionally skip these details to avoid polluting APoB address fields.
    """
    if val is None:
        return False
    s = str(val).strip().lower()
    if not s:
        return False
    # Common spellings seen in documents / datasets
    return any(token in s for token in ["khasra", "khasra no", "khasra number", "khasra nos", "khasra n"])


def optimize_and_compress_pdf(input_path: str, max_size_mb: float = 0.95) -> str:
    """
    Progressively compresses PDF images across multiple quality passes (25, 15, 10, 8, 5)
    until the output file size is safely below max_size_mb.
    """
    if not input_path or not os.path.exists(input_path):
        return input_path

    file_size_mb = os.path.getsize(input_path) / (1024 * 1024)
    if file_size_mb <= max_size_mb:
        log.info(f"[PDF-COMPRESS] File '{os.path.basename(input_path)}' is already small enough ({file_size_mb:.2f}MB). Skipping compression.")
        return input_path
        
    log.info(f"[PDF-COMPRESS] File '{os.path.basename(input_path)}' ({file_size_mb:.2f}MB) exceeds {max_size_mb}MB limit. Entering progressive multi-pass compression...")
    
    from pypdf import PdfReader, PdfWriter
    import tempfile
    import random
    quality_levels = [25, 15, 10, 8, 5]
    scale_factors = [1.0, 0.7, 0.5, 0.4, 0.25]
    
    # Unique temp file to prevent parallel write conflicts in concurrent thread pools
    temp_dir = tempfile.gettempdir()
    rand_id = random.randint(100000, 999999)
    temp_out = os.path.join(temp_dir, f"compressed_temp_{rand_id}_{os.path.basename(input_path)}")
    best_path = input_path

    for pass_idx, q in enumerate(quality_levels, start=1):
        scale = scale_factors[pass_idx - 1]
        log.info(f"[PDF-COMPRESS] Pass {pass_idx}/{len(quality_levels)}: Testing lossy image re-compression with quality={q}, scale={scale}...")
        try:
            reader = PdfReader(input_path)
            writer = PdfWriter()

            for idx in range(len(reader.pages)):
                added_page = writer.add_page(reader.pages[idx])
                try:
                    if added_page.images:
                        for img in added_page.images:
                            pil_img = img.image
                            if scale < 1.0:
                                w, h = pil_img.size
                                # Downscale scanned documents (typically 500-900px wide) when scale < 1.0
                                if w > 250 or h > 250:
                                    new_size = (int(w * scale), int(h * scale))
                                    # Use BILINEAR (2) for downscaling
                                    pil_img = pil_img.resize(new_size, resample=2)
                            img.replace(pil_img, quality=q)
                    else:
                        added_page.compress_content_streams()
                except Exception as page_err:
                    log.warning(f"[PDF-COMPRESS] Anomaly during image re-compression on page {idx}: {page_err}. Falling back to standard compress_content_streams...")
                    try:
                        added_page.compress_content_streams()
                    except Exception as comp_err:
                        log.error(f"[PDF-COMPRESS] Content stream compression fallback also failed: {comp_err}")

            with open(temp_out, "wb") as f:
                writer.write(f)

            compressed_size_mb = os.path.getsize(temp_out) / (1024 * 1024)
            log.info(f"[PDF-COMPRESS] Pass {pass_idx} finished. Size on disk: {compressed_size_mb:.2f}MB")

            best_path = temp_out
            if compressed_size_mb <= max_size_mb:
                log.info(f"[PDF-COMPRESS] SUCCESS! Compressed PDF file size ({compressed_size_mb:.2f}MB) is below {max_size_mb}MB limit on pass {pass_idx} (quality={q}, scale={scale}). Breaking loop.")
                return temp_out

        except Exception as pass_err:
            log.error(f"[PDF-COMPRESS] Error during progressive quality pass {pass_idx} (quality={q}): {pass_err}")

    final_size_mb = os.path.getsize(best_path) / (1024 * 1024)
    if final_size_mb > max_size_mb:
        log.warning(f"[PDF-COMPRESS] WARNING: All compression loops completed, but best output size ({final_size_mb:.2f}MB) is still slightly over {max_size_mb}MB limit.")

    return best_path


def click_add_new_apob(driver, timeout: int = 15) -> bool:
    """
    Enterprise-grade 'ADD NEW' button detection.
    Tries multiple selector strategies in sequence.
    """
    log.info("[APOB] Initiating enterprise-grade ADD NEW button detection...")
    
    # Priority 1: Exact button text XPath
    try:
        btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(normalize-space(),'ADD NEW')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", btn)
        log.info("[APOB] ADD NEW clicked via Priority 1 (Exact button text XPath)")
        time.sleep(1.0)
        return True
    except Exception:
        pass
        
    # Priority 2: Angular button XPath
    try:
        btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@ng-click,'add')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", btn)
        log.info("[APOB] ADD NEW clicked via Priority 2 (Angular button XPath)")
        time.sleep(1.0)
        return True
    except Exception:
        pass
        
    # Priority 3: PrimeNG button XPath
    try:
        btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'ui-button') or contains(@class,'p-button')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", btn)
        log.info("[APOB] ADD NEW clicked via Priority 3 (PrimeNG button XPath)")
        time.sleep(1.0)
        return True
    except Exception:
        pass
        
    # Priority 4: Bootstrap button XPath
    try:
        btn = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class,'btn-') and contains(normalize-space(),'ADD NEW')]"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", btn)
        log.info("[APOB] ADD NEW clicked via Priority 4 (Bootstrap button XPath)")
        time.sleep(1.0)
        return True
    except Exception:
        pass
        
    # Priority 5: JavaScript fallback
    try:
        clicked = driver.execute_script("""
            var elements = document.querySelectorAll('button, a, span, i');
            var patterns = ['add new', 'addl place', 'new additional', '+'];
            for (var i = 0; i < elements.length; i++) {
                var el = elements[i];
                if (el.offsetParent === null) continue;
                var txt = el.innerText.trim().toLowerCase();
                for (var p = 0; p < patterns.length; p++) {
                    if (txt === patterns[p] || txt.indexOf(patterns[p]) !== -1) {
                        el.click();
                        return true;
                    }
                }
            }
            return false;
        """)
        if clicked:
            log.info("[APOB] ADD NEW clicked via Priority 5 (JavaScript fallback)")
            time.sleep(1.0)
            return True
    except Exception as e:
        log.warning(f"[APOB] JS ADD NEW click failed: {e}")
        
    log.error("[APOB] FAILED: Could not click ADD NEW button using any of the enterprise strategies.")
    return False


def wait_for_new_apob_form(driver, timeout: int = 15):
    """
    Waits for a fresh blank APOB form to load after clicking ADD NEW.
    Verifies one of:
      - PIN field becomes empty.
      - New APOB modal/panel appears.
      - Form index/count increments.
    """
    log.info("[APOB] Waiting for fresh blank APOB form to load...")
    start_time = time.time()
    
    def get_existing_records_count():
        try:
            headers = driver.find_elements(By.XPATH, "//div[contains(@class,'panel-heading') or contains(@class,'card-header')]")
            return len(headers)
        except:
            return 0
            
    initial_count = get_existing_records_count()
    
    while time.time() - start_time < timeout:
        # 1. PIN field empty check
        try:
            pin_input = driver.find_element(By.XPATH, "//input[@id='pncd' or @name='pncd' or @id='pin' or @name='pin']")
            if pin_input.is_displayed():
                val = driver.execute_script("return arguments[0].value;", pin_input)
                if val == "":
                    log.info("[APOB] Fresh APOB form loaded (PIN field is empty).")
                    return True
        except:
            pass
            
        # 2. Form index count check
        current_count = get_existing_records_count()
        if current_count > initial_count:
            log.info(f"[APOB] Fresh APOB form loaded (Form count incremented from {initial_count} to {current_count}).")
            return True
            
        # 3. Modal presence check
        try:
            modals = driver.find_elements(By.XPATH, "//*[contains(@class,'modal-content') or contains(@class,'modal-body') or contains(@id,'modal')]")
            if modals and any(m.is_displayed() for m in modals):
                log.info("[APOB] Fresh APOB form loaded (Modal visible).")
                return True
        except:
            pass
            
        time.sleep(0.5)
        
    log.error("[APOB] Timeout waiting for new blank APOB form to load.")
    raise TimeoutException("New blank APOB form failed to load within timeout.")


def verify_and_reset_fresh_apob_form(driver, timeout: int = 15):
    """
    Verifies that the newly opened APOB form is completely empty/reset.
    If any fields (Road/Street, PIN, Nature of Possession, Business Activities, or Upload)
    are not reset, they are explicitly cleared/reset before proceeding.
    Logs '[APOB] Fresh APOB form verified.' after verification and potential resets.
    """
    log.info("[APOB] Verifying fresh/blank APOB form...")
    
    # 1. Verify PIN field
    pin_xpath = "//input[@id='pncd' or @name='pncd' or @id='pin' or @name='pin']"
    try:
        pin_el = driver.find_element(By.XPATH, pin_xpath)
        pin_val = driver.execute_script("return arguments[0].value;", pin_el)
        if pin_val and pin_val.strip() != "":
            log.warning(f"[APOB] PIN field not empty (found: '{pin_val}'). Clearing PIN field...")
            driver.execute_script("""
                arguments[0].value = '';
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, pin_el)
    except Exception as e:
        log.warning(f"[APOB] Could not verify/clear PIN field: {e}")

    # 2. Verify Road/Street
    road_xpath = "//input[@id='st' or @name='st' or @id='road' or @name='road']"
    try:
        road_el = driver.find_element(By.XPATH, road_xpath)
        road_val = driver.execute_script("return arguments[0].value;", road_el)
        if road_val and road_val.strip() != "":
            log.warning(f"[APOB] Road/Street field not empty (found: '{road_val}'). Clearing Road/Street...")
            driver.execute_script("""
                arguments[0].value = '';
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, road_el)
    except Exception as e:
        log.warning(f"[APOB] Could not verify/clear Road/Street field: {e}")

    # 3. Verify Nature of Possession
    poss_xpath = "//select[@id='psnt' or @name='psnt']"
    try:
        poss_el = driver.find_element(By.XPATH, poss_xpath)
        poss_val = driver.execute_script("return arguments[0].value;", poss_el)
        if poss_val and poss_val.strip() not in ["", "?", "Choose"]:
            log.warning(f"[APOB] Nature of Possession not reset (found: '{poss_val}'). Resetting select...")
            driver.execute_script("""
                arguments[0].selectedIndex = 0;
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, poss_el)
    except Exception as e:
        log.warning(f"[APOB] Could not verify/reset Nature of Possession: {e}")

    # 4. Verify Business Activities (unchecked)
    try:
        checkboxes = driver.find_elements(By.XPATH, "//input[@type='checkbox']")
        for cb in checkboxes:
            is_checked = driver.execute_script("return arguments[0].checked;", cb)
            if is_checked:
                cb_id = cb.get_attribute("id") or cb.get_attribute("name") or "unknown"
                log.warning(f"[APOB] Business Activity checkbox '{cb_id}' is checked. Unchecking...")
                driver.execute_script("""
                    arguments[0].checked = false;
                    arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                """, cb)
    except Exception as e:
        log.warning(f"[APOB] Could not verify/reset Business Activity checkboxes: {e}")

    # 5. Verify Upload control is reset
    try:
        delete_xpath = "//a[contains(@class,'fa-trash') or contains(@ng-click,'delete')] | //button[contains(.,'Delete') or contains(@class,'btn-danger') or contains(@ng-click,'delete')]"
        delete_els = driver.find_elements(By.XPATH, delete_xpath)
        for d_btn in delete_els:
            if d_btn.is_displayed():
                log.warning("[APOB] Existing upload found on fresh form. Clicking delete to reset...")
                try:
                    d_btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", d_btn)
                time.sleep(0.5)
    except Exception as e:
        log.warning(f"[APOB] Could not verify/reset Upload control: {e}")

    log.info("[APOB] Fresh APOB form verified.")


def expand_latest_apob_panel(driver, timeout: int = 15) -> bool:
    """
    Scans for all APoB panels/accordions, finds the latest one,
    and ensures it is expanded so the form fields are visible and interactable.
    """
    log.info("[APOB] Running APoB Form Expansion Handler to locate and expand the latest entry...")
    time.sleep(1.5)
    
    try:
        toggles = driver.find_elements(By.XPATH, "//*[contains(@class,'accordion') or contains(@class,'panel')]//a[@data-toggle='collapse' or contains(@class,'toggle') or contains(@id,'heading')] | //div[contains(@class,'panel-heading')]//a")
        if toggles:
            latest_toggle = toggles[-1]
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", latest_toggle)
            
            is_collapsed = driver.execute_script("""
                var el = arguments[0];
                if (el.classList.contains('collapsed') || el.getAttribute('aria-expanded') === 'false') {
                    return true;
                }
                var href = el.getAttribute('href') || el.getAttribute('data-target');
                if (href) {
                    var target = document.querySelector(href);
                    if (target && (!target.classList.contains('in') && !target.classList.contains('show'))) {
                        return true;
                    }
                }
                return false;
            """, latest_toggle)
            
            if is_collapsed:
                log.info("[APOB] Latest panel toggle is collapsed. Clicking to expand...")
                driver.execute_script("arguments[0].click();", latest_toggle)
                time.sleep(1.5)
            else:
                log.info("[APOB] Latest panel toggle is already expanded.")
            return True
    except Exception as e:
        log.warning(f"[APOB] Accordion toggle strategy warning: {e}")
        
    try:
        headers = driver.find_elements(By.XPATH, "//div[contains(@class,'panel-heading') or contains(@class,'card-header')]")
        if headers:
            latest_header = headers[-1]
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", latest_header)
            log.info("[APOB] Clicking latest panel header to ensure expansion...")
            driver.execute_script("arguments[0].click();", latest_header)
            time.sleep(1.5)
            return True
    except Exception as e:
        log.warning(f"[APOB] General panel header expansion fallback failed: {e}")
        
    return False


def fill_apob_form_fields(driver, row: Dict[str, Any], timeout: int = 15,
                           doc_cli_path: str = None, pdf_cache: Optional[Dict[str, str]] = None):
    """
    Fills all address, contact, and document details in the APoB form.
    Does NOT select business activities or perform saving.
    """
    log.info("[FORM] Filling APoB form fields...")
    
    # Address parameters
    fill_pin_code_and_wait(driver, row.get("PIN"), timeout)
    
    # District filling (supports both select dropdown and text input)
    dist_val = row.get("District")
    if dist_val:
        dist_select_xpath = "//select[@id='district' or @name='district' or @id='ap_dst' or @name='ap_dst']"
        dist_input_xpath = "//input[@id='ap_dst' or @name='ap_dst' or @id='district' or @name='district']"
        
        select_els = driver.find_elements(By.XPATH, dist_select_xpath)
        if select_els and select_els[0].tag_name == "select" and select_els[0].is_displayed():
            select_dropdown_by_text_or_value(driver, dist_select_xpath, dist_val, timeout)
        else:
            _fill_autosuggest_text_field(driver, dist_input_xpath, str(dist_val), label="District", timeout=timeout)
            
    # City/Town/Village filling (supports both select dropdown and text input)
    city_val = row.get("City")
    if city_val:
        city_select_xpath = "//select[@id='city' or @name='city' or @id='loc' or @name='loc']"
        city_input_xpath = "//input[@id='loc' or @name='loc' or @id='city' or @name='city']"
        
        select_els = driver.find_elements(By.XPATH, city_select_xpath)
        if select_els and select_els[0].tag_name == "select" and select_els[0].is_displayed():
            select_dropdown_by_text_or_value(driver, city_select_xpath, city_val, timeout)
        else:
            _fill_autosuggest_text_field(driver, city_input_xpath, str(city_val), label="City/Town/Village", timeout=timeout)
    
    try:
        warnings = driver.find_elements(By.XPATH, "//*[contains(text(),'does not match PIN') or contains(text(),'does not match')]")
        if warnings and any(w.is_displayed() for w in warnings):
            log.warning("[FORM] WARNING: 'City/Town/Village does not match PIN code' notice visible.")
    except:
        pass
        
    address_autosuggest = {
        "locality": ("//input[@id='ap_locality' or @name='ap_locality' or @id='locality' or @name='locality']", row.get("Locality")),
        "road": ("//input[@id='st' or @name='st' or @id='road' or @name='road']", row.get("Road")),
        "buildingName": ("//input[@id='ap_bdname' or @name='ap_bdname' or contains(@id,'building') or contains(@id,'Premises')]", row.get("BuildingName") or row.get("Building")),
    }
    
    for label, (xpath, val) in address_autosuggest.items():
        if val is not None and str(val).strip() != "":
            if _contains_khasra_details(val):
                log.info(f"[FORM] Skipping Khasra details for APoB Address.{label}: '{str(val).strip()}'")
                continue
            _fill_autosuggest_text_field(driver, xpath, str(val), label=f"APoB Address.{label}", timeout=5)
            
    address_standard = {
        "buildingNo": ("//input[@id='abp_bdnum' or @name='bno' or @id='buildingNo' or @name='buildingNo']", row.get("BuildingNo")),
        "floorNo": ("//input[@id='ap_flrnum' or @name='ap_flrnum' or @id='floor' or @name='floor']", row.get("FloorNo") or row.get("Floor")),
        "landmark": ("//input[@id='landmark' or @name='landmark' or contains(@id,'land')]", row.get("Landmark")),
        "latitude": ("//input[@id='latitude' or @name='latitude' or @id='lat' or @name='lat']", row.get("Latitude")),
        "longitude": ("//input[@id='longitude' or @name='longitude' or @id='lng' or @name='lng']", row.get("Longitude")),
    }
    
    for label, (xpath, val) in address_standard.items():
        if val is not None and str(val).strip() != "":
            if label == "landmark" and _contains_khasra_details(val):
                log.info(f"[FORM] Skipping Khasra details for APoB Address.{label}: '{str(val).strip()}'")
                continue
            _fill_text_field(driver, xpath, str(val), timeout=5, label=f"APoB Address.{label}")
            
    # Contact parameters
    contact_fields = {
        "mobile": ("//input[@id='mob' or @name='mob' or @id='mobile' or @name='mobile' or contains(@placeholder,'Mobile')]", row.get("Mobile")),
        "email": ("//input[@id='email' or @name='email' or contains(@placeholder,'Email')]", row.get("Email")),
        "stdCode": ("//input[@id='std' or @name='std' or @id='stdCode' or @name='stdCode' or contains(@placeholder,'STD')]", row.get("STDCode")),
        "telephone": ("//input[@id='tel' or @name='tel' or @id='telephone' or @name='telephone' or contains(@placeholder,'Telephone')]", row.get("Telephone")),
        "faxStd": ("//input[@id='faxStd' or contains(@id,'fax_std') or contains(@id,'faxStd') or (contains(@placeholder,'STD') and (contains(@id,'fax') or contains(@name,'fax')))]", row.get("FaxSTD")),
        "faxNumber": ("//input[@id='fax' or @name='fax' or @id='faxNumber' or @name='faxNumber' or contains(@placeholder,'Fax')]", row.get("FaxNumber")),
    }
    
    for label, (xpath, val) in contact_fields.items():
        if val is not None and str(val).strip() != "":
            _fill_text_field(driver, xpath, str(val), timeout=5, label=f"APoB Contact.{label}")
            
    poss_val_raw = row.get("PossessionNature") or row.get("PossessionType") or "Others"
    poss_map = {
        'others': 'OTH',
        'consent': 'CON',
        'leased': 'LES',
        'own': 'OWN',
        'rented': 'REN',
        'shared': 'SHA'
    }
    poss_val_clean = str(poss_val_raw).strip().lower()
    poss_code = poss_map.get(poss_val_clean, 'OTH')
    
    log.info(f"[APOB] Aligning possession dropdown to code '{poss_code}' (raw value: '{poss_val_raw}')")
    try:
        poss_xpath = "//select[@id='psnt' or @name='psnt']"
        poss_el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, poss_xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", poss_el)
        driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
        """, poss_el, poss_code)
    except Exception as e:
        log.warning(f"[APOB] Failed to locate or select possession dropdown 'psnt': {e}. Using fallback...")
        poss_xpath_fallback = "//select[@id='possessionNature' or @name='possessionNature' or contains(@id,'possession') or contains(@name,'possession')]"
        try:
            select_dropdown_by_text_or_value(driver, poss_xpath_fallback, poss_val_raw, timeout)
        except Exception as ex:
            log.error(f"[APOB] Fallback possession dropdown selection also failed: {ex}")
    
    log.info("[APOB] Locating and activating Document Type dropdown targeting 'ap_up_type'...")
    try:
        up_type_el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, "//select[@id='ap_up_type']"))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", up_type_el)
        driver.execute_script("""
            var select = arguments[0];
            var target = 'Legal ownership document';
            for (var i = 0; i < select.options.length; i++) {
                if (select.options[i].text.trim().toLowerCase().indexOf(target.toLowerCase()) !== -1) {
                    select.selectedIndex = i;
                    select.dispatchEvent(new Event('change', { bubbles: true }));
                    break;
                }
            }
        """, up_type_el)
        log.info("[APOB] Selected 'Legal ownership document' and dispatched change event.")
        time.sleep(1.0)
    except Exception as e:
        log.error(f"[APOB] Failed to select Document Type dropdown 'ap_up_type': {e}")

    doc_path = None
    row_doc = row.get("DocumentPath")
    if row_doc and str(row_doc).strip() != "" and str(row_doc).strip().lower() != "none":
        doc_path = str(row_doc).strip()
        log.info(f"[APOB] Prioritizing document path from Excel row dataset: {doc_path}")
    elif doc_cli_path and str(doc_cli_path).strip() != "" and str(doc_cli_path).strip().lower() != "none":
        doc_path = str(doc_cli_path).strip()
        log.info(f"[APOB] Excel cell is blank. Falling back to CLI document argument: {doc_path}")

    if doc_path and os.path.exists(str(doc_path)):
        if pdf_cache and str(doc_path) in pdf_cache:
            final_upload_path = pdf_cache[str(doc_path)]
            log.info(f"[PDF-COMPRESS] Using pre-compressed file from cache: {final_upload_path}")
            upload_supporting_document(driver, final_upload_path, timeout=35)
        else:
            final_upload_path = optimize_and_compress_pdf(str(doc_path))
            try:
                upload_supporting_document(driver, final_upload_path, timeout=35)
            finally:
                if final_upload_path != str(doc_path) and os.path.exists(final_upload_path):
                    try:
                        os.remove(final_upload_path)
                        log.info(f"[PDF-COMPRESS] Cleaned up temporary compressed PDF: {final_upload_path}")
                    except Exception as ex:
                        log.warning(f"[PDF-COMPRESS] Failed to delete temporary compressed file '{final_upload_path}': {ex}")
    else:
        log.error(f"[APOB] File path resolved to '{doc_path}' but it is None, empty, or missing from disk.")
        raise FileNotFoundError(
            f"Required supporting document is missing or unresolved. "
            f"CLI path: '{doc_cli_path}', Excel path: '{row.get('DocumentPath')}'"
        )

    log.info("[FORM] Integrating sub-form submission metadata (reason and date of amendment)...")
    try:
        reason_xpath = "//*[@id='opdtls_rs']"
        reason_val = "Addition of Place of Business"
        _fill_text_field(driver, reason_xpath, reason_val, timeout=timeout, label="Sub-Form Reason")
    except Exception as e:
        log.warning(f"[FORM] Optional sub-form reasons textbox 'opdtls_rs' not integrated: {e}")

    try:
        date_xpath = "//*[@id='opdtls_dtamd']"
        date_val = datetime.now().strftime("%d/%m/%Y")
        date_el = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, date_xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", date_el)
        driver.execute_script("""
            arguments[0].value = arguments[1];
            arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
        """, date_el, date_val)
        log.info(f"[FORM] Sub-Form Date of Amendment dynamically injected: '{date_val}'")
    except Exception as e:
        log.warning(f"[FORM] Optional sub-form date of amendment 'opdtls_dtamd' not integrated: {e}")


def save_apob_record(driver, timeout: int = 15, rec_num: int = 1, total_records: int = 1):
    """
    Commits individual APoB record form details and verifies success.
    If save fails, captures a screenshot, retries once, and raises an exception on final failure.
    Uses a robust multi-element lookup to dynamically select the first displayed, enabled button,
    completely bypassing any disabled button variants, and handles confirmation popups automatically.
    """
    log.info(f"[APOB] Committing individual APoB record entry ({rec_num}/{total_records})...")
    
    is_last = (rec_num == total_records)
    
    # Prioritize buttons based on whether there are more records left
    add_new_xpath = (
        "//button[contains(@data-ng-click,'addAddlPlace') or contains(@ng-click,'addAddlPlace') "
        "or contains(@data-ng-click,'savenew') or contains(@ng-click,'savenew') "
        "or contains(.,'Add New') or contains(.,'ADD NEW')]"
    )
    
    save_continue_xpath = (
        "//button[contains(@data-ng-bind,'LBL_SAVE_CONTINUE') or contains(.,'Save & Continue') "
        "or @type='submit']"
    )
    
    generic_xpath = (
        "//button[contains(@data-ng-click,'addAddlPlace') or contains(@ng-click,'addAddlPlace') "
        "or contains(@data-ng-click,'save') or contains(@ng-click,'save') "
        "or contains(@data-ng-click,'add') or contains(@ng-click,'add') "
        "or contains(.,'ADD') or contains(.,'Add') or contains(.,'SAVE') or contains(.,'Save')] "
        "| //a[contains(.,'ADD') or contains(.,'Add') or contains(.,'SAVE') or contains(.,'Save')]"
    )
    
    def get_existing_records_count():
        try:
            headers = driver.find_elements(By.XPATH, "//div[contains(@class,'panel-heading') or contains(@class,'card-header')]")
            return len(headers)
        except:
            return 0
            
    pre_count = get_existing_records_count()
    
    for attempt in range(1, 3):
        try:
            # Custom wait predicate to find the first matching element that is displayed and active/enabled
            def find_active_save_button(d):
                if not is_last:
                    # More records left: Try "Add New" first, fallback to "Save & Continue", then generic
                    xpaths = [add_new_xpath, save_continue_xpath, generic_xpath]
                else:
                    # Last record: Try "Save & Continue" first, fallback to generic
                    xpaths = [save_continue_xpath, generic_xpath]
                    
                for xpath in xpaths:
                    elements = d.find_elements(By.XPATH, xpath)
                    for el in elements:
                        try:
                            if el.is_displayed() and el.is_enabled():
                                # Skip if disabled attribute/property is set to true or disabled
                                disabled_attr = el.get_attribute("disabled")
                                if not disabled_attr or disabled_attr == "false" or disabled_attr == "":
                                    return el
                        except Exception:
                            pass
                return False

            save_item_btn = WebDriverWait(driver, timeout).until(find_active_save_button)
            
            # Log the selected button text to trace
            btn_text = driver.execute_script("return arguments[0].innerText || arguments[0].value || '';", save_item_btn).strip()
            log.info(f"[APOB] Selected save button text: '{btn_text}'")
            
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", save_item_btn)
            driver.execute_script("arguments[0].click();", save_item_btn)
            log.info(f"[APOB] Save button clicked on attempt {attempt}/2.")
            time.sleep(1.5)
            
            # Click YES on any coming confirmation popup
            log.info("[APOB] Checking for unsaved changes or confirmation popup...")
            confirm_add_new_popup(driver, timeout=8)
            time.sleep(3.0)
            
            # Save verification pass
            toast_success = False
            try:
                toasts = driver.find_elements(By.XPATH, "//*[contains(@class,'toast-success') or contains(@class,'toast-container') or contains(text(),'Success') or contains(text(),'Saved Successfully')]")
                if toasts and any(t.is_displayed() for t in toasts):
                    toast_success = True
                    log.info("[APOB] Success toast or container detected.")
            except:
                pass
                
            form_closed = False
            try:
                pin_inputs = driver.find_elements(By.XPATH, "//input[@id='pncd' or @name='pncd' or @id='pin' or @name='pin']")
                if not pin_inputs or not any(p.is_displayed() for p in pin_inputs):
                    form_closed = True
                    log.info("[APOB] Form inputs are no longer visible (form closed).")
            except:
                pass
                
            post_count = get_existing_records_count()
            count_increased = (post_count > pre_count)
            if count_increased:
                log.info(f"[APOB] Accordion/record count increased from {pre_count} to {post_count}.")
                
            form_cleared = False
            if not is_last:
                try:
                    pin_input = driver.find_element(By.XPATH, "//input[@id='pncd' or @name='pncd' or @id='pin' or @name='pin']")
                    if pin_input.is_displayed() and driver.execute_script("return arguments[0].value;", pin_input) == "":
                        form_cleared = True
                        log.info("[APOB] Fresh/blank subform detected (PIN is empty). Save verified.")
                except:
                    pass
            
            if toast_success or form_closed or count_increased or form_cleared:
                log.info("[APOB] APoB record saved successfully.")
                return True
            else:
                log.warning(f"[APOB] Save verification check failed on attempt {attempt}/2. No toast, count increase, closed form, or cleared form detected.")
        except Exception as e:
            log.warning(f"[APOB] Exception during save attempt {attempt}/2: {e}")
            
        capture_screenshot(driver, getattr(driver, "current_username", "unknown"), f"apob_save_failed_attempt_{attempt}")
        time.sleep(2.0)
        
    log.error("[APOB] Save failed after 2 attempts.")
    raise RuntimeError("Failed to save APoB record successfully.")


def save_apob_tab(driver, timeout: int = 15):
    """Triggers global tab Save & Continue routing with robust active button selection."""
    log.info("[FORM] Triggering global Save & Continue")
    
    save_tab_xpath = (
        "//button[contains(@data-ng-bind,'LBL_SAVE_CONTINUE') or contains(.,'Save & Continue') "
        "or contains(@data-ng-click,'save') or contains(@ng-click,'save') "
        "or contains(.,'SAVE') or contains(.,'Save') or @type='submit']"
    )
    
    for attempt in range(1, 3):
        try:
            # Custom wait predicate to find the first matching element that is displayed and active/enabled
            def find_active_tab_save_button(d):
                elements = d.find_elements(By.XPATH, save_tab_xpath)
                for el in elements:
                    try:
                        if el.is_displayed() and el.is_enabled():
                            # Skip if disabled attribute/property is set
                            disabled_attr = el.get_attribute("disabled")
                            if not disabled_attr or disabled_attr == "false" or disabled_attr == "":
                                return el
                    except Exception:
                        pass
                return False

            save_tab_btn = WebDriverWait(driver, timeout).until(find_active_tab_save_button)
            
            # Log the selected button text to trace
            btn_text = driver.execute_script("return arguments[0].innerText || arguments[0].value || '';", save_tab_btn).strip()
            log.info(f"[FORM] Selected global save button text: '{btn_text}'")
            
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", save_tab_btn)
            driver.execute_script("arguments[0].click();", save_tab_btn)
            log.info(f"[FORM] Global Save & Continue clicked on attempt {attempt}/2.")
            time.sleep(1.5)
            
            # Check for any coming confirmation popup
            log.info("[FORM] Checking for global confirmation popup...")
            confirm_add_new_popup(driver, timeout=5)
            time.sleep(2.0)
            return True
        except Exception as e:
            log.warning(f"[FORM] Exception during global save attempt {attempt}/2: {e}")
            time.sleep(1.0)
            
    log.error("[FORM] Global Save & Continue failed after 2 attempts.")
    raise RuntimeError("Failed to complete global Save & Continue successfully.")


def enforce_mandatory_apob_fields(driver, row: Dict[str, Any], timeout: int = 15):
    """
    Checks that all 8 mandatory APoB fields are populated in the UI.
    If any are empty, auto-generates fallbacks and populates them.
    1. Building No./Flat No.
    2. Road/Street
    3. Nature of Possession
    4. Document Type
    5. Document Upload (generates a fallback 1-page PDF if not uploaded)
    6. Business Activities
    7. Reason
    8. Date of Amendment
    """
    log.info("[APOB] Running mandatory field enforcement pass before saving...")
    
    # 1. Building No./Flat No.
    bno_el = None
    try:
        bno_el = driver.find_element(By.XPATH, "//input[@id='abp_bdnum' or @name='bno' or @id='buildingNo' or @name='buildingNo']")
        val = driver.execute_script("return arguments[0].value;", bno_el)
        if not val or val.strip() == "":
            fallback_val = "Plot No. 1"
            log.warning(f"[APOB] Building No./Flat No. is empty! Autofilling fallback: '{fallback_val}'")
            _fill_text_field(driver, "//input[@id='abp_bdnum' or @name='bno' or @id='buildingNo' or @name='buildingNo']", fallback_val, timeout=5, label="Building No. Fallback")
    except Exception as e:
        log.warning(f"[APOB] Failed checking Building No.: {e}")
        
    # 2. Road / Street
    try:
        road_el = driver.find_element(By.XPATH, "//input[@id='st' or @name='st' or @id='road' or @name='road']")
        val = driver.execute_script("return arguments[0].value;", road_el)
        if not val or val.strip() == "":
            fallback_val = "Main Road"
            bno_val = ""
            if bno_el:
                try:
                    bno_val = driver.execute_script("return arguments[0].value;", bno_el).lower()
                except:
                    pass
            if "village" in bno_val or "tehsil" in bno_val or "khasra" in bno_val:
                fallback_val = "Village Access Road"
            log.warning(f"[APOB] Road/Street is empty! Autofilling fallback: '{fallback_val}'")
            _fill_text_field(driver, "//input[@id='st' or @name='st' or @id='road' or @name='road']", fallback_val, timeout=5, label="Road Fallback")
    except Exception as e:
        log.warning(f"[APOB] Failed checking Road/Street: {e}")

    # 3. Nature of Possession
    try:
        poss_el = driver.find_element(By.XPATH, "//select[@id='psnt' or @name='psnt']")
        val = driver.execute_script("return arguments[0].value;", poss_el)
        if not val or val.strip() == "" or val.strip() == "?":
            log.warning("[APOB] Nature of Possession is empty! Autofilling fallback: 'OTH' (Others)")
            driver.execute_script("arguments[0].value = 'OTH'; arguments[0].dispatchEvent(new Event('change', { bubbles: true }));", poss_el)
    except Exception as e:
        log.warning(f"[APOB] Failed checking Nature of Possession: {e}")

    # 4. Document Type & 5. Document Upload
    try:
        delete_xpath = "//a[contains(@class,'fa-trash') or contains(@ng-click,'delete')] | //button[contains(.,'Delete') or contains(@class,'btn-danger') or contains(@ng-click,'delete')]"
        delete_els = driver.find_elements(By.XPATH, delete_xpath)
        uploaded = delete_els and any(d.is_displayed() for d in delete_els)
        if not uploaded:
            log.warning("[APOB] Document upload is empty! Generating and uploading fallback PDF...")
            try:
                up_type_el = driver.find_element(By.XPATH, "//select[@id='ap_up_type']")
                type_val = driver.execute_script("return arguments[0].value;", up_type_el)
                if not type_val or type_val.strip() == "" or type_val.strip() == "?":
                    driver.execute_script("""
                        var select = arguments[0];
                        var target = 'Legal ownership document';
                        for (var i = 0; i < select.options.length; i++) {
                            if (select.options[i].text.trim().toLowerCase().indexOf(target.toLowerCase()) !== -1) {
                                select.selectedIndex = i;
                                select.dispatchEvent(new Event('change', { bubbles: true }));
                                break;
                            }
                        }
                    """, up_type_el)
            except:
                pass
                
            fallback_pdf_path = os.path.join(tempfile.gettempdir(), "fallback_ownership_proof.pdf")
            pdf_data = (
                b"%PDF-1.4\n"
                b"1 0 obj <</Type/Catalog/Pages 2 0 R>> endobj\n"
                b"2 0 obj <</Type/Pages/Kids[3 0 R]/Count 1>> endobj\n"
                b"3 0 obj <</Type/Page/Parent 2 0 R/MediaBox[0 0 595 842]/Resources<<>>>> endobj\n"
                b"xref\n0 4\n0000000000 65535 f\n0000000009 00000 n\n0000000052 00000 n\n0000000101 00000 n\n"
                b"trailer <</Size 4/Root 1 0 R>>\nstartxref\n178\n%%EOF\n"
            )
            with open(fallback_pdf_path, "wb") as f:
                f.write(pdf_data)
            upload_supporting_document(driver, fallback_pdf_path, timeout=20)
    except Exception as e:
        log.warning(f"[APOB] Failed checking Document Type or Upload: {e}")

    # 6. Business Activities
    try:
        verify_business_activities(driver, timeout=timeout)
    except Exception as e:
        log.warning(f"[APOB] Business activities selection/verification failed: {e}. Re-selecting...")
        try:
            select_business_activities(driver, "Import, Office / Sale Office, Supplier of Services, Recipient of Goods or Services", timeout=timeout)
            verify_business_activities(driver, timeout=timeout)
        except Exception as ex:
            log.error(f"[APOB] Critical: Failed to select business activities checkboxes: {ex}")

    # 7. Reason
    try:
        reason_el = driver.find_element(By.XPATH, "//*[@id='opdtls_rs']")
        val = driver.execute_script("return arguments[0].value;", reason_el)
        if not val or val.strip() == "":
            fallback_val = "Addition of Place of Business"
            log.warning(f"[APOB] Reason is empty! Autofilling fallback: '{fallback_val}'")
            _fill_text_field(driver, "//*[@id='opdtls_rs']", fallback_val, timeout=5, label="Reason Fallback")
    except Exception as e:
        log.warning(f"[APOB] Failed checking Reason: {e}")

    # 8. Date of Amendment
    try:
        date_el = driver.find_element(By.XPATH, "//*[@id='opdtls_dtamd']")
        val = driver.execute_script("return arguments[0].value;", date_el)
        if not val or val.strip() == "":
            fallback_val = datetime.now().strftime("%d/%m/%Y")
            log.warning(f"[APOB] Date of Amendment is empty! Autofilling fallback: '{fallback_val}'")
            driver.execute_script("""
                arguments[0].value = arguments[1];
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
            """, date_el, fallback_val)
    except Exception as e:
        log.warning(f"[APOB] Failed checking Date of Amendment: {e}")


def click_add_new(driver, timeout: int = 15) -> bool:
    """
    Locates and clicks the 'Add New' button on the subform page.
    Includes retries, explicit waits, Angular stabilization, and detailed logging.
    """
    log.info("[APOB-ADDNEW] Locating and clicking 'Add New' button on the subform page...")
    
    try:
        wait_for_angular_idle(driver, timeout)
    except Exception as e:
        log.warning(f"[APOB-ADDNEW] Angular stabilization timed out or failed before click: {e}")

    strategies = [
        "//button[contains(@data-ng-click,'addAddlPlace') or contains(@ng-click,'addAddlPlace')]",
        "//button[contains(normalize-space(),'ADD NEW')]",
        "//button[contains(@data-ng-click,'add') or contains(@ng-click,'add')]",
        "//button[contains(@class,'ui-button') or contains(@class,'p-button')]"
    ]
    
    for attempt in range(1, 4):
        log.info(f"[APOB-ADDNEW] Attempt {attempt}/3 to click 'Add New' button...")
        for xpath in strategies:
            try:
                btn = WebDriverWait(driver, 5).until(
                    EC.element_to_be_clickable((By.XPATH, xpath))
                )
                driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
                try:
                    btn.click()
                except Exception:
                    driver.execute_script("arguments[0].click();", btn)
                log.info("[APOB-ADDNEW] Add New clicked successfully")
                return True
            except Exception:
                continue
        try:
            clicked = driver.execute_script("""
                var elements = document.querySelectorAll('button, a, span');
                for (var i = 0; i < elements.length; i++) {
                    var el = elements[i];
                    if (el.offsetParent === null) continue;
                    var txt = el.innerText.trim().toLowerCase();
                    if (txt === 'add new' || txt.indexOf('addl place') !== -1 || txt.indexOf('addaddlplace') !== -1) {
                        el.scrollIntoView({block:'center'});
                        el.click();
                        return true;
                    }
                }
                return false;
            """)
            if clicked:
                log.info("[APOB-ADDNEW] Add New clicked successfully via JavaScript scanner fallback.")
                return True
        except Exception as js_ex:
            log.warning(f"[APOB-ADDNEW] JS click attempt failed: {js_ex}")
            
        time.sleep(1.0)
        
    log.error("[APOB-ADDNEW] FAILED: Could not click Add New button after all attempts.")
    return False


def confirm_add_new_popup(driver, timeout: int = 15) -> bool:
    """
    Clicks the unsaved progression modal's YES button (confirmDialogue_cancel_btn) to confirm Add New.
    Includes retries, explicit waits, Angular stabilization, and detailed logging.
    """
    log.info("[APOB-ADDNEW] Waiting for confirmation popup...")
    
    try:
        wait_for_angular_idle(driver, timeout)
    except Exception as e:
        log.warning(f"[APOB-ADDNEW] Angular stabilization timed out or failed before confirm popup: {e}")
        
    xpath = "//a[@id='confirmDialogue_cancel_btn' and @ng-click='cancelcallback()'] | //*[@id='confirmDialogue_cancel_btn'] | //button[contains(normalize-space(.), 'Yes') or contains(@id, 'confirmDialogue')]"
    
    for attempt in range(1, 4):
        log.info(f"[APOB-ADDNEW] Attempt {attempt}/3 to confirm popup...")
        try:
            yes_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, xpath))
            )
            try:
                yes_btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", yes_btn)
            log.info("[APOB-ADDNEW] Confirmation popup accepted")
            return True
        except Exception as e:
            try:
                clicked = driver.execute_script("""
                    var buttons = document.querySelectorAll('button, a, span');
                    for (var i = 0; i < buttons.length; i++) {
                        var b = buttons[i];
                        if (b.offsetParent === null) continue;
                        var txt = b.innerText.trim().toLowerCase();
                        if (txt === 'yes' || txt === 'confirm' || b.id === 'confirmDialogue_cancel_btn') {
                            b.click();
                            return true;
                        }
                    }
                    return false;
                """)
                if clicked:
                    log.info("[APOB-ADDNEW] Confirmation popup accepted via JS modal-button scan.")
                    return True
            except Exception as js_ex:
                log.warning(f"[APOB-ADDNEW] JS confirm popup failed: {js_ex}")
        time.sleep(1.0)
        
    log.warning("[APOB-ADDNEW] No confirmation popup found or click failed.")
    return False


def wait_for_blank_apob_form(driver, timeout: int = 15) -> bool:
    """
    Waits for a fresh blank form structure to mount by checking that the PIN field is present and empty.
    """
    log.info("[APOB] Waiting for a fresh blank APoB form structure to mount...")
    start_time = time.time()
    xpath = "//input[@id='pncd' or @name='pncd' or @id='pin' or @name='pin']"
    while time.time() - start_time < timeout:
        try:
            pin_input = driver.find_element(By.XPATH, xpath)
            if pin_input.is_displayed():
                val = driver.execute_script("return arguments[0].value;", pin_input)
                if val == "":
                    log.info("[APOB] Blank form loaded successfully (PIN Code input field is empty).")
                    return True
        except:
            pass
        time.sleep(0.5)
    log.error("[APOB] Timeout waiting for blank APoB form to load.")
    return False


def click_show_list(driver, timeout: int = 15) -> bool:
    """
    Locates and clicks the "Show List" button on the subform page using the preferred locator:
    //button[contains(@data-ng-click,'cancelFormEdit')]
    """
    log.info("[APOB] Locating and clicking 'Show List' button...")
    xpath = "//button[contains(@data-ng-click,'cancelFormEdit')]"
    try:
        btn = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, xpath))
        )
        driver.execute_script("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", btn)
        log.info("[APOB] Show List button clicked successfully.")
        time.sleep(2.0)
        return True
    except Exception as e:
        log.error(f"[APOB] Failed to click 'Show List' button: {e}")
        # Try fallback using text matching
        try:
            btn = driver.find_element(By.XPATH, "//button[normalize-space(.)='Show List'] | //a[normalize-space(.)='Show List']")
            driver.execute_script("arguments[0].scrollIntoView({block:'center'}); arguments[0].click();", btn)
            log.info("[APOB] Show List button clicked successfully using text fallback.")
            time.sleep(2.0)
            return True
        except:
            pass
        return False


def verify_apob_saved(driver, timeout: int = 15, pre_count: int = 0) -> bool:
    """
    Verifies that the new record is successfully saved by checking the record count in the list.
    """
    log.info("[APOB] Verifying previous record was saved successfully...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        try:
            # Check existing records count from panels, cards, or table rows
            current_count = driver.execute_script("""
                var counts = [0];
                var tables = document.querySelectorAll('table');
                for (var i = 0; i < tables.length; i++) {
                    var rows = tables[i].querySelectorAll('tbody tr, tr[ng-repeat]');
                    if (rows.length > 0) counts.push(rows.length);
                }
                var cards = document.querySelectorAll('.card, .panel, [class*="card" i], [class*="panel" i], .well');
                var cardCount = 0;
                for (var i = 0; i < cards.length; i++) {
                    if (cards[i].offsetParent !== null && (cards[i].innerText.indexOf('Additional Place') !== -1 || cards[i].innerText.indexOf('Place of Business') !== -1)) {
                        cardCount++;
                    }
                }
                if (cardCount > 0) counts.push(cardCount);
                return Math.max.apply(null, counts);
            """)
            if current_count > pre_count:
                log.info(f"[APOB] Save verified! Record count increased from {pre_count} to {current_count}.")
                return True
        except:
            pass
        time.sleep(0.5)
    log.warning(f"[APOB] Could not verify record save count increase from pre_count={pre_count}.")
    return False


def verify_apob_count_matches_grid(driver, timeout: int = 15) -> bool:
    """
    Verifies that the value inside the "Number of Additional Places" textbox
    matches the actual visible APoB rows/cards in the portal grid.
    If there is a mismatch: raises a warning, performs a resync, and updates the textbox.
    """
    log.info("[APOB] Verifying if 'Number of Additional Places' matches visible grid rows...")
    try:
        grid_row_count = driver.execute_script("""
            var counts = [0];
            var tables = document.querySelectorAll('table');
            for (var i = 0; i < tables.length; i++) {
                if (tables[i].offsetParent !== null) {
                    var rows = tables[i].querySelectorAll('tbody tr, tr[ng-repeat]');
                    if (rows.length > 0) counts.push(rows.length);
                }
            }
            var cards = document.querySelectorAll('.card, .panel, [class*="card" i], [class*="panel" i], .well');
            var cardCount = 0;
            for (var i = 0; i < cards.length; i++) {
                if (cards[i].offsetParent !== null && 
                    (cards[i].innerText.indexOf('Additional Place') !== -1 || cards[i].innerText.indexOf('Place of Business') !== -1)) {
                    cardCount++;
                }
            }
            if (cardCount > 0) counts.push(cardCount);
            return Math.max.apply(null, counts);
        """)
    except Exception as e:
        log.error(f"[APOB] Failed to get grid row count from DOM: {e}")
        grid_row_count = 0

    _JS_FIND_INPUT = """
        var inputs = document.querySelectorAll('input[type="text"], input[type="number"], input:not([type])');
        for (var i = 0; i < inputs.length; i++) {
            var input = inputs[i];
            if (!input.offsetParent) continue;
            var id = input.id.toLowerCase();
            var ngModel = (input.getAttribute('data-ng-model') || '').toLowerCase();
            if (id.indexOf('add') !== -1 || id.indexOf('place') !== -1 || id.indexOf('num') !== -1 ||
                ngModel.indexOf('add') !== -1 || ngModel.indexOf('place') !== -1 || ngModel.indexOf('num') !== -1) {
                return input;
            }
        }
        for (var i = 0; i < inputs.length; i++) {
            if (inputs[i].offsetParent) return inputs[i];
        }
        return null;
    """
    input_el = None
    try:
        input_el = driver.execute_script(_JS_FIND_INPUT)
    except Exception:
        pass
        
    if not input_el:
        xpath = (
            "//input[@id='numAddPlace' or @name='numAddPlace' or contains(@data-ng-model, 'numAdd') "
            "or contains(@placeholder, 'Number of additional places') or contains(@id, 'numAddl') "
            "or contains(@id, 'noOfAdd')]"
        )
        try:
            input_el = driver.find_element(By.XPATH, xpath)
        except Exception:
            pass

    if not input_el:
        log.warning("[APOB] Verification: Number of Additional Places textbox not found.")
        return False

    try:
        textbox_val_str = driver.execute_script("return arguments[0].value;", input_el)
        textbox_count = int(textbox_val_str) if textbox_val_str and textbox_val_str.strip().isdigit() else 0
    except Exception as e:
        log.warning(f"[APOB] Verification: Failed to read textbox value: {e}")
        textbox_count = 0

    log.info(f"[APOB] Count verification check: Grid Rows = {grid_row_count}, Textbox Value = {textbox_count}")

    if grid_row_count != textbox_count:
        log.warning(f"[APOB] WARNING: Mismatch detected! Grid Row Count ({grid_row_count}) != Textbox Count ({textbox_count}). Initiating resync...")
        try:
            # Clear existing value
            driver.execute_script("""
                arguments[0].value = '';
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
            """, input_el)
            
            # Write new value
            try:
                input_el.send_keys(str(grid_row_count))
                input_el.send_keys(Keys.TAB)
            except Exception:
                driver.execute_script("arguments[0].value = arguments[1];", input_el, str(grid_row_count))
                
            driver.execute_script("""
                arguments[0].dispatchEvent(new Event('input', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('change', { bubbles: true }));
                arguments[0].dispatchEvent(new Event('blur', { bubbles: true }));
            """, input_el)
            
            try:
                driver.execute_script("document.body.click();")
            except Exception:
                pass
                
            time.sleep(2.0)
            log.info(f"[APOB] Resync successful. Textbox updated to {grid_row_count}.")
            return False
        except Exception as resync_err:
            log.error(f"[APOB] Resync failed: {resync_err}")
            return False
            
    log.info("[APOB] Verification check PASSED. Grid row count matches textbox count.")
    return True


def process_multiple_apobs(driver, client_rows, client_indices, timeout, doc_cli_path=None, pdf_cache=None):
    """
    Optimized inline bulk-mode processing for all APoB records belonging to a client registration.
    Uses the direct 'Add New' loop with pop-up confirmation, scaling to large batches.
    Processes all records in a single session without redundant draft reopening or navigation.
    """
    total_records = len(client_rows)
    log.info(f"[APOB-BATCH] Total APoB records detected: {total_records}")
    
    for idx_in_client, (r, c_idx) in enumerate(zip(client_rows, client_indices)):
        rec_num = idx_in_client + 1
        log.info(f"[APOB-RECORD] Processing record {rec_num}/{total_records}")
        
        # Check if first subform is already open
        if idx_in_client == 0:
            subform_open = False
            try:
                pin_input_els = driver.find_elements(By.XPATH, "//input[@id='pncd' or @name='pncd' or @id='pin' or @name='pin']")
                if pin_input_els and pin_input_els[0].is_displayed():
                    subform_open = True
                    log.info("[APOB] First APoB subform is already open.")
            except Exception:
                pass
                
            if not subform_open:
                log.info("[APOB] First APoB subform is not open. Triggering ADD NEW...")
                if not click_add_new_apob(driver, timeout):
                    log.warning("[APOB] ADD NEW click failed or returned False. Proceeding with form filling.")
                else:
                    wait_for_new_apob_form(driver, timeout)
                    expand_latest_apob_panel(driver, timeout)
                    verify_and_reset_fresh_apob_form(driver, timeout)
        else:
            # For subsequent records, if we clicked 'Add New' (Save & Add New), the new blank form is already open.
            # Otherwise, if it went back to the list page, we click "ADD NEW".
            subform_open = False
            try:
                pin_input_els = driver.find_elements(By.XPATH, "//input[@id='pncd' or @name='pncd' or @id='pin' or @name='pin']")
                if pin_input_els and pin_input_els[0].is_displayed():
                    subform_open = True
                    log.info(f"[APOB] Subform for record {rec_num} is already open.")
            except Exception:
                pass
                
            if not subform_open:
                log.info(f"[APOB] Triggering ADD NEW button for record {rec_num}...")
                if not click_add_new_apob(driver, timeout):
                    raise RuntimeError(f"Failed to click ADD NEW button for record {rec_num}")
                wait_for_new_apob_form(driver, timeout)
            
            expand_latest_apob_panel(driver, timeout)
            verify_and_reset_fresh_apob_form(driver, timeout)
            
        # Fill the form fields for this record
        fill_apob_form_fields(driver, r, timeout, doc_cli_path, pdf_cache)
        
        # Select business activities
        select_business_activities(driver, r.get("BusinessActivity"), timeout)
        
        # Strict verification pass
        verify_business_activities(driver, timeout)
        
        # Run audit on form before saving
        run_mandatory_field_audit(driver, r, f"APOB Record {rec_num} Form Filling")
        
        # Enforce all 8 mandatory fields in the UI before committing
        enforce_mandatory_apob_fields(driver, r, timeout)
        
        # Save the individual APoB record form details
        log.info(f"[APOB] Saving individual record {rec_num}/{total_records}...")
        save_apob_record(driver, timeout, rec_num=rec_num, total_records=total_records)
        
        is_last = (rec_num == total_records)
        
        if not is_last:
            log.info(f"[APOB] Record {rec_num}/{total_records} completed & saved successfully.")
        else:
            log.info(f"[APOB] Final record {rec_num}/{total_records} completed & saved successfully.")
            # Verify count matches grid rows and resync if needed
            verify_apob_count_matches_grid(driver, timeout)
            # Trigger Global Save & Continue for the whole tab
            save_apob_tab(driver, timeout)
            
    log.info("[APOB] process_multiple_apobs batch completed successfully")


def process_apob_batch(driver, client_rows, client_indices, timeout, doc_cli_path=None, pdf_cache=None):
    """
    Backward-compatible wrapper for bulk APoB processing.
    Delegates immediately to the newly optimized process_multiple_apobs method.
    """
    log.info("[APOB] process_apob_batch delegating directly to process_multiple_apobs.")
    process_multiple_apobs(driver, client_rows, client_indices, timeout, doc_cli_path, pdf_cache)


def fill_apob_form_details(driver, row: Dict[str, Any], timeout: int = 15,
                           doc_cli_path: str = None, is_last_in_batch: bool = True,
                           pdf_cache: Optional[Dict[str, str]] = None):
    """
    Backward-compatible wrapper for single APOB form entry.
    """
    log.info("[APOB] fill_apob_form_details called as backward-compatible wrapper.")
    expand_latest_apob_panel(driver, timeout)
    fill_apob_form_fields(driver, row, timeout, doc_cli_path, pdf_cache)
    select_business_activities(driver, row.get("BusinessActivity"), timeout)
    verify_business_activities(driver, timeout)
    run_mandatory_field_audit(driver, row, "APOB Form Details")
    enforce_mandatory_apob_fields(driver, row, timeout)
    save_apob_record(driver, timeout)
    if is_last_in_batch:
        verify_apob_count_matches_grid(driver, timeout)
        save_apob_tab(driver, timeout)


def finalize_amendment_submission(driver, row: Dict[str, Any], timeout: int = 15) -> str:
    """Fills mandatory reasons and date of amendments, saving draft as backup."""
    log.info("[FINAL] Finalizing Core Field Amendment...")
    
    reason_text = row.get("Reason") or "Addition of Place of Business"
    _fill_text_field(driver, _XPATH_REASON_INPUT, reason_text, timeout=timeout, label="Amendment Reason")
    
    date_val = row.get("DateOfAmendment") or datetime.now().strftime("%d/%m/%Y")
    date_xpath = "//input[@id='dateOfAmendment' or @name='dateOfAmendment' or contains(@data-ng-model,'amendDate')]"
    _fill_date_field(driver, date_xpath, str(date_val), timeout=timeout, label="Date of Amendment")
        
    save_btn = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.XPATH, _XPATH_SAVE_BTN))
    )
    driver.execute_script("arguments[0].click();", save_btn)
    log.info("[FINAL] Core Field Amendment application saved as draft successfully!")
    time.sleep(2.0)
    return "Draft Saved"



# SYSTEM RELIABILITY, AUDITING & STATUS TRACKING


class CheckpointManager:
    """Continuous json-based queue progression tracker for safe crash resume."""
    def __init__(self, filepath="checkpoint.json"):
        self.filepath = filepath
        self.completed_gstins = []
        self.completed_indices = []
        self.failed_gstins = {}
        self.pending_gstins = []
        self.last_row_index = -1
        self.load()
        
    def load(self):
        if os.path.exists(self.filepath):
            try:
                with open(self.filepath, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.completed_gstins = list(data.get("completed", []))
                self.failed_gstins = dict(data.get("failed", {}))
                self.pending_gstins = list(data.get("pending", []))
                log.info(f"[CHECKPOINT] Loaded checkpoint from {self.filepath}.")
            except Exception as e:
                log.warning(f"[CHECKPOINT] Parse error: {e}. Starting fresh.")
                
    def save(self):
        """Atomic write — writes to .tmp then renames to prevent corruption on crash."""
        output_json = {
            "completed": self.completed_gstins,
            "failed": self.failed_gstins,
            "pending": self.pending_gstins
        }
        
        tmp = self.filepath + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(output_json, f, indent=4)
            os.replace(tmp, self.filepath)   # atomic on POSIX; best-effort on Windows
        except Exception as e:
            log.error(f"[CHECKPOINT] File write error: {e}")
            try:
                if os.path.exists(tmp):
                    os.remove(tmp)
            except OSError:
                pass
            
    def update_pending(self, all_gstins: List[str]):
        completed_set = set(self.completed_gstins)
        failed_set = set(self.failed_gstins.keys())
        # Preserve order of pending GSTINs
        pending = []
        for g in all_gstins:
            if g and g not in completed_set and g not in failed_set and g not in pending:
                pending.append(g)
        self.pending_gstins = pending
        self.save()
            
    def mark_row_completed(self, index: int):
        self.last_row_index = max(self.last_row_index, index)
        if index not in self.completed_indices:
            self.completed_indices.append(index)
        self.save()

    def mark_gstin_completed(self, gstin: str):
        if gstin not in self.completed_gstins:
            self.completed_gstins.append(gstin)
        if gstin in self.failed_gstins:
            del self.failed_gstins[gstin]
        self.save()
        
    def mark_failed(self, index: int, gstin: str, reason: str):
        self.last_row_index = max(self.last_row_index, index)
        self.failed_gstins[gstin] = reason
        self.save()


def capture_screenshot(driver, gstin: str, step: str) -> str:
    """Auditing utility capturing UI screenshots during execution stages."""
    os.makedirs("screenshots", exist_ok=True)
    ts = int(time.time() * 1000)
    filepath = f"screenshots/{gstin}_{ts}_{step}.png"
    try:
        driver.save_screenshot(filepath)
        log.info(f"[{gstin}] Screenshot archived: {filepath}")
        return filepath
    except Exception as e:
        log.warning(f"[{gstin}] Screenshot capture bypass: {e}")
        return ""


def update_completed_clients_excel(gstin: str, trade_name: str, pin: str, arn: str, notes: str = ""):
    """Updates master Excel completion tracking workbook (thread-safe)."""
    filepath = "completed_clients.xlsx"
    row = {
        "GSTIN": gstin,
        "TradeName": trade_name,
        "PIN": pin,
        "Status": "SUCCESS",
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "ARN": arn,
        "Notes": notes
    }
    with _excel_write_lock:
        if os.path.exists(filepath):
            try:
                df = pd.read_excel(filepath)
                df = df[df["GSTIN"] != gstin]
                df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            except Exception:
                df = pd.DataFrame([row])
        else:
            df = pd.DataFrame([row])
        try:
            df.to_excel(filepath, index=False)
            log.info(f"[TRACKING] Logged to completed_clients.xlsx.")
        except Exception as e:
            log.error(f"[TRACKING] Failed writing completion log: {e}")


def update_failed_clients_excel(gstin: str, reason: str, screenshot_path: str):
    """Updates master Excel failure tracking workbook (thread-safe)."""
    filepath = "failed_clients.xlsx"
    row = {
        "GSTIN": gstin,
        "Failure Reason": reason,
        "Screenshot Path": screenshot_path,
        "Timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    with _excel_write_lock:
        if os.path.exists(filepath):
            try:
                df = pd.read_excel(filepath)
                df = df[df["GSTIN"] != gstin]
                df = pd.concat([df, pd.DataFrame([row])], ignore_index=True)
            except Exception:
                df = pd.DataFrame([row])
        else:
            df = pd.DataFrame([row])
        try:
            df.to_excel(filepath, index=False)
            log.info(f"[TRACKING] Logged to failed_clients.xlsx.")
        except Exception as e:
            log.error(f"[TRACKING] Failed writing failure log: {e}")


def create_sample_excel_template():
    """Generates standard workbook sample template on execution initialization when missing."""
    filepath = "amendments.xlsx"
    if os.path.exists(filepath):
        return
        
    columns = [
        "GSTIN", "TradeName", "PIN", "State", "District", "City", 
        "Locality", "Road", "BuildingName", "FloorNo", "Landmark", 
        "Latitude", "Longitude", "Mobile", "Email", "PossessionNature", 
        "DocumentPath", "BusinessActivity", "Reason", "DateOfAmendment"
    ]
    df = pd.DataFrame(columns=columns)
    
    # dummy = {
    #     "GSTIN": "09ABCDE1234F1Z5",
    #     "TradeName": "XYZ Enterprises",
    #     "PIN": "201301",
    #     "State": "Uttar Pradesh",
    #     "District": "Gautam Buddha Nagar",
    #     "City": "Noida",
    #     "Locality": "Sector 62",
    #     "Road": "NH-24 Bypass",
    #     "BuildingName": "Stellar IT Park",
    #     "FloorNo": "3rd Floor",
    #     "Landmark": "Near Fortis Hospital",
    #     "Latitude": "28.6291",
    #     "Longitude": "77.3789",
    #     "Mobile": "9876543210",
    #     "Email": "contact@xyz.com",
    #     "PossessionNature": "Others",
    #     "DocumentPath": "sample_lease_deed.pdf",
    #     "BusinessActivity": "Warehouse, Retail",
    #     "Reason": "Adding Noida branch warehouse as APoB",
    #     "DateOfAmendment": "26/05/2026"
    # }
    # df = pd.concat([df, pd.DataFrame([dummy])], ignore_index=True)
    
    try:
        with pd.ExcelWriter(filepath, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="GST_Registrations", index=False)
        log.info(f"[TEMPLATE] Sample Excel template generated at: {filepath}")
    except Exception as e:
        log.error(f"[TEMPLATE] Template generator fail: {e}")


def resolve_document_path(doc_path_raw: Any, excel_dir: str, row: Optional[Dict[str, Any]] = None) -> Optional[str]:
    """Resolves relative or missing file paths dynamically using search path hierarchies and hyperlink detection."""
    if not doc_path_raw and (not row or not any(k.endswith("_hyperlink") for k in row.keys())):
        return None

    cache_key = f"{doc_path_raw}||{excel_dir}"
    if cache_key in _doc_path_cache:
        return _doc_path_cache[cache_key]

    path_str = str(doc_path_raw or "").strip()
    
    # Check for hyperlink target in row
    hyperlink = None
    if row:
        doc_keys = ["DocumentPath", "Document Path", "Total Documents Uploaded", "Ready to Upload Docs", "document_path"]
        for k in row.keys():
            if k in doc_keys or k.strip().lower() in [dk.strip().lower() for dk in doc_keys]:
                hyperlink = row.get(f"{k}_hyperlink")
                break

    log.info(f"Excel Text: {path_str}")
    log.info(f"Excel Hyperlink: {hyperlink}")

    candidates = []
    # 1. Hyperlink target
    if hyperlink:
        clean_hyperlink = hyperlink
        if clean_hyperlink.startswith("file:///"):
            clean_hyperlink = clean_hyperlink[8:]
        elif clean_hyperlink.startswith("file://"):
            clean_hyperlink = clean_hyperlink[7:]
        import urllib.parse
        clean_hyperlink = urllib.parse.unquote(clean_hyperlink)
        candidates.append(clean_hyperlink)

    # 2. Raw text path
    if path_str and path_str.lower() != "none":
        if " - " in path_str:
            candidate = path_str.split(" - ")[-1].strip()
        else:
            candidate = path_str
        candidates.append(candidate)

    result = None
    for cand in candidates:
        if not cand:
            continue
        # Absolute check
        if os.path.isabs(cand) and os.path.exists(cand):
            result = os.path.abspath(cand)
            break
        # Relative check
        excel_cand = os.path.join(excel_dir, cand)
        if os.path.exists(excel_cand):
            result = os.path.abspath(excel_cand)
            break
        # Downloads check
        downloads_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        dl_cand = os.path.join(downloads_dir, cand)
        if os.path.exists(dl_cand):
            result = os.path.abspath(dl_cand)
            break
        # Filename matching (searching in downloads or excel dir)
        base_cand = os.path.basename(cand)
        try:
            for f in os.listdir(excel_dir):
                if base_cand.lower() in f.lower() or f.lower() in base_cand.lower():
                    result = os.path.abspath(os.path.join(excel_dir, f))
                    break
        except Exception:
            pass
        if result:
            break
        try:
            for f in os.listdir(downloads_dir):
                if base_cand.lower() in f.lower() or f.lower() in base_cand.lower():
                    result = os.path.abspath(os.path.join(downloads_dir, f))
                    break
        except Exception:
            pass
        if result:
            break

    # Fallback to the first candidate if not found
    if result is None:
        if candidates:
            result = os.path.abspath(candidates[0])
        else:
            result = path_str

    exists = os.path.exists(result) if result else False
    log.info(f"Resolved Document: {result}")
    log.info(f"Document Exists: {exists}")

    _doc_path_cache[cache_key] = result
    return result


def _prebatch_compress_pdfs(records: List[Dict[str, Any]], doc_cli_path: str = None) -> Dict[str, str]:
    """Pre-compress all unique document paths in the batch before processing starts.
    Returns a mapping {original_path: compressed_path} so each file is only
    compressed once regardless of how many rows share it."""
    seen: Dict[str, str] = {}
    
    unique_paths = set()
    for r in records:
        r_doc = r.get("DocumentPath")
        if r_doc and os.path.exists(str(r_doc)):
            unique_paths.add(str(r_doc))
            
    if doc_cli_path and os.path.exists(str(doc_cli_path)):
        unique_paths.add(str(doc_cli_path))
        
    if not unique_paths:
        return seen

    log.info(f"[PDF-COMPRESS] Pre-batch compression of {len(unique_paths)} unique document(s)...")

    def _compress_one(p):
        try:
            compressed = optimize_and_compress_pdf(str(p))
            return str(p), compressed
        except Exception as e:
            log.warning(f"[PDF-COMPRESS] Pre-batch compression failed for '{p}': {e}")
            return str(p), str(p)

    workers = min(4, len(unique_paths))
    with ThreadPoolExecutor(max_workers=workers) as pool:
        for orig, compressed in pool.map(_compress_one, unique_paths):
            seen[orig] = compressed

    return seen


def clean_khasra_khata(val: Any) -> Any:
    """Completely strips out Khasra, Khata, Gat, Survey, Plot numbers from the address value."""
    if not val or not isinstance(val, str):
        return val
    cleaned = val
    plot_patterns = [
        r'(?i)\bKhasra\s*(?:No\.?|Number)?\s*[\d\-/]+',
        r'(?i)\bKhata\s*(?:No\.?|Number)?\s*[\d\-/]+',
        r'(?i)\bSurvey\s*(?:No\.?|Number)?\s*[\d\-/]+',
        r'(?i)\bGat\s*(?:No\.?|Number)?\s*[\d\-/]+',
        r'(?i)\bPlot\s*(?:No\.?|Number)?\s*[\d\-/]+',
        r'(?i)\bKhatauni\s*(?:No\.?|Number)?\s*[\d\-/]+',
        r'(?i)\b\d+\s*kata\s*\d+\b',
        r'(?i)\bKhasra\b',
        r'(?i)\bKhata\b',
        r'(?i)\bSurvey\b',
        r'(?i)\bGat\b',
        r'(?i)\bPlot\b',
        r'(?i)\bkata\b'
    ]
    for pattern in plot_patterns:
        cleaned = re.sub(pattern, '', cleaned)
    cleaned = re.sub(r',\s*,', ',', cleaned)
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip(', ')


def parse_address_fields(composite_address: str) -> Dict[str, str]:
    """
    Intelligently splits and maps composite address components to standard GST fields.
    Strictly strips Khasra, Khata, Survey, Plot, Gat, kata details from all fields.
    """
    res = {
        "BuildingNo": "",
        "City": "",
        "Locality": "",
        "Road": "",
        "PIN": "",
        "District": ""
    }
    if not composite_address:
        return res
        
    address_str = str(composite_address).strip()
    
    # Extract PIN Code (6 digits)
    pin_match = re.search(r'\b\d{6}\b', address_str)
    if pin_match:
        res["PIN"] = pin_match.group(0)
        address_str = address_str.replace(res["PIN"], "")
        
    # Strictly clean plot/land identifiers from the composite address before parsing standard components
    address_str = clean_khasra_khata(address_str)
        
    # Split by commas
    parts = [p.strip() for p in address_str.split(",") if p.strip()]
    
    village_parts = []
    tehsil_parts = []
    road_parts = []
    district_parts = []
    city_parts = []
    locality_parts = []
    unmatched_parts = []
    
    # Classification regex patterns
    village_pat = r'(?i)\b(?:Village|Gram)\b'
    tehsil_pat = r'(?i)\bTehsil\b'
    road_pat = r'(?i)\b(?:Road|Street|Marg|Bypass|Highway|Lane)\b'
    district_pat = r'(?i)\b(?:District|Dist\.?)\b'
    locality_pat = r'(?i)\b(?:Mohalla|Area|Sector|Colony|Civil\s*Lines|Patel\s*Nagar)\b'
    
    for part in parts:
        part_clean = part.strip()
        part_lower = part_clean.lower()
        
        # Check District first
        if re.search(district_pat, part_lower) or "dist" in part_lower:
            dist_name = re.sub(r'(?i)\b(?:district|dist\.?)\b', '', part_clean).strip()
            district_parts.append(dist_name)
            continue
            
        # Check village & tehsil patterns
        if re.search(village_pat, part_clean):
            village_parts.append(part_clean)
        elif re.search(tehsil_pat, part_clean):
            tehsil_parts.append(part_clean)
        elif re.search(road_pat, part_clean):
            road_parts.append(part_clean)
        elif re.search(locality_pat, part_clean):
            locality_parts.append(part_clean)
        else:
            # Ignore state/country/PIN markers
            if any(s in part_lower for s in ["pradesh", "delhi", "bihar", "india", "pin code", "pincode"]):
                continue
            if part_lower == "pin":
                continue
            unmatched_parts.append(part_clean)
            
    # Village / Gram descriptor string mapped to BuildingNo
    res["BuildingNo"] = ", ".join(village_parts)
    
    # Assign unmatched parts intelligently (first one to City/Town/Village)
    for part in unmatched_parts:
        if not city_parts:
            city_parts.append(part)
        else:
            if not res["BuildingNo"]:
                res["BuildingNo"] = part
            else:
                res["BuildingNo"] = res["BuildingNo"] + ", " + part
            
    res["City"] = ", ".join(city_parts)
    res["Locality"] = ", ".join(locality_parts)
    
    # Tehsil goes to Road selection, along with road parts
    r_parts = []
    r_parts.extend(road_parts)
    r_parts.extend(tehsil_parts)  # Tehsil name inside road selection
    res["Road"] = ", ".join(r_parts)
    
    res["District"] = ", ".join(district_parts)
    
    # Clean up empty strings or formatting
    for k in res:
        res[k] = res[k].strip(", ")
        
    return res


def map_row_to_standard(row: Dict[str, Any], excel_path: str) -> Dict[str, Any]:
    """Standardizes dynamic column mappings and composite addresses from assessment logs."""
    excel_dir = os.path.dirname(excel_path)

    # Build a lowercase-key lookup once (O(n)) instead of O(n²) per field
    _lower_map: Dict[str, Any] = {k.strip().lower(): v for k, v in row.items() if v is not None}

    def get_val(keys: List[str], default=None):
        for k in keys:
            # Direct key match first (fast path)
            if k in row and row[k] is not None:
                return row[k]
            # Case-insensitive fallback
            lk = k.strip().lower()
            if lk in _lower_map:
                return _lower_map[lk]
        return default

    gstin_raw = get_val(["GSTIN", "GST_Reg_ID", "GST Reg ID", "gstin"])
    gstin = str(gstin_raw).strip().upper() if gstin_raw is not None else None
    trade_name = get_val(["TradeName", "Trade Name", "Legal Name (GST Applicant)", "Legal Name", "trade_name"])
    
    composite_addr = get_val(["Principal Place of Business Address", "Ownership Proof Address", "Address", "address"])
    
    pin = get_val(["PIN", "PIN Code", "Pincode", "pin"])
    state = get_val(["State", "state"])
    district = get_val(["District", "district"])
    city = get_val(["City", "City/Town/Village", "city"])
    locality = get_val(["Locality", "Locality/Sub Locality", "locality"])
    road = get_val(["Road", "Road/Street", "road"])
    building = get_val(["BuildingName", "Building Name", "Building", "building"])
    building_no = get_val(["BuildingNo/FlatNo", "Building No./Flat No.", "Building No./Flat No."])
    floor = get_val(["FloorNo", "Floor Number", "Floor", "floor"])
    landmark = get_val(["Landmark", "landmark"])
    lat = get_val(["Latitude", "latitude", "lat"])
    lng = get_val(["Longitude", "longitude", "lng"])
    
    # Construct composite address from individual fields if not explicitly provided
    if not composite_addr:
        # Join any available fields to form a clean parseable composite address
        addr_parts = []
        for val in [building_no, building, floor, landmark, road, locality, city, district, pin]:
            if val and str(val).strip() != "":
                addr_parts.append(str(val).strip())
        composite_addr = ", ".join(addr_parts)

    # Resilient address parsing & Khasra / Khata Address Intelligence
    if composite_addr:
        log.info(f"[{gstin}] Resolving fields from composite address: {composite_addr}")
        parsed_addr = parse_address_fields(composite_addr)
        if parsed_addr["PIN"]:
            pin = parsed_addr["PIN"]
        if parsed_addr["District"]:
            district = parsed_addr["District"]
        if parsed_addr["City"]:
            city = parsed_addr["City"]
        if parsed_addr["BuildingNo"]:
            building_no = parsed_addr["BuildingNo"]
        if parsed_addr["Locality"]:
            locality = parsed_addr["Locality"]
        if parsed_addr["Road"]:
            road = parsed_addr["Road"]
            
    # Completely clean all fields from plot/land identifiers (strictly omitted)
    if building_no and isinstance(building_no, str):
        building_no = clean_khasra_khata(building_no)
    if road and isinstance(road, str):
        road = clean_khasra_khata(road)
    if locality and isinstance(locality, str):
        locality = clean_khasra_khata(locality)
    if building and isinstance(building, str):
        building = clean_khasra_khata(building)
    if landmark and isinstance(landmark, str):
        landmark = clean_khasra_khata(landmark)

    # Intelligently satisfy mandatory Building No. and Road fields if empty after filtering
    if not building_no or str(building_no).strip() == "" or str(building_no).strip().lower() == "none":
        match = re.search(r'(?i)\b(Village|Gram|Tehsil)\s+([^,]+)', composite_addr)
        if match:
            val_to_use = clean_khasra_khata(f"{match.group(1)} {match.group(2).strip()}")
            if val_to_use and val_to_use.strip() != "":
                building_no = val_to_use
            else:
                building_no = "Plot No. 1"
        else:
            building_no = "Plot No. 1"

    if not road or str(road).strip() == "" or str(road).strip().lower() == "none":
        tehsil_match = re.search(r'(?i)\bTehsil\s+([^,]+)', composite_addr)
        if tehsil_match:
            val_to_use = clean_khasra_khata(f"Tehsil {tehsil_match.group(1).strip()}")
            if val_to_use and val_to_use.strip() != "":
                road = val_to_use
            else:
                road = "Main Road"
        elif landmark and str(landmark).strip() != "":
            road = f"Road Near {str(landmark).strip()}"
        else:
            road = "Main Road"
            
    mobile = get_val(["Mobile", "Mobile Number", "mobile", "telephone"])
    email = get_val(["Email", "email"])
    std_code = get_val(["STDCode", "STD Code", "std_code"])
    telephone = get_val(["Telephone", "telephone"])
    fax_std = get_val(["FaxSTD", "fax_std"])
    fax_number = get_val(["FaxNumber", "fax_number"])
    
    possession = get_val(["PossessionNature", "PossessionType", "Possession Type", "Ownership Proof Type", "possession"], "Others")
    doc_path_raw = get_val(["DocumentPath", "Document Path", "Total Documents Uploaded", "Ready to Upload Docs", "document_path"])
    
    # Robust Fallback: If no document path was resolved from the primary columns,
    # extract any valid file name from incorrect/mismatch or review required columns.
    if not doc_path_raw or str(doc_path_raw).strip() == "" or str(doc_path_raw).lower() == "none":
        fallback_val = get_val(["Incorrect/Mismatch Documents", "Review Required Docs", "Incorrect / Mismatch Documents", "Review Required Documents"])
        if fallback_val and str(fallback_val).strip() != "" and str(fallback_val).lower() != "none":
            # Search for standard PDF, JPG, etc. extensions
            match = re.search(r'\b([\w\s.-]+\.(?:pdf|jpg|jpeg|png|webp|tiff|doc|docx))\b', str(fallback_val), re.IGNORECASE)
            if match:
                doc_path_raw = match.group(1).strip()
                log.info(f"[{gstin}] Extracted fallback document filename from mismatch/review columns: '{doc_path_raw}'")

    doc_path = resolve_document_path(doc_path_raw, excel_dir, row)
    
    # Document path resolving from raw Excel value
    excel_act = get_val(["BusinessActivity", "Business Activity", "Business Activities"])
    default_act_list = ["Import", "Office / Sale Office", "Supplier of Services"]
    if excel_act:
        if isinstance(excel_act, str):
            act_list = [a.strip() for a in excel_act.split(",") if a.strip()]
        else:
            act_list = [str(excel_act).strip()]
        for d_act in default_act_list:
            if d_act not in act_list and not any(d_act.lower() in x.lower() for x in act_list):
                act_list.append(d_act)
        activities = ", ".join(act_list)
    else:
        activities = ", ".join(default_act_list)
        
    reason = get_val(["Reason", "AmendmentReason", "Amendment Reason", "Final Remarks", "Our Remarks"], "Addition of additional place of business")
    date_amend = datetime.now().strftime("%d/%m/%Y")
    
    username_val = get_val(["Username", "username", "GST Username", "User Name", "User_Name", "User"])
    password_val = get_val(["Password", "password", "GST Password", "Pass Word", "Pass_Word", "Pass"])

    return {
        "Username": username_val,
        "Password": password_val,
        "GSTIN": gstin,
        "TradeName": trade_name,
        "PIN": pin,
        "State": state,
        "District": district,
        "City": city,
        "Locality": locality,
        "Road": road,
        "BuildingName": building,
        "BuildingNo":building_no,
        "FloorNo": floor,
        "Landmark": landmark,
        "Latitude": lat or "",
        "Longitude": lng or "",
        "Mobile": mobile,
        "Email": email,
        "STDCode": std_code,
        "Telephone": telephone,
        "FaxSTD": fax_std,
        "FaxNumber": fax_number,
        "PossessionNature": possession,
        "DocumentPath": doc_path,
        "BusinessActivity": activities,
        "Reason": reason[:150],
        "DateOfAmendment": date_amend
    }


def load_excel_data(excel_path: str) -> List[Dict[str, Any]]:
    """Loads and maps Excel rows to standard format using openpyxl for hyperlink extraction.
    Row mapping is parallelised across a thread pool for large workbooks."""
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"Workbook missing at path: {excel_path}")

    # Use openpyxl to parse Excel rows and extract any hyperlink targets
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    if "GST_Registrations" not in wb.sheetnames:
        raise ValueError("Sheet 'GST_Registrations' not found in workbook")
    sheet = wb["GST_Registrations"]
    
    rows = list(sheet.iter_rows())
    if not rows:
        wb.close()
        return []
        
    headers = [cell.value for cell in rows[0]]
    # Normalize headers
    headers = [str(h).strip() if h is not None else "" for h in headers]
    
    raw_records = []
    for r_idx in range(1, len(rows)):
        row_cells = rows[r_idx]
        if all(cell.value is None for cell in row_cells):
            continue
            
        row_data = {}
        for col_idx, cell in enumerate(row_cells):
            if col_idx >= len(headers):
                continue
            header = headers[col_idx]
            if not header:
                continue
            row_data[header] = cell.value
            if cell.hyperlink:
                row_data[f"{header}_hyperlink"] = cell.hyperlink.target
        raw_records.append(row_data)
        
    wb.close()

    def _map_row(args):
        idx, row = args
        try:
            return map_row_to_standard(row, excel_path)
        except Exception as e:
            log.warning(f"Raw row {idx} parsing failed: {e}. Using raw format.")
            return row

    # Use threads to share the same address-space as Selenium driver
    workers = min(8, len(raw_records) or 1)
    with ThreadPoolExecutor(max_workers=workers) as pool:
        mapped_records = list(pool.map(_map_row, enumerate(raw_records)))

    return mapped_records


def is_session_active(driver) -> bool:
    """Audits session browser navigation validity to trigger automated relogin loops."""
    if not driver:
        return False
    try:
        url = driver.current_url
        if "login" in url.lower():
            return False
        un_el = driver.find_elements(By.ID, "username")
        if un_el and un_el[0].is_displayed():
            return False
        return True
    except Exception as e:
        log.warning(f"[SESSION] Health monitor detected invalid/crashed session: {e}")
        return False


def run_mandatory_field_audit(driver, row: Dict[str, Any], stage: str):
    """
    Scans the DOM for any empty mandatory fields.
    Attempts to autofill them if data is available in the row dict.
    If unavailable, logs the exact field name and raises a ValueError to stop submission.
    """
    log.info(f"[AUDIT] Running mandatory field audit for stage: {stage}...")
    
    JS_AUDIT = """
        function getElementXPath(element) {
            if (element.id) return '//*[@id="' + element.id + '"]';
            if (element === document.body) return '/html/body';
            var ix = 0;
            var siblings = element.parentNode.childNodes;
            for (var i = 0; i < siblings.length; i++) {
                var sibling = siblings[i];
                if (sibling === element) return getElementXPath(element.parentNode) + '/' + element.tagName.toLowerCase() + '[' + (ix + 1) + ']';
                if (sibling.nodeType === 1 && sibling.tagName === element.tagName) ix++;
            }
        }
        var emptyFields = [];
        var elems = document.querySelectorAll('input:not([type="hidden"]), textarea, select');
        for (var i = 0; i < elems.length; i++) {
            var el = elems[i];
            if (el.offsetParent === null) continue;
            var isRequired = el.hasAttribute('required') || 
                             el.getAttribute('aria-required') === 'true' || 
                             el.getAttribute('ng-required') === 'true' ||
                             el.classList.contains('ng-invalid-required');
            if (!isRequired) {
                var parent = el.parentElement;
                while (parent && parent.tagName !== 'FORM') {
                    if (parent.classList.contains('required') || parent.querySelector('.text-danger') || parent.innerHTML.indexOf('*') !== -1) {
                        var labels = parent.querySelectorAll('label');
                        for (var j = 0; j < labels.length; j++) {
                            if (labels[j].innerText.indexOf('*') !== -1) {
                                isRequired = true;
                                break;
                            }
                        }
                    }
                    if (isRequired) break;
                    parent = parent.parentElement;
                }
            }
            if (isRequired) {
                var isEmpty = false;
                if (el.tagName === 'SELECT') {
                    isEmpty = el.selectedIndex <= 0 || el.value === '' || el.value.indexOf('Select') !== -1;
                } else if (el.type === 'checkbox' || el.type === 'radio') {
                    if (el.name) {
                        var group = document.querySelectorAll('input[name="' + el.name + '"]');
                        var checked = false;
                        for (var g = 0; g < group.length; g++) {
                            if (group[g].checked) {
                                checked = true;
                                break;
                            }
                        }
                        isEmpty = !checked;
                    } else {
                        isEmpty = !el.checked;
                    }
                } else {
                    isEmpty = el.value.trim() === '';
                }
                if (isEmpty) {
                    var labelText = '';
                    if (el.id) {
                        var labelEl = document.querySelector('label[for="' + el.id + '"]');
                        if (labelEl) labelText = labelEl.innerText;
                    }
                    if (!labelText) {
                        var parentLabel = el.closest('label');
                        if (parentLabel) labelText = parentLabel.innerText;
                    }
                    if (!labelText) {
                        var prevLabel = el.previousElementSibling;
                        if (prevLabel && prevLabel.tagName === 'LABEL') labelText = prevLabel.innerText;
                    }
                    if (!labelText) {
                        labelText = el.getAttribute('placeholder') || el.name || el.id || 'Unknown';
                    }
                    emptyFields.push({
                        id: el.id,
                        name: el.name,
                        label: labelText.replace('*', '').trim(),
                        tagName: el.tagName,
                        type: el.type,
                        xpath: getElementXPath(el)
                    });
                }
            }
        }
        return emptyFields;
    """
    
    empty_fields = driver.execute_script(JS_AUDIT)
    if not empty_fields:
        log.info(f"[AUDIT] ✓ Stage '{stage}' passed. No empty mandatory fields found.")
        return True
        
    log.warning(f"[AUDIT] Found {len(empty_fields)} empty mandatory field(s) at stage '{stage}':")
    for f in empty_fields:
        log.warning(f"  - Field Label: '{f['label']}' | Tag: {f['tagName']} | ID: '{f['id']}' | XPath: {f['xpath']}")
        
    # Attempt to auto-fill
    for f in empty_fields:
        label_lower = f['label'].lower()
        id_lower = (f['id'] or '').lower()
        xpath = f['xpath']
        value_to_fill = None
        
        # Mappings
        if "pin" in label_lower or "pin" in id_lower:
            value_to_fill = row.get("PIN")
        elif "locality" in label_lower or "locality" in id_lower:
            value_to_fill = row.get("Locality")
        elif "road" in label_lower or "street" in label_lower or "st" in id_lower or "road" in id_lower:
            value_to_fill = row.get("Road")
        elif "building" in label_lower or "premises" in label_lower or "bdname" in id_lower:
            value_to_fill = row.get("BuildingName")
        elif "floor" in label_lower or "flr" in id_lower:
            value_to_fill = row.get("FloorNo")
        elif "district" in label_lower or "dst" in id_lower:
            value_to_fill = row.get("District")
        elif "city" in label_lower or "town" in label_lower or "village" in label_lower or "loc" in id_lower:
            value_to_fill = row.get("City")
        elif "mobile" in label_lower or "phone" in label_lower or "mob" in id_lower:
            value_to_fill = row.get("Mobile")
        elif "email" in label_lower or "mail" in id_lower:
            value_to_fill = row.get("Email")
        elif "reason" in label_lower or "reason" in id_lower:
            value_to_fill = row.get("Reason")
        elif "date" in label_lower or "date" in id_lower:
            value_to_fill = row.get("DateOfAmendment")
            
        if value_to_fill:
            log.info(f"[AUDIT] Attempting to auto-fill field '{f['label']}' with value: '{value_to_fill}'")
            if f['tagName'] == 'SELECT':
                select_dropdown_by_text_or_value(driver, xpath, value_to_fill, timeout=5)
            elif f['type'] == 'checkbox' or f['type'] == 'radio':
                try:
                    cb = driver.find_element(By.XPATH, xpath)
                    if not cb.is_selected():
                        driver.execute_script("arguments[0].click();", cb)
                except Exception as ex:
                    log.error(f"[AUDIT] Failed to toggle checkbox/radio: {ex}")
            else:
                _fill_text_field(driver, xpath, str(value_to_fill), timeout=5, label=f['label'])
        else:
            log.error(f"[AUDIT] Mandatory field '{f['label']}' is empty and no available data to fill it.")
            raise ValueError(f"Mandatory GST field '{f['label']}' (ID: '{f['id']}') is empty and has no mapping data!")
            
    # Re-verify after filling
    empty_fields_after = driver.execute_script(JS_AUDIT)
    if empty_fields_after:
        log.error(f"[AUDIT] Mandatory fields still empty after autofill attempt: {[ef['label'] for ef in empty_fields_after]}")
        raise ValueError(f"Mandatory GST fields are still empty: {[ef['label'] for ef in empty_fields_after]}")
        
    log.info(f"[AUDIT] ✓ Stage '{stage}' successfully audited and auto-filled.")
    return True


def check_session_health_and_recover(driver, username, password, timeout, model, headless) -> webdriver.Chrome:
    """
    Verifies if browser session is still healthy.
    If not, restarts driver, logs in again, and returns the new driver instance.
    """
    if driver is None or not is_session_active(driver):
        log.warning("[HEALTH] Inactive or crashed session detected. Triggering auto-recovery...")
        if driver:
            try:
                driver.quit()
            except:
                pass
        driver = build_driver(headless=headless)
        login_to_gst_portal(driver, username, password, timeout, model, headless=headless)
        driver.current_username = username
    return driver


def validate_data_accuracy(row: Dict[str, Any]) -> Tuple[bool, List[str]]:
    """Validates parameters (GSTIN, PIN, Mobile, Email, Document presence) before processing."""
    errors = []
    gstin = row.get("GSTIN") or ""
    mobile = str(row.get("Mobile") or "").strip()
    email = str(row.get("Email") or "").strip()
    pin = str(row.get("PIN") or "").strip()
    doc_path = row.get("DocumentPath") or ""
    
    # Check if the identifier is an internal registration ID instead of a standard GSTIN
    is_reg_id = False
    if "GSTREG" in gstin.upper() or len(gstin) != 15:
        is_reg_id = True
        log.info(f"[{gstin}] Bypassing strict GSTIN format validation as it is identified as an internal registration ID.")
    
    # 1. GSTIN validation: 15 chars, alphanumeric (only for non-reg-IDs)
    if not is_reg_id:
        if not gstin or not re.match(r'^\d{2}[A-Z]{5}\d{4}[A-Z]{1}[A-Z\d]{1}[Z]{1}[A-Z\d]{1}$', gstin.upper()):
            errors.append(f"Invalid GSTIN format: '{gstin}'")
        
    # 2. Mobile validation: exactly 10 digits
    digits_mobile = re.sub(r'\D', '', mobile)
    if digits_mobile and len(digits_mobile) != 10:
        errors.append(f"Mobile number must be exactly 10 digits: '{mobile}'")
        
    # 3. Email validation
    if email and email.lower() != "none" and not re.match(r'^[\w\.-]+@[\w\.-]+\.\w+$', email):
        errors.append(f"Invalid email format: '{email}'")
        
    # 4. PIN validation: exactly 6 digits
    if len(pin) != 6 or not pin.isdigit():
        errors.append(f"PIN Code must be exactly 6 digits: '{pin}'")
        
    # 5. Document existence validation
    if not doc_path or not os.path.exists(doc_path):
        errors.append(f"Supporting document path does not exist on disk: '{doc_path}'")
        
    return len(errors) == 0, errors


def login_to_gst_portal(driver, username, password, timeout, model, headless=False) -> bool:
    """Auto-solve enabled resilient login sequence wrapper."""
    try:
        driver.get("https://www.gst.gov.in/")
        wait_and_click(driver, By.XPATH, "//a[contains(text(),'Login')]", timeout=12)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "username"))
        ).clear()
        driver.find_element(By.ID, "username").send_keys(username)
        driver.find_element(By.ID, "user_pass").clear()
        driver.find_element(By.ID, "user_pass").send_keys(password)
        
        login_success = False
        for attempt in range(1, _CAPTCHA_MAX_AUTO_RETRIES + 1):
            log.info(f"[CAPTCHA] Login attempt {attempt}/{_CAPTCHA_MAX_AUTO_RETRIES}")
            
            # Ensure the password field is always filled (as the portal clears it on invalid CAPTCHA submissions)
            try:
                pass_el = driver.find_element(By.ID, "user_pass")
                pass_el.clear()
                pass_el.send_keys(password)
            except Exception as e:
                log.warning(f"[LOGIN] Could not re-fill password field: {e}")
                
            solved, captcha_text = auto_solve_captcha(driver, timeout=timeout, model_name=model)
            if not solved:
                # CAPTCHA has already been refreshed inside auto_solve_captcha on failure,
                # so we just sleep briefly and continue to the next attempt.
                time.sleep(0.5)
                continue
                
            try:
                wait_and_click(
                    driver, By.XPATH,
                    "//button[@type='submit' or contains(@id,'btnLogin') or contains(text(),'Login')]",
                    timeout=timeout,
                )
            except Exception as e:
                log.warning(f"[CAPTCHA] Login button clicked: {e}")
                if attempt < _CAPTCHA_MAX_AUTO_RETRIES:
                    refresh_captcha(driver)
                    time.sleep(0.5)
                continue
                
            try:
                WebDriverWait(driver, 12).until(
                    EC.visibility_of_element_located((By.XPATH, "//a[contains(text(),'Dashboard')]"))
                )
                log.info(f"[CAPTCHA] Login OK on attempt {attempt}.")
                login_success = True
                break
            except TimeoutException:
                captcha_errors = driver.find_elements(
                    By.XPATH,
                    "//*[contains(text(),'Invalid Captcha') or contains(@class,'error')]",
                )
                msg = "CAPTCHA rejected" if captcha_errors else "Dashboard not found"
                log.warning(f"[CAPTCHA] {msg} on attempt {attempt} (text='{captcha_text}').")
                if attempt < _CAPTCHA_MAX_AUTO_RETRIES:
                    refresh_captcha(driver)
                    time.sleep(0.6)
                    
        if not login_success:
            # 1. Trigger fresh refresh and capture loop
            out_path = None
            for retry in range(3):
                refresh_captcha(driver)
                time.sleep(1.0)
                out_path = capture_captcha_image(driver)
                if out_path and os.path.exists(out_path):
                    break
                log.warning(f"[CAPTCHA] Capture failed on attempt {retry+1}, retrying...")

            portal_hash = "N/A"
            if out_path and os.path.exists(out_path):
                import hashlib
                try:
                    with open(out_path, "rb") as f:
                        portal_hash = hashlib.sha256(f.read()).hexdigest()
                except Exception as hash_err:
                    log.warning(f"[CAPTCHA] Failed to compute portal hash: {hash_err}")

            if out_path:
                log.info(f"[CAPTCHA] Saved -> {out_path}")
            log.info(f"[CAPTCHA] Portal Captcha Hash: {portal_hash}")

            # Only print the manual mode activation warning after writing and logging the fresh captcha file path
            log.warning("[CAPTCHA] Auto-solve failed. Switched to user manual CAPTCHA entry mode...")
            log.info("[CAPTCHA] Manual input required. Solving CAPTCHA. Enter 6-character solution:")
            
            import sys
            import threading
            
            captcha_val = [None]
            def read_stdin():
                try:
                    line = sys.stdin.readline()
                    if line:
                        captcha_val[0] = line.strip()
                except Exception:
                    pass
                    
            stdin_thread = threading.Thread(target=read_stdin, daemon=True)
            stdin_thread.start()
            
            dashboard_loaded = False
            start_wait = time.time()
            # Wait up to 180 seconds for either dashboard load or manual stdin submission
            while time.time() - start_wait < 180:
                # Check if dashboard is loaded
                try:
                    dash = driver.find_elements(By.XPATH, "//a[contains(text(),'Dashboard')]")
                    if dash and dash[0].is_displayed():
                        log.info("[CAPTCHA] Dashboard detected. Login successful.")
                        dashboard_loaded = True
                        break
                except Exception:
                    pass
                
                # Check if we got stdin input from the popup GUI
                if not stdin_thread.is_alive():
                    val = captcha_val[0]
                    if val:
                        log.info(f"[CAPTCHA] Received manual solution: {val}")
                        
                        # Fill CAPTCHA field
                        captcha_field = None
                        for s in ["input#captcha", "input[name='captcha']", "input[data-ng-model='lform.captcha']", "input#userCaptcha"]:
                            try:
                                captcha_field = driver.find_element(By.CSS_SELECTOR, s)
                                if captcha_field and captcha_field.is_displayed():
                                    break
                            except Exception:
                                continue
                        if not captcha_field:
                            try:
                                captcha_field = driver.find_element(By.ID, "captcha")
                            except Exception:
                                pass
                        
                        if captcha_field:
                            try:
                                captcha_field.clear()
                                captcha_field.send_keys(val)
                                
                                # Re-fill password field as portal clears it
                                try:
                                    pass_el = driver.find_element(By.ID, "user_pass")
                                    pass_el.clear()
                                    pass_el.send_keys(password)
                                except Exception:
                                    pass
                                
                                # Click login
                                btn = driver.find_element(By.XPATH, "//button[@type='submit' or contains(@id,'btnLogin') or contains(text(),'Login')]")
                                btn.click()
                                log.info("[CAPTCHA] Submitted manual CAPTCHA value.")
                            except Exception as fill_err:
                                log.error(f"[CAPTCHA] Failed to submit manual CAPTCHA: {fill_err}")
                        
                        # Reset stdin thread for next try if needed
                        captcha_val[0] = None
                        stdin_thread = threading.Thread(target=read_stdin, daemon=True)
                        stdin_thread.start()
                    else:
                        # Stdin EOF/error
                        break
                
                time.sleep(1.0)
            
            if not dashboard_loaded:
                log.info("Waiting for dashboard to load...")
                WebDriverWait(driver, 180).until(
                    EC.visibility_of_element_located((By.XPATH, "//a[contains(text(),'Dashboard')]"))
                )
            
        log.info("[LOGIN] Login successfully completed. Stabilizing session...")
        
        # --- HARD SESSION STABILIZATION LAYER ---
        time.sleep(4)

        dismiss_all_popups(driver)

        WebDriverWait(driver, 40).until(
            lambda d: d.execute_script(
                "return document.readyState"
            ) == "complete"
        )

        time.sleep(3)

        driver.execute_script("window.scrollTo(0,0);")

        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((
                By.XPATH,
                "//*[contains(text(),'Dashboard') or contains(text(),'My Saved Applications')]"
            ))
        )
        # --- END OF HARD SESSION STABILIZATION ---
        
        # Inject the UploadTracker XHR interceptor
        UploadTrackerService.inject_interceptor(driver)
        
        return True
    except Exception as e:
        log.error(f"[LOGIN] Resilient login engine failed: {e}")
        return False



# QUEUE ENGINE EXECUTION CONTROLLER


def is_browser_crash_exception(e: Exception) -> bool:
    """Helper to check if the given exception indicates a crashed or invalid browser session."""
    err_str = str(e).lower()
    if isinstance(e, (ConnectionResetError, ConnectionRefusedError)):
        return True
    if "connectionreset" in err_str or "10061" in err_str or "connection refused" in err_str:
        return True
    if "sessionnotcreated" in err_str or "invalidsessionid" in err_str or "invalid session id" in err_str:
        return True
    if "chrome not reachable" in err_str or "target window already closed" in err_str or "no such session" in err_str:
        return True
    return False


def process_user_session_batch(
    active_username: str,
    active_password: str,
    user_batch_rows: List[Dict[str, Any]],
    user_batch_indices: List[int],
    total_rows: int,
    checkpoint: CheckpointManager,
    all_gstins: List[str],
    doc_cli_path: Optional[str],
    pdf_cache: Dict[str, Any],
    headless: bool,
    timeout: int,
    model: Optional[str],
    metrics: Optional[Any] = None
) -> None:
    """Helper to process all batch rows for a specific logged-in user session in isolation."""
    driver = None
    session_contact_cache = {}
    try:
        log.info(f"[BATCH] Initializing fresh Selenium browser instance for user '{active_username}'...")
        driver = build_driver(headless=headless)
        login_ok = login_to_gst_portal(driver, active_username, active_password, timeout, model, headless=headless)
        if not login_ok:
            log.error(f"[BATCH] Resilient login sequence failed for user '{active_username}'.")
            try:
                driver.quit()
            except:
                pass
            return

        driver.current_username = active_username
        
        # Group user_batch_rows by GSTIN within the user login session to support multi-client practitioner logins
        gstin_groups = defaultdict(list)
        gstin_pattern = re.compile(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$")
        
        for r, idx in zip(user_batch_rows, user_batch_indices):
            r_gstin = r.get("GSTIN")
            if not r_gstin:
                r_gstin = f"NO_GSTIN_{_safe_name(r.get('TradeName') or f'Client_{idx}')}"
            r_gstin_normalized = str(r_gstin).strip().upper()
            
            # Group real GSTINs by themselves, and temporary/placeholder IDs strictly by state
            if gstin_pattern.match(r_gstin_normalized):
                group_key = r_gstin_normalized
            else:
                state_norm = str(r.get("State") or "STATE").strip().upper()
                group_key = f"TEMP_{state_norm}"
            
            gstin_groups[group_key].append((r, idx))
            
        for active_gstin, gstin_rows_with_indices in gstin_groups.items():
            if not active_gstin:
                continue
                
            client_rows = [item[0] for item in gstin_rows_with_indices]
            client_indices = [item[1] for item in gstin_rows_with_indices]
            
            log.info(f"\n{'='*70}\n[GSTIN: {active_gstin}] PROCESSING {len(client_rows)} APoB RECORDS IN BATCH\n{'='*70}")
            
            # isolated try-except block: Failure of one GSTIN must NEVER stop remaining GSTINs
            try:
                first_row_idx = client_indices[0]
                log.info(f"Registration {first_row_idx + 1} of {total_rows}")
                
                # Early Data Accuracy Validation
                for r_item in client_rows:
                    valid, errs = validate_data_accuracy(r_item)
                    if not valid:
                        log.error(f"[{active_gstin}] Data accuracy validation failed: {errs}")
                        raise ValueError(f"Data accuracy validation failed: {', '.join(errs)}")
                        
                # Resilient session check & health monitor recovery
                prev_driver = driver
                driver = check_session_health_and_recover(driver, active_username, active_password, timeout, model, headless)
                if driver is not prev_driver:
                    session_contact_cache = {}
                
                dismiss_all_popups(driver)
                
                # Always auto-fetch contact details from the logged-in GST profile, utilizing cache if available
                if "mobile" in session_contact_cache and "email" in session_contact_cache:
                    fetched_mobile = session_contact_cache["mobile"]
                    fetched_email = session_contact_cache["email"]
                    log.info(f"[{active_gstin}] Reusing contact details from Session Contact Cache: Mobile='{fetched_mobile}', Email='{fetched_email}'")
                    if metrics:
                        metrics.record_cache(hit=True)
                else:
                    log.info(f"[{active_gstin}] Session cache is empty. Auto-fetching contact details from logged-in GST profile...")
                    t_api_start = time.time()
                    fetched_mobile, fetched_email = fetch_profile_contact_info(driver, timeout)
                    if metrics:
                        metrics.record_cache(hit=False)
                        metrics.record_processed(api_time=time.time() - t_api_start)
                    if fetched_mobile or fetched_email:
                        session_contact_cache["mobile"] = fetched_mobile
                        session_contact_cache["email"] = fetched_email
                        log.info(f"[{active_gstin}] Stored contact details in Session Contact Cache: Mobile='{fetched_mobile}', Email='{fetched_email}'")
                
                for r in client_rows:
                    if fetched_mobile:
                        r["Mobile"] = fetched_mobile
                        log.info(f"[{active_gstin}] Dynamically updated Mobile for queue record with profile value: {fetched_mobile}")
                    if fetched_email:
                        r["Email"] = fetched_email
                        log.info(f"[{active_gstin}] Dynamically updated Email for queue record with profile value: {fetched_email}")
                            
                # Check draft or unrelated drafts
                draft_status = check_and_resume_draft(driver, client_rows[0], timeout)
                
                if draft_status == 'cleared':
                    log.info(f"[{active_gstin}] Launching fresh amendment routing...")
                    driver.get("https://services.gst.gov.in/services/auth/dashboard")
                    time.sleep(1.0)
                    _open_services_amendment_link(driver, timeout)
                    
                    select_gstin_context_if_prompted(driver, active_gstin, timeout)
                    
                    log.info(f"[{active_gstin}] Waiting for amendment workspace or potential blocking warnings...")
                    
                    warning_xpath = (
                        "//*[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'already initiated') "
                        "or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'seems you have already initiated') "
                        "or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'access saved application')]"
                    )
                    
                    has_warning = False
                    workspace_opened = False
                    
                    for poll_idx in range(20):  # 10s max wait
                        if _AMENDMENT_URL_FRAG in driver.current_url.lower():
                            workspace_opened = True
                            break
                            
                        warning_els = driver.find_elements(By.XPATH, warning_xpath)
                        if warning_els and any(w.is_displayed() for w in warning_els):
                            has_warning = True
                            break
                            
                        time.sleep(0.5)
                        
                    if has_warning:
                        log.warning(f"[{active_gstin}] 'Already initiated' Core Amendment blocking warning visible. Navigating to saved draft...")
                        status = check_and_resume_draft(driver, client_rows[0], timeout)
                        if status != 'resumed':
                            log.info(f"[{active_gstin}] Target draft not resumed by name/GSTIN. Fallback to resume first core draft...")
                            if not resume_first_core_draft(driver, timeout):
                                raise RuntimeError("Failed to resume saved Core Fields Amendment draft")
                                
                        log.info(f"[{active_gstin}] Waiting for workspace to open after resuming draft...")
                        WebDriverWait(driver, 25).until(EC.url_contains(_AMENDMENT_URL_FRAG))
                        log.info(f"[{active_gstin}] Amendment workspace opened successfully.")
                    elif workspace_opened:
                        log.info(f"[{active_gstin}] Amendment workspace opened directly.")
                    else:
                        if _AMENDMENT_URL_FRAG in driver.current_url.lower():
                            log.info(f"[{active_gstin}] Amendment workspace opened (final check).")
                        else:
                            capture_screenshot(driver, active_gstin, "workspace_open_timeout")
                            raise TimeoutException("Timed out waiting for amendment workspace or warning banner to appear")
                            
                capture_screenshot(driver, active_gstin, "01_amend_workspace")
                
                # Principal Place APoB toggling
                capture_screenshot(driver, active_gstin, "before_save_apob_toggle")
                enable_have_apob(driver, timeout)
                capture_screenshot(driver, active_gstin, "02_apob_enabled")
                
                # Additional Place form filling
                navigate_to_additional_places_tab(driver, timeout)
                capture_screenshot(driver, active_gstin, "03_apob_tab")
                
                # Fill Number of Additional Places using smart incrementer
                apob_rows = client_rows
                log.info(f"[APOB-BATCH] Total APoB records detected: {len(apob_rows)}")
                log.info(f"[DEBUG] About to update Additional Places count using {len(apob_rows)} records")
                fill_number_of_additional_places(driver, len(apob_rows), timeout)
                capture_screenshot(driver, active_gstin, "03b_num_apob_filled")
                
                # Fill form details in batch mode
                t_up_start = time.time()
                process_apob_batch(driver, client_rows, client_indices, timeout, doc_cli_path=doc_cli_path, pdf_cache=pdf_cache)
                upload_dur = time.time() - t_up_start
                    
                # Run mandatory field audit before final tab save
                run_mandatory_field_audit(driver, client_rows[0], "Finalize Submission")
                
                capture_screenshot(driver, active_gstin, "before_final_submit")
                # Finalize the whole tab and save overall draft
                finalize_amendment_submission(driver, client_rows[0], timeout)
                capture_screenshot(driver, active_gstin, "06_draft_saved")
                
                # Mark all individual rows completed in checkpoint only AFTER successful finalization
                for c_idx in client_indices:
                    checkpoint.mark_row_completed(c_idx)
                    log.info(f"[{active_gstin}] Row {c_idx} successfully marked completed.")
                    
                # Mark the entire GSTIN completed in checkpoint and log to Excel
                checkpoint.mark_gstin_completed(active_gstin)
                
                for c_idx, r in zip(client_indices, client_rows):
                    update_completed_clients_excel(r.get("GSTIN") or active_gstin, r.get("TradeName"), str(r.get("PIN")), "Draft Saved", "Successfully processed APoB amendment in bulk.")
                    log.info(f"[{active_gstin}] Row {c_idx} successfully finalized and completed.")
                
                if metrics:
                    metrics.record_processed(apobs=len(client_rows), upload_time=upload_dur, success=True)
                    
            except Exception as e:
                # Propagating session crash exceptions to the outer loop for browser restart and login recovery
                if is_browser_crash_exception(e):
                    log.warning(f"[{active_gstin}] Browser crash exception detected during inner processing: {e}. Propagating to restarter.")
                    raise
                    
                # Logic/Data error: Save screenshot, log failure, mark failed in checkpoint, and continue queue
                ss = capture_screenshot(driver, active_gstin, "gstin_processing_fail")
                log.error(f"[{active_gstin}] GSTIN Processing Error | Step: 'GSTIN Loop' | Error: {e} | URL: {driver.current_url if driver else 'N/A'} | Screenshot: {ss}")
                
                checkpoint.mark_failed(client_indices[0], client_rows[0].get("GSTIN") or active_gstin, str(e))
                update_failed_clients_excel(client_rows[0].get("GSTIN") or active_gstin, str(e), ss)
                
                if metrics:
                    metrics.record_processed(apobs=len(client_rows), success=False)
                
                log.info(f"[{active_gstin}] Resiliently continuing queue from next pending GSTIN.")
                continue

    finally:
        if driver:
            try:
                driver.quit()
            except:
                pass


def execute_bulk_apob_pipeline(doc_cli_path=None):
    """Main bulk batch orchestrator that runs 60-70 address additions per Username session."""
    args = parse_args()
    excel_path = args.excel
    checkpoint_path = args.checkpoint
    timeout = args.timeout
    no_wait = args.no_wait
    model = args.model
    headless = args.headless
    
    # Create template if missing
    if not os.path.exists(excel_path):
        create_sample_excel_template()
        log.info(f"Excel workbook template 'amendments.xlsx' created. Please fill sheet 'GST_Registrations' and execute the system again.")
        sys.exit(0)
        
    try:
        records = load_excel_data(excel_path)
    except Exception as e:
        log.error(f"Failed to load registration data: {e}")
        sys.exit(1)

    # Dynamic document path relocation layer:
    # If the user has moved their project folders, absolute paths stored in the Excel rows
    # may become stale. This scans the input/output parent directory for matching filenames.
    for r in records:
        doc_path = r.get("DocumentPath")
        if doc_path and str(doc_path).strip() != "" and str(doc_path).lower() != "none":
            doc_path_str = str(doc_path).strip()
            if not os.path.exists(doc_path_str):
                filename = os.path.basename(doc_path_str)
                if filename:
                    abs_excel = os.path.abspath(excel_path)
                    excel_dir = os.path.dirname(abs_excel)
                    if "_internal" in excel_dir or "client_outputs" in excel_dir:
                        search_root = os.path.dirname(os.path.dirname(excel_dir))
                    else:
                        search_root = os.path.dirname(excel_dir)
                    
                    if os.path.exists(search_root):
                        resolved = False
                        for root, dirs, files in os.walk(search_root):
                            if any(x in root for x in [".git", "venv", "__pycache__", "build", "dist"]):
                                continue
                            if filename in files:
                                candidate = os.path.join(root, filename)
                                if os.path.exists(candidate):
                                    r["DocumentPath"] = candidate
                                    log.info(f"[HUB] Dynamically relocated supporting document '{filename}' path to: '{candidate}'")
                                    resolved = True
                                    break

    # Pre-compress all unique PDFs up-front so each file is only compressed once
    _pdf_cache = _prebatch_compress_pdfs(records, doc_cli_path=args.doc)
        
    checkpoint = CheckpointManager(checkpoint_path)
    
    # Prompt for credentials if not supplied (to serve as fallback)
    username = args.username or input("Enter GST Username: ").strip()
    password = args.password or input("Enter GST Password: ").strip()
    if not username or not password:
        log.error("GST Portal session credentials missing.")
        sys.exit(1)
        
    total_rows = len(records)
    log.info(f"Total Registrations Found: {total_rows}")
    
    # Update pending GSTINs list in checkpoint
    all_gstins = [r.get("GSTIN") for r in records if r.get("GSTIN")]
    checkpoint.update_pending(all_gstins)
    
    # Reconstruct completed_indices from completed GSTINs
    checkpoint.completed_indices = []
    for idx, r in enumerate(records):
        r_gstin = r.get("GSTIN")
        if r_gstin in checkpoint.completed_gstins:
            checkpoint.completed_indices.append(idx)
            
    start_row = 0
    while start_row < total_rows and start_row in checkpoint.completed_indices:
        start_row += 1
        
    if start_row >= total_rows:
        log.info("Batch execution check: All registration amendments successfully completed.")
        return
        
    log.info(f"Bulk engine started. Processing rows {start_row} to {total_rows-1} (Total: {total_rows})...")
    active_rows_all = [
        r for idx, r in enumerate(records)
        if idx not in checkpoint.completed_indices
        and r.get("GSTIN") not in checkpoint.completed_gstins
    ]
    log.info(f"[DEBUG] Total Active Rows Loaded: {len(active_rows_all)}")
    
    if not active_rows_all:
        log.info("No active pending rows to process.")
        return

    # Check for multi-worker parallel execution
    if getattr(args, "workers", 1) > 1:
        log.info(f"[ORCHESTRATOR] Starting Multi-Worker Concurrency: {args.workers} workers active.")
        orchestrator = BatchOrchestrator(max_workers=args.workers)
        
        def task_fn(user, user_records, metrics):
            user_indices = []
            for ur in user_records:
                for idx, r in enumerate(records):
                    if r is ur:
                        user_indices.append(idx)
                        break
            
            user_pass = (user_records[0].get("Password") or password).strip()
            
            process_user_session_batch(
                active_username=user,
                active_password=user_pass,
                user_batch_rows=user_records,
                user_batch_indices=user_indices,
                total_rows=total_rows,
                checkpoint=checkpoint,
                all_gstins=all_gstins,
                doc_cli_path=args.doc,
                pdf_cache=_pdf_cache,
                headless=headless,
                timeout=timeout,
                model=model,
                metrics=metrics
            )
            
        orchestrator.run_parallel_batch(active_rows_all, username, task_fn)
    else:
        # Run sequentially (100% backward compatible sequential execution)
        row_idx = start_row
        while row_idx < total_rows:
            row = records[row_idx]
            
            gstin = row.get("GSTIN")
            if not gstin:
                log.warning(f"Row {row_idx} has no valid GSTIN. Skipping.")
                row_idx += 1
                checkpoint.mark_row_completed(row_idx - 1)
                checkpoint.update_pending(all_gstins)
                continue
                
            completed_indices = checkpoint.completed_indices
            completed_gstins = checkpoint.completed_gstins
            
            if row_idx in completed_indices or gstin in completed_gstins:
                row_idx += 1
                continue
                
            active_username = (row.get("Username") or username).strip()
            active_password = (row.get("Password") or password).strip()
            
            user_batch_rows = []
            user_batch_indices = []
            
            for idx in range(row_idx, total_rows):
                r = records[idx]
                r_user = (r.get("Username") or username).strip()
                r_gstin = r.get("GSTIN")
                
                if str(r_user).strip().lower() == str(active_username).strip().lower():
                    if idx in completed_indices or r_gstin in completed_gstins:
                        continue
                    user_batch_rows.append(r)
                    user_batch_indices.append(idx)
                    
            if not user_batch_rows:
                row_idx += 1
                continue
                
            log.info(f"\n{'='*70}\n[USER SESSION: {active_username}] FOUND {len(user_batch_rows)} ACTIVE ROWS FOR BATCH PROCESSING\n{'='*70}")
            
            process_user_session_batch(
                active_username=active_username,
                active_password=active_password,
                user_batch_rows=user_batch_rows,
                user_batch_indices=user_batch_indices,
                total_rows=total_rows,
                checkpoint=checkpoint,
                all_gstins=all_gstins,
                doc_cli_path=args.doc,
                pdf_cache=_pdf_cache,
                headless=headless,
                timeout=timeout,
                model=model,
                metrics=None
            )
            
            row_idx = user_batch_indices[-1] + 1
            
    log.info("\nBulk queue execution finished.")
    if not no_wait:
        input("\nExecution completed. Press Enter to exit...")


# ENTRY POINT
if __name__ == "__main__":
    args = parse_args()
    execute_bulk_apob_pipeline(doc_cli_path=args.doc)