import time
import threading
from typing import Optional, Tuple
from src.logger import log

class OCRService:
    """
    Global thread-safe OCR Service implementing the Singleton pattern.
    Loads heavy models (TrOCR / PaddleOCR) exactly once and shares memory.
    """
    _instance = None
    _lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(OCRService, cls).__new__(cls)
                cls._instance._initialized = False
            return cls._instance

    def initialize(self, model_name: Optional[str] = None):
        """
        Thread-safe warm startup of the OCR model.
        """
        with self._lock:
            if self._initialized:
                return
            
            log.info("[OCR] Initializing global thread-safe OCR model...")
            t0 = time.perf_counter()
            
            # Dynamic imports to prevent startup delay when importing src
            global cv2, np, Image, torch, TrOCRProcessor, VisionEncoderDecoderModel, _OCR_AVAILABLE
            try:
                import cv2
                import numpy as np
                from PIL import Image
                import torch
                from transformers import TrOCRProcessor, VisionEncoderDecoderModel
                _OCR_AVAILABLE = True
            except ImportError as e:
                _OCR_AVAILABLE = False
                log.warning(f"[OCR] OCR libraries are unavailable. Using manual entry. Error: {e}")
                self._initialized = True
                self.ocr_engine = None
                return

            if not _OCR_AVAILABLE:
                self.ocr_engine = None
                self._initialized = True
                return

            # Warm startup: load the model
            name = model_name or "microsoft/trocr-base-printed"
            device = "cuda" if torch.cuda.is_available() else "cpu"
            log.info(f"[OCR] Loading TrOCR model candidate '{name}' on {device.upper()}...")
            
            try:
                self.processor = TrOCRProcessor.from_pretrained(name)
                self.model = VisionEncoderDecoderModel.from_pretrained(name, low_cpu_mem_usage=False)
                self.model.to(device)
                self.model.eval()
                self.device = device
                log.info(f"[OCR] Global TrOCR model '{name}' loaded successfully in {time.perf_counter() - t0:.2f}s.")
            except Exception as e:
                log.error(f"[OCR] FAILED to load TrOCR model candidate '{name}': {e}. Loading fallback base-printed...")
                try:
                    self.processor = TrOCRProcessor.from_pretrained("microsoft/trocr-base-printed")
                    self.model = VisionEncoderDecoderModel.from_pretrained("microsoft/trocr-base-printed", low_cpu_mem_usage=False)
                    self.model.to(device)
                    self.model.eval()
                    self.device = device
                    log.info(f"[OCR] Global Fallback TrOCR model 'microsoft/trocr-base-printed' loaded successfully.")
                except Exception as ex:
                    log.error(f"[OCR] Fallback also failed: {ex}. Automated OCR is disabled.")
                    self.ocr_engine = None
                    self._initialized = True
                    return

            self._initialized = True

    def run_ocr(self, bgr_img) -> Tuple[str, float]:
        """
        Executes TrOCR inference on preprocessed image.
        """
        if not self._initialized:
            self.initialize()
            
        if not hasattr(self, "model") or self.model is None:
            return "", 0.0

        t0 = time.perf_counter()
        try:
            rgb = cv2.cvtColor(bgr_img, cv2.COLOR_BGR2RGB)
            pil_img = Image.fromarray(rgb)

            pixel_values = self.processor(images=pil_img, return_tensors="pt").pixel_values.to(self.device)

            with torch.no_grad():
                outputs = self.model.generate(
                    pixel_values,
                    num_beams=3,
                    max_new_tokens=20,
                    output_scores=True,
                    return_dict_in_generate=True,
                )

            sequences = outputs.sequences
            raw_text = self.processor.batch_decode(sequences, skip_special_tokens=True)[0].strip()

            if hasattr(outputs, "sequences_scores") and outputs.sequences_scores is not None:
                conf = float(torch.exp(outputs.sequences_scores[0]).item())
                conf = max(0.0, min(1.0, conf))
            else:
                conf = 0.5

            log.info(f"[OCR] Captcha solve result: '{raw_text}' | Confidence: {conf*100:.1f}% | Time: {(time.perf_counter()-t0)*1000:.1f}ms")
            return raw_text, conf

        except Exception as e:
            log.error(f"[OCR] TrOCR inference exception: {e}")
            return "", 0.0
