import os
import json
import hashlib
import threading
from pathlib import Path
from typing import Dict, Any, Optional
from src.logger import log

class DocumentCacheManager:
    """
    Optimized Cryptographic SHA256 document cache manager.
    Caches compressed PDFs and metadata to prevent redundant processing.
    """
    _instance = None
    _lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        with cls._lock:
            if cls._instance is None:
                cls._instance = super(DocumentCacheManager, cls).__new__(cls)
                cls._instance._initialized = False
            return cls._instance

    def __init__(self, cache_file_path: str = "document_cache.json"):
        if self._initialized:
            return
        self.cache_file_path = cache_file_path
        self.cache: Dict[str, Dict[str, Any]] = {}
        self.lock = threading.Lock()
        self.load()
        self._initialized = True

    def calculate_sha256(self, filepath: str) -> str:
        """
        Calculates SHA256 hash of a file efficiently by streaming chunks.
        """
        sha256 = hashlib.sha256()
        try:
            with open(filepath, "rb") as f:
                while True:
                    data = f.read(65536) # 64kb chunks
                    if not data:
                        break
                    sha256.update(data)
            return sha256.hexdigest()
        except Exception as e:
            log.error(f"[CACHE] Error hashing file '{filepath}': {e}")
            return filepath # fallback to filename if file cannot be read

    def load(self) -> None:
        """
        Loads document cache from local JSON store.
        """
        with self.lock:
            if os.path.exists(self.cache_file_path):
                try:
                    with open(self.cache_file_path, "r", encoding="utf-8") as f:
                        self.cache = json.load(f)
                    log.info(f"[CACHE] Loaded {len(self.cache)} entries from '{self.cache_file_path}'.")
                except Exception as e:
                    log.warning(f"[CACHE] Failed to load cache: {e}. Resetting.")
                    self.cache = {}
            else:
                self.cache = {}

    def save(self) -> None:
        """
        Saves document cache to local JSON store using atomic write.
        """
        with self.lock:
            tmp_path = self.cache_file_path + ".tmp"
            try:
                with open(tmp_path, "w", encoding="utf-8") as f:
                    json.dump(self.cache, f, indent=4)
                os.replace(tmp_path, self.cache_file_path)
            except Exception as e:
                log.error(f"[CACHE] Failed to write cache: {e}")
                if os.path.exists(tmp_path):
                    try:
                        os.remove(tmp_path)
                    except OSError:
                        pass

    def get_compressed_pdf(self, file_hash: str) -> Optional[str]:
        """
        Returns cached compressed PDF path if exists and file still exists.
        """
        with self.lock:
            entry = self.cache.get(file_hash)
            if entry:
                compressed_path = entry.get("compressed_path")
                if compressed_path and os.path.exists(compressed_path):
                    return compressed_path
            return None

    def set_compressed_pdf(self, file_hash: str, original_path: str, compressed_path: str) -> None:
        """
        Stores the compressed PDF path in cache.
        """
        with self.lock:
            self.cache[file_hash] = {
                "original_path": original_path,
                "compressed_path": compressed_path,
                "timestamp": time.time()
            }
        self.save()
