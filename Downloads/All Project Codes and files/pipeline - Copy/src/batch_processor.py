import os
import time
from typing import Dict, List, Any
from concurrent.futures import ThreadPoolExecutor
from src.logger import log, track_execution_time
from src.doc_cache import DocumentCacheManager
from src.state_machine import RegistrationStateMachine, RegistrationState

class BatchProcessor:
    """
    Enterprise Batch Processing Engine.
    Handles queue management, batch orchestration, and concurrent preprocessing
    of non-portal operations like PDF compression and validation.
    """
    def __init__(self, records: List[Dict[str, Any]], checkpoint_mgr = None):
        self.records = records
        self.checkpoint_mgr = checkpoint_mgr
        self.state_machine = RegistrationStateMachine(RegistrationState.DASHBOARD)
        self.doc_cache = DocumentCacheManager()

    @track_execution_time("Pre-batch PDF Cryptographic Preprocessing")
    def preprocess_documents(self, doc_cli_path: str = None) -> Dict[str, str]:
        """
        Parallelized SHA256 hash-based PDF pre-compression.
        Runs concurrent compression using ProcessPool/ThreadPool for non-UI work.
        """
        seen: Dict[str, str] = {}
        unique_paths = set()
        
        for r in self.records:
            r_doc = r.get("DocumentPath")
            if r_doc and os.path.exists(str(r_doc)):
                unique_paths.add(str(r_doc))
                
        if doc_cli_path and os.path.exists(str(doc_cli_path)):
            unique_paths.add(str(doc_cli_path))
            
        if not unique_paths:
            return seen

        log.info(f"[BATCH] Initiating concurrent PDF pre-compression for {len(unique_paths)} unique files...")
        
        uncached_paths = set()
        for p in unique_paths:
            file_hash = self.doc_cache.calculate_sha256(str(p))
            cached_compressed = self.doc_cache.get_compressed_pdf(file_hash)
            if cached_compressed:
                log.info(f"[CACHE] Hash matched for '{p}'. Reusing cached PDF: '{cached_compressed}'")
                seen[str(p)] = cached_compressed
            else:
                uncached_paths.add(str(p))

        if not uncached_paths:
            return seen

        # ThreadPoolExecutor is ideal as PDF compression inside pypdf is I/O & C-bound
        workers = min(4, len(uncached_paths))
        
        def _compress_one(p):
            try:
                from amend_req import optimize_and_compress_pdf
                compressed = optimize_and_compress_pdf(str(p))
                file_hash = self.doc_cache.calculate_sha256(str(p))
                self.doc_cache.set_compressed_pdf(file_hash, str(p), compressed)
                return str(p), compressed
            except Exception as e:
                log.warning(f"[PDF-COMPRESS] Pre-batch compression failed for '{p}': {e}")
                return str(p), str(p)

        with ThreadPoolExecutor(max_workers=workers) as executor:
            results = list(executor.map(_compress_one, uncached_paths))
            for orig, compressed in results:
                seen[orig] = compressed
            
        return seen

