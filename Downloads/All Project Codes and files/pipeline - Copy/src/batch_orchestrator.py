import time
import logging
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from typing import Dict, Any, List, Callable, Optional

log = logging.getLogger("GST-Amendment.Orchestrator")

# ── METRICS COLLECTOR ─────────────────────────────────────────────────────────

class MetricsCollector:
    """Thread-safe metrics accumulator for monitoring and reporting."""
    def __init__(self):
        self._lock = threading.Lock()
        self.start_time = time.time()
        self.records_processed = 0
        self.apobs_processed = 0
        self.total_upload_time = 0.0
        self.total_api_time = 0.0
        self.cache_hits = 0
        self.cache_misses = 0
        self.failures = 0
        self.worker_times: Dict[str, float] = {}

    def record_processed(self, apobs: int = 1, upload_time: float = 0.0, api_time: float = 0.0, success: bool = True):
        with self._lock:
            self.records_processed += 1
            self.apobs_processed += apobs
            self.total_upload_time += upload_time
            self.total_api_time += api_time
            if not success:
                self.failures += 1

    def record_cache(self, hit: bool = True):
        with self._lock:
            if hit:
                self.cache_hits += 1
            else:
                self.cache_misses += 1

    def record_worker_time(self, worker_name: str, duration: float):
        with self._lock:
            self.worker_times[worker_name] = duration

    def generate_summary(self) -> Dict[str, Any]:
        with self._lock:
            total_duration = time.time() - self.start_time
            avg_time = total_duration / max(1, self.records_processed)
            hit_ratio = (self.cache_hits / max(1, self.cache_hits + self.cache_misses)) * 100
            
            return {
                "total_duration_sec": round(total_duration, 2),
                "records_processed": self.records_processed,
                "apobs_processed": self.apobs_processed,
                "total_upload_time_sec": round(self.total_upload_time, 2),
                "total_api_response_time_sec": round(self.total_api_time, 2),
                "average_record_processing_time_sec": round(avg_time, 2),
                "cache_hit_ratio_pct": round(hit_ratio, 2),
                "failures_count": self.failures,
                "worker_utilization": {k: f"{v:.1f}s" for k, v in self.worker_times.items()}
            }

# ── WORKER MANAGER ────────────────────────────────────────────────────────────

class WorkerManager:
    """Manages worker threads, browser cleanup, and failure isolation."""
    def __init__(self):
        self._active_workers: List[str] = []
        self._lock = threading.Lock()

    def register_worker(self, name: str):
        with self._lock:
            self._active_workers.append(name)
            log.info(f"[WORKER-MGR] Worker '{name}' registered and starting browser session.")

    def unregister_worker(self, name: str):
        with self._lock:
            if name in self._active_workers:
                self._active_workers.remove(name)
            log.info(f"[WORKER-MGR] Worker '{name}' finished and browser session closed.")

    def get_active_count(self) -> int:
        with self._lock:
            return len(self._active_workers)

# ── BATCH ORCHESTRATOR ────────────────────────────────────────────────────────

class BatchOrchestrator:
    """
    Groups records strictly by GST username credentials, spawns and coordinates parallel workers
    using a thread pool, aggregates progress statistics, and generates the final metrics report.
    """
    def __init__(self, max_workers: int = 2):
        self.max_workers = max_workers
        self.worker_manager = WorkerManager()
        self.metrics = MetricsCollector()

    def partition_records(self, records: List[Dict[str, Any]], global_username: str) -> Dict[str, List[Dict[str, Any]]]:
        """Groups uncompleted records strictly by target GST username to prevent session overlaps."""
        groups: Dict[str, List[Dict[str, Any]]] = {}
        for r in records:
            user = (r.get("Username") or global_username).strip()
            if not user:
                continue
            groups.setdefault(user, []).append(r)
        return groups

    def run_parallel_batch(
        self,
        records: List[Dict[str, Any]],
        global_username: str,
        worker_task_fn: Callable[[str, List[Dict[str, Any]], MetricsCollector], None]
    ) -> Dict[str, Any]:
        """Partitions the records by username and runs them concurrently up to the configured worker limit."""
        partitions = self.partition_records(records, global_username)
        total_users = len(partitions)
        log.info(f"[ORCHESTRATOR] Partitions identified: {total_users} users in parallel batch.")
        
        t0 = time.time()
        with ThreadPoolExecutor(max_workers=min(self.max_workers, total_users)) as executor:
            futures = {}
            for user, user_records in partitions.items():
                self.worker_manager.register_worker(user)
                
                # Wrap the task function to perform setup, teardown, and duration recording
                def run_worker(u=user, recs=user_records):
                    start = time.time()
                    try:
                        worker_task_fn(u, recs, self.metrics)
                    except Exception as e:
                        log.error(f"[ORCHESTRATOR] Worker thread for '{u}' encountered critical error: {e}")
                    finally:
                        self.worker_manager.unregister_worker(u)
                        self.metrics.record_worker_time(u, time.time() - start)
                
                fut = executor.submit(run_worker)
                futures[fut] = user
                
            for fut in as_completed(futures):
                u = futures[fut]
                try:
                    fut.result()
                    log.info(f"[ORCHESTRATOR] Worker session successfully finished: {u}")
                except Exception as e:
                    log.error(f"[ORCHESTRATOR] Worker session failed for '{u}' with error: {e}")

        summary = self.metrics.generate_summary()
        self._log_summary_report(summary)
        return summary

    def _log_summary_report(self, summary: Dict[str, Any]):
        """Logs a beautiful summary report of the batch run."""
        log.info("\n" + "="*80 + "\n" +
                 "                BULK PORTAL OPTIMIZED PERFORMANCE REPORT\n" +
                 "="*80 + f"\n" +
                 f"● Total Duration            : {summary['total_duration_sec']}s\n" +
                 f"● Records Audited           : {summary['records_processed']}\n" +
                 f"● APoBs Finalized           : {summary['apobs_processed']}\n" +
                 f"● Failures Encountered      : {summary['failures_count']}\n" +
                 f"● Average Processing / Rec  : {summary['average_record_processing_time_sec']}s\n" +
                 f"● Total API Network Time    : {summary['total_api_response_time_sec']}s\n" +
                 f"● Total Document Upload Time: {summary['total_upload_time_sec']}s\n" +
                 f"● Core API Cache Hit Ratio  : {summary['cache_hit_ratio_pct']}%\n" +
                 f"● Worker Session Durations  : {summary['worker_utilization']}\n" +
                 "="*80)
