import os
import json
from pathlib import Path
from typing import Dict, List, Any, Optional

class CheckpointManager:
    """
    Manages checkpoints, run states, and batch states under the _internal folder
    for resilient crash recovery and batch continuation.
    """
    def __init__(self, internal_dir: Path):
        self.internal_dir = internal_dir
        self.internal_dir.mkdir(parents=True, exist_ok=True)
        
        self.checkpoint_path = self.internal_dir / "checkpoint.json"
        self.run_state_path = self.internal_dir / "run_state.json"
        self.batch_state_path = self.internal_dir / "batch_state.json"

    def initialize_batch(self, batch_id: str, records: List[Dict[str, Any]], operator: str = "Operator") -> None:
        """Initializes a new batch session state and clears old completed caches."""
        # Create fresh batch_state
        batch_state = {
            "batch_id": batch_id,
            "batch_size": len(records),
            "timestamp": os.getenv("GST_RUN_TIME", ""),
            "operator": operator,
            "records": records
        }
        self.batch_state_path.write_text(json.dumps(batch_state, indent=2, ensure_ascii=False), encoding="utf-8")

        # Initialize progress tracker checkpoint
        checkpoint = {
            "batch_id": batch_id,
            "completed_gstins": [],
            "failed_records": {},
            "pending_gstins": [r.get("GSTIN") for r in records if r.get("GSTIN")]
        }
        self.checkpoint_path.write_text(json.dumps(checkpoint, indent=2, ensure_ascii=False), encoding="utf-8")

        # Initialize run stats
        run_state = {
            "batch_id": batch_id,
            "total": len(records),
            "completed": 0,
            "failed": 0,
            "skipped": 0,
            "stage": "Initialized"
        }
        self.run_state_path.write_text(json.dumps(run_state, indent=2, ensure_ascii=False), encoding="utf-8")

    def load_checkpoint(self) -> Dict[str, Any]:
        """Loads the current checkpoint dictionary from disk."""
        if self.checkpoint_path.exists():
            try:
                return json.loads(self.checkpoint_path.read_text(encoding="utf-8"))
            except Exception:
                pass
        return {"batch_id": "", "completed_gstins": [], "failed_records": {}, "pending_gstins": []}

    def load_batch_state(self) -> Dict[str, Any]:
        """Loads the saved batch records state."""
        if self.batch_state_path.exists():
            try:
                return json.loads(self.batch_state_path.read_text(encoding="utf-8"))
            except Exception:
                pass
        return {"batch_id": "", "batch_size": 0, "records": []}

    def load_run_state(self) -> Dict[str, Any]:
        """Loads the runtime telemetry parameters."""
        if self.run_state_path.exists():
            try:
                return json.loads(self.run_state_path.read_text(encoding="utf-8"))
            except Exception:
                pass
        return {"total": 0, "completed": 0, "failed": 0, "skipped": 0, "stage": "Unknown"}

    def update_record_success(self, gstin: str) -> None:
        """Marks a specific GSTIN registration as successfully processed."""
        cp = self.load_checkpoint()
        gstin_norm = str(gstin).strip().upper()
        
        if gstin_norm not in cp.get("completed_gstins", []):
            cp.setdefault("completed_gstins", []).append(gstin_norm)
        
        if gstin_norm in cp.get("pending_gstins", []):
            cp["pending_gstins"].remove(gstin_norm)
            
        if gstin_norm in cp.get("failed_records", {}):
            cp.get("failed_records", {}).pop(gstin_norm, None)
            
        self.checkpoint_path.write_text(json.dumps(cp, indent=2, ensure_ascii=False), encoding="utf-8")
        self._sync_run_stats()

    def update_record_failure(self, gstin: str, error_msg: str, screenshot_path: Optional[str] = None) -> None:
        """Flags a specific GSTIN registration as failed with error details."""
        cp = self.load_checkpoint()
        gstin_norm = str(gstin).strip().upper()
        
        cp.setdefault("failed_records", {})[gstin_norm] = {
            "error": error_msg,
            "timestamp": os.getenv("GST_RUN_TIME", ""),
            "screenshot": screenshot_path or ""
        }
        
        if gstin_norm in cp.get("pending_gstins", []):
            cp["pending_gstins"].remove(gstin_norm)
            
        self.checkpoint_path.write_text(json.dumps(cp, indent=2, ensure_ascii=False), encoding="utf-8")
        self._sync_run_stats()

    def update_stage(self, stage: str) -> None:
        """Updates the active run state stage name."""
        rs = self.load_run_state()
        rs["stage"] = stage
        self.run_state_path.write_text(json.dumps(rs, indent=2, ensure_ascii=False), encoding="utf-8")

    def _sync_run_stats(self) -> None:
        """Syncs run_state count registers with the physical checkpoints."""
        cp = self.load_checkpoint()
        rs = self.load_run_state()
        
        rs["completed"] = len(cp.get("completed_gstins", []))
        rs["failed"] = len(cp.get("failed_records", {}))
        
        # Recalculate skipped or total if required
        self.run_state_path.write_text(json.dumps(rs, indent=2, ensure_ascii=False), encoding="utf-8")

    def get_remaining_records(self) -> List[Dict[str, Any]]:
        """Filters the initial batch state to yield only uncompleted (pending/failed) rows."""
        bs = self.load_batch_state()
        cp = self.load_checkpoint()
        
        completed = set(cp.get("completed_gstins", []))
        remaining = []
        
        for record in bs.get("records", []):
            gstin = str(record.get("GSTIN") or "").strip().upper()
            if gstin not in completed:
                remaining.append(record)
                
        return remaining

    def is_batch_complete(self) -> bool:
        """Checks if all records in the active batch have completed successfully."""
        cp = self.load_checkpoint()
        bs = self.load_batch_state()
        
        total_records = len(bs.get("records", []))
        if total_records == 0:
            return True
            
        return len(cp.get("completed_gstins", [])) >= total_records

    def clear(self) -> None:
        """Clears all session recovery check files."""
        for p in [self.checkpoint_path, self.run_state_path, self.batch_state_path]:
            if p.exists():
                try:
                    p.unlink()
                except Exception:
                    pass
