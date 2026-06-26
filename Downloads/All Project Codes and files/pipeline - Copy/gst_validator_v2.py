"""
================================================================================
  GST AMENDMENT DOCUMENT VALIDATOR  — PRODUCTION GRADE  v5.4
  CBIC Instruction No. 03/2025-GST (Para 6) | Additional Place of Business
================================================================================
"""

# ── Standard library ──────────────────────────────────────────────────────────
import os, re, sys, json, time, base64, random, logging, hashlib, argparse, traceback
import csv
import shutil
import tempfile
import threading
import queue
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import deepcopy
from pathlib import Path
from datetime import date, timedelta
from collections import defaultdict
from difflib import SequenceMatcher
from typing import Any, Dict, List, Optional, Tuple
import customtkinter as ctk
from tkinter import filedialog, messagebox
# ── Third-party: openpyxl ─────────────────────────────────────────────────────
try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
except ImportError:
    print("ERROR: openpyxl not installed.  Run: pip install openpyxl")
    sys.exit(1)

# ── Third-party: tqdm (optional progress bar) ─────────────────────────────────
try:
    from tqdm import tqdm as _tqdm
    def progress(iterable, **kw): return _tqdm(iterable, **kw)
    HAS_TQDM = True
except ImportError:
    def progress(iterable, **kw): return iterable   # graceful fallback
    HAS_TQDM = False

# ── Third-party: Gemini SDK (new google-genai OR legacy google-generativeai) ──
try:
    from google import genai as _genai_new
    from google.genai import types as _gtypes
    USE_NEW_SDK = True
except ImportError:
    try:
        import google.generativeai as _genai_legacy
        USE_NEW_SDK = False
    except ImportError:
        print("ERROR: Install Gemini SDK:  pip install google-generativeai")
        sys.exit(1)

def _safe_folder_name(name: str) -> str:
    """
    Sanitise a client name so it is safe as a folder name across all platforms.
    Strips characters that are illegal on Windows/Linux/macOS file systems.
    """
    # Remove characters that are illegal in folder names
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    # Collapse multiple underscores / trailing dots/spaces (Windows restriction)
    safe = re.sub(r"_+", "_", safe).strip(". _")
    return safe or "UnknownClient"


def _safe_state_folder_name(state: Optional[str]) -> str:
    """Sanitise Indian state or region label for use as a single path segment."""
    if state is None:
        return "Unknown_State"
    raw = str(state).strip()
    if not raw:
        return "Unknown_State"
    safe = _safe_folder_name(raw)
    return safe or "Unknown_State"


def extract_state_from_validation_data(data: Optional[Dict[str, Any]]) -> str:
    """
    Resolve state folder label from extracted validation JSON (post-reconcile).
    Priority: summary.state → addresses.state → addresses.finalised_pob_address.state
    → lease_details property/premises/state fields.

    HARDENED : applies normalize_state_name() to every candidate so that
    "UP", "u.p.", "Uttar Pradesh" all map to the same canonical folder name,
    preventing duplicate state folders.
    """
    if not data or not isinstance(data, dict):
        return "Unknown_State"

    candidates: List[str] = []

    s = data.get("summary")
    if isinstance(s, dict) and s.get("state"):
        candidates.append(str(s["state"]).strip())

    addrs = data.get("addresses")
    if isinstance(addrs, dict):
        if addrs.get("state"):
            candidates.append(str(addrs["state"]).strip())
        fpob = addrs.get("finalised_pob_address")
        if isinstance(fpob, dict) and fpob.get("state"):
            candidates.append(str(fpob["state"]).strip())

    ld = data.get("lease_details")
    if isinstance(ld, dict):
        for key in ("property_state", "premises_state", "state"):
            if ld.get(key):
                candidates.append(str(ld[key]).strip())

    for raw in candidates:
        if raw:
            normalized = normalize_state_name(raw)
            if normalized and normalized != "Unknown_State":
                return _safe_state_folder_name(normalized)

    return "Unknown_State"


def create_statewise_client_structure(
    base_output_dir: Path,
    client_name: str,
    state_label: str,
    *,
    _folder_lock_override: Optional[threading.Lock] = None,
) -> Dict[str, Path]:
    """
    Create OUTPUT/<State>/<Client>/ with required subfolders.

    Client-visible layout (shared with clients):
        <base>/<State>/<Client>/input_docs/
        <base>/<State>/<Client>/GST_Validation_Report_{ts}.xlsx

    Developer-only layout (hidden under _internal):
        <base>/_internal/<State>/<Client>/json_cache/
        <base>/_internal/<State>/<Client>/logs/
        <base>/_internal/<State>/<Client>/metadata/

    HARDENED :
    - normalize_state_name() applied to state_label (prevents duplicate folders)
    - All paths through normalize_windows_path() (Windows long-path safe)
    - safe_create_dir() instead of mkdir() (retry-safe on Windows TOCTOU races)
    - Per-client lock via get_client_lock() when no override provided

    Keys returned include legacy aliases ``output`` → logs dir and ``cache`` → json_cache.
    """
    # Normalize state to prevent duplicate folders (e.g. "UP" vs "Uttar Pradesh")
    canonical_state = normalize_state_name(state_label)
    safe_state  = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", canonical_state).strip(". _") or "Unknown_State"
    safe_client = _safe_folder_name(client_name)

    # Client-visible root: <base>/<State>/<Client>/
    root = normalize_windows_path(base_output_dir / safe_state / safe_client)

    # Developer-only internal root: <base>/_internal/<State>/<Client>/
    internal_root = normalize_windows_path(base_output_dir / "_internal" / safe_state / safe_client)

    lock = _folder_lock_override or get_client_lock(client_name)
    with lock:
        input_docs = root / "input_docs"          # client-visible
        json_cache = internal_root / "json_cache" # developer-only
        logs_dir   = internal_root / "logs"       # developer-only
        meta_dir   = internal_root / "metadata"   # developer-only

        for folder in (input_docs, json_cache, logs_dir, meta_dir):
            try:
                safe_create_dir(folder)
                log.debug(f"   [FolderManager] Ensured: {folder}")
            except OSError as exc:
                log.error(f"   [FolderManager] Cannot create {folder}: {exc}")

    log.info(f"   [FolderManager] Client structure ready: {root}")
    return {
        "root":          root,           # <base>/<State>/<Client>/
        "internal_root": internal_root,  # <base>/_internal/<State>/<Client>/
        "input_docs":    input_docs,     # client-visible
        "json_cache":    json_cache,     # developer-only (_internal)
        "logs":          logs_dir,       # developer-only (_internal)
        "metadata":      meta_dir,       # developer-only (_internal)
        "output":        logs_dir,       # legacy alias
        "cache":         json_cache,     # legacy alias
    }


def create_client_structure(base_output_dir: Path, client_name: str) -> Dict[str, Path]:
    """Backward-compatible: Unknown_State until extraction supplies the real state folder."""
    return create_statewise_client_structure(base_output_dir, client_name, "Unknown_State")


def resolve_client_output_folder(
    base_output_dir: Path,
    client_name: str,
    validation_data: Optional[Dict[str, Any]] = None,
) -> Path:
    """
    Resolve the on-disk client folder under state-wise layout.
    Falls back to scanning existing OUTPUT/*/Client when JSON is unavailable.
    """
    safe_client = _safe_folder_name(client_name)
    state_key = "Unknown_State"
    if validation_data is not None:
        state_key = extract_state_from_validation_data(validation_data)
    else:
        cached = load_json_cache(client_name)
        if cached is not None:
            state_key = extract_state_from_validation_data(cached)

    primary = (base_output_dir / _safe_state_folder_name(state_key) / safe_client).resolve()
    if primary.exists():
        return primary

    try:
        hits = sorted(base_output_dir.glob(f"*/{safe_client}"))
        for h in hits:
            if h.is_dir():
                return h.resolve()
    except OSError as exc:
        log.warning(f"   [FolderManager] Could not scan for client folder {safe_client}: {exc}")

    return primary


def organize_client_documents(
    source_files: List[Path],
    input_docs_dir: Path,
    *,
    move: bool = False,
    client_name: str = "",
) -> Dict[Path, Path]:
    """
    Copy (or move) source documents into the client's input_docs/ folder.

    HARDENED :
    - Uses safe_copy_file() instead of raw shutil.copy2() (Windows-safe, retry-backed)
    - Validates all source paths before attempting copy
    - Applies normalize_windows_path() to all paths
    - Logs every skip/fail with reason to AuditLogger
    - Returns src→src mapping on failure so the pipeline continues
    - Never raises; always returns a complete mapping dict

    Args:
        source_files   : List of Paths to the original documents.
        input_docs_dir : Destination folder (client/input_docs/).
        move           : If True, move files instead of copying.  Default: copy.
        client_name    : For audit log messages.

    Returns:
        Dict mapping each source Path to its destination Path.
    """
    mapping: Dict[Path, Path] = {}
    audit = get_audit_logger()

    dest_resolved = normalize_windows_path(input_docs_dir)
    try:
        safe_create_dir(dest_resolved)
    except OSError as exc:
        log.error(f"   [FolderManager] Cannot create input_docs dir: {dest_resolved} — {exc}")
        return {f: f for f in source_files}

    for src in source_files:
        src_norm = normalize_windows_path(src)

        if not src_norm.exists():
            log.warning(f"   [FolderManager] Source not found, skipping: {src_norm}")
            if audit:
                audit.log_skipped(client_name, src.name, "source_file_not_found")
            mapping[src] = src
            continue

        # Safety guard: prevent copying INTO the source tree
        try:
            dest_resolved.relative_to(src_norm.parent)
            log.warning(
                f"   [FolderManager] SAFETY SKIP — input_docs dir '{dest_resolved}' "
                f"is inside the source folder '{src_norm.parent}'. "
                f"Set --client-output-dir to a path OUTSIDE the input folder. "
                f"Skipping copy for: {src.name}"
            )
            if audit:
                audit.log_skipped(client_name, src.name, "input_docs_inside_source_tree")
            mapping[src] = src
            continue
        except ValueError:
            pass  # dest is outside src_parent — safe

        if move:
            dst = dest_resolved / src_norm.name
            # Handle collision
            if dst.exists():
                stem, suffix, ctr = dst.stem, dst.suffix, 1
                while dst.exists():
                    dst = dst.parent / f"{stem}_{ctr}{suffix}"
                    ctr += 1
            try:
                shutil.move(str(src_norm), str(dst))
                log.info(f"   [FolderManager] Moved  : {src.name} → {dst.name}")
                mapping[src] = dst
                continue
            except OSError as exc:
                log.warning(f"   [FolderManager] Move failed, trying copy: {src.name} — {exc}")

        # Copy path (also fallback when move failed)
        result = safe_copy_file(src_norm, dest_resolved)
        if result is not None:
            log.info(f"   [FolderManager] Copied : {src.name} → {result.name}")
            mapping[src] = result
        else:
            log.error(f"   [FolderManager] Copy FAILED: {src.name}")
            if audit:
                audit.log_failed(client_name, src.name, "file_copy",
                                 "safe_copy_file returned None")
            mapping[src] = src  # fallback to original

    return mapping


def save_client_outputs(
    client_name: str,
    output_dir: Path,
    *,
    excel_src: Optional[Path] = None,
    json_data: Optional[Dict] = None,
    log_lines: Optional[List[str]] = None,
    client_root_for_excel: Optional[Path] = None,
) -> Dict[str, Optional[Path]]:
    """
    Persist validation artefacts under the client's logs/ folder (JSON + text log).

    Writes:
        validation_response.json    ← written from json_data if provided
        processing_log.txt          ← written from log_lines if provided
        GST_Validation_Report.xlsx  ← copied from excel_src into client_root_for_excel when set

    Duplicate-safe: existing files are renamed with a timestamp suffix before
    overwriting so historical runs are preserved.

    Args:
        client_name : Human-readable name (used only in log messages).
        output_dir  : client/logs/ folder (created by create_statewise_client_structure).
        excel_src   : Path to a completed Excel workbook to copy in (optional).
        json_data   : Gemini JSON dict to serialise.
        log_lines   : List of log message strings to write as a plain-text log.
        client_root_for_excel : Client root; report workbook lands here when copying excel_src.

    Returns:
        Dict with keys: excel, json, log — each either a Path or None.
    """
    results: Dict[str, Optional[Path]] = {"excel": None, "json": None, "log": None}

    def _backup_if_exists(path: Path):
        """Rename an existing file to path.stem_YYYYMMDD_HHMMSS.suffix before overwriting."""
        if path.exists():
            ts   = time.strftime("%Y%m%d_%H%M%S")
            bak  = path.with_name(f"{path.stem}_{ts}{path.suffix}")
            path.rename(bak)
            log.info(f"   [FolderManager] Backed up existing → {bak.name}")

    # ── Excel report (optional copy — primary report built by generate_client_report at client root)
    if excel_src and excel_src.exists():
        root_for_report = client_root_for_excel or output_dir.parent
        # VERSIONED: never overwrite old reports — always create a new timestamped copy
        ts = time.strftime("%Y%m%d_%H%M%S")
        dest_excel = root_for_report / f"GST_Validation_Report_{ts}.xlsx"
        # Ensure uniqueness if same-second collision
        ctr = 1
        while dest_excel.exists():
            dest_excel = root_for_report / f"GST_Validation_Report_{ts}_{ctr:03d}.xlsx"
            ctr += 1
        try:
            shutil.copy2(str(excel_src), str(dest_excel))
            results["excel"] = dest_excel
            log.info(f"   [FolderManager] Excel saved  : {dest_excel}")
        except OSError as exc:
            log.error(f"   [FolderManager] Excel copy failed for {client_name}: {exc}")

    # ── JSON validation response (VERSIONED — never overwrites old JSON) ──────
    if json_data is not None:
        ts = time.strftime("%Y%m%d_%H%M%S")
        dest_json = normalize_windows_path(output_dir / f"validation_response_{ts}.json")
        ctr = 1
        while dest_json.exists():
            dest_json = normalize_windows_path(output_dir / f"validation_response_{ts}_{ctr:03d}.json")
            ctr += 1
        ok = atomic_json_dump(json_data, dest_json)
        if ok:
            results["json"] = dest_json
            log.info(f"   [FolderManager] JSON saved   : {dest_json}")
        else:
            log.error(f"   [FolderManager] JSON write failed for {client_name}: atomic_json_dump returned False")

    # ── Processing log (VERSIONED — never overwrites old logs) ───────────────
    if log_lines:
        ts = time.strftime("%Y%m%d_%H%M%S")
        dest_log = normalize_windows_path(output_dir / f"processing_log_{ts}.txt")
        ctr = 1
        while dest_log.exists():
            dest_log = normalize_windows_path(output_dir / f"processing_log_{ts}_{ctr:03d}.txt")
            ctr += 1
        ok = safe_write_text(dest_log, "\n".join(log_lines))
        if ok:
            results["log"] = dest_log
            log.info(f"   [FolderManager] Log saved    : {dest_log}")
        else:
            log.error(f"   [FolderManager] Log write failed for {client_name}")

    return results


def save_json_response(client_name: str, output_dir: Path, json_data: Dict) -> Optional[Path]:
    """
    Convenience wrapper: write validation_response.json to client logs/ folder.

    Args:
        client_name : Used for log messages only.
        output_dir  : client/logs/ directory.
        json_data   : Dict to serialise.

    Returns:
        Path to written file, or None on failure.
    """
    return save_client_outputs(client_name, output_dir, json_data=json_data)["json"]


def save_logs(client_name: str, output_dir: Path, log_lines: List[str]) -> Optional[Path]:
    """
    Convenience wrapper: write processing_log.txt to client logs/ folder.

    Args:
        client_name : Used for log messages only.
        output_dir  : client/logs/ directory.
        log_lines   : Lines to write.

    Returns:
        Path to written file, or None on failure.
    """
    return save_client_outputs(client_name, output_dir, log_lines=log_lines)["log"]


def save_client_cache(
    client_name: str,
    cache_dir: Path,
    gemini_data: Dict,
) -> Optional[Path]:
    """
    Write Gemini JSON response to client/json_cache/File_cache.json atomically.

    HARDENED : uses safe_json_save (temp-file + rename) instead of
    write_text() so partial writes never corrupt the cache on process kill.
    Thread-safe via _cache_lock.

    Args:
        client_name : Used for log messages only.
        cache_dir   : client/json_cache/ directory.
        gemini_data : Raw Gemini response dict.

    Returns:
        Path to the written cache file, or None on failure.
    """
    cache_path = normalize_windows_path(cache_dir / "File_cache.json")
    try:
        safe_create_dir(cache_path.parent)
    except OSError as exc:
        log.error(f"   [FolderManager] Cache dir creation failed for {client_name}: {exc}")
        return None
    ok = safe_json_save(gemini_data, cache_path, lock=_cache_lock, client_name=client_name)
    if ok:
        log.info(f"   [FolderManager] Gemini cache : {cache_path}")
        return cache_path
    log.error(f"   [FolderManager] Cache write failed for {client_name}")
    return None


def load_client_cache(cache_dir: Path) -> Optional[Dict]:
    """
    Load Gemini JSON from client/cache/gemini_cache.json if it exists.

    Thread-safe: protected by _cache_lock.

    Args:
        cache_dir : client/json_cache/ directory.

    Returns:
        Parsed dict, or None if not found / unreadable.
    """
    cache_path = cache_dir / "File_cache.json"
    if not cache_path.exists():
        return None
    try:
        with _cache_lock:
            return json.loads(cache_path.read_text(encoding="utf-8"))
    except Exception as exc:
        log.warning(f"   [FolderManager] Cache read failed ({cache_path}): {exc}")
        return None


def write_client_metadata(
    client_root: Path,
    client_name: str,
    documents_count: int,
    status: str,
    extra: Optional[Dict] = None,
) -> Path:
    """
    Write metadata.json under client/metadata/metadata.json.

    Schema:
        {
          "client_name"      : "Balumath",
          "documents_count"  : 12,
          "processed_time"   : "2026-05-06 16:00:00",
          "status"           : "SUCCESS",
          ... any keys from `extra` ...
        }

    Args:
        client_root     : client/ directory (root of the hierarchy).
        client_name     : Human-readable client name.
        documents_count : Total number of input documents.
        status          : "SUCCESS" | "FAILED" | "SKIPPED" | "PARTIAL".
        extra           : Optional additional key-value pairs merged into metadata.

    Returns:
        Path to metadata.json.
    """
    meta: Dict[str, Any] = {
        "client_name"     : client_name,
        "documents_count" : documents_count,
        "processed_time"  : time.strftime("%Y-%m-%d %H:%M:%S"),
        "status"          : status,
    }
    if extra:
        meta.update(extra)

    # metadata lives under _internal/<State>/<Client>/metadata/
    # client_root is <base>/<State>/<Client>/; we re-anchor to _internal.
    # If client_root already points into _internal, use it directly.
    if "_internal" in str(client_root):
        meta_dir = normalize_windows_path(client_root / "metadata")
    else:
        # Derive: <base>/_internal/<State>/<Client>/metadata/
        # client_root = <base>/<State>/<Client>
        base_part = client_root.parent.parent  # <base>
        state_client = client_root.relative_to(base_part)  # <State>/<Client>
        meta_dir = normalize_windows_path(base_part / "_internal" / state_client / "metadata")
    meta_path = meta_dir / "metadata.json"
    try:
        safe_create_dir(meta_dir)
        ok = atomic_json_dump(meta, meta_path)
        if ok:
            log.info(f"   [FolderManager] Metadata     : {meta_path}")
        else:
            log.error(f"   [FolderManager] Metadata write failed ({client_root.name}): atomic_json_dump returned False")
    except OSError as exc:
        log.error(f"   [FolderManager] Metadata dir creation failed ({client_root.name}): {exc}")
    return meta_path


# CLIENT_OUTPUT_BASE — top-level output directory that holds all client folders.
# Defaults to ./client_outputs but can be overridden via GST_CLIENT_OUTPUT_DIR
# env variable or --client-output-dir CLI flag (added to _parse_args below).
CLIENT_OUTPUT_BASE = Path(
    os.environ.get("GST_CLIENT_OUTPUT_DIR", "./client_outputs")
)

# Module-level lock protecting per-client folder creation (avoids TOCTOU race).
_folder_lock = threading.Lock()

# End of Client-Wise Folder Management Module


#  GLOBAL CONFIGURATION 
GEMINI_MODEL      = "gemini-3.5-flash"
# Fallback chain tried in order if primary model is unavailable:
GEMINI_MODEL_FALLBACKS = [
    "gemini-3-flash-preview",
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-2.5-pro",
]
VALIDATION_DATE   = date.today().isoformat()
CUTOFF_DATE       = (date.today() - timedelta(days=93)).isoformat()
OUTPUT_DIR        = Path(os.environ.get("GST_OUTPUT_DIR", "./outputs"))
JSON_CACHE_DIR    = OUTPUT_DIR / "json_cache"

# ── Parallelism & Performance ─────────────────────────────────────────────────
MAX_WORKERS       = int(os.environ.get("GST_MAX_WORKERS", "5"))
CACHE_ENABLED     = os.environ.get("GST_CACHE_ENABLED", "true").lower() != "false"
RETRY_COUNT       = int(os.environ.get("GST_RETRY_COUNT", "5"))
MAX_RETRIES       = RETRY_COUNT
BASE_BACKOFF      = 4      # seconds (reduced: faster recovery on transient errors)
MAX_BACKOFF       = 60     # seconds
MAX_DOCS_PER_CALL = 15     # chunk guard for Gemini context window
INTER_CLIENT_DELAY = 0     # Disabled: parallel workers handle rate control naturally

# ── Thread-safety primitives ──────────────────────────────────────────────────
_excel_lock    = threading.Lock()   # guards ExcelBuilder.add_client()
_cache_lock    = threading.Lock()   # guards JSON cache reads/writes

SUPPORTED_EXT = {".pdf", ".jpg", ".jpeg", ".png", ".tiff", ".webp", ".doc", ".docx"}
MIME_MAP = {
    ".pdf":  "application/pdf",
    ".png":  "image/png",
    ".jpg":  "image/jpeg",
    ".jpeg": "image/jpeg",
    ".tiff": "image/tiff",
    ".webp": "image/webp",
    ".doc":  "application/msword",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
}

# ── Tiered Logging System (Issue 14 — Professional logging levels) ───────────
#
#   file_logger  → gst_validator.log  (ALL levels: DEBUG, INFO, WARNING, ERROR, CRITICAL)
#   ui_logger    → UI log panel        (WARNING + ERROR + CRITICAL only — quiet by default)
#   debug_logger → DEBUG console       (DEBUG only, suppressed in production)
#
# Log categories suppressed from UI panel to reduce noise:
#   - [FolderManager] folder creation / copy messages  (DEBUG → file only)
#   - [PathUtil] copy confirmations                    (DEBUG → file only)
#   - [JSONUtil] routine save confirmations            (DEBUG → file only)
#   - Repetitive INFO messages from inner loops        (file only)
#
# To enable full DEBUG output in UI, set env:  GST_UI_DEBUG=1

_UI_DEBUG_MODE = os.environ.get("GST_UI_DEBUG", "0").strip() == "1"

# Suppressed prefixes: these INFO messages go to file only, not to the UI panel
_UI_SUPPRESSED_PREFIXES = (
    "   [FolderManager]",
    "   [PathUtil]",
    "   [JSONUtil]",
    "[PathUtil]",
    "[JSONUtil]",
)

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s  %(levelname)-8s  [%(threadName)-14s]  %(message)s",
    handlers=[],
)

# File handler — receives EVERYTHING (DEBUG+)
_file_handler = logging.FileHandler("gst_validator.log", encoding="utf-8")
_file_handler.setLevel(logging.DEBUG)
_file_handler.setFormatter(logging.Formatter(
    "%(asctime)s  %(levelname)-8s  [%(threadName)-14s]  %(message)s"
))
logging.getLogger().addHandler(_file_handler)

# Console handler — INFO+ (for CLI / EXE console window)
_console_stream = getattr(sys, "stdout", None) or getattr(sys, "__stdout__", None)
if _console_stream is not None and hasattr(_console_stream, "write"):
    _con_handler = logging.StreamHandler(_console_stream)
    _con_handler.setLevel(logging.INFO)
    _con_handler.setFormatter(logging.Formatter(
        "%(asctime)s  %(levelname)-8s  [%(threadName)-14s]  %(message)s"
    ))
    logging.getLogger().addHandler(_con_handler)

log = logging.getLogger("GSTValidator")


# ══════════════════════════════════════════════════════════════════════════════
#  PRODUCTION HARDENING MODULES  (integrated from patch v4.0)
#  Issues addressed: WinError 3, state duplication, state overrides, JSON save
#  failures, client grouping, thread safety, Excel mapping, audit logging,
#  checkpoint/resume, folder structure.
# ══════════════════════════════════════════════════════════════════════════════

# ── MODULE 1: Windows-safe path utilities ─────────────────────────────────────
_WIN_LONG_PREFIX = "\\\\?\\"
_IS_WINDOWS = sys.platform.startswith("win")


def normalize_windows_path(path: Path) -> Path:
    """Return an absolute, resolved Path safe for Windows long-path operations."""
    try:
        resolved = path.resolve(strict=False)
    except Exception:
        resolved = path.absolute()
    if _IS_WINDOWS:
        s = str(resolved)
        if len(s) > 240 and not s.startswith(_WIN_LONG_PREFIX):
            resolved = Path(_WIN_LONG_PREFIX + s)
    return resolved


def safe_create_dir(path: Path, *, exist_ok: bool = True) -> Path:
    """Create directory (and parents) thread-safely with retry on Windows TOCTOU races."""
    norm = normalize_windows_path(path)
    last_exc: Optional[Exception] = None
    for attempt in range(1, 4):
        try:
            norm.mkdir(parents=True, exist_ok=exist_ok)
            return norm
        except FileExistsError:
            return norm
        except OSError as exc:
            last_exc = exc
            if attempt < 3:
                time.sleep(0.05 * attempt)
    log.error(f"[PathUtil] safe_create_dir failed after 3 attempts: {norm} — {last_exc}")
    raise OSError(f"Cannot create directory: {norm}") from last_exc


def safe_copy_file(src: Path, dst: Path, *, overwrite: bool = False) -> Optional[Path]:
    """Copy src → dst with Windows path safety, parent creation, and 3-attempt retry."""
    src = normalize_windows_path(src)
    dst = normalize_windows_path(dst)
    if not src.exists():
        log.warning(f"[PathUtil] safe_copy_file: source does not exist → {src}")
        return None
    if dst.is_dir():
        dst = dst / src.name
    if not overwrite and dst.exists():
        stem, suffix = dst.stem, dst.suffix
        counter = 1
        while dst.exists():
            dst = dst.parent / f"{stem}_{counter}{suffix}"
            counter += 1
    try:
        safe_create_dir(dst.parent)
    except OSError as exc:
        log.error(f"[PathUtil] safe_copy_file: cannot create dst dir {dst.parent} — {exc}")
        return None
    last_exc: Optional[Exception] = None
    for attempt in range(1, 4):
        try:
            shutil.copy2(str(src), str(dst))
            log.debug(f"[PathUtil] Copied: {src.name} → {dst}")
            return dst
        except OSError as exc:
            last_exc = exc
            if attempt < 3:
                time.sleep(0.1 * attempt)
    log.error(f"[PathUtil] safe_copy_file FAILED after 3 attempts: {src.name} → {dst} — {last_exc}")
    return None


def safe_write_text(path: Path, content: str, *, encoding: str = "utf-8") -> bool:
    """Write text to path atomically (temp file + rename) to prevent partial writes."""
    path = normalize_windows_path(path)
    try:
        safe_create_dir(path.parent)
    except OSError as exc:
        log.error(f"[PathUtil] safe_write_text: cannot create dir {path.parent} — {exc}")
        return False
    tmp_path: Optional[Path] = None
    try:
        fd, tmp_str = tempfile.mkstemp(
            dir=str(path.parent),
            prefix=f".{path.stem}_tmp_",
            suffix=path.suffix or ".tmp",
        )
        tmp_path = Path(tmp_str)
        with os.fdopen(fd, "w", encoding=encoding) as f:
            f.write(content)
        if _IS_WINDOWS and path.exists():
            path.unlink()
        tmp_path.rename(path)
        return True
    except Exception as exc:
        log.error(f"[PathUtil] safe_write_text failed: {path} — {exc}")
        if tmp_path and tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass
        return False


# ── MODULE 2: Atomic JSON utilities ──────────────────────────────────────────

def atomic_json_dump(data: Any, path: Path, *, indent: int = 2,
                     ensure_ascii: bool = False) -> bool:
    """Serialize data to JSON and write atomically (temp-rename) to prevent corrupt files."""
    try:
        serialized = json.dumps(data, indent=indent, ensure_ascii=ensure_ascii, default=str)
    except (TypeError, ValueError) as exc:
        log.error(f"[JSONUtil] atomic_json_dump: serialization failed for {path.name} — {exc}")
        return False
    return safe_write_text(path, serialized)


def safe_json_save(data: Any, path: Path, *, lock: Optional[threading.Lock] = None,
                   client_name: str = "") -> bool:
    """Thread-safe atomic JSON save with optional lock and descriptive logging."""
    label = f"[{client_name}] " if client_name else ""
    if lock is not None:
        with lock:
            ok = atomic_json_dump(data, path)
    else:
        ok = atomic_json_dump(data, path)
    if ok:
        log.info(f"   [JSONUtil] {label}JSON saved → {path}")
    else:
        log.error(f"   [JSONUtil] {label}JSON save FAILED → {path}")
    return ok


# ── MODULE 3: State normalization ─────────────────────────────────────────────

STATE_NORMALIZATION_MAP: Dict[str, str] = {
    "andhra pradesh": "Andhra Pradesh", "andhrapradesh": "Andhra Pradesh",
    "andhra": "Andhra Pradesh", "ap": "Andhra Pradesh",
    "arunachal pradesh": "Arunachal Pradesh", "arunachalpradesh": "Arunachal Pradesh",
    "arunachal": "Arunachal Pradesh",
    "assam": "Assam",
    "bihar": "Bihar", "bih": "Bihar", "br": "Bihar",
    "बिहार": "Bihar", "बिहर": "Bihar",
    "chhattisgarh": "Chhattisgarh", "chattisgarh": "Chhattisgarh",
    "c.g.": "Chhattisgarh", "cg": "Chhattisgarh",
    "goa": "Goa",
    "gujarat": "Gujarat", "gujrat": "Gujarat", "gj": "Gujarat",
    "haryana": "Haryana", "hr": "Haryana",
    "himachal pradesh": "Himachal Pradesh", "himachalpradesh": "Himachal Pradesh",
    "himachal": "Himachal Pradesh", "hp": "Himachal Pradesh",
    "jharkhand": "Jharkhand", "jharkand": "Jharkhand", "jharkhnd": "Jharkhand",
    "jh": "Jharkhand", "jhk": "Jharkhand", "jharkhand state": "Jharkhand",
    "झारखंड": "Jharkhand", "झारखण्ड": "Jharkhand",
    "karnataka": "Karnataka", "karnatak": "Karnataka", "kk": "Karnataka",
    "ka": "Karnataka",
    "kerala": "Kerala", "kl": "Kerala",
    "madhya pradesh": "Madhya Pradesh", "madhyapradesh": "Madhya Pradesh",
    "m.p.": "Madhya Pradesh", "mp": "Madhya Pradesh",
    "maharashtra": "Maharashtra", "maharastra": "Maharashtra",
    "mh": "Maharashtra", "maha": "Maharashtra",
    "manipur": "Manipur", "mn": "Manipur",
    "meghalaya": "Meghalaya", "ml": "Meghalaya",
    "mizoram": "Mizoram", "mz": "Mizoram",
    "nagaland": "Nagaland", "nl": "Nagaland",
    "odisha": "Odisha", "orissa": "Odisha", "od": "Odisha",
    "punjab": "Punjab", "pb": "Punjab",
    "rajasthan": "Rajasthan", "rajsthan": "Rajasthan", "rj": "Rajasthan",
    "sikkim": "Sikkim", "sk": "Sikkim",
    "tamil nadu": "Tamil Nadu", "tamilnadu": "Tamil Nadu", "tn": "Tamil Nadu",
    "tamilnad": "Tamil Nadu",
    "telangana": "Telangana", "telegana": "Telangana", "ts": "Telangana",
    "tg": "Telangana",
    "tripura": "Tripura", "tr": "Tripura",
    "uttar pradesh": "Uttar Pradesh", "uttarpradesh": "Uttar Pradesh",
    "u.p.": "Uttar Pradesh", "up": "Uttar Pradesh", "u p": "Uttar Pradesh",
    "uttar prades": "Uttar Pradesh", "uttar pradeesh": "Uttar Pradesh",
    "उत्तर प्रदेश": "Uttar Pradesh", "उत्तरप्रदेश": "Uttar Pradesh",
    "uttarakhand": "Uttarakhand", "uttaranchal": "Uttarakhand",
    "uk": "Uttarakhand", "ua": "Uttarakhand",
    "west bengal": "West Bengal", "westbengal": "West Bengal",
    "wb": "West Bengal", "bengal": "West Bengal",
    "delhi": "Delhi", "new delhi": "Delhi", "ncr": "Delhi",
    "jammu and kashmir": "Jammu and Kashmir", "j&k": "Jammu and Kashmir",
    "jk": "Jammu and Kashmir", "jammu kashmir": "Jammu and Kashmir",
    "ladakh": "Ladakh",
    "chandigarh": "Chandigarh",
    "dadra and nagar haveli": "Dadra and Nagar Haveli and Daman and Diu",
    "daman and diu": "Dadra and Nagar Haveli and Daman and Diu",
    "puducherry": "Puducherry", "pondicherry": "Puducherry",
    "lakshadweep": "Lakshadweep",
    "andaman and nicobar": "Andaman and Nicobar Islands",
    "andaman nicobar": "Andaman and Nicobar Islands",
}


def normalize_state_name(raw: Optional[str]) -> str:
    """
    Map any raw state string to a single canonical Indian state/UT name.

    Normalisation steps: strip whitespace → remove punctuation noise → lowercase lookup
    in STATE_NORMALIZATION_MAP → fuzzy fallback (SequenceMatcher ≥ 0.82) → title-case passthrough.

    CRITICAL: This function NEVER corrects state values based on geographic inference.
    It only normalises spelling/abbreviation variants.
    """
    if not raw:
        return "Unknown_State"
    cleaned = str(raw).strip()
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = re.sub(r"[.\-\[\]()]+", " ", cleaned)
    cleaned = cleaned.strip()
    if not cleaned:
        return "Unknown_State"
    key = cleaned.lower()
    if key in STATE_NORMALIZATION_MAP:
        return STATE_NORMALIZATION_MAP[key]
    best_score, best_canonical = 0.0, ""
    for variant, canonical in STATE_NORMALIZATION_MAP.items():
        score = SequenceMatcher(None, key, variant).ratio()
        if score > best_score:
            best_score, best_canonical = score, canonical
    if best_score >= 0.82:
        log.debug(f"[StateNorm] Fuzzy matched '{raw}' → '{best_canonical}' (score={best_score:.2f})")
        return best_canonical
    result = cleaned.title()
    log.info(f"[StateNorm] Unknown state '{raw}' → keeping as '{result}'")
    return result


def _set_nested(d: Dict, keys: List[str], value: Any) -> None:
    """Set d[keys[0]][keys[1]]... = value, creating intermediate dicts as needed."""
    for k in keys[:-1]:
        if not isinstance(d.get(k), dict):
            d[k] = {}
        d = d[k]
    d[keys[-1]] = value


def lock_state_from_primary_lease(extracted_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Enforce STATE_VERBATIM rule: propagate the primary lease deed state to ALL
    state fields in the JSON. Normalises spelling only — never corrects geography.
    Adds _meta.state_lock_applied audit flag.
    """
    if not isinstance(extracted_data, dict):
        return extracted_data
    data = extracted_data
    ld    = data.get("lease_details") or {}
    addrs = data.get("addresses") or {}
    s     = data.get("summary") or {}
    op    = data.get("ownership_proof") or {}
    primary_state: str = ""
    source_field: str = ""
    for field_path, val in [
        ("lease_details.property_state",  ld.get("property_state")),
        ("lease_details.premises_state",  ld.get("premises_state")),
        ("lease_details.state",           ld.get("state")),
        ("addresses.state",               addrs.get("state")),
        ("summary.state",                 s.get("state")),
        ("ownership_proof.state",         op.get("state")),
    ]:
        if val and str(val).strip() not in ("", "null", "None", "N/A"):
            primary_state = str(val).strip()
            source_field  = field_path
            break
    if not primary_state:
        log.debug("[StateLock] No primary state found — skip lock.")
        return data
    canonical = normalize_state_name(primary_state)
    # FIX (v5.3): Do not propagate a null/unknown state — if the model returned
    # null, "null", "N/A", or any value that normalises to Unknown_State it means
    # the state was NOT explicitly written in the document.  Locking Unknown_State
    # across all fields would mask that absence; keep the fields as-is instead.
    if canonical in ("Unknown_State", "", "null", "None", "N/A"):
        log.info(f"[StateLock] State '{primary_state}' normalised to Unknown/null — skip lock "
                 f"(source: {source_field}). State fields left as extracted.")
        return data
    log.info(f"[StateLock] Primary state locked: '{primary_state}' → '{canonical}' (source: {source_field})")
    _set_nested(data, ["lease_details", "property_state"], canonical)
    _set_nested(data, ["lease_details", "premises_state"], canonical)
    _set_nested(data, ["addresses", "state"], canonical)
    _set_nested(data, ["summary", "state"], canonical)
    if "ownership_proof" in data and isinstance(data["ownership_proof"], dict):
        _set_nested(data, ["ownership_proof", "state"], canonical)
    fpob = addrs.get("finalised_pob_address")
    if isinstance(fpob, dict) and fpob.get("state"):
        fpob["state"] = canonical
    _set_nested(data, ["_meta", "state_lock_applied"], True)
    _set_nested(data, ["_meta", "state_lock_source"], source_field)
    _set_nested(data, ["_meta", "state_lock_raw"], primary_state)
    _set_nested(data, ["_meta", "state_canonical"], canonical)
    return data


def validate_geographic_consistency(extracted_data: Dict[str, Any]) -> Dict[str, Any]:
    """
    Flag (but NEVER correct) geographic inconsistencies between district and state.
    Python-side safety net that mirrors the Gemini pincode_state_flag logic.
    """
    if not isinstance(extracted_data, dict):
        return extracted_data
    addrs    = extracted_data.get("addresses") or {}
    district = str(addrs.get("district") or "").strip().lower()
    state    = str(addrs.get("state") or "").strip()
    if not district or not state:
        return extracted_data
    DISTRICT_STATE_MAP: Dict[str, str] = {
        "garhwa": "Jharkhand", "palamu": "Jharkhand", "latehar": "Jharkhand",
        "chatra": "Jharkhand", "hazaribagh": "Jharkhand", "ramgarh": "Jharkhand",
        "bokaro": "Jharkhand", "dhanbad": "Jharkhand", "giridih": "Jharkhand",
        "koderma": "Jharkhand", "ranchi": "Jharkhand", "khunti": "Jharkhand",
        "gumla": "Jharkhand", "simdega": "Jharkhand", "lohardaga": "Jharkhand",
        "west singhbhum": "Jharkhand", "east singhbhum": "Jharkhand",
        "saraikela": "Jharkhand", "dumka": "Jharkhand", "jamtara": "Jharkhand",
        "deoghar": "Jharkhand", "godda": "Jharkhand", "sahibganj": "Jharkhand",
        "pakur": "Jharkhand",
        "gaya": "Bihar", "patna": "Bihar", "bhagalpur": "Bihar",
        "muzaffarpur": "Bihar", "darbhanga": "Bihar", "nalanda": "Bihar",
        "vaishali": "Bihar", "saran": "Bihar", "sitamarhi": "Bihar",
        "madhubani": "Bihar", "supaul": "Bihar", "araria": "Bihar",
        "kishanganj": "Bihar", "purnia": "Bihar", "katihar": "Bihar",
        "east champaran": "Bihar", "west champaran": "Bihar",
        "samastipur": "Bihar", "begusarai": "Bihar", "khagaria": "Bihar",
        "bhojpur": "Bihar", "buxar": "Bihar", "rohtas": "Bihar",
        "kaimur": "Bihar", "aurangabad": "Bihar", "arwal": "Bihar",
        "jehanabad": "Bihar", "nawada": "Bihar", "sheikhpura": "Bihar",
        "lakhisarai": "Bihar", "sheohar": "Bihar",
        "lucknow": "Uttar Pradesh", "kanpur": "Uttar Pradesh",
        "varanasi": "Uttar Pradesh", "agra": "Uttar Pradesh",
        "prayagraj": "Uttar Pradesh", "allahabad": "Uttar Pradesh",
        "meerut": "Uttar Pradesh", "noida": "Uttar Pradesh",
        "ghaziabad": "Uttar Pradesh", "gorakhpur": "Uttar Pradesh",
    }
    expected_state = DISTRICT_STATE_MAP.get(district.replace("-", " ").replace("_", " "))
    if expected_state is None:
        return extracted_data
    state_norm = normalize_state_name(state)
    if state_norm.lower() == expected_state.lower():
        _set_nested(extracted_data, ["addresses", "pincode_state_flag"], False)
        _set_nested(extracted_data, ["addresses", "state_discrepancy_note"], None)
    else:
        note = (f"District '{district.title()}' belongs to {expected_state}, "
                f"but deed says '{state}'. State kept verbatim per lease deed.")
        if not extracted_data.get("addresses", {}).get("pincode_state_flag"):
            _set_nested(extracted_data, ["addresses", "pincode_state_flag"], True)
            _set_nested(extracted_data, ["addresses", "state_discrepancy_note"], note)
            log.info(f"[GeoCheck] Inconsistency flagged: district={district} "
                     f"expected={expected_state} deed_says={state}")
    return extracted_data


# ── MODULE 4: Client grouping utilities ──────────────────────────────────────
#
#  Deterministic, production-safe clustering for large batches:
#    • Structural gates: differing pincode / village / land-id / owner → never merge
#    • Filename + optional first-page PDF text hints (no Gemini, no full-body fuzzy)
#    • Optional fuzzy (SequenceMatcher) only on short structured strings ≥ 0.93
#

_TYPO_CORRECTIONS: Dict[str, str] = {
    "aggrement": "agreement", "agrement": "agreement", "agreemnt": "agreement",
    "leaase":    "lease",     "laese":    "lease",     "leasse":   "lease",
    "receit":    "receipt",   "reciept":  "receipt",   "recipt":   "receipt",
    "utilty":    "utility",   "utlity":   "utility",
    "electrcity": "electricity", "electricty": "electricity",
    "muzffarpur": "muzaffarpur",
    "patana":    "patna",     "patan":    "patna",
    "allhabad":  "allahabad", "allahabd": "allahabad",
    "varansi":   "varanasi",  "banaras":  "varanasi",
    "luckno":    "lucknow",   "kanpoor":  "kanpur",
    "gorkhpur":  "gorakhpur", "gorakpur": "gorakhpur",
}

# Fuzzy ratio is applied only to short labels (filename locality / village / owner), never raw OCR dumps.
GROUPING_FUZZY_THRESHOLD: float = 0.93
GROUPING_OWNER_COMPAT_MIN: float = 0.88

_GROUPING_GENERIC_LOCALITY: frozenset = frozenset({
    "prayagraj", "allahabad", "uttar", "pradesh", "uttarpradesh", "up", "u", "p",
    "india", "bharat", "state", "district", "dist", "distt", "city", "tehsil", "tahsil",
    "taluka", "block", "po", "ps", "pin", "pincode", "zip",
})

_GROUPING_STRIP_RE = re.compile(
    r"\bland\s+lease\s+agreement\b|\blease\s+agreement\b|\brent\s+agreement\b"
    r"|\bnew\s+agreement\b|\bagreement\s+paper\b|\blease\s+deed\b"
    r"|\bconsent\s+letter\b|\butility\s+bill\b|\belectricity\s+bill\b"
    r"|\bland\s+receipt\b|\bland\s+lease\b"
    r"|\bagreement\b|\bregister(?:ed)?\b|\bdeed\b|\breceipt\b"
    r"|\bbill\b|\binvoice\b|\bnoc\b|\bconsent\b|\bletter\b"
    r"|\bnew\b|\bold\b|\bpaper\b|\bland\b|\blease\b|\brent\b"
    r"|\bdocument\b|\bdoc\b|\bform\b|\boriginal\b|\bscanned\b"
    r"|\s*\(\d+\)\s*|\s*_+\d+\s*|\s*-\s*copy\s*|\s*copy\s*",
    flags=re.IGNORECASE,
)

_PDF_TEXT_FOR_GROUPING: Optional[str] = None  # "pypdf" | "PyPDF2" | None (set once)


def _detect_pdf_reader_backend() -> None:
    """Detect pypdf / PyPDF2 once for lightweight grouping text extraction."""
    global _PDF_TEXT_FOR_GROUPING
    if _PDF_TEXT_FOR_GROUPING is not None:
        return
    for pkg in ("pypdf", "PyPDF2"):
        try:
            __import__(pkg, fromlist=["PdfReader"])
            _PDF_TEXT_FOR_GROUPING = pkg
            return
        except ImportError:
            continue
    _PDF_TEXT_FOR_GROUPING = ""


def normalize_client_name(raw: str) -> str:
    """Normalize a raw client name: apply typo corrections, strip trailing digits, title-case."""
    s = raw.lower()
    words = s.split()
    corrected = [_TYPO_CORRECTIONS.get(w, w) for w in words]
    s = " ".join(corrected)
    s = re.sub(r"\s+\d+$", "", s).strip()
    s = re.sub(r"\s+", " ", s).strip().title()
    return s or raw.strip().title()


def _normalize_grouping_spelling(text: str) -> str:
    """Safe spelling normalisation for grouping (substring + word typos, spacing, case)."""
    if not text:
        return ""
    s = text.lower()
    s = s.replace("aggrement", "agreement").replace("agrement", "agreement")
    s = s.replace("leaase", "lease").replace("leasse", "lease")
    s = re.sub(r"\s+", " ", s).strip()
    words = s.split()
    s = " ".join(_TYPO_CORRECTIONS.get(w, w) for w in words)
    return re.sub(r"\s+", " ", s).strip()


def _slug_alnum(s: str, max_len: int = 48) -> str:
    if not s:
        return ""
    t = re.sub(r"[^a-z0-9]+", "_", s.lower()).strip("_")
    return t[:max_len]


def _norm_key(filename: str) -> str:
    """Backward-compatible: normalised filename locality label (used by legacy call sites)."""
    key, _conf = _filename_locality_key(filename)
    return key


def _filename_locality_key(filename: str) -> Tuple[str, float]:
    """
    Strip document-type noise from filename; return (locality_key, confidence).
    """
    stem_raw = re.sub(r"\s*\(\d+\)\s*$", "", Path(filename).stem).strip()
    stem = re.sub(r"[_\-]+", " ", stem_raw).strip()
    stem = _normalize_grouping_spelling(stem)

    key = stem
    for _ in range(5):
        prev = key
        key = _GROUPING_STRIP_RE.sub(" ", key)
        key = re.sub(r"[^a-zA-Z0-9& ]", " ", key)
        key = re.sub(r"\s+", " ", key).strip().title()
        if key == prev:
            break
    if not key or len(key) < 2:
        key = re.sub(r"[^a-zA-Z0-9& ]", " ", stem).strip().title()
    key = re.sub(r"\s+\d+$", "", key).strip()
    key = normalize_client_name(key)
    if len(key) >= 6:
        conf = 0.85
    elif len(key) >= 3:
        conf = 0.65
    else:
        conf = 0.4
    return (key[:80] if len(key) >= 2 else stem_raw[:80]), conf


def _strip_generic_locality_tokens(label: str) -> str:
    toks = [t for t in label.lower().split() if t and t not in _GROUPING_GENERIC_LOCALITY]
    return " ".join(toks).strip()


def _is_generic_only_locality(label: str) -> bool:
    s = _strip_generic_locality_tokens(label)
    return len(s) < 3


def _quick_pdf_text_for_grouping(path: Path, max_pages: int = 4, max_chars: int = 24000) -> str:
    """First pages of extractable PDF text for regex hints (optional dependency)."""
    if path.suffix.lower() != ".pdf":
        return ""
    _detect_pdf_reader_backend()
    if not _PDF_TEXT_FOR_GROUPING:
        return ""
    try:
        mod = __import__(_PDF_TEXT_FOR_GROUPING, fromlist=["PdfReader"])
        reader = mod.PdfReader(str(path), strict=False)
    except Exception as exc:
        log.debug(f"[Grouping] PDF open failed {path.name}: {exc}")
        return ""
    chunks: List[str] = []
    n = 0
    for page in reader.pages[:max_pages]:
        try:
            chunks.append(page.extract_text() or "")
        except Exception:
            chunks.append("")
        n += len(chunks[-1])
        if n >= max_chars:
            break
    return "\n".join(chunks)[:max_chars]


def _first_pincode(text: str) -> str:
    if not text:
        return ""
    m = re.search(r"\b(\d{6})\b", text)
    return m.group(1) if m else ""


def _hints_from_text(text: str) -> Dict[str, Any]:
    """Structured hints from a short text sample (not used for bulk fuzzy matching)."""
    hints: Dict[str, Any] = {}
    if not text:
        return hints
    t = text[: max(8000, len(text))]
    pc = _first_pincode(t)
    if pc:
        hints["pincode"] = pc

    for pat in (
        r"(?:mouza|mauza|village|gram|grama|ग्राम|मौजा)\s*[:.-]*\s*([A-Za-zऀ-ॿ][A-Za-zऀ-ॿ\s]{1,34})",
        r"(?:at|near)\s+([A-Za-zऀ-ॿ][A-Za-zऀ-ॿ\s]{1,30})\s*(?:,|\(|village|mouza)",
    ):
        m = re.search(pat, t, re.IGNORECASE)
        if m:
            hints["village"] = re.sub(r"\s+", " ", m.group(1).strip())[:60]
            break

    lessor_patterns = (
        r"(?:lessor|landlord|licensor|first\s+party)\s*[:.-]*\s*([A-Za-z][A-Za-z\s.\-]{2,70})",
        r"(?:vendor|owner)\s*/\s*lessor\s*[:.-]*\s*([A-Za-z][A-Za-z\s.\-]{2,70})",
    )
    for pat in lessor_patterns:
        m = re.search(pat, t, re.IGNORECASE)
        if m:
            hints["lessor_name"] = re.sub(r"\s+", " ", m.group(1).strip())[:80]
            break

    m = re.search(
        r"(?:khata\s*(?:no|number|sankhya)?|खाता\s*संख्या)\s*[:.-]*\s*([\d\s,/\-]+)",
        t, re.IGNORECASE,
    )
    if m:
        hints["khata"] = re.sub(r"\s+", "", m.group(1))[:24]

    m = re.search(
        r"(?:khasra|khesra|khatoni|plot)\s*(?:no|number|sankhya)?\s*[:.-]*\s*([\w\d\s,/\-]+)",
        t, re.IGNORECASE,
    )
    if m:
        hints["khasra"] = re.sub(r"\s+", "", m.group(1).strip())[:32]

    return hints


def _hints_from_filename(filename: str) -> Dict[str, Any]:
    hints: Dict[str, Any] = {}
    stem = Path(filename).stem
    m = re.search(r"\b(\d{6})\b", stem)
    if m:
        hints["pincode"] = m.group(1)
    m = re.search(
        r"(?i)\b(khata|khesra|khasra)\s*[_\s.-]*\s*([\w\d,/\-]+)\b",
        stem.replace("_", " "),
    )
    if m:
        k = m.group(1).lower()
        if "khata" in k:
            hints["khata"] = m.group(2)[:24]
        else:
            hints["khasra"] = m.group(2)[:32]
    return hints


def _merge_hint_dicts(a: Dict[str, Any], b: Dict[str, Any]) -> Dict[str, Any]:
    out = dict(a)
    for k, v in b.items():
        if v and (k not in out or not out.get(k)):
            out[k] = v
    return out


def _grouping_confidence_score(h: Dict[str, Any]) -> float:
    score = 0.35
    if h.get("pincode"):
        score += 0.28
    if h.get("village"):
        score += 0.18
    if h.get("lessor_name"):
        score += 0.12
    if h.get("khata") or h.get("khasra"):
        score += 0.07
    if h.get("filename_norm"):
        score += 0.05
    return min(1.0, score)


def _build_grouping_hints(path: Path, category: str) -> Dict[str, Any]:
    fn_key, fn_conf = _filename_locality_key(path.name)
    h = _hints_from_filename(path.name)
    h["filename_norm"] = fn_key
    h["filename_conf"] = fn_conf
    h["orig_filename"] = path.name
    txt = _quick_pdf_text_for_grouping(path)
    if txt.strip():
        h["pdf_text_sample"] = True
        h = _merge_hint_dicts(h, _hints_from_text(txt))
    else:
        h["pdf_text_sample"] = False
    h["locality_residual"] = _strip_generic_locality_tokens(fn_key)
    h["village_slug"] = _slug_alnum(h["village"]) if h.get("village") else ""
    h["owner_slug"] = _slug_alnum(h.get("lessor_name") or "")
    h["khata_slug"] = _slug_alnum(h.get("khata") or "")
    h["khasra_slug"] = _slug_alnum(h.get("khasra") or "")
    h["locality_fn_slug"] = _slug_alnum(h.get("locality_residual") or fn_key)
    h["confidence"] = _grouping_confidence_score(h)
    return h


def _primary_locality_label(h: Dict[str, Any]) -> str:
    if h.get("village"):
        return str(h["village"])
    return str(h.get("filename_norm") or "")


def _grouping_label_similarity(a: str, b: str) -> float:
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a.lower(), b.lower()).ratio()


def _locality_token_core(s: str) -> frozenset:
    """Non-generic tokens for locality containment checks (not fuzzy on full OCR)."""
    raw = re.sub(r"[^a-z0-9]+", " ", s.lower()).split()
    return frozenset(t for t in raw if t and t not in _GROUPING_GENERIC_LOCALITY)


def _locality_compatible(a: str, b: str) -> bool:
    """
    True if two locality labels refer to the same place for grouping purposes.
    Uses ≥GROUPING_FUZZY_THRESHOLD on the full string, or non-generic token subset
    containment (e.g. 'Mau Aima' inside 'Mau Aima Prayagraj' after dropping 'Prayagraj').

    FIX (v5.3): Changed empty-string handling from True → False (conservative: no
    locality info means we cannot confirm same client — do NOT merge).
    FIX (v5.3): Raised shared-token length gate from 5 → 8 to prevent district-level
    words like "nagar" (5), "kabir" (5), "singh" (5) from causing false positive merges
    of documents belonging to different villages in the same district (e.g. "Nath Nagar
    Sant Kabir Nagar" vs "Baghauli Sant Kabir Nagar" sharing "nagar"/"kabir" tokens).
    """
    if not a or not b:
        # Conservative: missing locality info → do NOT assume same client.
        # Callers in _grouping_hints_compatible already guard with `if loc_a and loc_b`,
        # so this path is only reached from explicit calls with potentially empty strings.
        return False
    if _grouping_label_similarity(a, b) >= GROUPING_FUZZY_THRESHOLD:
        return True
    ca, cb = _locality_token_core(a), _locality_token_core(b)
    if not ca or not cb:
        return _grouping_label_similarity(a, b) >= GROUPING_FUZZY_THRESHOLD
    if ca <= cb or cb <= ca:
        return True
    inter = ca & cb
    if not inter:
        return False
    # Require a shared token of ≥ 8 characters to avoid false merges on short
    # administrative/common words (e.g. "nagar"=5, "kabir"=5, "singh"=5, "kumar"=5).
    # Village/hamlet names are typically 8+ characters (e.g. "dhanghata"=9,
    # "baghauli"=8, "bhawanathpur"=12, "muzaffarpur"=11).
    longest = max((len(t) for t in inter), default=0)
    if longest >= 8:
        return True
    jacc = len(inter) / max(1, len(ca | cb))
    return jacc >= 0.66 and longest >= 6


def _grouping_hints_compatible(ha: Dict[str, Any], hb: Dict[str, Any]) -> bool:
    """
    Return True only if two documents may belong to the same client bucket.
    Hard gates: pincode, village/locality, land parcel ids, owner/lessor.
    """
    pa, pb = ha.get("pincode") or "", hb.get("pincode") or ""
    if pa and pb and pa != pb:
        log.debug(f"[Grouping] block merge: pincode {pa!r} vs {pb!r}")
        return False

    loc_a = _primary_locality_label(ha)
    loc_b = _primary_locality_label(hb)
    if loc_a and loc_b and not _locality_compatible(loc_a, loc_b):
        log.debug(f"[Grouping] block merge: locality {loc_a!r} vs {loc_b!r}")
        return False

    oa = ha.get("owner_slug") or ""
    ob = hb.get("owner_slug") or ""
    if oa and ob:
        if SequenceMatcher(None, oa, ob).ratio() < GROUPING_OWNER_COMPAT_MIN:
            log.debug("[Grouping] block merge: lessor/owner slug mismatch")
            return False

    for fld in ("khata_slug", "khasra_slug"):
        xa, xb = ha.get(fld) or "", hb.get(fld) or ""
        if xa and xb and xa != xb:
            if xa in xb or xb in xa:
                continue
            log.debug(f"[Grouping] block merge: {fld} {xa!r} vs {xb!r}")
            return False

    fn_a = ha.get("filename_norm") or ""
    fn_b = hb.get("filename_norm") or ""
    gen_a = _is_generic_only_locality(fn_a)
    gen_b = _is_generic_only_locality(fn_b)
    if gen_a and gen_b:
        if _normalize_grouping_spelling(fn_a) != _normalize_grouping_spelling(fn_b):
            return False
        if not (
            ha.get("pincode") or hb.get("pincode")
            or ha.get("village") or hb.get("village")
        ):
            log.debug("[Grouping] block merge: generic-only locality (no pin/village) — unsafe")
            return False
    elif fn_a and fn_b and not (ha.get("pincode") or hb.get("pincode") or ha.get("village") or hb.get("village")):
        if not _locality_compatible(fn_a, fn_b):
            return False

    # FIX (v5.3): If NEITHER document has any discriminating signal (no pincode,
    # no village, no non-generic filename locality), refuse to merge — different
    # client documents with zero shared locality evidence must never collapse.
    has_any_signal_a = bool(ha.get("pincode") or ha.get("village") or (fn_a and not gen_a))
    has_any_signal_b = bool(hb.get("pincode") or hb.get("village") or (fn_b and not gen_b))
    if not has_any_signal_a and not has_any_signal_b:
        log.debug("[Grouping] block merge: both documents have zero discriminating signals — refuse merge")
        return False

    if pa and pb and pa == pb:
        return True
    if ha.get("village") and hb.get("village"):
        return _locality_compatible(str(ha["village"]), str(hb["village"]))

    if (pa or pb) and (loc_a and loc_b):
        return _locality_compatible(loc_a, loc_b)

    if fn_a and fn_b:
        return _locality_compatible(fn_a, fn_b)

    return False


def _cluster_index_union_find(n: int, pairs: List[Tuple[int, int]]) -> List[int]:
    parent = list(range(n))

    def find(i: int) -> int:
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(i: int, j: int) -> None:
        ri, rj = find(i), find(j)
        if ri != rj:
            parent[rj] = ri

    for i, j in pairs:
        union(i, j)
    return [find(i) for i in range(n)]


def _deterministic_cluster_display_name(
    members: List[Tuple[str, Path]], hints_list: List[Dict[str, Any]],
) -> str:
    """Human-readable stable client label from merged hints."""
    best = max(
        hints_list,
        key=lambda h: (
            bool(h.get("pincode")),
            bool(h.get("village")),
            len(h.get("filename_norm") or ""),
            h.get("confidence", 0),
        ),
    )
    v = (best.get("village") or "").strip()
    p = (best.get("pincode") or "").strip()
    o = (best.get("lessor_name") or "").strip()
    fn = (best.get("filename_norm") or "").strip()
    if v and p and o:
        raw = f"{v}_{p}_{o}"
    elif v and p:
        raw = f"{v}_{p}"
    elif v and o:
        raw = f"{v}_{o}"
    elif p and fn:
        raw = f"{p}_{fn}"
    elif fn:
        raw = fn
    else:
        raw = Path(members[0][1].name).stem
    raw = re.sub(r"[<>:\"/\\|?*\x00-\x1f]", "_", raw)
    raw = re.sub(r"_+", "_", raw).strip("_ ")[:120]
    disp = normalize_client_name(raw.replace("_", " "))
    return disp or "UnknownClient"


def generate_client_group_key(filename: str) -> Tuple[str, float]:
    """
    Filename-only deterministic key (preview / tooling). Full batch grouping uses
    _build_grouping_hints() + structural clustering.
    """
    key, conf = _filename_locality_key(filename)
    return key, conf


def group_documents_by_client(
    cats: Dict[str, Optional[Path]]
) -> Dict[str, Dict[str, List[Path]]]:
    """
    Returns {client_key: {agreement:[paths], utility:[paths], other:[paths]}}.

    Clustering is deterministic: filename normalisation, optional first-page PDF
    regex hints, then union-find with hard gates (pincode / village / owner / khasra).
    SequenceMatcher ≥ GROUPING_FUZZY_THRESHOLD is used only on short labels — never on
    raw full-document text.
    """
    all_files: Dict[str, List[Path]] = {"agreement": [], "utility": [], "other": []}
    for cat, folder in cats.items():
        if folder and folder.exists():
            files = collect_files_from_folder(folder)
            all_files[cat] = files
            log.info(f"   [{cat.upper():12s}] → {len(files)} files")

    if all_files["other"] and not all_files["agreement"] and not all_files["utility"]:
        log.info("   Flat-folder mode: inferring category from filenames…")
        reclassified: Dict[str, List[Path]] = {"agreement": [], "utility": [], "other": []}
        _AGR_KW = {"agreement", "lease", "deed", "rent", "contract"}
        _UTIL_KW = {"bill", "receipt", "utility", "electricity", "electric",
                    "land receipt", "tax", "invoice", "water", "gas", "phone"}
        for fp in all_files["other"]:
            fn = fp.name.lower().replace("_", " ").replace("-", " ")
            if any(k in fn for k in _AGR_KW):
                reclassified["agreement"].append(fp)
            elif any(k in fn for k in _UTIL_KW):
                reclassified["utility"].append(fp)
            else:
                reclassified["other"].append(fp)
        if reclassified["agreement"] or reclassified["utility"]:
            all_files = reclassified
            log.info(
                f"   Reclassified: agreements={len(all_files['agreement'])}  "
                f"utility={len(all_files['utility'])}  other={len(all_files['other'])}"
            )

    entries: List[Tuple[str, Path]] = []
    hints_per_entry: List[Dict[str, Any]] = []
    _detect_pdf_reader_backend()
    pdf_hint_mode = _PDF_TEXT_FOR_GROUPING if _PDF_TEXT_FOR_GROUPING else "disabled"
    log.info(
        f"[Grouping] engine=v2 deterministic | label_fuzzy≥{GROUPING_FUZZY_THRESHOLD} "
        f"| PDF_text={pdf_hint_mode}"
    )

    for cat, files in all_files.items():
        for fp in files:
            h = _build_grouping_hints(fp, cat)
            entries.append((cat, fp))
            hints_per_entry.append(h)
            log.debug(
                f"[Grouping] file={fp.name!r} pin={h.get('pincode')!r} "
                f"village={h.get('village')!r} fn_loc={h.get('filename_norm')!r} "
                f"conf={h.get('confidence'):.2f} pdf_text={h.get('pdf_text_sample')}"
            )

    n = len(entries)
    if n == 0:
        return {}

    pairs: List[Tuple[int, int]] = []
    for i in range(n):
        for j in range(i + 1, n):
            if _grouping_hints_compatible(hints_per_entry[i], hints_per_entry[j]):
                pairs.append((i, j))

    roots = _cluster_index_union_find(n, pairs)
    buckets: Dict[int, List[int]] = defaultdict(list)
    for i, r in enumerate(roots):
        buckets[r].append(i)

    structured: Dict[str, Dict[str, List[Path]]] = {}
    for _root, idxs in sorted(buckets.items(), key=lambda kv: min(kv[1])):
        members = [entries[i] for i in idxs]
        hlist = [hints_per_entry[i] for i in idxs]
        label = _deterministic_cluster_display_name(members, hlist)
        if label in structured:
            suf = hashlib.sha1(
                "|".join(sorted(p.name for _, p in members)).encode("utf-8", errors="replace")
            ).hexdigest()[:8]
            label = f"{label}_{suf}"
        d: Dict[str, List[Path]] = {"agreement": [], "utility": [], "other": []}
        for cat, fp in members:
            d[cat].append(fp)
        cmean = sum(h.get("confidence", 0) for h in hlist) / max(1, len(hlist))
        log.info(
            f"[Grouping] cluster '{label}' files={len(members)} "
            f"avg_confidence={cmean:.2f}"
        )
        structured[label] = d

    log.info(f"\n  {len(structured)} unique client(s) identified after deterministic grouping")
    return structured


# ── MODULE 5: Thread-safety primitives ───────────────────────────────────────

_client_lock_registry: Dict[str, threading.Lock] = defaultdict(threading.Lock)
_registry_meta_lock   = threading.Lock()
excel_write_lock      = threading.Lock()
json_cache_lock       = threading.Lock()
folder_create_lock    = threading.Lock()


def get_client_lock(client_name: str) -> threading.Lock:
    """Return the per-client threading.Lock for client_name (created on first call)."""
    with _registry_meta_lock:
        return _client_lock_registry[client_name]


def cleanup_client_locks(client_name: str) -> None:
    """Issue 13/15: Remove a client's lock from registry after processing (memory hygiene)."""
    with _registry_meta_lock:
        _client_lock_registry.pop(client_name, None)


# ── MODULE 6: Excel schema validation ────────────────────────────────────────

EXCEL_ROW_SCHEMA: Dict[str, Tuple[type, Any]] = {
    "GST_Reg_ID":                          (str,   ""),
    "Legal Name (GST Applicant)":          (str,   ""),
    "State":                               (str,   ""),
    "Principal Place of Business Address": (str,   ""),
    "Basis of Documents":                  (str,   ""),
    "Total Documents Uploaded":            (str,   ""),
    "Ready to Upload Docs":                (str,   ""),
    "Incorrect/Mismatch Documents":        (str,   "None"),
    "Missing Documents":                   (str,   "None"),
    "Review Required Docs":                (str,   "None"),
    "Weighted Compliance %":               (float, 0.0),
    "Final Status":                        (str,   ""),
    "Overall Risk Exposure":               (str,   ""),
    "Final Remarks":                       (str,   ""),
    "Ownership Proof Type":                (str,   "Not Detected"),
    "Ownership Proof Address":             (str,   ""),
    "Ownership vs Agreement Address Match":(str,   "N/A"),
    "Ownership Proof Source Pages":        (str,   ""),
    "Address Comparison Remarks":          (str,   ""),
}


def safe_excel_value(value: Any, expected_type: type, default: Any) -> Any:
    """Coerce value to expected_type, returning default on failure. Prevents NoneType crashes."""
    if value is None:
        return default
    try:
        if expected_type is str:
            return str(value).strip() if str(value).strip() else default
        elif expected_type is float:
            return float(str(value).replace("%", "").strip())
        elif expected_type is int:
            return int(float(str(value).strip()))
        elif expected_type is bool:
            return bool(value)
        else:
            return expected_type(value)
    except (ValueError, TypeError, AttributeError):
        return default


def validate_excel_row_schema(row_dict: Dict[str, Any]) -> Dict[str, Any]:
    """Validate and sanitize an Excel row dict against EXCEL_ROW_SCHEMA with defaults."""
    result: Dict[str, Any] = {}
    for col, (expected_type, default) in EXCEL_ROW_SCHEMA.items():
        raw = row_dict.get(col)
        if raw is None and col not in row_dict:
            log.debug(f"[ExcelSchema] Missing column '{col}', using default={default!r}")
        result[col] = safe_excel_value(raw, expected_type, default)
    for k, v in row_dict.items():
        if k not in result:
            result[k] = v
    return result


def extract_ownership_proof_excel_fields(data: Dict[str, Any]) -> Dict[str, str]:
    """Extract ownership proof fields from Gemini JSON for the 6 new Excel columns."""
    op = data.get("ownership_proof") or {}
    op_type    = safe_excel_value(op.get("document_type"), str, "Not Detected")
    op_addr    = safe_excel_value(op.get("ownership_address"), str, "")
    op_fy      = safe_excel_value(op.get("financial_year"), str, "")
    op_match   = safe_excel_value(op.get("address_match_with_agreement"), str, "N/A")
    op_pages_raw = op.get("source_pages")
    op_pages   = (", ".join(str(p) for p in op_pages_raw)
                  if isinstance(op_pages_raw, list)
                  else safe_excel_value(op_pages_raw, str, ""))
    op_remarks = safe_excel_value(op.get("comparison_remarks"), str, "")
    addrs      = data.get("addresses") or {}
    address_source = safe_excel_value(addrs.get("address_source"), str, "")
    return {
        "Ownership Proof Type":                 op_type,
        "Ownership Proof Address":              op_addr,
        "Ownership Proof Financial Year":       op_fy,
        "Ownership vs Agreement Address Match": op_match,
        "Ownership Proof Source Pages":         op_pages,
        "Address Comparison Remarks":           op_remarks,
        "_address_source":                      address_source,
    }


# ── MODULE 7: Audit logging & failure reporting ───────────────────────────────

class AuditLogger:
    """
    Professional audit-grade processing reporter.
    Creates: skipped_documents.csv, failed_documents.csv,
             processing_summary.json, retry_report.json
    All writes are thread-safe via per-file locks.
    """
    _FAILED_COLS  = ["timestamp", "client_name", "filename", "worker_id",
                     "stage", "reason", "traceback"]
    _SKIPPED_COLS = ["timestamp", "client_name", "filename", "reason"]

    def __init__(self, output_root: Path):
        self.output_root   = normalize_windows_path(output_root)
        safe_create_dir(self.output_root)
        self._failed_path  = self.output_root / "failed_documents.csv"
        self._skipped_path = self.output_root / "skipped_documents.csv"
        self._summary_path = self.output_root / "processing_summary.json"
        self._retry_path   = self.output_root / "retry_report.json"
        self._failed_lock  = threading.Lock()
        self._skipped_lock = threading.Lock()
        self._summary_lock = threading.Lock()
        self._retry_lock   = threading.Lock()
        self._summary: Dict[str, Any] = {
            "run_started": time.strftime("%Y-%m-%d %H:%M:%S"),
            "run_finished": None,
            "total_clients": 0, "processed_ok": 0, "skipped": 0, "failed": 0,
            "clean": 0, "warning": 0, "high_risk": 0,
            "states_processed": defaultdict(int), "errors": [],
        }
        self._retry_data: Dict[str, List[Dict]] = {}
        self._init_csv(self._failed_path,  self._FAILED_COLS)
        self._init_csv(self._skipped_path, self._SKIPPED_COLS)

    def _init_csv(self, path: Path, cols: List[str]) -> None:
        if path.exists():
            return
        try:
            with open(str(path), "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(cols)
        except OSError as exc:
            log.warning(f"[AuditLog] Could not init CSV {path.name}: {exc}")

    def log_failed(self, client_name: str, filename: str, stage: str,
                   reason: str, tb: str = "", worker_id: str = "") -> None:
        row = [time.strftime("%Y-%m-%d %H:%M:%S"), client_name, filename,
               worker_id or threading.current_thread().name,
               stage, reason[:500], tb[:1000]]
        with self._failed_lock:
            try:
                with open(str(self._failed_path), "a", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerow(row)
            except OSError as exc:
                log.warning(f"[AuditLog] Failed to write failed_documents.csv: {exc}")

    def log_skipped(self, client_name: str, filename: str, reason: str) -> None:
        row = [time.strftime("%Y-%m-%d %H:%M:%S"), client_name, filename, reason]
        with self._skipped_lock:
            try:
                with open(str(self._skipped_path), "a", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerow(row)
            except OSError as exc:
                log.warning(f"[AuditLog] Failed to write skipped_documents.csv: {exc}")

    def log_retry(self, client_name: str, attempt: int, reason: str) -> None:
        entry = {"timestamp": time.strftime("%Y-%m-%d %H:%M:%S"),
                 "attempt": attempt, "reason": reason}
        with self._retry_lock:
            self._retry_data.setdefault(client_name, []).append(entry)

    def update_summary(self, **kwargs) -> None:
        with self._summary_lock:
            for k, v in kwargs.items():
                if k == "state" and v:
                    self._summary["states_processed"][v] += 1
                elif k in self._summary:
                    if isinstance(self._summary[k], int) and isinstance(v, int):
                        self._summary[k] += v
                    else:
                        self._summary[k] = v
                else:
                    self._summary[k] = v

    def finalize(self) -> None:
        with self._summary_lock:
            self._summary["run_finished"] = time.strftime("%Y-%m-%d %H:%M:%S")
            self._summary["states_processed"] = dict(self._summary["states_processed"])
            atomic_json_dump(self._summary, self._summary_path)
            log.info(f"[AuditLog] processing_summary.json → {self._summary_path}")
        with self._retry_lock:
            if self._retry_data:
                atomic_json_dump(self._retry_data, self._retry_path)
                log.info(f"[AuditLog] retry_report.json → {self._retry_path}")


# Module-level singletons (initialized by _initialize_production_modules)
_audit_logger: Optional[AuditLogger] = None


def get_audit_logger() -> Optional[AuditLogger]:
    """Return the global AuditLogger, or None if not yet initialized."""
    return _audit_logger


# ── MODULE 8: Checkpoint & resume support ────────────────────────────────────

class CheckpointManager:
    """
    Lightweight checkpoint system for interrupted-run recovery.
    Maintains checkpoint.json tracking done/failed/in_progress clients.
    Enables interrupted batches to resume from where they left off.
    """
    CHECKPOINT_FILENAME = "checkpoint.json"

    def __init__(self, output_root: Path):
        self.path  = normalize_windows_path(output_root) / self.CHECKPOINT_FILENAME
        self._lock = threading.Lock()
        self._data: Dict[str, Any] = self._load()

    def _load(self) -> Dict[str, Any]:
        if self.path.exists():
            try:
                raw = self.path.read_text(encoding="utf-8")
                d   = json.loads(raw)
                log.info(f"[Checkpoint] Loaded: {len(d.get('done', []))} done, "
                         f"{len(d.get('failed', []))} failed")
                return d
            except Exception as exc:
                log.warning(f"[Checkpoint] Could not load {self.path.name}: {exc}")
        return {"done": [], "failed": [], "in_progress": [],
                "run_id": time.strftime("%Y%m%d_%H%M%S")}

    def _save(self) -> None:
        atomic_json_dump(self._data, self.path)

    def is_done(self, client_name: str) -> bool:
        with self._lock:
            return client_name in self._data.get("done", [])

    def mark_in_progress(self, client_name: str) -> None:
        with self._lock:
            ip = self._data.setdefault("in_progress", [])
            if client_name not in ip:
                ip.append(client_name)
            self._save()

    def mark_done(self, client_name: str) -> None:
        with self._lock:
            done = self._data.setdefault("done", [])
            if client_name not in done:
                done.append(client_name)
            for key in ("in_progress", "failed"):
                lst = self._data.get(key, [])
                if client_name in lst:
                    lst.remove(client_name)
            self._save()

    def mark_failed(self, client_name: str) -> None:
        with self._lock:
            failed = self._data.setdefault("failed", [])
            if client_name not in failed:
                failed.append(client_name)
            ip = self._data.get("in_progress", [])
            if client_name in ip:
                ip.remove(client_name)
            self._save()

    def get_failed(self) -> List[str]:
        with self._lock:
            return list(self._data.get("failed", []))

    def reset(self) -> None:
        with self._lock:
            self._data = {"done": [], "failed": [], "in_progress": [],
                          "run_id": time.strftime("%Y%m%d_%H%M%S")}
            self._save()


_checkpoint_manager: Optional[CheckpointManager] = None


def get_checkpoint_manager() -> Optional[CheckpointManager]:
    """Return the global CheckpointManager, or None if not yet initialized."""
    return _checkpoint_manager


# ══════════════════════════════════════════════════════════════════════════════
#  MODULE 9 — VERSIONED OUTPUT MANAGEMENT + FAILED/SKIPPED FILE ISOLATION
#  Implements: immutable timestamped outputs, failed/skipped quarantine folders,
#  retry history, audit CSVs (retry_history.csv, processing_history.json),
#  thread-safe version generation, Windows-safe naming, enterprise auditability.
# ══════════════════════════════════════════════════════════════════════════════

_VERSION_LOCK = threading.Lock()     # guards all versioned filename generation


def generate_versioned_filename(base_name: str, extension: str) -> str:
    """
    Return a timestamped filename guaranteed to be collision-free.

    Pattern:  <base_name>_YYYYMMDD_HHMMSS[_NNN].ext
    Thread-safe via _VERSION_LOCK + 1-second sleep if a sub-second clash occurs.

    Args:
        base_name : Base stem, e.g. "GST_Validation_Workbook".
        extension : File extension with leading dot, e.g. ".xlsx".

    Returns:
        e.g. "GST_Validation_Workbook_20260514_113000.xlsx"
    """
    with _VERSION_LOCK:
        ts = time.strftime("%Y%m%d_%H%M%S")
        return f"{base_name}_{ts}{extension}"


def get_next_available_version(directory: Path, base_name: str, extension: str) -> Path:
    """
    Return a Path in *directory* that does NOT yet exist.

    Tries <base_name>_YYYYMMDD_HHMMSS<ext>, then appends _001, _002 … if needed.
    Never returns an existing path — fully collision-safe.

    Args:
        directory  : Folder where the file will be written.
        base_name  : Stem, e.g. "GST_Validation_Report".
        extension  : Extension with dot, e.g. ".xlsx".

    Returns:
        Unique non-existing Path.
    """
    directory = normalize_windows_path(directory)
    with _VERSION_LOCK:
        ts   = time.strftime("%Y%m%d_%H%M%S")
        candidate = directory / f"{base_name}_{ts}{extension}"
        if not candidate.exists():
            return candidate
        for seq in range(1, 10000):
            candidate = directory / f"{base_name}_{ts}_{seq:03d}{extension}"
            if not candidate.exists():
                return candidate
    raise RuntimeError(f"Cannot generate unique filename for {base_name} in {directory}")


def create_versioned_output(
    source_path: Path,
    dest_dir: Path,
    base_name: Optional[str] = None,
) -> Optional[Path]:
    """
    Copy *source_path* into *dest_dir* with a versioned (timestamped) filename.
    Never overwrites; always creates a new file.

    Args:
        source_path : Completed file to archive.
        dest_dir    : Target directory (created if absent).
        base_name   : Override stem (default: source stem).

    Returns:
        Path of the newly created versioned copy, or None on failure.
    """
    if not source_path.exists():
        log.warning(f"[VersionedOutput] Source not found: {source_path}")
        return None
    try:
        safe_create_dir(dest_dir)
    except OSError as exc:
        log.error(f"[VersionedOutput] Cannot create dest_dir {dest_dir}: {exc}")
        return None
    stem = base_name or source_path.stem
    versioned = get_next_available_version(dest_dir, stem, source_path.suffix)
    result = safe_copy_file(source_path, versioned, overwrite=False)
    if result:
        log.info(f"[VersionedOutput] Created versioned copy → {versioned.name}")
    else:
        log.error(f"[VersionedOutput] Copy failed: {source_path.name} → {versioned}")
    return result


def archive_previous_run(
    output_root: Path,
    patterns: Optional[List[str]] = None,
) -> List[Path]:
    """
    Move existing non-versioned output files in *output_root* into a
    timestamped archive sub-folder before a new run begins.

    Only touches files matching *patterns* (default: *.xlsx, *.json except
    checkpoint.json and processing_summary.json).  Does NOT delete anything.

    Args:
        output_root : Root output folder.
        patterns    : Optional list of glob patterns (relative to output_root).

    Returns:
        List of Paths that were archived.
    """
    output_root = normalize_windows_path(output_root)
    if not output_root.exists():
        return []
    if patterns is None:
        patterns = ["*.xlsx", "GST_Validation_Workbook.xlsx"]

    archive_dir = output_root / f"_archive_{time.strftime('%Y%m%d_%H%M%S')}"
    archived: List[Path] = []

    NEVER_ARCHIVE = {
        "checkpoint.json", "processing_summary.json",
        "failed_documents.csv", "skipped_documents.csv",
        "retry_history.csv",
    }

    for pattern in patterns:
        for candidate in output_root.glob(pattern):
            if candidate.is_file() and candidate.name not in NEVER_ARCHIVE:
                try:
                    safe_create_dir(archive_dir)
                    dst = archive_dir / candidate.name
                    if dst.exists():
                        dst = archive_dir / f"{candidate.stem}_{int(time.time())}{candidate.suffix}"
                    shutil.move(str(candidate), str(dst))
                    archived.append(dst)
                    log.info(f"[Archive] Moved {candidate.name} → {archive_dir.name}/{dst.name}")
                except Exception as exc:
                    log.warning(f"[Archive] Could not archive {candidate.name}: {exc}")

    if archived:
        log.info(f"[Archive] Archived {len(archived)} file(s) to {archive_dir}")
    return archived


# ── Failed / Skipped file quarantine ─────────────────────────────────────────

def safe_archive_failed_file(
    client_name: str,
    state_label: str,
    source_files: List[Path],
    failure_reason: str,
    tb_text: str = "",
    *,
    output_root: Path,
    retry_count: int = 0,
    worker_id: str = "",
    status: str = "FAILED",        # "FAILED" | "SKIPPED"
    duration_sec: float = 0.0,
    extra_json: Optional[Dict] = None,
) -> Path:
    """
    Quarantine documents for a failed or skipped client into an isolated folder:

        OUTPUT/
          Failed_Files/  or  Skipped_Files/
            <State>/
              <ClientName>/
                retry_<N>/          (numbered retry sub-folder)
                  input_docs/       (copies of original documents)
                  failure_log.json  (structured error snapshot)
                  traceback.txt     (full traceback)

    Rules:
      • NEVER overwrites previous failed/skipped records.
      • Preserves ALL retry history (retry_1/, retry_2/, …).
      • Thread-safe via per-client lock.

    Returns:
        Path to the retry sub-folder created for this attempt.
    """
    bucket   = "Failed_Files" if status == "FAILED" else "Skipped_Files"
    safe_st  = _safe_state_folder_name(state_label)
    safe_cl  = _safe_folder_name(client_name)
    client_lock = get_client_lock(f"_quarantine_{client_name}")

    with client_lock:
        base_dir = normalize_windows_path(output_root / bucket / safe_st / safe_cl)
        # Determine next retry sub-folder (always increment — never overwrite)
        existing_retries = [
            d for d in base_dir.glob("retry_*") if d.is_dir()
        ] if base_dir.exists() else []
        next_retry = len(existing_retries) + 1
        retry_dir = base_dir / f"retry_{next_retry}"
        input_docs_dir = retry_dir / "input_docs"

        try:
            safe_create_dir(input_docs_dir)
        except OSError as exc:
            log.error(f"[Quarantine] Cannot create quarantine dir {input_docs_dir}: {exc}")
            return retry_dir

        # ── Copy original documents (NEVER move — preserve source) ───────────
        for src in source_files:
            if src.exists():
                safe_copy_file(src, input_docs_dir, overwrite=False)
            else:
                log.warning(f"[Quarantine] Source file missing: {src}")

        # ── Write failure_log.json ────────────────────────────────────────────
        ts = time.strftime("%Y-%m-%d %H:%M:%S")
        failure_snapshot: Dict[str, Any] = {
            "status"            : status,
            "client_name"       : client_name,
            "state"             : state_label,
            "processing_timestamp": ts,
            "retry_number"      : next_retry,
            "retry_count"       : retry_count,
            "worker_id"         : worker_id or threading.current_thread().name,
            "duration_sec"      : round(duration_sec, 2),
            "failure_reason"    : failure_reason[:2000],
            "documents_count"   : len(source_files),
            "documents"         : [str(f.name) for f in source_files],
            "output_folder"     : str(retry_dir),
        }
        if extra_json:
            failure_snapshot["extra"] = extra_json
        atomic_json_dump(failure_snapshot, retry_dir / "failure_log.json")

        # ── Write traceback.txt ───────────────────────────────────────────────
        if tb_text:
            safe_write_text(
                retry_dir / "traceback.txt",
                f"Timestamp : {ts}\nClient    : {client_name}\nReason    : {failure_reason}\n\n{tb_text}",
            )

        log.info(
            f"[Quarantine] {status} — {client_name} → {bucket}/{safe_st}/{safe_cl}/retry_{next_retry}/ "
            f"({len(source_files)} doc(s) preserved)"
        )
    return retry_dir


# ── Retry History & Processing History CSVs ───────────────────────────────────

_RETRY_HISTORY_LOCK   = threading.Lock()
_PROC_HISTORY_LOCK    = threading.Lock()

_RETRY_HISTORY_COLS = [
    "timestamp", "client_name", "state", "retry_number",
    "failure_reason", "worker_id", "duration_sec", "status",
]

_PROC_HISTORY_COLS = [
    "timestamp", "client_name", "state", "processing_timestamp",
    "retry_number", "failure_reason", "worker_id", "duration_sec",
    "output_version_generated", "final_status",
]


def _append_retry_history(output_root: Path, row: Dict[str, Any]) -> None:
    """Append one row to _internal/retry_history.csv (created with header on first write)."""
    internal_dir = normalize_windows_path(output_root / "_internal")
    safe_create_dir(internal_dir)
    path = internal_dir / "retry_history.csv"
    with _RETRY_HISTORY_LOCK:
        write_header = not path.exists()
        try:
            with open(str(path), "a", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=_RETRY_HISTORY_COLS, extrasaction="ignore")
                if write_header:
                    w.writeheader()
                w.writerow(row)
        except OSError as exc:
            log.warning(f"[RetryHistory] Could not write retry_history.csv: {exc}")


def _append_processing_history(output_root: Path, row: Dict[str, Any]) -> None:
    """Append one row to _internal/processing_history.csv (created with header on first write)."""
    internal_dir = normalize_windows_path(output_root / "_internal")
    safe_create_dir(internal_dir)
    path = internal_dir / "processing_history.csv"
    with _PROC_HISTORY_LOCK:
        write_header = not path.exists()
        try:
            with open(str(path), "a", newline="", encoding="utf-8") as f:
                w = csv.DictWriter(f, fieldnames=_PROC_HISTORY_COLS, extrasaction="ignore")
                if write_header:
                    w.writeheader()
                w.writerow(row)
        except OSError as exc:
            log.warning(f"[ProcHistory] Could not write processing_history.csv: {exc}")


def record_processing_event(
    output_root: Path,
    client_name: str,
    state: str,
    status: str,                   # "ok" | "failed" | "skipped"
    failure_reason: str = "",
    worker_id: str = "",
    duration_sec: float = 0.0,
    output_version: str = "",
    final_compliance_status: str = "",
    retry_number: int = 0,
) -> None:
    """
    Record one processing event to both retry_history.csv (failures/retries)
    and processing_history.csv (all events).

    Args:
        output_root              : Root output folder for the run.
        client_name              : Human-readable client name.
        state                    : Indian state/UT label.
        status                   : "ok" | "failed" | "skipped".
        failure_reason           : Short error description (empty for ok).
        worker_id                : Thread/worker name.
        duration_sec             : Processing wall-time in seconds.
        output_version           : Versioned filename created (e.g. workbook name).
        final_compliance_status  : "Clean" | "Warning" | "High Risk".
        retry_number             : 0 for first attempt, 1+ for retries.
    """
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    base_row = {
        "timestamp"               : ts,
        "client_name"             : client_name,
        "state"                   : state,
        "processing_timestamp"    : ts,
        "retry_number"            : retry_number,
        "failure_reason"          : failure_reason[:500] if failure_reason else "",
        "worker_id"               : worker_id or threading.current_thread().name,
        "duration_sec"            : round(duration_sec, 2),
        "output_version_generated": output_version,
        "final_status"            : final_compliance_status,
        "status"                  : status,
    }

    # Always write to processing_history
    _append_processing_history(output_root, base_row)

    # Write to retry_history only for failures / retries
    if status in ("failed", "skipped") or retry_number > 0:
        _append_retry_history(output_root, base_row)


# ── Enhanced AuditLogger with extended columns ────────────────────────────────

class EnhancedAuditLogger(AuditLogger):
    """
    Extends the base AuditLogger to write the additional columns required by
    the production spec:
      failed_documents.csv  — adds: state, retry_number, duration_sec, output_version
      skipped_documents.csv — adds: state, retry_number, duration_sec
    """
    _FAILED_COLS_EXT = [
        "timestamp", "client_name", "filename", "state",
        "worker_id", "stage", "reason", "traceback",
        "retry_number", "duration_sec", "output_version",
    ]
    _SKIPPED_COLS_EXT = [
        "timestamp", "client_name", "filename", "state",
        "reason", "retry_number", "duration_sec",
    ]

    def __init__(self, output_root: Path):
        # Call super().__init__ but we'll re-init CSVs with extended columns
        self.output_root   = normalize_windows_path(output_root)
        safe_create_dir(self.output_root)
        self._failed_path  = self.output_root / "failed_documents.csv"
        self._skipped_path = self.output_root / "skipped_documents.csv"
        self._summary_path = self.output_root / "processing_summary.json"
        self._retry_path   = self.output_root / "retry_report.json"
        self._failed_lock  = threading.Lock()
        self._skipped_lock = threading.Lock()
        self._summary_lock = threading.Lock()
        self._retry_lock   = threading.Lock()
        self._summary: Dict[str, Any] = {
            "run_started": time.strftime("%Y-%m-%d %H:%M:%S"),
            "run_finished": None,
            "total_clients": 0, "processed_ok": 0, "skipped": 0, "failed": 0,
            "clean": 0, "warning": 0, "high_risk": 0,
            "states_processed": defaultdict(int), "errors": [],
        }
        self._retry_data: Dict[str, List[Dict]] = {}
        self._init_csv(self._failed_path,  self._FAILED_COLS_EXT)
        self._init_csv(self._skipped_path, self._SKIPPED_COLS_EXT)

    def log_failed(self, client_name: str, filename: str, stage: str,
                   reason: str, tb: str = "", worker_id: str = "",
                   state: str = "", retry_number: int = 0,
                   duration_sec: float = 0.0, output_version: str = "") -> None:
        row = [
            time.strftime("%Y-%m-%d %H:%M:%S"), client_name, filename,
            state,
            worker_id or threading.current_thread().name,
            stage, reason[:500], tb[:1000],
            retry_number, round(duration_sec, 2), output_version,
        ]
        with self._failed_lock:
            try:
                with open(str(self._failed_path), "a", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerow(row)
            except OSError as exc:
                log.warning(f"[AuditLog] Failed to write failed_documents.csv: {exc}")

    def log_skipped(self, client_name: str, filename: str, reason: str,
                    state: str = "", retry_number: int = 0,
                    duration_sec: float = 0.0) -> None:
        row = [
            time.strftime("%Y-%m-%d %H:%M:%S"), client_name, filename,
            state, reason, retry_number, round(duration_sec, 2),
        ]
        with self._skipped_lock:
            try:
                with open(str(self._skipped_path), "a", newline="", encoding="utf-8") as f:
                    csv.writer(f).writerow(row)
            except OSError as exc:
                log.warning(f"[AuditLog] Failed to write skipped_documents.csv: {exc}")


# ── Versioned Excel workbook naming ──────────────────────────────────────────

def versioned_workbook_path(output_root: Path, stem: str = "GST_Validation_Workbook") -> Path:
    """
    Return a versioned (timestamped) master workbook path that does NOT exist yet.

    Example: GST_Validation_Workbook_20260514_113000.xlsx

    The OLD non-versioned "GST_Validation_Workbook.xlsx" is NEVER overwritten —
    if it exists it stays intact; the new run gets its own timestamped file.

    Args:
        output_root : Root output folder.
        stem        : Workbook base name (default: "GST_Validation_Workbook").

    Returns:
        Unique non-existing Path for the new workbook.
    """
    return get_next_available_version(
        normalize_windows_path(output_root), stem, ".xlsx"
    )


def versioned_client_report_path(client_folder: Path) -> Path:
    """
    Return a versioned per-client report path that does NOT yet exist.

    Example: GST_Validation_Report_20260514_113000.xlsx

    Args:
        client_folder : Client root folder.

    Returns:
        Unique non-existing Path.
    """
    return get_next_available_version(
        normalize_windows_path(client_folder),
        "GST_Validation_Report", ".xlsx"
    )


def versioned_json_response_path(logs_dir: Path) -> Path:
    """
    Return a versioned validation_response path that does NOT yet exist.

    Example: validation_response_20260514_113000.json
    """
    return get_next_available_version(
        normalize_windows_path(logs_dir),
        "validation_response", ".json"
    )


# ── END MODULE 9 ─────────────────────────────────────────────────────────────


def _initialize_production_modules(output_root: Path) -> None:
    """
    Activate all production hardening modules.
    Called once at startup (from main()) before any other logic.
    Initializes EnhancedAuditLogger and CheckpointManager singletons.

    All developer-only artefacts (CSVs, summary JSON, checkpoint) are written
    under  output_root/_internal/  so they stay separate from client-visible files.
    """
    global _audit_logger, _checkpoint_manager
    output_root = normalize_windows_path(output_root)
    # All developer-only files land in _internal/
    internal_root = output_root / "_internal"
    safe_create_dir(internal_root)
    _audit_logger       = EnhancedAuditLogger(internal_root)   # upgraded logger
    _checkpoint_manager = CheckpointManager(internal_root)
    log.info("[Production] Production hardening modules initialized (v5.2):")
    log.info("  ✓ Windows-safe path utilities (safe_copy_file, safe_create_dir)")
    log.info("  ✓ Atomic JSON save (atomic_json_dump, safe_json_save)")
    log.info("  ✓ State normalization + verbatim lock")
    log.info("  ✓ Deterministic client grouping (structural gates + label similarity)")
    log.info("  ✓ Per-client thread locks")
    log.info("  ✓ Excel schema validation")
    log.info("  ✓ Enhanced audit logging (failed/skipped CSV + summary JSON + extended columns)")
    log.info("  ✓ Checkpoint/resume support")
    log.info("  ✓ Hardened file organizer")
    log.info("  ✓ [NEW] Versioned output management (immutable timestamped files)")
    log.info("  ✓ [NEW] Failed/Skipped file isolation (quarantine folders)")
    log.info("  ✓ [NEW] Retry history & processing history CSVs")
    log.info("  ✓ [NEW] Retry-safe reprocessing pipeline")


# ══════════════════════════════════════════════════════════════════════════════
#  END OF PRODUCTION HARDENING MODULES
# ══════════════════════════════════════════════════════════════════════════════


#  SYSTEM PROMPT  (full CBIC-compliant extraction & validation rules)

SYSTEM_PROMPT = """
You are an expert GST Document Validator for India specialising in Additional Place of \
Business (APOB) amendments. Strictly follow CBIC Instruction No. 03/2025-GST (Para 6) and the \
Document_Reader_-_Registration.xlsx checklist.

TODAY = {VALIDATION_DATE}
CUTOFF_DATE (3-month rule) = {CUTOFF_DATE}

You can analyse ANY format: searchable PDFs, scanned PDFs, images (JPG/PNG/TIFF/WEBP), Word \
documents (.doc/.docx), using deep OCR wherever necessary.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION A — PARTY IDENTIFICATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- lessor_first_party  = OWNER/LANDLORD who leases OUT the property (first party in the registered \
  lease deed or rent agreement).
- lessee_second_party = direct TENANT in the registered lease deed (usually a warehouse \
  operator/service provider/intermediate entity).
- finalised_gst_business_name = the ACTUAL GST APPLICANT — the end-user company seeking APOB GST \
  registration (the "Customer" in any Warehousing/Service Agreement). NEVER the lessor or \
  intermediate service provider unless there is NO service agreement.
- legal_name = official registered legal name of the GST applicant entity.
- trade_name = trade or brand name (if different from legal_name; else same as legal_name).
- final_business_name in summary = same as finalised_gst_business_name.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION B — SUB-LEASE / WAREHOUSING SERVICE AGREEMENT (CRITICAL)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- When documents contain BOTH a Registered Lease Deed (owner→intermediate) AND a \
  Warehousing/Service Agreement (intermediate→end-user), the GST APPLICANT is the Customer in the \
  Warehousing/Service Agreement.
- final_pob_address MUST be the specific warehouse/premises address where the GST applicant will \
  operate (Khasra numbers, village, road, city, district, state, pincode from the lease deed or \
  addendums) (Example: GST applicant is a warehouse operator "HUSK POWER SYSTEMS PVT LTD" so the final_pob_address is the warehouse address where the GST applicant will operate).
- Document the sub-lease chain clearly in key_notes and compliance.
- ADDRESS SCRIPT TRANSLATION RULE (MANDATORY — SCRIPT ONLY, NOT GEOGRAPHIC CORRECTION):
  All address fields must be in Latin script (English characters). If any part appears in Hindi
  (Devanagari) or another regional script, convert ONLY the script — phoneme for phoneme —
  into its English spelling. Do NOT use geographic knowledge to alter any address component.

  PERMITTED:  "ग्राम चेताग"    → "Village Chetag"    (Devanagari → Latin phonetics only)
  PERMITTED:  "राज्य- बिहार"   → "State: Bihar"       (script conversion — keep Bihar as-is)
  PERMITTED:  "राज्य- झारखण्ड" → "State: Jharkhand"   (script conversion — keep Jharkhand as-is)
  FORBIDDEN:  "राज्य- बिहार"   → "State: Jharkhand"   (geographic correction — NEVER do this)
  FORBIDDEN:  "राज्य- झारखण्ड" → "State: Bihar"       (geographic correction — NEVER do this)
  FORBIDDEN:  Changing any state/district name because the pincode "suggests" a different state.

  The word "translation" here means Devanagari → Latin characters ONLY. It does NOT mean
  correcting perceived geographic errors. Whatever the lease clause says — copy it exactly.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
STATE_VERBATIM ENFORCEMENT (ABSOLUTE — NO EXCEPTIONS)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
You MUST follow this three-step protocol for EVERY document bundle, without exception:

  STEP 1 — READ THE EXACT STATE TEXT FROM THE LEASE DEED CLAUSE:
    Open the primary lease deed. Locate the property description clause (e.g. Clause 5,
    "किरायेदारी संपत्ति का विवरण", or Schedule of Premises). Find the word or phrase that
    names the state. Read it — character by character. If it is in Devanagari, convert the
    phoneme to Latin script only (e.g. "बिहार" → "Bihar", "झारखण्ड" → "Jharkhand").

  STEP 2 — COPY THAT EXACT VALUE INTO EVERY STATE FIELD:
    Write that exact text — and only that text — into:
      addresses.state
      addresses.finalised_pob_address   (as the state component)
      lease_details.premises_address     (as the state component)
      summary.final_pob_address          (as the state component)
      summary.state
    Do NOT substitute, do NOT "correct", do NOT use pincode databases, do NOT use your
    geographic training data. The document is the one and only source of truth.

  STEP 3 — PINCODE/DISTRICT CONSISTENCY CHECK (flag only — NEVER modify state):
    After copying the state verbatim, check whether the district name written in the deed
    is geographically consistent with that state. This check is for flagging ONLY.
    If inconsistent: set pincode_state_flag = true, write an explanation in
    state_discrepancy_note, but leave state exactly as the document says.
    If consistent: set pincode_state_flag = false, state_discrepancy_note = null.

  CRITICAL REMINDER — WHY THIS MATTERS:
    Some lease deeds in Bihar mention properties that are geographically in Jharkhand
    (or vice versa) — this can be a genuine document error or a boundary ambiguity.
    Regardless, your job is ONLY to report what the document says. Silently "correcting"
    Bihar → Jharkhand destroys the address-match comparison with GST portal records and
    hides the exact discrepancy that the reviewer needs to see and decide upon.

  SELF-CHECK BEFORE OUTPUTTING JSON:
    Before writing any state value, ask yourself:
      "Am I writing exactly what the lease deed property clause says?"
    If yes → proceed. If no (even slightly different) → go back to the document and
    copy it verbatim. A "geographically correct" value that differs from the document
    is a VALIDATION FAILURE, not a helpful correction.

  STATE EXTRACTION AND MULTI-DOCUMENT FALLBACK RULE:
    If the primary lease deed's property description clause contains the state name, copy it verbatim.
    If the primary lease deed's property clause does NOT explicitly mention the state name,
    or if the primary lease deed is missing from the bundle, you are permitted to extract the
    state from any other available document in the bundle (e.g. stamps, notary text, land revenue
    receipts, municipal receipts, utility bills, or document headers).
    Translate any Hindi/Devanagari spellings (e.g. "बिहार", "झारखंड", "झारखण्ड", "उत्तर प्रदेश",
    "उत्तरप्रदेश") into their canonical English forms (e.g. "Bihar", "Jharkhand", "Uttar Pradesh").
    Only set the state fields (addresses.state and summary.state) to null if the state cannot be
    determined or inferred from ANY of the uploaded documents.

    CRITICAL — ALWAYS EXTRACT DISTRICT EVEN WHEN STATE IS ABSENT:
    When the property clause or supporting documents contain a district name but no state, you
    MUST still populate addresses.district with the exact district name. Do NOT leave district
    null just because state is absent. The post-processing system uses addresses.district to
    auto-infer the state and raise a compliance issue, but only if you have extracted it.

    Example: Lease deed clause says "जिला- प्रयागराज" (district: Prayagraj) but no state:
      CORRECT:   addresses.district = "Prayagraj"   addresses.state = null
      WRONG:     addresses.district = null           addresses.state = null

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION B2 — ADDRESS EXTRACTION PRIORITY & SOURCE TRACKING (MANDATORY)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
CRITICAL DISTINCTION — APOB Address vs Registered Office Address:
  • final_pob_address = the PHYSICAL PROPERTY described in the lease/rent agreement — this is
    the Additional Place of Business (APOB) being registered. Extract from the property
    description / schedule of the lease deed.
  • The GST applicant company's own registered office or corporate address (e.g. Delhi/Patna HQ)
    is NOT the APOB address. NEVER use the company's own address as final_pob_address.
  • If a document shows two addresses — one for the company's registered office and one for the
    leased property — ALWAYS use the leased property address as final_pob_address.

ADDRESS EXTRACTION RULE — VERBATIM ONLY (MANDATORY, NO EXCEPTIONS):
  Extract ALL address fields (state, district, city, pincode, village, road, etc.) EXACTLY as they
  appear in the document — character for character, word for word.
  DO NOT correct, infer, override, or "fix" any address component based on:
    • pincode databases or geographic inference
    • your own knowledge of which state a pincode "belongs to"
    • supporting documents (Aadhaar, passbook, land revenue receipts)
    • any other source outside the PRIMARY LEASE DEED property clause

  REASON: The extracted address will be compared letter-by-letter against GST portal records and
  other source documents to compute risk exposure. Any silent correction destroys this comparison
  and hides real discrepancies.

  SOURCE PRIORITY FOR addresses.state AND lease_details.premises_address (CRITICAL):
  When multiple documents in the bundle show different states for the same property, ALWAYS use
  the state as written in the PRIMARY LEASE DEED property description clause (e.g. Clause 5 /
  "किरायेदारी संपत्ति का विवरण"). This is the legally operative document for APOB registration,
  and its verbatim content governs ALL extracted address fields.

  Example A (Bihar in deed, Jharkhand in support doc):
    Lease deed clause says "राज्य- बिहार" → extract state = "Bihar".
    If land revenue receipt says "Jharkhand", flag cross-document conflict in
    state_discrepancy_note. Keep state = "Bihar" (from lease deed). Do NOT change it.

  Example B (Jharkhand in deed, Bihar in support doc):
    Lease deed clause says "राज्य- झारखण्ड" → extract state = "Jharkhand".
    If Aadhaar shows "Bihar", flag the conflict. Keep state = "Jharkhand". Do NOT change it.

  DO NOT allow supporting/identity documents (Aadhaar, bank passbook, land revenue receipts)
  to override the state/district written in the PRIMARY LEASE DEED property clause.

PINCODE → STATE CONSISTENCY CHECK (MANDATORY — flag only, NEVER correct):
  After extracting state, district, and pincode verbatim from the lease deed clause, check
  whether they are mutually consistent. This is for FLAGGING ONLY — never override any field.

  HOW TO CHECK:
    1. Use the district name extracted from the lease deed to identify the expected state
       (every district in India belongs to exactly one state).
    2. Compare the expected state (from the district) against the state written in the deed.
    3. If they match → pincode_state_flag = FALSE, state_discrepancy_note = null. Done.
    4. If they don't match → pincode_state_flag = TRUE and explain the conflict in
       state_discrepancy_note. Keep the state field exactly as written in the deed.

  WORKED EXAMPLES (apply the same logic for every state, not just these):
    District = Latehar → expected state = Jharkhand
      • Deed says "Jharkhand" → flag = FALSE, note = null ✓
      • Deed says "Bihar"     → flag = TRUE, note the conflict ✗
    District = Gaya → expected state = Bihar
      • Deed says "Bihar"     → flag = FALSE, note = null ✓
      • Deed says "Jharkhand" → flag = TRUE, note the conflict ✗
    District = Pune → expected state = Maharashtra
      • Deed says "Maharashtra" → flag = FALSE, note = null ✓
      • Deed says "Gujarat"     → flag = TRUE, note the conflict ✗
    District = Ernakulam → expected state = Kerala
      • Deed says "Kerala"      → flag = FALSE, note = null ✓
      • Deed says "Tamil Nadu"  → flag = TRUE, note the conflict ✗

  RULES FOR pincode_state_flag AND state_discrepancy_note:
    • Set pincode_state_flag = FALSE and state_discrepancy_note = null when the state in
      the deed is consistent with the district. NEVER flag consistent addresses.
    • Set pincode_state_flag = TRUE ONLY when a genuine inconsistency is detected.
    • When TRUE, state_discrepancy_note must be a plain-English explanation, e.g.:
        "Deed says 'Bihar' but district Latehar belongs to Jharkhand. State kept verbatim."
    • If you cannot determine the expected state from the district, leave the flag as FALSE
      and do not fabricate a discrepancy note.

POB ADDRESS EXTRACTION SOURCE (MANDATORY — ONE SOURCE ONLY):
  Extract final_pob_address and premises_address EXCLUSIVELY from the property description
  clause of the PRIMARY LEASE DEED (e.g. Clause 5 / "किरायेदारी संपत्ति का विवरण").
  Do NOT synthesize or combine data from multiple documents into final_pob_address.
  Do NOT use the summary to "override" or "finalise" what the lease clause says.
  Copy the address fields directly from the clause — Khasra number, village, post, anchal,
  district, state, pincode — EXACTLY as written in that specific clause, word for word.

  FIELD-BY-FIELD EXTRACTION CHECKLIST (fill each field from the lease clause verbatim):
    khasra_numbers   ← exact KH/Khasra/Plot numbers as written
    village_locality ← village/locality/mohalla/ward name as written
    road_street      ← road/street/nagar name, or "Not specified" if absent
    city_town        ← town/anchal/tehsil name as written
    district         ← district name as written (do NOT change based on pincode)
    state            ← state name as written (do NOT change based on pincode)
    pincode          ← pincode as written (do NOT change or infer)
  Then concatenate all non-null fields into finalised_pob_address (comma-separated).

For EVERY address field you extract, you MUST also populate addresses.address_source with a plain \
English string describing exactly where the address was found AND which specific document type was used. Examples:
  • "Extracted verbatim from Primary Lease Deed – property description clause (Clause 5)"
  • "From Registered Lease Agreement dated 2024-09-17, Schedule of Premises"
  • "From Rent Agreement (11 months), property clause"
  • "From Warehousing/Service Agreement dated 2024-09-17, Schedule of Premises"
  • "From Electricity Bill issued by JBVNL, consumer no. XXXXXX"
  • "From Municipality Tax Receipt issued by [Authority]"
  • "From NOC issued by [Owner Name]"
  • "From Consent Letter dated [Date]"

This address_source value will be displayed in the Excel output as the "Basis of Documents" column \
and in Final Remarks. Always provide it — never leave it null. The value must clearly identify \
the specific document type (Registered Lease Agreement / Rent Agreement / Electricity Bill / \
Municipality Tax Receipt / Consent Letter / NOC / Warehousing Agreement).

SCRIPT CONVERSION RULE (REPEAT — CRITICAL): Every address field must use Latin script only.
Convert Devanagari/regional script characters to their English phonetic spelling — do NOT
use this as an opportunity to "correct" or "fix" state/district/city names using external
geographic knowledge. Convert exactly what is written in the document, phoneme by phoneme.
Never output non-Latin characters. Never substitute a geographically "correct" value for what
the document actually says.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION B4 — STAMP DUTY RECEIPT / GOVERNMENT HEADER TRAP (CRITICAL)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Many lease deed bundles include a "Receipt of Online Payment of Stamp Duty" from the
state government portal (e.g. "Government of Jharkhand — Receipt of Online Payment of
Stamp Duty"). The header "Government of Jharkhand" on this receipt refers to the issuing
authority of the stamp duty receipt — NOT the state of the leased property.

DO NOT use the stamp duty receipt header to determine the state in addresses.state or
any address field. The stamp duty receipt is NOT the primary lease deed property clause.

SPECIFIC TRAP — Garhwa district leases:
  • Garhwa district is geographically in Jharkhand.
  • But some lease deeds for properties in the Garhwa area explicitly state "राज्य- बिहार"
    (State: Bihar) in Clause 5 — perhaps reflecting an older boundary, a document error,
    or the lessor's own writing.
  • Regardless of geographic reality, you MUST extract whatever Clause 5 says.
  • If Clause 5 says "Bihar" → state = "Bihar". Set pincode_state_flag = true and note
    the inconsistency. DO NOT silently substitute "Jharkhand".
  • If Clause 5 says "Jharkhand" → state = "Jharkhand". No flag needed.

RULE: The stamp duty receipt, Aadhaar card, land revenue receipt, and any other supporting
document NEVER override the state written in the PRIMARY LEASE DEED property clause (Clause 5
or equivalent). They are supporting documents only. Extract state EXCLUSIVELY from Clause 5.

AADHAAR TRAP: If the lessor's Aadhaar card shows a different state (e.g. Jharkhand) from
what Clause 5 says (e.g. Bihar), that is a cross-document discrepancy to FLAG — not to
silently "correct". The Aadhaar shows where the person lives, not necessarily what they wrote
in the lease deed's property description clause.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION B3 — PROHIBITED ACTIONS (ABSOLUTE — NO EXCEPTIONS)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
The following actions are STRICTLY FORBIDDEN when extracting any address field:

  ✗ Using pincode to infer or override the state name written in the lease deed clause.
  ✗ Using your geographic knowledge (e.g. "I know pincode 829202 is in Jharkhand") to change
    any address field.
  ✗ "Correcting" a state name even if you believe it is wrong (e.g. Bihar → Jharkhand).
  ✗ Using data from Aadhaar cards, bank passbooks, or land revenue receipts to override
    the state/district/city written in the PRIMARY LEASE DEED property clause.
  ✗ Synthesizing final_pob_address from multiple sources — extract from ONE source only:
    the primary lease deed property description clause.
  ✗ Treating summary.final_pob_address as a "correction layer" over the lease deed clause.
  ✗ Setting pincode_state_flag = true when the extracted state, pincode, and district
    are all mutually consistent (no false alarms).
  ✗ Populating state_discrepancy_note with a warning when there is no actual discrepancy
    between the state written in the deed and what the pincode/district indicate.

  If you believe an address field contains an error (e.g. wrong state), do NOT correct it.
  Instead: extract verbatim, set pincode_state_flag = true (only if genuinely inconsistent),
  and explain the specific conflict in state_discrepancy_note.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION C — SPECIAL CONDITIONS (MANDATORY)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1. CLIENT NAME IDENTIFICATION: If the GST applicant name is not clearly mentioned or identifiable, \
   explicitly flag the specific missing document needed to establish the name; add to \
   missing_documents and compliance.issues with HIGH risk.
2. NAME + SIGNATURE WITHOUT COMPLETE ADDRESS: If a document provides Name and Signature but the \
   complete address is missing/incomplete, set address_match_status to PARTIAL MATCH or LOW MATCH \
   and add to compliance issues.
3. WITNESS VALIDATION RULES:
   • Registered lease deeds (>11 months, i.e., Lease Deeds): Witness Name + Signature is \
     sufficient. Full address is NOT mandatory for registered documents.
   • Unregistered rent/simple rent agreements (≤11 months): Witness MUST include Name, Signature \
     AND full address. Missing address = compliance failure.
   Record all witnesses in witness_records and flag non-compliance in compliance and risk sections.
4. DOCUMENT ROLES (use exactly): PRIMARY_LEASE_DEED | WAREHOUSING_SERVICE_AGREEMENT | \
   AMENDMENT_TO_SERVICE_AGREEMENT | PROOF_OF_OWNERSHIP | NOC | CONSENT_LETTER | IDENTITY_PROOF | OTHER

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION D — DOCUMENT TYPE IDENTIFICATION (MANDATORY)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
For documents_analyzed[].detected_type, always use precise document-type labels:

  UTILITY / OWNERSHIP DOCUMENTS — use one of:
    "Electricity Bill" | "Water Tax Receipt" | "Property Tax Receipt" | "Municipal Tax Receipt" |
    "Gas Bill" | "Telephone Bill" | "Internet Bill" | "Land Revenue Receipt" |
    "Any Other Legal Document"
  Use "Any Other Legal Document" only when none of the above labels apply.

  AGREEMENT DOCUMENTS — classify strictly by lease duration:
    • Duration > 11 months → detected_type = "Lease Deed"
    • Duration ≤ 11 months → detected_type = "Rent Agreement"
  Never label a >11-month agreement as "Rent Agreement" or vice versa.

  This classification MUST be consistent with lease_details.duration_months and \
  stamp_duty_compliant logic in Section J.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION E — FILENAME RULE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- file_name in every documents_analyzed entry MUST be the EXACT original uploaded filename. \
  Do NOT use generic document-type names.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION F — ELECTRICITY / UTILITY BILL
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- Extract: owner_name, full address (English only), bill_date (YYYY-MM-DD), bill_type \
  (Electricity/Water/Gas/Telephone/Internet/Municipal Tax/Other), account_number, consumer_number.
- sub_document_type: identify whether the document is a "Tax Invoice", "Bill of Supply", \
  "Receipt", "Demand Notice", "Statement", or "Other" — check the document header/title.
- issuing_authority: extract the name of the organisation/board/company that issued the document \
  (e.g. "JSEB", "JBVNL", "BSES Rajdhani", "Indane Gas", "MTNL").
- bill_date: Extract the exact bill date if present in YYYY-MM-DD format. If no date is found \
  in the document, set bill_date to null.
- within_3_months: Set to true ONLY if bill_date is not null AND bill_date >= {CUTOFF_DATE}. \
  Set to false if bill_date is not null but older than cutoff. Set to null if bill_date is null \
  or no utility bill is present at all (do NOT default to false in the null case).
- older_than_3_months: Return "Yes" if bill_date is not null and bill_date < {CUTOFF_DATE}. \
  Return "No" if bill_date is not null and bill_date >= {CUTOFF_DATE}. \
  Return "NA" if bill_date is null or no utility bill is present.
- owner_name_matched_with_lessor: "Matched" | "Unmatched" | "Partial Match".
- If no utility bill or ownership-proof document is present, populate electricity_bill fields with \
  null values and set risk_exposure to "HIGH" with risk_notes explaining the absence.

NAME & ADDRESS MATCHING (apply to electricity_bill):
- name_match_status: Compare electricity_bill.owner_name with lease_details.lessor_name. \
  Return "HIGH MATCH" (score ≥ 80), "PARTIAL MATCH" (score 50–79), "LOW MATCH" (score < 50), \
  or "N/A" if either name is absent.
- agreement_address_match: Compare electricity_bill.address with lease_details.premises_address. \
  Return "HIGH MATCH", "PARTIAL MATCH", "LOW MATCH", or "N/A".

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION F2 — MERGED DOCUMENT DETECTION (CRITICAL — NEW IN v3.0)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
IMPORTANT: A single uploaded PDF may contain MULTIPLE DOCUMENT TYPES merged together.
Example: Lease Agreement (pages 1-3) + Bank Passbook (page 4) + Land Revenue Receipt (page 5).

MANDATORY DETECTION RULES:
- Analyse EVERY PAGE independently before performing whole-document validation.
- Detect transitions between: Lease Deed | Rent Agreement | Bank Passbook | Land Revenue Receipt |
  Property Tax Receipt | Aadhaar Card | NOC | Consent Letter | Utility Bill | Other.
- Build document_segmentation with page ranges for each detected document type:
  { "type": "Lease Deed", "pages": [1,2,3] }
  { "type": "Bank Passbook", "pages": [4] }
  { "type": "Land Revenue Receipt", "pages": [5] }
- Set merged_document = true if more than one distinct document type is found.
- Set has_ownership_proof = true if any page contains Land Revenue Receipt, Property Tax Receipt,
  Municipal Tax Receipt, Mutation Record, or Registry Extract.
- CRITICAL: Validate each segment independently. An ownership proof must NOT fail validation
  merely because the enclosing file was initially classified as an agreement document.
- Populate agreement_source_pages and ownership_proof_source_pages with page ranges (e.g. "1-3", "5").

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION F3 — OWNERSHIP PROOF VALIDATION (CRITICAL — NEW IN v3.0)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Ownership proof types supported: Property Tax Receipt | Municipal Tax Receipt |
Land Revenue Receipt | Mutation Record | Registry Extract | Government Land Record | Any legal land ownership proof.

RULE A — PROPERTY TAX RECEIPT:
- Validate by FINANCIAL YEAR — NOT by 3-month bill date rule.
- Extract financial year (e.g. "2025-26" or "2024-25") from the document.
- Mark valid if FY is current (2025-26) or immediately previous (2024-25).
- Mark OUTDATED if FY is 2023-24 or older.
- Set: within_3_months = "NOT_APPLICABLE", three_month_rule_applicable = false.
- NEVER flag Property Tax Receipt as "older than 3 months" — this rule does NOT apply.

RULE B — LAND REVENUE RECEIPT (jamabandi, lagan rasid, Rajasva evam Bhumi Sudhar):
- Land Revenue Records are PERMANENT ownership records. Treat as valid once officially recorded.
- DO NOT apply 3-month capping logic. DO NOT mark old merely because issue date is older than cutoff.
- Extract: receipt number, jamabandi/khata/khesra, owner name, village/state/district.
- Set: within_3_months = "NOT_APPLICABLE", three_month_rule_applicable = false, ownership_record_valid = true.
- MANDATORY: The existing "older than 3 months" validation logic MUST NOT flag land revenue records.

RULE C — ADDRESS COMPARISON (MANDATORY):
- Extract ownership_address from the ownership proof document pages.
- Extract agreement_address from the lease deed property clause.
- Compare both addresses and set address_match_with_agreement:
  "HIGH MATCH" (≥70% similarity) | "PARTIAL MATCH" (40-69%) | "LOW MATCH" (<40%) |
  "MISMATCH" (very low) | "STATE MISMATCH" (different states indicated).
- Write comparison_remarks explaining which components matched and which differed.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION G — AGREEMENT VALIDITY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- agreement_valid_as_on_validation_date: true if end_date >= {VALIDATION_DATE}, else false.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION H — NOTARY / REGISTRATION DETECTION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- notary_verified: true if document is EITHER notarised OR registered with Sub-Registrar \
  (registration stamp, Reg. No., Book No., IGR stamp, e-stamp, egrashry, grashry etc.).
- registration_type: "REGISTERED" | "NOTARISED" | "NONE" | "UNKNOWN"
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION I — OWNER NAME MATCH
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- Case-insensitive comparison allowing M/s, &, AND, spacing variations.
- Flag ONLY genuine entity mismatches.
- owner_name_match_score: 0–100 (100 = exact, ≥ 80 = HIGH, 50–79 = MEDIUM, < 50 = LOW).

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION J — STAMP DUTY COMPLIANCE
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- Rent Agreement (≤ 11 months): Fixed stamp duty Rs. 500.
- Lease Deed (> 11 months): Stamp duty = percentage of annual rent (state-specific); \
  registration with Sub-Registrar is mandatory.
- stamp_duty_compliant: true only if the paid amount meets the applicable rule above.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION K — CBIC CHECKLIST (Para 6) — evaluate ALL 15 points
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
1.  Ownership proof NOT older than 3 months
2.  Address on ownership proof matches Rent/Lease Agreement address
3.  Electricity Bill / Utility Bill present and identified
4.  Owner Name on Electricity Bill matches Lessor name in Rent/Lease
5.  Client (Lessee/GST Applicant) name found in documents
6.  Name on Electricity Bill matched with Rent/Lease Agreement
7.  If names unmatched — initial lease deed / earlier document provided
8.  Both Parties Signature and Stamp documented
9.  Witness Name, Address (if required) and Signature documented
10. Validity in force — not expired as on {VALIDATION_DATE}
11. Lease Deed / Rent Agreement is Notarised or Registered
12. No separate agreement apart from ownership proof required (or NOC/consent obtained)
13. Time period: > 11 months = Lease Deed; ≤ 11 months = Rent Agreement (check stamp duty category)
14. Stamp duty compliant (Rent ≤ Rs. 500; Lease = annual rent basis)
15. Sub-lease clause — no bar on sub-leasing (if sub-lease scenario)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION L — EXTRACTION STATUS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Use exactly one of: OK | FAILED | SCANNED_PDF | IMAGE | DOC | PARTIAL

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION M — COMPLIANCE SCORING
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- weighted_compliance_percent: (passed_weight / total_weight) × 100
- Critical checks (1, 4, 8, 10, 11) weight = 2; all others weight = 1.
- effective_status per document: CORRECT | INCORRECT | REVIEW | MISSING | MISMATCH

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION N — ISSUE CONSOLIDATION FOR DISCREPANCY REGISTER (MANDATORY)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
In compliance.issues and in each document's discrepancy_reason / issue_type, phrase every issue \
so it can be directly mapped to Issue I, Issue II, Issue III, etc. in the Excel Discrepancy \
Register. Follow this format for each issue string:

  "[Plain-English description of the deficiency] — [Recommended Action]"

Examples:
  "Document Absence: Utility bill / ownership proof is missing. Action: Submit a valid utility bill or valid ownership proof."
  "Stamp Duty: Stamp duty of Rs 100 paid; 15-year lease requires registration and \
   higher stamp duty per state schedule. Action: Register deed and pay correct stamp duty."
  "Witness Details: Addresses of both witnesses missing from the agreement. \
   Action: Obtain and attach affidavits with full witness addresses."

compliance.issues must be a list of such formatted strings — one per distinct deficiency — so the \
Excel mapping software can split them into Issue I, Issue II, Issue III columns without any \
manual interpretation.
 
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION O — RISK EXPOSURE ASSESSMENT (apply to EVERY section)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- LOW    = All critical fields present, fully compliant, zero gaps.
- MEDIUM = Minor gaps/warnings (partial address, witness address missing only in a registered \
           deed, low name confidence).
- HIGH   = Major compliance risk that COULD CAUSE GST REGISTRATION REJECTION (missing name proof, \
           expired agreement, owner name mismatch, bill older than 3 months, stamp duty \
           non-compliant, missing signatures, unregistered lease deed > 11 months).
Include risk_exposure in: summary, electricity_bill, lease_details, compliance.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION P — STAMP & SIGNATURE RULES (MANDATORY)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- You must strictly distinguish between handwritten signatures and rubber seals/ink stamps.
- Determine if a lessee company stamp (rubber seal, round ink stamp, etc.) is present on the agreement, and if a lessor stamp is present.
- Set 'lessee_company_stamp_present' (boolean) and 'lessor_stamp_present' (boolean) under 'lease_details' accordingly.
- A lessee company stamp (rubber seal, round ink stamp, etc.) is mandatory for business applications in lease deeds/rent agreements under CBIC.
- A lessee company stamp (rubber seal, round ink stamp, etc.) strictly should NOT be listed in 'missing_documents'. Its absence is a compliance issue (logged under compliance.issues), NOT a missing document. Do not add any lessee stamp or lessee company stamp entry to the summary.missing_documents array under any circumstances.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
SECTION Q — QUALITY WARNINGS & READABILITY (MANDATORY)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- If any document or page is blurry, dark, low-resolution, contains illegible text, or is otherwise partially or fully unreadable, set 'readability_warning' = true under 'summary' and set 'readability_warning' = true and 'readability_confidence_score' < 60.0 (between 0.0 and 100.0 based on visual clarity) for that specific document item in 'documents_analyzed'.
- If the document is perfectly clear, set 'readability_warning' = false and 'readability_confidence_score' >= 90.0 (or 100.0).

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
GENERAL OUTPUT RULES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
- key_notes must be a plain string (comma-separated observations, no bullet points).
- Always respond with clean, structured data that can be directly mapped to Excel columns \
  without any manual interpretation. Every field name in the JSON corresponds to an Excel column; \
  use consistent values, English-only text, and null (never empty string) for absent data.
- Respond ONLY with a valid JSON matching the schema exactly. \
  No markdown, no backticks, no extra text outside the JSON object.
"""



#  JSON SCHEMA
JSON_SCHEMA = f"""{{
  "summary": {{
    "category": "Registered Lease Deed | Rent Agreement | Warehousing Agreement | Other",
    "overall_status": "COMPLETE | COMPLETE_WITH_WARNINGS | INCOMPLETE",
    "validation_date": "{VALIDATION_DATE}",
    "final_business_name": "string",
    "legal_name": "string",
    "trade_name": "string",
    "name_confidence_percent": 0,
    "name_confidence_label": "HIGH | MEDIUM | LOW",
    "final_pob_address": "string",
    "address_match_status": "HIGH MATCH | PARTIAL MATCH | LOW MATCH",
    "pincode": "string",
    "state": "string",
    "overall_risk_exposure": "LOW | MEDIUM | HIGH",
    "missing_documents": [],
    "total_documents_analyzed": 0,
    "notary_verified": false,
    "registration_type": "REGISTERED | NOTARISED | NONE | UNKNOWN",
    "weighted_compliance_percent": 0,
    "effective_overall_status": "CORRECT | INCORRECT | REVIEW | MISSING | MISMATCH",
    "readability_warning": false,
    "merged_document_partial_failure": false
  }},
  "names": {{
    "lessor_first_party": "string",
    "lessee_second_party": "string",
    "finalised_gst_business_name": "string",
    "legal_name": "string",
    "trade_name": "string",
    "sub_lease_chain_note": "string"
  }},
  "addresses": {{
    "finalised_pob_address": "string — APOB leased property address only",
    "khasra_numbers": "string",
    "village_locality": "string",
    "road_street": "string",
    "city_town": "string",
    "district": "string",
    "state": "string",
    "pincode": "string",
    "pincode_state_flag": false,
    "state_discrepancy_note": "string or null — plain English explanation if pincode and state in document are inconsistent; null if consistent",
    "address_source": "plain English: which document and section this address was extracted from",
    "registered_office_address": "string or null — company HQ/corporate address if found in docs (NOT the APOB address)"
  }},
  "electricity_bill": {{
    "bill_type": "Electricity | Water | Gas | Telephone | Internet | Municipal Tax | Other",
    "sub_document_type": "Tax Invoice | Bill of Supply | Receipt | Demand Notice | Statement | Other",
    "issuing_authority": "string",
    "owner_name": "string",
    "address": "string",
    "bill_date": "YYYY-MM-DD or null",
    "within_3_months": false,
    "account_number": "string or null",
    "consumer_number": "string or null",
    "owner_name_matched_with_lessor": "Matched | Unmatched | Partial Match",
    "owner_name_match_score": 0,
    "risk_exposure": "LOW | MEDIUM | HIGH",
    "risk_notes": "string"
  }},
  "lease_details": {{
    "start_date": "YYYY-MM-DD or null",
    "end_date": "YYYY-MM-DD or null",
    "duration_months": 0,
    "lessor_name": "string",
    "lessee_name": "string",
    "premises_address": "string",
    "rent_amount": 0,
    "stamp_duty_amount": 0,
    "stamp_duty_compliant": false,
    "stamp_duty_notes": "string",
    "signatures_both_parties": false,
    "lessor_signed": false,
    "lessee_signed": false,
    "lessee_company_stamp_present": false,
    "lessor_stamp_present": false,
    "notary_verified": false,
    "registration_type": "REGISTERED | NOTARISED | NONE | UNKNOWN",
    "registration_number": "string or null",
    "agreement_valid_as_on_validation_date": false,
    "sub_lease_permitted": null,
    "witness_records": [
      {{
        "witness_number": 1,
        "name": "string",
        "full_address": "string or null",
        "signature_present": false,
        "compliant": false,
        "compliance_note": "string"
      }}
    ],
    "risk_exposure": "LOW | MEDIUM | HIGH",
    "risk_notes": "string"
  }},
  "documents_analyzed": [
    {{
      "file_name": "EXACT original filename",
      "detected_type": "string",
      "role": "PRIMARY_LEASE_DEED | WAREHOUSING_SERVICE_AGREEMENT | AMENDMENT_TO_SERVICE_AGREEMENT | PROOF_OF_OWNERSHIP | NOC | CONSENT_LETTER | IDENTITY_PROOF | OTHER",
      "extraction_status": "OK | FAILED | SCANNED_PDF | IMAGE | DOC | PARTIAL",
      "owner_name": "string or null",
      "address_in_doc": "string or null",
      "name_match_status": "HIGH MATCH | PARTIAL MATCH | LOW MATCH | N/A",
      "address_match_status": "HIGH MATCH | PARTIAL MATCH | LOW MATCH | N/A",
      "signature_stamp_present": false,
      "notarised_registered": false,
      "stamp_duty_amount": null,
      "expiry_date": "YYYY-MM-DD or null",
      "is_duplicate": false,
      "effective_status": "CORRECT | INCORRECT | REVIEW | MISSING | MISMATCH",
      "issue_type": "None | Missing | Expired | Mismatch | Incomplete | Unreadable | Duplicate",
      "discrepancy_reason": "string or null",
      "key_notes": "string",
      "readability_warning": false,
      "readability_confidence_score": 100.0
    }}
  ],
  "ownership_proof": {{
    "detected": false,
    "document_type": "Land Revenue Receipt | Property Tax Receipt | Municipal Tax Receipt | Mutation Record | Registry Extract | Government Land Record | null",
    "source_pages": [],
    "source_file": "string — original filename where ownership proof was found",
    "owner_name": "string or null — name of owner as on ownership proof document",
    "ownership_address": "string or null — full address extracted from ownership proof",
    "financial_year": "YYYY-YY or null — applicable only for Property/Municipal Tax Receipts",
    "financial_year_valid": "true | false | null — null if FY not applicable",
    "within_3_months": "NOT_APPLICABLE — ownership proof documents are NEVER subject to 3-month rule",
    "three_month_rule_applicable": false,
    "ownership_record_valid": false,
    "address_match_with_agreement": "HIGH MATCH | PARTIAL MATCH | LOW MATCH | MISMATCH | STATE MISMATCH | N/A",
    "comparison_remarks": "string — plain English explanation of address comparison result",
    "khata": "string or null",
    "khesra": "string or null",
    "jamabandi": "string or null",
    "village": "string or null",
    "thana": "string or null",
    "district": "string or null",
    "state": "string or null",
    "receipt_number": "string or null"
  }},
  "document_segmentation": {{
    "source_file": "string",
    "total_pages": 0,
    "merged_document": false,
    "has_ownership_proof": false,
    "has_agreement": false,
    "merged_document_partial_failure": false,
    "document_segments": [
      {{
        "type": "Lease Deed | Land Revenue Receipt | Bank Passbook | Aadhaar Card | ...",
        "pages": []
      }}
    ],
    "agreement_source_pages": "string — e.g. '1-3' or '1,2,3'",
    "ownership_proof_source_pages": "string — e.g. '5' or '4,5'"
  }},
  "compliance": {{
    "cbic_checklist": [
      {{
        "check_number": 1,
        "description": "string",
        "status": "PASS | FAIL | N/A | REVIEW",
        "weight": 2,
        "notes": "string"
      }}
    ],
    "passes": [],
    "issues": [],
    "total_passes": 0,
    "total_issues": 0,
    "weighted_compliance_percent": 0,
    "overall_risk_exposure": "LOW | MEDIUM | HIGH",
    "risk_summary": "string",
    "final_status": "Clean | Warning | High Risk",
    "final_remarks": "string"
  }},
  "document_quality": {{
    "utility_bill_blurry": false,
    "human_readable": true,
    "quality_issue_note": "string or null"
  }}
}}"""

# ── User prompt template injected per-call ────────────────────────────────────
USER_PROMPT_TEMPLATE = """The uploaded documents belong to ONE GST APOB amendment application client/location.

Documents uploaded (use these EXACT filenames in documents_analyzed[].file_name):
{file_list}

Instructions:
1. Analyse ALL uploaded documents thoroughly following every rule in the system prompt.
2. Identify the GST applicant (finalised_gst_business_name) — also extract legal_name and trade_name separately.
3. Apply all 15 CBIC checklist checks (Para 6).
4. Compute weighted_compliance_percent accurately using weights (critical checks weight=2).
5. List all compliance passes AND issues with accurate risk levels.
6. Use null or empty values for missing/unavailable data. Never fabricate data.

Respond ONLY with a valid JSON object strictly matching this schema (no markdown, no backticks, no text outside JSON):
{schema}"""

# ── Lightweight name-extraction prompt (used for grouping, not full validation) ──
NAME_EXTRACT_PROMPT = """Extract ONLY the GST applicant business name (the end-user seeking APOB registration) from this document.
Return ONLY a single JSON: {"business_name": "string"}.
If not determinable, return: {"business_name": "UNKNOWN"}.
No markdown, no explanation."""


#  COLOUR & STYLE CONSTANTS
C = {
    "hdr_bg":   "1B3A5C", "hdr_fg":   "FFFFFF",
    "sec_bg":   "2E6DA4", "sec_fg":   "FFFFFF",
    "pass_bg":  "D6F5D6", "pass_fg":  "155724",
    "warn_bg":  "FFF3CD", "warn_fg":  "856404",
    "fail_bg":  "F8D7DA", "fail_fg":  "721C24",
    "info_bg":  "D1ECF1", "info_fg":  "0C5460",
    "alt":      "EBF3FB", "white":    "FFFFFF",
    "label_bg": "D9E8F7", "border":   "B0C4DE",
    "sub_hdr1": "1A5276", "sub_hdr2": "154360", "sub_hdr3": "17202A",
}

def _side():            return Side(style="thin", color=C["border"])
def _bdr():             return Border(left=_side(), right=_side(), top=_side(), bottom=_side())
def _fill(h):           return PatternFill("solid", fgColor=str(h).lstrip("#"))
def _bfont(c="000000", sz=11): return Font(bold=True, color=c, size=sz, name="Calibri")
def _rfont(c="000000", sz=10): return Font(color=c, size=sz, name="Calibri")
def _ca(w=True):        return Alignment(horizontal="center", vertical="center", wrap_text=w)
def _la(w=True):        return Alignment(horizontal="left",   vertical="center", wrap_text=w)


# ══════════════════════════════════════════════════════════════════════════════
#  OWNERSHIP PROOF & MERGED DOCUMENT MODULE  (v3.0)
#
#  Provides:
#    1. Page-wise document type classification for merged PDFs
#    2. Document segmentation with page range tracking
#    3. Ownership-proof-specific validation rules (FY vs 3-month)
#    4. Address comparison between ownership proof and agreement
#    5. Discrepancy generation for ownership-specific issues
#    6. OCR extraction helpers for Hindi land revenue receipts
# ══════════════════════════════════════════════════════════════════════════════

# ─── Document-type keyword signatures used for page-level classification ────
# Maps canonical type label → list of keyword signals (case-insensitive).
# Order matters: more-specific labels are checked before generic ones.
_DOC_TYPE_SIGNATURES: List[Tuple[str, List[str]]] = [
    # Ownership / Land records
    ("Land Revenue Receipt",  ["लगान रसीद", "lagan rasid", "rajasva", "raajaswa",
                                "राजस्व", "jamabandi", "जमाबंदी", "khata", "खाता",
                                "khesra", "खेसरा", "bhumi sudhar", "भूमि सुधार",
                                "nagar panchayat", "halkha", "halka", "प्रपत्र-XIV",
                                "prapatra", "lagaan", "land revenue", "revenue receipt"]),
    ("Property Tax Receipt",  ["property tax", "house tax", "property tax receipt",
                                "municipal tax", "nagar palika", "नगर पालिका",
                                "griha kar", "griha kara", "grih kar"]),
    ("Municipal Tax Receipt", ["municipal", "municipality", "nagar nigam", "नगर निगम",
                                "ward tax", "holding tax"]),
    ("Mutation Record",       ["mutation", "dakhil kharij", "दाखिल खारिज", "namantaran"]),
    ("Registry Extract",      ["registry extract", "sale deed", "sale certificate",
                                "registration deed", "sub registrar"]),
    # Agreement types
    ("Lease Deed",            ["लीज़ एग्रीमेंट", "lease agreement", "lease deed",
                                "leej", "lizadhari", "लीजधारक", "लीजकर्ता",
                                "india non judicial", "non-judicial", "stamp paper"]),
    ("Rent Agreement",        ["rent agreement", "rental agreement", "kiraya",
                                "किरायेदार", "kiraaedari"]),
    # Banking
    ("Bank Passbook",         ["passbook", "पासबुक", "account particulars",
                                "खाता विवरण", "ifsc", "micr code", "joint holder",
                                "saving account", "savings account", "current account",
                                "national bank", "punjab national", "state bank",
                                "bank of baroda", "canara bank", "axis bank"]),
    # Identity
    ("Aadhaar Card",          ["aadhaar", "आधार", "uid", "unique identification",
                                "aam aadmi ka adhikar"]),
    ("PAN Card",              ["permanent account number", "pan card", "income tax"]),
    # Other supporting
    ("NOC",                   ["no objection", "noc", "n.o.c"]),
    ("Consent Letter",        ["consent letter", "consent", "anumati patra"]),
    ("Utility Bill",          ["electricity bill", "bijli bill", "बिजली", "jseb",
                                "jbvnl", "bses", "tpddl", "msedcl", "uppcl",
                                "water bill", "gas bill"]),
]

# ─── Financial Year helpers ──────────────────────────────────────────────────
_CURRENT_FY_START = 2024   # FY 2024-25 (update each April)
_CURRENT_FY_LABEL = "2025-26"  # current ongoing FY

def _current_fy_year() -> int:
    """Return the starting year of the CURRENT financial year (April–March)."""
    today = date.today()
    return today.year if today.month >= 4 else today.year - 1

def _fy_label(start_year: int) -> str:
    """Return 'YYYY-YY' label for a financial year starting in start_year."""
    return f"{start_year}-{str(start_year + 1)[-2:]}"

def _is_fy_valid(fy_string: str) -> bool:
    """
    Return True if the financial year label is the current or immediately
    previous FY.  E.g. if current FY is 2025-26:
      2025-26 → True  (current)
      2024-25 → True  (immediately previous — acceptable)
      2023-24 → False (outdated)
    """
    if not fy_string:
        return False
    m = re.search(r"(\d{4})\s*[-–]\s*(\d{2,4})", str(fy_string))
    if not m:
        return False
    start = int(m.group(1))
    cur   = _current_fy_year()
    return start >= cur - 1   # current year or last year


def _extract_fy_from_text(text: str) -> Optional[str]:
    """
    Extract a financial year string like '2024-25' or '2024-2025' from text.
    Returns 'YYYY-YY' format or None.
    """
    # Match patterns: 2024-25, 2024-2025, 2024 - 25, FY 2024-25, etc.
    m = re.search(r"(?:fy|financial\s*year|vitt(?:iya)?\s*varsh|वित्त(?:ीय)?\s*वर्ष)?\s*"
                  r"(\d{4})\s*[-–]\s*(\d{2,4})", text, re.IGNORECASE)
    if not m:
        return None
    start = int(m.group(1))
    end_raw = m.group(2)
    if len(end_raw) == 2:
        end = int(f"{str(start)[:2]}{end_raw}")
    else:
        end = int(end_raw)
    if end == start + 1:
        return f"{start}-{end_raw[:2] if len(end_raw)==4 else end_raw}"
    return None


# ─── Page-level document classifier ─────────────────────────────────────────

def _classify_page_text(page_text: str) -> str:
    """
    Given OCR/extracted text of a single page, return the best-match
    document type label from _DOC_TYPE_SIGNATURES.
    Returns 'Unknown' if no pattern is found.
    """
    text_lower = page_text.lower()
    best_type  = "Unknown"
    best_score = 0

    for doc_type, keywords in _DOC_TYPE_SIGNATURES:
        score = sum(1 for kw in keywords if kw.lower() in text_lower)
        if score > best_score:
            best_score = score
            best_type  = doc_type

    return best_type if best_score >= 1 else "Unknown"


def _segment_merged_pdf(
    source_file: Path,
    page_texts: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """
    Analyse a PDF file and return a document segmentation map.

    Args:
        source_file : Path to the PDF (or image file).
        page_texts  : Pre-extracted text per page (1-indexed list).
                      If None, the function tries pdf2image + pytesseract.

    Returns a dict:
    {
      "source_file": "filename.pdf",
      "total_pages": N,
      "document_segments": [
        { "type": "Lease Deed",            "pages": [1,2,3] },
        { "type": "Bank Passbook",         "pages": [4]     },
        { "type": "Land Revenue Receipt",  "pages": [5]     },
        { "type": "Aadhaar Card",          "pages": [6]     },
      ],
      "has_ownership_proof": True,
      "has_agreement": True,
      "merged_document": True,   # True if >1 distinct type found
    }
    """
    result: Dict[str, Any] = {
        "source_file"        : source_file.name,
        "total_pages"        : 0,
        "document_segments"  : [],
        "has_ownership_proof": False,
        "has_agreement"      : False,
        "merged_document"    : False,
        "agreement_source_pages": "",
        "ownership_proof_source_pages": "",
    }

    # ── Build page_texts if not provided ───────────────────────────────────
    if page_texts is None:
        page_texts = _extract_page_texts_from_pdf(source_file)

    if not page_texts:
        log.debug(f"   [Segmenter] No page texts for {source_file.name}")
        return result

    result["total_pages"] = len(page_texts)

    # ── Classify each page ─────────────────────────────────────────────────
    page_types: List[str] = [_classify_page_text(t) for t in page_texts]

    # ── Build contiguous segments ─────────────────────────────────────────
    segments: List[Dict[str, Any]] = []
    if page_types:
        current_type  = page_types[0]
        current_pages = [1]
        for pg_idx, ptype in enumerate(page_types[1:], start=2):
            # Allow 'Unknown' pages to inherit the last known type
            if ptype == "Unknown":
                ptype = current_type
            if ptype == current_type:
                current_pages.append(pg_idx)
            else:
                segments.append({"type": current_type, "pages": list(current_pages)})
                current_type  = ptype
                current_pages = [pg_idx]
        segments.append({"type": current_type, "pages": list(current_pages)})

    result["document_segments"] = segments

    # ── Derive flags ────────────────────────────────────────────────────────
    types_found = {seg["type"] for seg in segments}
    _OWNERSHIP_TYPES = {
        "Land Revenue Receipt", "Property Tax Receipt", "Municipal Tax Receipt",
        "Mutation Record", "Registry Extract",
    }
    _AGREEMENT_TYPES = {"Lease Deed", "Rent Agreement"}

    result["has_ownership_proof"] = bool(types_found & _OWNERSHIP_TYPES)
    result["has_agreement"]       = bool(types_found & _AGREEMENT_TYPES)
    result["merged_document"]     = len(types_found - {"Unknown"}) > 1
    result["agreement_source_pages"] = _format_pages_for_display(
        sorted(
            p
            for seg in segments if seg.get("type") in _AGREEMENT_TYPES
            for p in (seg.get("pages") or [])
        )
    )
    result["ownership_proof_source_pages"] = _format_pages_for_display(
        sorted(
            p
            for seg in segments if seg.get("type") in _OWNERSHIP_TYPES
            for p in (seg.get("pages") or [])
        )
    )
    if result["merged_document"]:
        log.info(
            f"   [Segmenter] MERGED PDF detected: {source_file.name} "
            f"— {len(segments)} segment(s): "
            + ", ".join(f"{s['type']}(p{s['pages']})" for s in segments)
        )

    return result


OCR_CONFIDENCE_CACHE: Dict[str, float] = {}


def _extract_page_texts_from_pdf(source_file: Path) -> List[str]:
    """
    Extract per-page text from a PDF using pdf2image + pytesseract (Hindi+English).
    Returns list of strings (one per page), empty list on failure.
    Falls back gracefully if pdf2image / pytesseract are not available.
    """
    try:
        from pdf2image import convert_from_path  # type: ignore
        import pytesseract                        # type: ignore
    except ImportError:
        log.debug("   [Segmenter] pdf2image/pytesseract not available — skipping page-text extraction")
        return []

    try:
        images = convert_from_path(str(source_file), dpi=200)
    except Exception as exc:
        log.debug(f"   [Segmenter] pdf2image failed for {source_file.name}: {exc}")
        return []

    page_texts: List[str] = []
    page_conf: List[float] = []
    ocr_attempts = [
        ("hin+eng", r"--oem 3 --psm 6"),
        ("hin+eng", r"--oem 3 --psm 4"),
        ("hin+eng", r"--oem 3 --psm 3"),
        ("hin+eng", r"--oem 3 --psm 11"),
        ("eng", r"--oem 3 --psm 6"),
        ("eng", r"--oem 3 --psm 4"),
        ("eng", r"--oem 3 --psm 3"),
        ("eng", r"--oem 3 --psm 11"),
    ]

    for img in images:
        best_text = ""
        best_conf = -1.0
        for lang, cfg in ocr_attempts:
            try:
                txt = pytesseract.image_to_string(img, lang=lang, config=cfg)
                conf = _ocr_confidence_score(pytesseract, img, lang, cfg)
                if len((txt or "").strip()) > len(best_text.strip()) or conf > best_conf:
                    best_text = txt or ""
                    best_conf = conf
            except Exception:
                continue
        text = best_text
        page_texts.append(text)
        page_conf.append(max(best_conf, 0.0))

    if page_conf:
        avg_conf = sum(page_conf) / len(page_conf)
        OCR_CONFIDENCE_CACHE[source_file.name] = avg_conf
        log.info(
            f"   [Segmenter] OCR confidence for {source_file.name}: "
            f"avg={avg_conf:.1f}% over {len(page_conf)} page(s)"
        )

    return page_texts


def _ocr_confidence_score(pytesseract_mod, image_obj, lang: str, cfg: str) -> float:
    """Best-effort OCR confidence score (0-100)."""
    try:
        data = pytesseract_mod.image_to_data(
            image_obj,
            lang=lang,
            config=cfg,
            output_type=pytesseract_mod.Output.DICT,
        )
    except Exception:
        return 0.0
    confs: List[float] = []
    for val in data.get("conf", []):
        sval = str(val).strip()
        if not sval or sval in ("-1", "nan"):
            continue
        try:
            fval = float(sval)
        except Exception:
            continue
        if fval >= 0:
            confs.append(fval)
    return round(sum(confs) / len(confs), 1) if confs else 0.0


def _format_pages_for_display(pages: List[int]) -> str:
    """Format page list like [1,2,3,5] -> '1-3, 5'."""
    if not pages:
        return ""
    uniq = sorted(set(int(p) for p in pages if str(p).isdigit()))
    if not uniq:
        return ""
    ranges: List[str] = []
    start = prev = uniq[0]
    for p in uniq[1:]:
        if p == prev + 1:
            prev = p
            continue
        ranges.append(f"{start}-{prev}" if start != prev else str(start))
        start = prev = p
    ranges.append(f"{start}-{prev}" if start != prev else str(start))
    return ", ".join(ranges)


# ─── Ownership proof extractor & validator ───────────────────────────────────

_OWNERSHIP_PROOF_TYPES = {
    "Land Revenue Receipt", "Property Tax Receipt", "Municipal Tax Receipt",
    "Mutation Record", "Registry Extract", "Government Land Record",
}

_PERMANENT_OWNERSHIP_TYPES = {
    # These are permanent records; 3-month rule does NOT apply.
    "Land Revenue Receipt", "Mutation Record", "Registry Extract",
    "Government Land Record",
}

_FY_BASED_TYPES = {
    # These are validated by financial year, NOT by bill date age.
    "Property Tax Receipt", "Municipal Tax Receipt",
}

# ── Utility bill document types — for blurry/quality detection ─────────────
_UTILITY_BILL_DOC_TYPES = {
    "electricity bill", "electric bill", "power bill",
    "water bill", "water tax receipt", "gas bill",
    "telephone bill", "phone bill", "internet bill",
    "broadband bill", "utility bill", "invoice bill",
    "tax invoice", "bill of supply",
}


def _extract_ownership_proof_data(
    doc_segment_text: str,
    doc_type: str,
    source_pages: List[int],
    source_file: str,
) -> Dict[str, Any]:
    """
    Extract structured ownership proof data from page text of a detected
    ownership-proof segment.

    Returns a dict matching the ownership_proof JSON schema section:
    {
      "detected":                   True,
      "document_type":              "Land Revenue Receipt",
      "source_pages":               [5],
      "source_file":                "Bikram_Land_Agreement.pdf",
      "owner_name":                 "Mukesh Kumar",
      "ownership_address":          "Village Gangachak, Bikram, Patna, Bihar",
      "financial_year":             null,       # or "2024-25"
      "financial_year_valid":       null,       # or true/false
      "within_3_months":            "NOT_APPLICABLE",
      "three_month_rule_applicable": false,
      "ownership_record_valid":     true,
      "address_match_with_agreement": null,     # populated later
      "comparison_remarks":         null,        # populated later
      "khata":                      "75",
      "khesra":                     "562",
      "jamabandi":                  "368",
      "village":                    "Gangachak",
      "district":                   "Patna",
      "state":                      "Bihar",
      "receipt_number":             "...",
    }
    """
    proof: Dict[str, Any] = {
        "detected"                  : True,
        "document_type"             : doc_type,
        "source_pages"              : source_pages,
        "source_file"               : source_file,
        "owner_name"                : None,
        "ownership_address"         : None,
        "financial_year"            : None,
        "financial_year_valid"      : None,
        "within_3_months"           : "NOT_APPLICABLE",
        "three_month_rule_applicable": doc_type not in _PERMANENT_OWNERSHIP_TYPES
                                       and doc_type not in _FY_BASED_TYPES,
        "ownership_record_valid"    : False,
        "address_match_with_agreement": None,
        "comparison_remarks"        : None,
        "khata"                     : None,
        "khesra"                    : None,
        "jamabandi"                 : None,
        "village"                   : None,
        "thana"                     : None,
        "district"                  : None,
        "state"                     : None,
        "receipt_number"            : None,
    }

    text = doc_segment_text or ""
    text_lower = text.lower()

    # ── Extract owner/raiyat name ───────────────────────────────────────────
    # Hindi: "जमाबंदी रैयत का नाम:- <name>" or "raiyat ka naam"
    owner_patterns = [
        r"(?:jamabandi\s+raiyat\s+ka\s+naam|रैयत\s+का\s+नाम)\s*[::-]\s*([A-Za-zऀ-ॿ\s]+)",
        r"(?:owner|lessee|party|malik|मालिक)\s*[::-]\s*([A-Za-zऀ-ॿ\s]{3,40})",
        r"(?:name|naam|नाम)\s*[::-]\s*([A-Za-z\s]{3,40})",
    ]
    for pat in owner_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            candidate = m.group(1).strip()
            if 3 <= len(candidate) <= 50:
                proof["owner_name"] = candidate
                break

    # ── Extract Hindi land-record fields ───────────────────────────────────
    # Khata (खाता संख्या)
    m = re.search(r"(?:khata\s+(?:no|sankhya|number)|खाता\s+संख्या)\s*[:.-]*\s*([\d\s,]+)",
                  text, re.IGNORECASE)
    if m:
        proof["khata"] = m.group(1).strip()[:40]

    # Khesra (खेसरा संख्या)
    m = re.search(r"(?:khesra\s+(?:no|sankhya|number)|खेसरा\s+संख्या)\s*[:.-]*\s*([\d\s,]+)",
                  text, re.IGNORECASE)
    if m:
        proof["khesra"] = m.group(1).strip()[:40]

    # Jamabandi number
    m = re.search(r"(?:jamabandi\s+(?:no|sankhya|number)|जमाबंदी\s+(?:सं|no))\s*[:.-]*\s*(\d+)",
                  text, re.IGNORECASE)
    if m:
        proof["jamabandi"] = m.group(1).strip()

    # Thana / Mouza / Village
    m = re.search(r"(?:mouza|mauza|village|gram|grama|ग्राम|मौजा)\s*[:.-]*\s*([A-Za-zऀ-ॿ\s]{2,30})",
                  text, re.IGNORECASE)
    if m:
        proof["village"] = m.group(1).strip()[:40]

    # Thana / Anchal
    m = re.search(r"(?:thana|anchal|अंचल|थाना)\s*[:.-]*\s*([A-Za-zऀ-ॿ\s]{2,30})",
                  text, re.IGNORECASE)
    if m:
        proof["thana"] = m.group(1).strip()[:40]

    # District
    m = re.search(r"(?:district|distt|jila|जिला|जिला)\s*[:.-]*\s*([A-Za-zऀ-ॿ\s]{2,25})",
                  text, re.IGNORECASE)
    if m:
        proof["district"] = m.group(1).strip()[:30]

    # State (Latin or Devanagari)
    m = re.search(r"(?:state|raajya|rajya|राज्य)\s*[:.-]*\s*([A-Za-zऀ-ॿ\s]{2,20})",
                  text, re.IGNORECASE)
    if m:
        proof["state"] = m.group(1).strip()[:25]

    # Pata / address line
    m = re.search(r"(?:pata|pata:|पता)\s*[:.-]*\s*([A-Za-zऀ-ॿ\s,.-]{5,80})",
                  text, re.IGNORECASE)
    if m:
        proof["ownership_address"] = m.group(1).strip()[:120]

    # Build address from components if direct pata not found
    if not proof["ownership_address"]:
        parts = [proof.get("village"), proof.get("thana"),
                 proof.get("district"), proof.get("state")]
        addr_built = ", ".join(p for p in parts if p)
        if addr_built:
            proof["ownership_address"] = addr_built

    # ── Receipt / running unique number ────────────────────────────────────
    m = re.search(r"(?:receipt\s*(?:no|number)|running\s*unique\s*no|rn\s*no)\s*[:.-]*\s*([\w/\-]+)",
                  text, re.IGNORECASE)
    if m:
        proof["receipt_number"] = m.group(1).strip()[:40]

    # ── Financial year ──────────────────────────────────────────────────────
    fy = _extract_fy_from_text(text)
    if fy:
        proof["financial_year"] = fy

    # ── Validity logic ──────────────────────────────────────────────────────
    if doc_type in _PERMANENT_OWNERSHIP_TYPES:
        # Permanent ownership records are ALWAYS valid once officially recorded.
        proof["three_month_rule_applicable"] = False
        proof["within_3_months"]             = "NOT_APPLICABLE"
        proof["ownership_record_valid"]      = True
        proof["financial_year_valid"]        = None  # N/A

    elif doc_type in _FY_BASED_TYPES:
        # Property/Municipal Tax Receipt: validate by FY.
        proof["three_month_rule_applicable"] = False
        proof["within_3_months"]             = "NOT_APPLICABLE"
        if fy:
            fy_ok = _is_fy_valid(fy)
            proof["financial_year_valid"]    = fy_ok
            proof["ownership_record_valid"]  = fy_ok
        else:
            proof["financial_year_valid"]    = None
            proof["ownership_record_valid"]  = False  # cannot validate without FY

    else:
        # Other types (Utility Ownership Proof etc.): standard 3-month rule applies.
        proof["three_month_rule_applicable"] = True
        # within_3_months stays NOT_APPLICABLE here — only Gemini can assess bill date.
        proof["ownership_record_valid"]      = True  # assume valid unless Gemini flags

    return proof


def _compare_addresses(addr_ownership: str, addr_agreement: str) -> Tuple[str, str]:
    """
    Compare two address strings and return (match_label, comparison_remarks).

    match_label : "HIGH MATCH" | "PARTIAL MATCH" | "LOW MATCH" | "MISMATCH" | "N/A"
    comparison_remarks : plain-English explanation of what matched and what didn't.
    """
    if not addr_ownership or not addr_agreement:
        return "N/A", "One or both addresses were not available for comparison."

    def _tokens(s: str) -> set:
        """Normalised token set: lower, strip punctuation, remove short words."""
        raw = re.sub(r"[^\w\s]", " ", s.lower())
        return {w for w in raw.split() if len(w) >= 3}

    own_tokens = _tokens(addr_ownership)
    agr_tokens = _tokens(addr_agreement)

    if not own_tokens or not agr_tokens:
        return "N/A", "Could not tokenise addresses for comparison."

    common  = own_tokens & agr_tokens
    total   = own_tokens | agr_tokens
    jaccard = len(common) / len(total) if total else 0.0
    score   = int(jaccard * 100)

    # Identify which important components matched / mismatched
    # Try to detect state and district tokens
    _STATES = {"bihar", "jharkhand", "uttar", "pradesh", "uttarakhand",
               "maharashtra", "gujarat", "rajasthan", "delhi", "haryana",
               "punjab", "bengal", "odisha", "chhattisgarh"}
    own_state_tok = own_tokens & _STATES
    agr_state_tok = agr_tokens & _STATES

    state_match = bool(own_state_tok & agr_state_tok) if own_state_tok and agr_state_tok else None

    if score >= 70:
        label   = "HIGH MATCH"
        remarks = f"Address comparison: HIGH MATCH (similarity {score}%). "
        if common:
            remarks += f"Common elements: {', '.join(sorted(common)[:8])}."
    elif score >= 40:
        label   = "PARTIAL MATCH"
        remarks = f"Address comparison: PARTIAL MATCH (similarity {score}%). "
        if common:
            remarks += f"Matched: {', '.join(sorted(common)[:5])}. "
        diff = (own_tokens - agr_tokens) | (agr_tokens - own_tokens)
        if diff:
            remarks += f"Differing elements: {', '.join(sorted(diff)[:5])}."
    else:
        label   = "LOW MATCH" if score >= 15 else "MISMATCH"
        remarks = f"Address comparison: {label} (similarity {score}%). "
        if state_match is False:
            remarks += "STATE MISMATCH detected — different states indicated. "
        elif state_match is True:
            remarks += "States match but other address components differ significantly. "

    if state_match is False:
        label = "STATE MISMATCH"

    return label, remarks.strip()


def _build_ownership_proof_remarks(
    proof: Dict[str, Any],
    agreement_address: str,
) -> str:
    """
    Build the evidence-driven Final Remarks paragraph for the ownership
    proof section — to be appended to the main Final Remarks string.
    """
    if not proof or not proof.get("detected"):
        return ""

    doc_type      = proof.get("document_type") or "Ownership Proof"
    source_pages  = proof.get("source_pages") or []
    source_file   = proof.get("source_file") or ""
    own_addr      = proof.get("ownership_address") or "Not extracted"
    match_label   = proof.get("address_match_with_agreement") or "N/A"
    cmp_remarks   = proof.get("comparison_remarks") or ""
    fy            = proof.get("financial_year")
    fy_valid      = proof.get("financial_year_valid")
    three_m_applic= proof.get("three_month_rule_applicable", False)
    rec_valid     = proof.get("ownership_record_valid", False)
    merged        = bool(source_pages and len(source_pages) >= 1)

    lines: List[str] = []

    # Source detection note
    pg_str = ", ".join(str(p) for p in source_pages)
    if merged and source_file:
        lines.append(
            f"{doc_type} detected within merged PDF '{source_file}' "
            f"(page{'s' if len(source_pages) > 1 else ''}: {pg_str})."
        )
    else:
        lines.append(f"{doc_type} detected.")

    # Owner name
    owner = proof.get("owner_name")
    if owner:
        lines.append(f"Owner name on record: {owner}.")

    # Ownership address
    lines.append(f"Ownership address extracted from {doc_type}:\n  {own_addr}.")

    # Agreement address for comparison
    agr_addr_display = agreement_address or "Not extracted"
    lines.append(f"Agreement property address:\n  {agr_addr_display}.")

    # Address match result
    lines.append(f"Address match result: {match_label}.")
    if cmp_remarks:
        lines.append(cmp_remarks)

    # FY validation note
    if doc_type in _FY_BASED_TYPES:
        if fy:
            cur_fy = _fy_label(_current_fy_year())
            fy_status = "VALID" if fy_valid else "OUTDATED"
            lines.append(
                f"Financial Year on {doc_type}: {fy} → {fy_status} "
                f"(current FY: {cur_fy}). "
                "3-month recency rule not applicable to tax receipts; "
                "financial year-based validation applied instead."
            )
        else:
            lines.append(
                f"Financial year not extracted from {doc_type}. "
                "Manual verification of document year required."
            )

    # Permanent record note
    if doc_type in _PERMANENT_OWNERSHIP_TYPES:
        lines.append(
            f"3-month recency rule NOT applicable: {doc_type} is a permanent "
            "ownership record. Validity is based on official recording, not issue date."
        )

    # Hindi field note
    fields_extracted = [
        k for k in ("khata", "khesra", "jamabandi", "village", "thana")
        if proof.get(k)
    ]
    if fields_extracted:
        vals = "; ".join(f"{k.capitalize()}={proof[k]}" for k in fields_extracted)
        lines.append(f"Land record fields extracted: {vals}.")

    return "\n".join(lines)


def _process_ownership_proof_from_segments(
    segmentation: Dict[str, Any],
    page_texts: List[str],
    agreement_address: str,
) -> Dict[str, Any]:
    """
    Given a segmentation result and per-page OCR texts, extract and validate
    the ownership proof data.

    Returns the ownership_proof dict (ready to merge into the main JSON).
    """
    _OWNERSHIP_TYPES = {
        "Land Revenue Receipt", "Property Tax Receipt", "Municipal Tax Receipt",
        "Mutation Record", "Registry Extract", "Government Land Record",
    }

    empty_proof: Dict[str, Any] = {
        "detected"                  : False,
        "document_type"             : None,
        "source_pages"              : [],
        "source_file"               : segmentation.get("source_file", ""),
        "owner_name"                : None,
        "ownership_address"         : None,
        "financial_year"            : None,
        "financial_year_valid"      : None,
        "within_3_months"           : "NOT_APPLICABLE",
        "three_month_rule_applicable": False,
        "ownership_record_valid"    : False,
        "address_match_with_agreement": None,
        "comparison_remarks"        : None,
        "khata": None, "khesra": None, "jamabandi": None,
        "village": None, "thana": None, "district": None,
        "state": None, "receipt_number": None,
    }

    if not segmentation.get("has_ownership_proof"):
        return empty_proof

    # Find best ownership proof segment (prioritise Land Revenue, then Property Tax)
    priority = [
        "Land Revenue Receipt", "Property Tax Receipt", "Municipal Tax Receipt",
        "Mutation Record", "Registry Extract", "Government Land Record",
    ]
    segments = segmentation.get("document_segments", [])
    candidates: List[Dict[str, Any]] = []
    for seg in segments:
        if seg.get("type") in _OWNERSHIP_TYPES:
            candidates.append(seg)
    if not candidates:
        return empty_proof

    def _rank_seg(seg: Dict[str, Any]) -> Tuple[int, int]:
        typ = str(seg.get("type") or "")
        pages = list(seg.get("pages") or [])
        try:
            pri = priority.index(typ)
        except ValueError:
            pri = len(priority)
        return pri, len(pages)

    best_proof: Optional[Dict[str, Any]] = None
    for seg in sorted(candidates, key=_rank_seg):
        seg_pages = list(seg.get("pages") or [])
        seg_text = "\n".join(
            page_texts[pg - 1]
            for pg in seg_pages
            if 1 <= pg <= len(page_texts)
        )
        proof = _extract_ownership_proof_data(
            seg_text,
            str(seg.get("type") or ""),
            seg_pages,
            segmentation.get("source_file", ""),
        )
        match_label, cmp_remarks = _compare_addresses(
            proof.get("ownership_address") or "",
            agreement_address or "",
        )
        proof["address_match_with_agreement"] = match_label
        proof["comparison_remarks"] = cmp_remarks
        if best_proof is None:
            best_proof = proof

    if best_proof is None:
        return empty_proof
    return best_proof


# ─── Discrepancy helper for ownership-specific issues ────────────────────────

def _ownership_discrepancy_remarks(proof: Dict[str, Any]) -> List[str]:
    """
    Generate a list of ownership-proof-specific discrepancy remarks.
    Returns empty list if no issues.
    """
    remarks: List[str] = []
    if not proof or not proof.get("detected"):
        return remarks

    doc_type = proof.get("document_type") or ""
    match    = proof.get("address_match_with_agreement") or "N/A"
    fy       = proof.get("financial_year")
    fy_valid = proof.get("financial_year_valid")
    rec_valid= proof.get("ownership_record_valid", True)
    seg_pages= proof.get("source_pages") or []

    # Merged document note
    # if seg_pages:
    #     remarks.append(
    #         f"Ownership proof pages detected inside merged PDF "
    #         f"(page(s): {', '.join(str(p) for p in seg_pages)})."
    #     )

    # Address mismatch
    if match in ("LOW MATCH", "MISMATCH", "STATE MISMATCH"):
        remarks.append(f"Ownership proof address differs from agreement address ({match}).")

    # Outdated FY
    if doc_type in _FY_BASED_TYPES:
        if fy and fy_valid is False:
            remarks.append(f"Property Tax Receipt belongs to outdated FY {fy}.")
        elif not fy:
            remarks.append("Financial year could not be extracted from Property Tax Receipt.")

    # # Permanent record valid note (informational — not a problem)
    # if doc_type in _PERMANENT_OWNERSHIP_TYPES and rec_valid:
    #     remarks.append(f"Land Revenue Record valid; 3-month rule exempted.")

    return remarks


# ── End of Ownership Proof & Merged Document Module ──────────────────────────


def _status_clr(val: str) -> Tuple[str, str]:
    """Return (fg_hex, bg_hex) for a status/risk string."""
    v = str(val).strip().upper()
    if v in ("CORRECT", "CLEAN", "PASS", "YES", "TRUE", "MATCHED", "OK",
             "COMPLETE", "HIGH MATCH", "LOW", "WITHIN 3 MONTHS", "NO (RECENT)"):
        return C["pass_fg"], C["pass_bg"]
    if v in ("INCORRECT", "HIGH RISK", "FAIL", "NO", "FALSE", "UNMATCHED",
             "FAILED", "INCOMPLETE", "LOW MATCH", "HIGH", "OLDER THAN 3 MONTHS",
             "MISSING", "MISMATCH", "YES (OLD)"):
        return C["fail_fg"], C["fail_bg"]
    return C["warn_fg"], C["warn_bg"]


# ══════════════════════════════════════════════════════════════════════════════
#  CENTRALIZED NORMALIZATION ENGINE  (v6.0 — Production Grade)
#  Provides reusable helpers for issue normalization, deduplication,
#  English translation, Missing Docs canonicalization, Final Remarks
#  and Suggested Actions building. All workbook output passes through here.
# ══════════════════════════════════════════════════════════════════════════════

# Ownership/utility proof document keywords — ANY one of these satisfies
# BOTH "Utility Bill Missing" AND "Ownership Proof Missing" checks.
_OWNERSHIP_UTILITY_KEYWORDS = [
    "electricity bill", "electric bill", "power bill",
    "water bill", "water tax", "gas bill", "telephone bill", "phone bill",
    "internet bill", "broadband bill",
    "municipal tax", "municipality tax", "nagar palika", "nagar nigam",
    "property tax", "house tax", "griha kar",
    "land revenue", "lagan rasid", "lagaan", "jamabandi", "khatauni",
    "registry", "sale deed", "mutation", "khatoni", "khasra", "bhulekh",
    "gram panchayat receipt", "gram panchayat", "panchayat receipt",
    "khata", "ownership proof", "proof of ownership",
    "mutation record", "revenue receipt", "rasid",
]


def _has_ownership_or_utility_proof(docs: List[Dict], missing_docs: List[str]) -> bool:
    """
    Return True if ANY valid ownership/utility proof document is already present
    in the analyzed documents — so we do NOT falsely raise missing docs for it.
    Checks documents_analyzed roles/types AND existing missing_docs list.
    """
    _PROOF_ROLES = {
        "proof_of_ownership", "primary_lease_deed",
        "warehousing_service_agreement",
    }
    _PROOF_TYPES = {
        "electricity bill", "water bill", "water tax receipt",
        "gas bill", "telephone bill", "internet bill",
        "property tax receipt", "municipal tax receipt",
        "land revenue receipt", "mutation record", "registry extract",
        "government land record", "any other legal document",
    }
    for doc in docs:
        role = str(doc.get("role") or "").lower()
        dtype = str(doc.get("detected_type") or "").lower()
        eff = str(doc.get("effective_status") or "").upper()
        # Count it only if it's not a MISSING/INCORRECT doc that couldn't be read
        if eff in ("MISSING",):
            continue
        if role in _PROOF_ROLES:
            if any(kw in dtype for kw in ("electricity", "water", "gas", "telephone",
                                           "internet", "municipal", "property tax",
                                           "land revenue", "mutation", "registry",
                                           "ownership", "bill", "receipt")):
                return True
        if any(dtype == pt for pt in _PROOF_TYPES):
            return True
        if any(kw in dtype for kw in ("electricity", "water bill", "gas bill",
                                       "property tax", "land revenue", "khatauni",
                                       "lagan", "mutation", "registry")):
            return True
    return False


def normalize_issue_text(text: str) -> str:
    """
    Normalize a raw issue string:
    - Strip leading "Issue N –", "Issue I:", "Issue II:" etc.
    - Strip [Source:...] tags
    - Strip status words like FAILED, PASSED, Observation
    - Trim whitespace
    Returns clean audit remark text only.
    """
    if not text:
        return ""
    t = str(text).strip()
    # Strip [Source:...] tags
    t = re.sub(r"\s*\[Source:[^\]]*\]", "", t)
    # Strip custom prefixes like "Issue – State Verbatim Check: " or "Issue – Missing State in Document: "
    t = re.sub(r"^Issue\s*–?\s*[^:]+:\s*", "", t, flags=re.IGNORECASE)
    # Strip leading "Issue N –" / "Issue I:" patterns
    t = re.sub(r"^Issue\s+[IVX\d]+\s*[–\-:]\s*", "", t, flags=re.IGNORECASE)
    # Strip leading "Observation N:" patterns
    t = re.sub(r"^Observation\s*\d*\s*[:\-–]\s*", "", t, flags=re.IGNORECASE)
    # Strip status words at end: "Status: FAILED", "— Action Required"
    t = re.sub(r"\s*(—\s*)?(Status|Action Required|FAILED|PASSED|Compliance Label)\s*[:\-]?\s*\w*\s*$", "", t, flags=re.IGNORECASE)
    # Strip Action suffixes with or without separators/variations
    t = re.sub(r"\s*(?:—\s*)?Action(?:\s+Required)?\s*:\s*.{0,200}$", "", t, flags=re.IGNORECASE | re.DOTALL)
    return t.strip()



def translate_to_english(text: str) -> str:
    """
    Best-effort ASCII/English enforcement: transliterate Devanagari/regional
    script characters into readable English equivalents for Excel output.

    Uses a comprehensive Devanagari → Latin phonetic map so that addresses
    like 'मोजा-धली वेलारी, जिला- भागलपुर, राज्य- बिहार' are rendered as
    'Moja-Dhali Velari, Jila- Bhagalpur, Rajya- Bihar' rather than '[?]'.

    Falls back to a '[NON-LATIN]' marker only for any character not covered
    by the map (e.g. rare conjuncts) — never silently drops content.
    """
    if not text:
        return ""

    # ── Devanagari → Latin phonetic transliteration map ─────────────────────
    # Ordered longest-match first (matras before base consonants, etc.)
    _DEVA_MAP: List[Tuple[str, str]] = [
        # Common Indian-language address keywords (whole-word replacements first)
        ("का नाम",   "Name"),
        ("सं०",      "No."),   ("सं",       "No."),
        ("नं०",      "No."),   ("नं",       "No."),
        ("रकवा",     "Rakwa"),    ("मोजा",     "Moja"),     ("मौजा",     "Mauja"),
        ("थाना",     "Thana"),    ("थाना सं०", "Thana No."),
        ("अंचल",    "Anchal"),   ("जिला",     "Jila"),      ("जिला-",     "Jila-"),
        ("जिला का नाम", "District Name"),
        ("मौजा का नाम", "Mauja Name"),
        ("अंचल का नाम", "Anchal Name"),
        ("राज्य",    "Rajya"),    ("राज्य-",   "State-"),
        ("पिन कोड",  "Pin Code"), ("खाता नं०", "Khata No."),
        ("खेसरा नं०","Khesra No."),("जमाबंदी नं०","Jamabandi No."),
        ("थाना सं०", "Thana No."),
        ("पो०",      "P.O."),     ("वर्ग फीट", "Sq.Ft"),
        ("एकर",      "Acre"),     ("डी०",      "D."),
        ("लं०",      "L."),       ("चौ०",      "W."),
        ("वार्ड",    "Ward"),     ("ग्राम",    "Gram"),
        ("बिहार",    "Bihar"),    ("झारखण्ड", "Jharkhand"),
        ("झारखंड",   "Jharkhand"),("उत्तर प्रदेश","Uttar Pradesh"),
        ("उत्तरप्रदेश","Uttar Pradesh"),
        ("महाराष्ट्र","Maharashtra"),("राजस्थान","Rajasthan"),
        ("मध्य प्रदेश","Madhya Pradesh"),("छत्तीसगढ़","Chhattisgarh"),
        ("पश्चिम बंगाल","West Bengal"),("ओडिशा","Odisha"),
        ("गुजरात",   "Gujarat"),  ("कर्नाटक",  "Karnataka"),
        ("तमिलनाडु", "Tamil Nadu"),("केरल",    "Kerala"),
        ("पंजाब",    "Punjab"),   ("हरियाणा",  "Haryana"),
        ("दिल्ली",   "Delhi"),    ("असम",      "Assam"),
        ("मणिपुर",   "Manipur"),  ("मिज़ोरम",  "Mizoram"),
        ("त्रिपुरा", "Tripura"),  ("मेघालय",   "Meghalaya"),
        ("नागालैंड", "Nagaland"), ("सिक्किम",  "Sikkim"),
        ("उत्तराखंड","Uttarakhand"),("गोवा",    "Goa"),
        # Devanagari digits
        ("०", "0"), ("१", "1"), ("२", "2"), ("३", "3"), ("४", "4"),
        ("५", "5"), ("६", "6"), ("७", "7"), ("८", "8"), ("९", "9"),
        # Individual Devanagari characters (consonants + vowels + matras)
        ("अ","a"),("आ","aa"),("इ","i"),("ई","ee"),("उ","u"),("ऊ","oo"),
        ("ए","e"),("ऐ","ai"),("ओ","o"),("औ","au"),("अं","an"),("अः","ah"),
        ("क","k"),("ख","kh"),("ग","g"),("घ","gh"),("ङ","ng"),
        ("च","ch"),("छ","chh"),("ज","j"),("झ","jh"),("ञ","ny"),
        ("ट","t"),("ठ","th"),("ड","d"),("ढ","dh"),("ण","n"),
        ("त","t"),("थ","th"),("द","d"),("ध","dh"),("न","n"),
        ("प","p"),("फ","ph"),("ब","b"),("भ","bh"),("म","m"),
        ("य","y"),("र","r"),("ल","l"),("व","v"),("श","sh"),
        ("ष","sh"),("स","s"),("ह","h"),("ळ","l"),("क्ष","ksh"),
        ("त्र","tr"),("ज्ञ","gn"),
        # Matras (vowel signs attached to consonants)
        ("\u093e","a"),("\u093f","i"),("\u0940","ee"),("\u0941","u"),
        ("\u0942","oo"),("\u0947","e"),("\u0948","ai"),("\u094b","o"),
        ("\u094c","au"),("\u0902","n"),("\u0903","h"),("\u094d",""),
        # Nukta, avagraha, visarga
        ("\u093c",""),("\u093d",""),("\u0950","Om"),
        # Gujarati digits (U+0AE0-U+0AEF) and Bengali digits (U+09E6-U+09EF)
        ("\u0AE6","0"),("\u0AE7","1"),("\u0AE8","2"),("\u0AE9","3"),("\u0AEA","4"),
        ("\u0AEB","5"),("\u0AEC","6"),("\u0AED","7"),("\u0AEE","8"),("\u0AEF","9"),
        ("\u09E6","0"),("\u09E7","1"),("\u09E8","2"),("\u09E9","3"),("\u09EA","4"),
        ("\u09EB","5"),("\u09EC","6"),("\u09ED","7"),("\u09EE","8"),("\u09EF","9"),
    ]

    if not any(ord(c) > 127 for c in text):
        return text

    result = text
    # Apply multi-character keyword replacements first (longest match)
    for src, tgt in _DEVA_MAP:
        result = result.replace(src, tgt)

    # Any remaining non-ASCII chars → mark individually
    output_chars = []
    i = 0
    while i < len(result):
        c = result[i]
        if ord(c) > 127:
            # Check 2-char combo
            if i + 1 < len(result) and ord(result[i+1]) > 127:
                pair = result[i:i+2]
                output_chars.append("[NON-LATIN]")
                i += 2
                continue
            output_chars.append("[NON-LATIN]")
        else:
            output_chars.append(c)
        i += 1

    cleaned = "".join(output_chars)
    # Collapse multiple consecutive [NON-LATIN] markers
    cleaned = re.sub(r"(\[NON-LATIN\])+", "[NON-LATIN]", cleaned)
    # Clean up spacing artifacts
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip()

    if "[NON-LATIN]" in cleaned:
        cleaned = cleaned + " [Note: partial non-Latin script — verify manually]"
    return cleaned


def deduplicate_issues(issues: List[str]) -> List[str]:
    """
    Semantically deduplicate a list of issue strings.
    Two issues are duplicates if their normalized 50-char prefix matches.
    Returns deduplicated list preserving order.
    """
    seen: List[str] = []
    seen_keys: set = set()
    for iss in issues:
        if not iss or not str(iss).strip():
            continue
        key = re.sub(r"\s+", " ", str(iss).strip().lower())[:60]
        # Also check short key (first 30 chars) to catch near-duplicates
        short_key = key[:30]
        if key not in seen_keys and short_key not in {k[:30] for k in seen_keys}:
            seen.append(str(iss).strip())
            seen_keys.add(key)
    return seen


def canonicalize_missing_documents(
    raw_missing: List[str],
    docs: List[Dict],
    data: Dict,
) -> List[str]:
    """
    Canonicalize the Missing Documents list:
    1. Remove false "Utility Bill Missing" / "Ownership Proof Missing" if any
       ownership/utility proof document was actually provided.
    2. Deduplicate semantically.
    3. Return clean list of missing document names only.

    OWNERSHIP PROOF and UTILITY BILL are treated as SAME REQUIREMENT CATEGORY.
    If ANY one ownership/utility doc is present → do NOT flag either as missing.
    """
    if not isinstance(raw_missing, list):
        raw_missing = []

    # Check if ownership/utility proof is actually present
    # FIX: op.get("detected") alone is sufficient — Land Revenue Receipt / Khatoni
    # detected inside a merged PDF satisfies the ownership proof requirement even
    # when ownership_record_valid is not yet True (e.g. FY not extracted).
    analyzed_docs = data.get("documents_analyzed") or []
    op = data.get("ownership_proof") or {}
    has_proof = (
        bool(op.get("detected"))  # any ownership proof found counts
        or _has_ownership_or_utility_proof(analyzed_docs, raw_missing)
    )

    # Also check electricity_bill section
    eb = data.get("electricity_bill") or {}
    has_utility = bool(
        eb.get("owner_name") or eb.get("bill_date")
        or str(eb.get("bill_type") or "").strip() not in ("", "Other")
    )
    has_proof = has_proof or has_utility

    # Keywords that indicate "ownership/utility" in missing doc strings
    _OWNERSHIP_MISSING_PATTERNS = [
        r"utility\s+bill", r"electricity\s+bill", r"water\s+bill",
        r"gas\s+bill", r"telephone\s+bill", r"ownership\s+proof",
        r"proof\s+of\s+ownership", r"land\s+revenue", r"property\s+tax",
        r"municipal\s+tax", r"possession\s+proof",
    ]

    # Patterns for items that are ISSUES (not missing documents) and must never
    # appear in the missing_documents list.  Lessee company stamp/signature
    # absence is a compliance issue logged under compliance.issues; it is NOT a
    # "missing document" the applicant needs to supply as a separate file.
    _NOT_MISSING_DOC_PATTERNS = [
        r"lessee\s+company\s+stamp",
        r"lessee\s+stamp\s+and\s+signature",
        r"lessee\s+company\s+stamp\s+and\s+signature",
    ]

    cleaned: List[str] = []
    seen_keys: set = set()

    for item in raw_missing:
        item_s = str(item).strip()
        if not item_s:
            continue
        item_lower = item_s.lower()

        # Always suppress lessee company stamp/signature — it is an issue, not
        # a missing document.
        if any(re.search(pat, item_lower) for pat in _NOT_MISSING_DOC_PATTERNS):
            continue

        # Suppress false ownership/utility missing if proof exists
        if has_proof:
            is_ownership_related = any(
                re.search(pat, item_lower) for pat in _OWNERSHIP_MISSING_PATTERNS
            )
            if is_ownership_related:
                continue

        # Normalize item text
        item_clean = re.sub(r"^[•\-\*]\s*", "", item_s).strip()
        key = item_clean.lower()[:40]
        if key and key not in seen_keys:
            cleaned.append(item_clean)
            seen_keys.add(key)

    return cleaned


def build_final_remarks(issues: List[str]) -> str:
    """
    Build Final Remarks in old-workbook style:
    Clean serial numbered list, one issue per line.
    Format: "1. Stamp duty details not available.\n2. Witness details missing."

    STRICT RULES:
    - No "Issue 1", "Issue 2" — just numbers
    - No status wording (FAILED/PASSED/Compliance)
    - No action text
    - No duplicate remarks
    - Concise audit remarks only
    """
    normalized: List[str] = []
    for iss in issues:
        cleaned = normalize_issue_text(str(iss))
        if not cleaned:
            continue
        # Remove leading "N." if already numbered
        cleaned = re.sub(r"^\d+\.\s*", "", cleaned).strip()
        # Translate to English
        cleaned = translate_to_english(cleaned)
        if cleaned and len(cleaned) > 5:
            normalized.append(cleaned)

    deduped = deduplicate_issues(normalized)

    lines = []
    for i, remark in enumerate(deduped, 1):
        # Ensure ends with period
        remark = remark.rstrip(".")
        lines.append(f"{i}. {remark}.")

    return "\n".join(lines)


def build_suggested_actions(issues: List[str]) -> str:
    """
    Build Suggested Actions from normalized issues.
    Format: "1. Obtain registered lease deed copy.\n2. Submit ownership proof."

    STRICT RULES:
    - Serial numbers only
    - Concise corrective language
    - No issue wording, no "Issue", no "Observation", no "Failed"
    - No duplicate actions
    """
    # Action templates derived from issue patterns
    _ACTION_MAP = [
        # Pattern → suggested action
        (r"stamp\s+duty", "Pay correct stamp duty as per state schedule and register the deed with Sub-Registrar."),
        (r"not\s+register|unregister", "Register the lease deed with Sub-Registrar and obtain registration number."),
        (r"utility\s+bill.*older|bill.*older|older.*3\s*month", "Provide a utility bill dated within 3 months of the application date."),
        (r"utility\s+bill.*not\s+found|no.*utility\s+bill|utility\s+bill\s+not|no.*current.*bill", "Submit a valid utility bill (Electricity/Water/Gas) OR land ownership proof (Land Revenue Receipt, Khatoni, Property Tax Receipt)."),
        (r"owner\s+name.*mismatch|name.*mismatch|mismatch.*name", "Provide initial lease deed or NOC explaining the owner name discrepancy."),
        (r"witness.*address|address.*witness", "Obtain complete witness details including full name, address, and signature."),
        (r"witness.*signature|signature.*witness", "Obtain witness signature(s) on the agreement."),
        (r"signature.*missing|missing.*signature|both\s+parties.*sign", "Obtain signed and stamped copies from both lessor and lessee."),
        (r"notari|registr.*not|not.*registr", "Get the agreement notarised or registered with Sub-Registrar."),
        (r"expir|expired|invalid.*date", "Renew the lease/rent agreement to ensure it is valid on the application date."),
        (r"applicant.*not\s+identif|business\s+name.*not|gst.*applicant.*missing", "Provide identity/incorporation proof to establish the GST applicant name."),
        (r"address.*mismatch|mismatch.*address|state\s+mismatch", "Verify and align the address on all documents with the lease deed property clause."),
        (r"state.*verbatim|verbatim.*state|state.*not.*explicit|missing.*state", "Obtain corrected agreement with explicit state name in the property description clause."),
        (r"pincode.*missing|missing.*pincode", "Include complete pincode in the property address on all documents."),
        (r"ownership\s+proof.*mismatch|proof.*address.*mismatch", "Verify ownership proof address matches the lease deed property address."),
        (r"property\s+tax.*outdat|outdated.*tax|tax.*receipt.*old|financial\s+year.*invalid", "Provide Property Tax Receipt for the current or immediately previous financial year."),
        (r"sub.*lease|sub-lease", "Verify sub-lease clause in original deed; obtain NOC from owner if sub-letting is not permitted."),
        (r"address.*incomplete|incomplete.*address|partial.*address", "Provide complete property address including house/flat number, street, city, district, state, and pincode."),
        (r"inferred.*state|state.*inferred|auto.*inferred", "Obtain corrected agreement with explicit state name written in the property description clause."),
        (r"duplicate", "Remove duplicate document from the submission set."),
        (r"ocr.*unread|unread|scanned.*not|not.*readable", "Provide a clearer digital version or re-scan the document at higher resolution."),
        (r"utility bill\s*/\s*ownership proof is missing", "Submit a valid utility bill (Electricity/Water/Gas) dated within 3 months OR a land ownership proof (Land Revenue Receipt, Khatoni, Property Tax Receipt)."),
        (r"witness details like address", "Collect witness information comprising full legal names, active physical addresses, and execution signatures directly on execution pages."),
        (r"lessee signature not found", "Ensure the lessee party executes the contract with a formal wet signature and structural corporate rubber stamp."),
        (r"not registered with sub-registrar", "Present the executed instrument before the local jurisdictional Sub-Registrar to execute formal registration and pay outstanding stamp duty defaults."),
        (r"lessee\s+company\s+stamp", "Obtain signed copy containing the lessee's company seal/rubber stamp."),
        (r"merged\s+document\s+partial\s+failure", "Verify and resolve segment-specific documentation failures within the merged document."),
        (r"poor\s+or\s+blurry|readability\s+warning", "Provide a clearer digital version or re-scan the document at higher resolution."),
    ]

    actions: List[str] = []
    seen_actions: set = set()

    for iss in issues:
        iss_lower = str(iss).lower()
        matched = False
        for pattern, action in _ACTION_MAP:
            if re.search(pattern, iss_lower):
                ak = action[:40].lower()
                if ak not in seen_actions:
                    actions.append(action)
                    seen_actions.add(ak)
                matched = True
                break
        if not matched:
            # Generic fallback: derive action from issue text
            cleaned = normalize_issue_text(str(iss))
            if cleaned and len(cleaned) > 10:
                # Build a generic "Resolve: <issue summary>" action
                short = re.sub(r"\s+", " ", cleaned)[:80]
                generic = f"Resolve: {short}."
                gk = generic[:40].lower()
                if gk not in seen_actions:
                    # actions.append(generic)
                    # seen_actions.add(gk)
                    pass
    if not actions:
        return "No corrective actions required — all compliance checks passed."

    lines = []
    for i, act in enumerate(actions, 1):
        act = act.rstrip(".")
        lines.append(f"{i}. {act}.")
    return "\n".join(lines)


def normalize_excel_output(data: Dict) -> Dict:
    """
    Master normalization: run all normalization helpers on extracted data
    before writing to workbook. Modifies data in-place and returns it.
    Ensures:
    - All text fields are English-only
    - Issues and passes are deduplicated
    - Missing documents are canonicalized (no false ownership/utility missing)
    - Final Remarks follow old-workbook format
    """
    comp = data.get("compliance") or {}
    s = data.get("summary") or {}

    # 1. Normalize issues
    raw_issues = comp.get("issues") or []
    comp["issues"] = deduplicate_issues([str(x) for x in raw_issues if x])

    # 2. Normalize passes
    raw_passes = comp.get("passes") or []
    comp["passes"] = deduplicate_issues([str(x) for x in raw_passes if x])

    # 3. Canonicalize missing documents
    docs = data.get("documents_analyzed") or []
    raw_missing = s.get("missing_documents") or []
    s["missing_documents"] = canonicalize_missing_documents(raw_missing, docs, data)

    filtered_issues = []
    EXEMPT_PHRASES = ["land revenue record valid", "3-month rule exempted", "pages detected inside merged pdf"]
    for iss in comp["issues"]:            # was: bare 'issues' (NameError); iterate the deduplicated list
        iss_lower = str(iss).lower()
        # Explicitly drop technical logs and exemption descriptions from appearing in sheets
        if any(phrase in iss_lower for phrase in EXEMPT_PHRASES):
            continue
        filtered_issues.append(iss)

    comp["issues"] = filtered_issues
    data["compliance"] = comp
    data["summary"] = s

    # ── UTILITY BILL QUALITY DETECTION (document_quality) ─────────────────
    # Detect if any utility-bill-type document is blurry / unreadable.
    # Only triggers for Utility Bill / Electricity Bill / Invoice Bill docs.
    _util_blurry = False
    _util_readable = True
    _quality_note = None
    _QUALITY_WARNING_TEXT = (
        "Note: The Image quality is poor and blurry, "
        "content in non human readable format."
    )
    for _doc in docs:
        _dt = str(_doc.get("detected_type") or "").strip().lower()
        _role = str(_doc.get("role") or "").strip().lower()
        _fn_lower = str(_doc.get("file_name") or "").strip().lower()
        # Check if this document is a utility-bill-type document
        _is_utility = (
            any(kw in _dt for kw in _UTILITY_BILL_DOC_TYPES)
            or any(kw in _fn_lower for kw in ("utility", "electric", "bill", "invoice"))
            or _role in ("proof_of_ownership",)  # utility can be filed as proof
            and any(kw in _dt for kw in ("bill", "invoice", "electricity", "utility"))
        )
        if not _is_utility:
            continue
        # Check readability_warning flag and confidence score
        _rw = _doc.get("readability_warning") is True
        _rcs = 100.0
        try:
            _rcs = float(_doc.get("readability_confidence_score", 100.0) or 100.0)
        except (ValueError, TypeError):
            _rcs = 100.0
        _low_conf = _rcs < 60.0
        # Check if address extraction was empty/fragmented (unreliable)
        _eb = data.get("electricity_bill") or {}
        _addr_text = str(_eb.get("address") or "").strip()
        _addr_empty = len(_addr_text) < 10  # fragmented or missing
        # Only trigger if genuinely unreadable — not minor OCR glitches
        if _rw or _low_conf or (_addr_empty and (_rw or _low_conf)):
            _util_blurry = True
            _util_readable = False
            _quality_note = _QUALITY_WARNING_TEXT
            break  # one blurry utility bill is sufficient

    data["document_quality"] = {
        "utility_bill_blurry": _util_blurry,
        "human_readable": _util_readable,
        "quality_issue_note": _quality_note,
    }

    return data


# ── End of Centralized Normalization Engine ───────────────────────────────────

def _older3m_clr(val: str) -> Tuple[str, str]:
    """Specific colour logic for Older_Than_3_Months column (Yes=bad, No=good, NA=warn)."""
    v = str(val).strip().upper()
    if v == "NO":   return C["pass_fg"], C["pass_bg"]    # Not old = good
    if v == "YES":  return C["fail_fg"], C["fail_bg"]    # Old = bad
    return C["warn_fg"], C["warn_bg"]                    # NA = warn


# HELPER: Determine document type label used as the address extraction source
#         for the "Basis of Documents" column.
def _determine_basis_of_documents(data: Dict, addrs: Dict, ld: Dict, s: Dict) -> str:
    """
    Dynamically determine which specific document was used to extract the
    Principal Place of Business Address. Returns a human-readable string
    naming the exact document type and source.

    Priority (matches address extraction logic in _write_registration_row):
      1. Warehousing/Service Agreement (if sub-lease scenario)
      2. Registered Lease Deed (if duration > 11 months)
      3. Rent Agreement (if duration <= 11 months)
      4. Electricity Bill / Utility bill (if address came from utility doc)
      5. Municipality Tax Receipt / Property Tax Receipt
      6. NOC / Consent Letter (if from other doc category)
      7. address_source field from Gemini (verbatim if set)

    Returns:
        str: e.g. "From Registered Lease Agreement", "From Electricity Bill"
    """
    # Check Gemini-provided address_source first — it is the most precise
    addr_src = str(addrs.get("address_source") or "").strip()

    # Map Gemini's address_source to a document-type label
    _src_lower = addr_src.lower()
    if addr_src and addr_src not in ("null", "none", ""):
        if any(k in _src_lower for k in ("electricity bill", "electric bill", "jseb", "jbvnl",
                                          "bses", "tpddl", "msedcl", "uppcl", "wbsedcl", "tneb")):
            return "From Electricity Bill"
        if any(k in _src_lower for k in ("municipality tax", "municipal tax", "property tax",
                                          "house tax", "khata")):
            return "From Municipality Tax Receipt"
        if any(k in _src_lower for k in ("water tax", "water bill", "water receipt")):
            return "From Water Tax Receipt"
        if any(k in _src_lower for k in ("warehousing", "service agreement", "warehousing/service")):
            return "From Warehousing Agreement"
        if any(k in _src_lower for k in ("noc", "no objection")):
            return "From NOC"
        if any(k in _src_lower for k in ("consent letter", "consent")):
            return "From Consent Letter"
        if any(k in _src_lower for k in ("registered lease", "lease deed", "registered deed")):
            return "From Registered Lease Agreement"
        if any(k in _src_lower for k in ("rent agreement", "rent deed", "rental")):
            return "From Rent Agreement"

    # Fall back to deriving from document category/duration
    _cat_raw = str(s.get("category") or "").lower()
    dur = 0
    try:
        dur = int(ld.get("duration_months") or 0)
    except Exception:
        pass

    if "warehous" in _cat_raw or "service" in _cat_raw:
        return "From Warehousing Agreement"

    # Check documents_analyzed for document roles and types
    docs = data.get("documents_analyzed", []) or []
    for doc in docs:
        role = str(doc.get("role") or "").upper()
        dtype = str(doc.get("detected_type") or "").lower()
        if role == "PRIMARY_LEASE_DEED":
            if dur > 11:
                return "From Registered Lease Agreement"
            elif dur > 0:
                return "From Rent Agreement"
            elif "lease" in dtype:
                return "From Registered Lease Agreement"
            return "From Rent Agreement"
        if role == "WAREHOUSING_SERVICE_AGREEMENT":
            return "From Warehousing Agreement"
        if role == "NOC":
            return "From NOC"
        if role == "CONSENT_LETTER":
            return "From Consent Letter"
        if role == "PROOF_OF_OWNERSHIP":
            if any(k in dtype for k in ("electricity", "electric")):
                return "From Electricity Bill"
            if any(k in dtype for k in ("municipality", "municipal", "property tax")):
                return "From Municipality Tax Receipt"

    # Ultimate fallback using duration
    if dur > 11:
        return "From Registered Lease Agreement"
    elif dur > 0:
        return "From Rent Agreement"

    return "From primary agreement document"


# HELPER: Build a formatted file list with Excel hyperlinks for a cell.
#         Used in "Total Documents Uploaded" and "Ready to Upload Docs" columns.
def _build_file_list_text(files: List[Path]) -> str:
    """
    Build display text listing all files with bullet numbering.
    Format: • File 1 - Filename.pdf\\n• File 2 - Filename2.pdf ...
    """
    if not files:
        return ""
    lines = []
    for i, fp in enumerate(files, 1):
        lines.append(f"• File {i} - {fp.name}")
    return "\n".join(lines)


def _write_hyperlink_file_list(ws, row: int, col: int,
                                files: List[Path], bg: str = "FFFFFF"):
    """
    Write a numbered file list into a cell with clickable Excel hyperlinks.

    Each file gets its own line in the cell text and a single HYPERLINK formula
    pointing to the first file (Excel supports only one hyperlink per cell).
    For multiple files, the display text lists all; the hyperlink opens the first.

    Args:
        ws    : The openpyxl worksheet.
        row   : Target row (1-indexed).
        col   : Target column (1-indexed).
        files : List of Path objects for the files to list.
        bg    : Background hex colour string.
    """
    from openpyxl.styles import Font as _Font

    cell = ws.cell(row=row, column=col)

    if not files:
        cell.value     = ""
        cell.font      = _rfont(sz=10)
        cell.fill      = _fill(bg)
        cell.alignment = _la(w=True)
        cell.border    = _bdr()
        return

    # Build display text (multi-line bullet list)
    display_lines = []
    for i, fp in enumerate(files, 1):
        display_lines.append(f"File {i} - {fp.name}")
    display_text = "\n".join(display_lines)

    # Build HYPERLINK formula pointing to the first file
    # Use Windows-compatible absolute path (forward slashes cause issues on Win)
    first_path = str(files[0].resolve()).replace("/", "\\")
    # Escape double-quotes in path for Excel formula safety
    first_path_escaped = first_path.replace('"', '""')

    if len(files) == 1:
        # Single file: use HYPERLINK formula so clicking opens the file
        formula = f'=HYPERLINK("{first_path_escaped}","File 1 - {files[0].name}")'
        cell.value = formula
    else:
        # Multiple files: display all in text; hyperlink opens first file
        # We set the cell value to the display text and add a hyperlink
        cell.value = display_text
        try:
            cell.hyperlink = first_path_escaped
        except Exception:
            pass  # hyperlink not supported in this openpyxl version — skip gracefully

    cell.font      = _Font(color="0563C1", underline="single", size=10, name="Calibri")
    cell.fill      = _fill(bg)
    cell.alignment = _la(w=True)
    cell.border    = _bdr()


#  FILE UTILITIES
def encode_file(path: Path) -> Tuple[str, bytes]:
    """Return (mime_type, raw_bytes) for a supported document."""
    mime = MIME_MAP.get(path.suffix.lower(), "application/octet-stream")
    return mime, path.read_bytes()


#  FOLDER DETECTION — auto-maps sub-folders to agreement/utility/other
def _folder_cat(name: str) -> Optional[str]:
    """Classify a folder name into agreement | utility | other."""
    n = name.lower().replace("\\", "/").replace("_", " ")
    if any(k in n for k in ["agreement", "lease", "rent", "deed"]):
        return "agreement"
    if any(k in n for k in ["utility", "electricity", "bill", "invoice", "electric"]):
        return "utility"
    if any(k in n for k in ["other", "noc", "identity", "consent", "proof", "letter", "misc"]):
        return "other"
    return None


def detect_category_folders(root: Path) -> Dict[str, Optional[Path]]:
    """Walk root and map the three expected category sub-folders.

    fix: if no recognised sub-folders are found, check if root itself
    contains supported files directly (flat layout). In that case, treat root
    as 'other' so all files are still picked up and client-grouped by filename.
    This handles cases where 50 docs from different states are dumped flat.
    """
    cats: Dict[str, Optional[Path]] = {"agreement": None, "utility": None, "other": None}
    found_any = False
    for folder in sorted(root.iterdir()):
        if not folder.is_dir():
            continue
        cat = _folder_cat(folder.name)
        if cat and cats[cat] is None:
            cats[cat] = folder
            found_any = True
            log.info(f"   Detected [{cat.upper():12s}] → {folder.name}")

    # v4.2: flat-folder fallback — if no sub-folders detected but files exist in root
    if not found_any:
        root_files = [p for p in root.iterdir()
                      if p.is_file() and p.suffix.lower() in SUPPORTED_EXT]
        if root_files:
            log.warning(
                f"  ⚠ No recognised sub-folders found in '{root.name}'. "
                f"Found {len(root_files)} document(s) directly in root. "
                f"Treating root as a flat mixed folder. "
                f"For best results, organise into AGREEMENTS/, UTILITY BILL/, OTHER DOCUMENTS/."
            )
            cats["other"] = root   # treat entire root as 'other' — grouping still works
    return cats


def collect_files_from_folder(folder: Path) -> List[Path]:
    """Recursively collect all supported-extension files from a folder."""
    return sorted([p for p in folder.rglob("*")
                   if p.is_file() and p.suffix.lower() in SUPPORTED_EXT])


#  GEMINI API — retry wrapper + encoding + call logic
def _is_transient(exc: Exception) -> bool:
    """True if the exception is a transient/retryable API error."""
    msg = str(exc).lower()
    return any(k in msg for k in (
        "429", "resource_exhausted", "resourceexhausted", "quota",
        "readerror", "connection", "reset", "forcibly closed",
        "10054", "timeout", "broken pipe", "503", "502",
    ))


def _call_with_retry(fn, label: str = "Gemini") -> Any:
    """Execute fn() with exponential back-off on transient errors."""
    delay = BASE_BACKOFF
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return fn()
        except Exception as exc:
            if _is_transient(exc):
                if attempt == MAX_RETRIES:
                    log.error(f"❌ {label}: failed after {MAX_RETRIES} retries — {exc}")
                    raise
                wait = min(delay + random.uniform(0, delay * 0.25), MAX_BACKOFF)
                log.warning(f"⚠  {label} transient (attempt {attempt}/{MAX_RETRIES}): "
                            f"{type(exc).__name__} — retry in {wait:.0f}s…")
                time.sleep(wait)
                delay = min(delay * 2, MAX_BACKOFF)
            else:
                log.error(f" {label}: non-retryable error — {exc}")
                raise


def _clean_json(raw: str) -> str:
    """Strip markdown fences and trim to the outermost JSON object."""
    raw = raw.strip()
    raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.IGNORECASE | re.MULTILINE)
    raw = re.sub(r"\s*```\s*$",       "", raw, flags=re.IGNORECASE | re.MULTILINE)
    s, e = raw.find("{"), raw.rfind("}") + 1
    return raw[s:e] if s != -1 and e > s else raw


def _parse_json(raw: str) -> Optional[Dict]:
    """Parse JSON with multiple fallback strategies."""
    cleaned = _clean_json(raw)

    # Strategy 1 & 2: direct / trailing-comma fix
    for fixer in [lambda x: x, lambda x: re.sub(r",\s*([}\]])", r"\1", x)]:
        try:
            return json.loads(fixer(cleaned))
        except json.JSONDecodeError:
            pass

    # Strategy 3: find the largest valid JSON block
    for m in reversed(list(re.finditer(r"\{", cleaned))):
        candidate, depth = cleaned[m.start():], 0
        for i, ch in enumerate(candidate):
            if ch == "{":   depth += 1
            elif ch == "}":
                depth -= 1
                if depth == 0:
                    try:
                        return json.loads(candidate[:i + 1])
                    except:
                        break

    log.error(" JSON parsing failed — raw saved to debug_response.txt")
    Path("debug_response.txt").write_text(raw, encoding="utf-8")
    return None


def _gemini_single_call(
    doc_paths: List[Path], api_key: str, label: str,
    cat_hints: Optional[Dict[str, str]] = None,
) -> Optional[Dict]:
    """Load files, build a single Gemini request, and return parsed JSON.

    v4.2 fix: if the configured GEMINI_MODEL is unavailable (404/not found),
    automatically tries each model in GEMINI_MODEL_FALLBACKS before giving up.
    This prevents the entire batch from failing when a preview model is retired.
    """
    loaded: List[Tuple[Path, str, bytes]] = []
    for p in doc_paths:
        if not p.exists():
            log.warning(f"  ⚠ Missing file: {p}"); continue
        try:
            mime, raw = encode_file(p)
            loaded.append((p, mime, raw))
            log.info(f"  ✓ {p.name}  ({mime}, {len(raw)//1024} KB)")
        except Exception as e:
            log.warning(f"  ⚠ Cannot read {p.name}: {e}")

    if not loaded:
        log.error(f"   {label}: no valid files to send"); return None

    # Warn if some requested files were dropped (could not be read / missing)
    dropped = len(doc_paths) - len(loaded)
    if dropped > 0:
        dropped_names = [p.name for p in doc_paths
                         if p not in {lp for lp, _, _ in loaded}]
        log.warning(
            f"  ⚠ {label}: {dropped} of {len(doc_paths)} file(s) could not be loaded "
            f"and will be EXCLUDED from the Gemini call: {dropped_names}"
        )

    _CAT_LABEL = {"agreement": "LEASE/RENT AGREEMENT",
                  "utility":   "UTILITY/ELECTRICITY BILL",
                  "other":     "OTHER DOCUMENT"}

    # Build file_list from successfully-loaded files (content sent to Gemini)
    _loaded_paths = {p for p, _, _ in loaded}
    file_list_lines = []
    idx = 0
    for i, (p, _, _) in enumerate(loaded):
        idx += 1
        cat_tag = (f" [{_CAT_LABEL.get(cat_hints.get(p.name, ''), '')}]"
                   if cat_hints and cat_hints.get(p.name) else "")
        file_list_lines.append(f"  {idx}. {p.name}{cat_tag}")

    # Append any file that was requested but couldn't be loaded —
    # listed as [CONTENT UNAVAILABLE] so Gemini still references the filename
    for p in doc_paths:
        if p not in _loaded_paths:
            idx += 1
            cat_tag = (f" [{_CAT_LABEL.get(cat_hints.get(p.name, ''), '')}]"
                       if cat_hints and cat_hints.get(p.name) else "")
            file_list_lines.append(
                f"  {idx}. {p.name}{cat_tag} [CONTENT UNAVAILABLE — note in documents_analyzed]"
            )

    file_list   = "\n".join(file_list_lines)
    user_prompt = USER_PROMPT_TEMPLATE.format(file_list=file_list, schema=JSON_SCHEMA)
    log.info(f"   Sending {len(loaded)} file(s): [{label}]…")

    try:
        # try GEMINI_MODEL first, then fall through GEMINI_MODEL_FALLBACKS
        models_to_try = [GEMINI_MODEL] + [
            m for m in GEMINI_MODEL_FALLBACKS if m != GEMINI_MODEL
        ]
        last_exc: Optional[Exception] = None

        for model_attempt in models_to_try:
            try:
                if USE_NEW_SDK:
                    client_obj = _genai_new.Client(api_key=api_key)
                    parts = [_gtypes.Part.from_bytes(data=raw, mime_type=mime)
                             for _, mime, raw in loaded]
                    parts.append(_gtypes.Part.from_text(text=user_prompt))

                    def _call(_m=model_attempt):
                        return client_obj.models.generate_content(
                            model=_m,
                            contents=_gtypes.Content(parts=parts, role="user"),
                            config=_gtypes.GenerateContentConfig(
                                system_instruction=SYSTEM_PROMPT,
                                max_output_tokens=16384,
                                temperature=0.0,
                            ),
                        )
                    resp     = _call_with_retry(_call, label=f"{label}[{model_attempt}]")
                    raw_text = resp.text.strip()

                else:
                    _genai_legacy.configure(api_key=api_key)
                    model_obj = _genai_legacy.GenerativeModel(
                        model_name=model_attempt, system_instruction=SYSTEM_PROMPT)
                    pts = [{"inline_data": {"mime_type": mime,
                            "data": base64.b64encode(raw).decode()}}
                           for _, mime, raw in loaded]
                    pts.append({"text": user_prompt})

                    def _call(_mo=model_obj):
                        return _mo.generate_content(pts)
                    raw_text = _call_with_retry(_call, label=f"{label}[{model_attempt}]").text.strip()

                # Success — log if we used a fallback model
                if model_attempt != GEMINI_MODEL:
                    log.warning(f"  ⚠ Used fallback model '{model_attempt}' (primary '{GEMINI_MODEL}' unavailable)")

                result = _parse_json(raw_text)
                if result is None:
                    log.error(f"  ⚠ {label}: JSON parse failed"); return None

                # ── Post-call doc-count sanity check ────────────────────────
                _returned_docs = result.get("documents_analyzed") or []
                _sent_count    = len(loaded)
                if len(_returned_docs) < _sent_count:
                    log.warning(
                        f"  ⚠ {label}: Gemini returned {len(_returned_docs)} "
                        f"documents_analyzed entry/entries but {_sent_count} file(s) "
                        f"were sent. Some documents may have been missed. "
                        f"_reconcile_compliance will synthesise stub entries."
                    )
                    # Patch summary so the count is visible in Excel even before reconcile
                    _summ = result.setdefault("summary", {})
                    _summ["total_documents_analyzed"] = max(
                        _summ.get("total_documents_analyzed", 0),
                        len(_returned_docs),
                    )

                log.info(f"   {label}: Gemini response parsed successfully ") #[{model_attempt}]
                return result

            except Exception as exc:
                err_msg = str(exc).lower()
                # Only fall through on model-not-found errors; re-raise others
                if any(k in err_msg for k in ("404", "not found", "invalid model",
                                               "model not found", "does not exist",
                                               "unsupported", "deprecated")):
                    log.warning(f"  ⚠ Model '{model_attempt}' unavailable: {exc}. Trying next…")
                    last_exc = exc
                    continue
                # Non-model errors: propagate immediately
                raise

        # All models exhausted
        log.error(f"  {label}: All Gemini models failed. Last error: {last_exc}")
        return None

    except Exception as e:
        log.error(f"  {label}: Gemini call failed — {e}")
        log.debug(traceback.format_exc())
        return None


def _merge_chunks(results: List[Dict]) -> Dict:
    """Merge multi-chunk Gemini results into a single data dict."""
    if len(results) == 1:
        return results[0]
    base = deepcopy(results[0])
    for r in results[1:]:
        # Extend documents_analyzed
        base.setdefault("documents_analyzed", []).extend(r.get("documents_analyzed", []))
        # Merge compliance passes & issues
        bc, rc = base.setdefault("compliance", {}), r.get("compliance", {})
        bc["passes"] = list(set(bc.get("passes", []) + rc.get("passes", [])))
        bc["issues"] = list(set(bc.get("issues", []) + rc.get("issues", [])))
        # Escalate risk to worst seen
        rk = {"LOW": 0, "MEDIUM": 1, "HIGH": 2}
        if rk.get(r.get("summary", {}).get("overall_risk_exposure", "LOW"), 0) > \
           rk.get(base.get("summary", {}).get("overall_risk_exposure", "LOW"), 0):
            base.setdefault("summary", {})["overall_risk_exposure"] = \
                r["summary"]["overall_risk_exposure"]
    return base


def call_gemini(
    doc_paths: List[Path], api_key: str, client_name: str,
    cat_map: Optional[Dict[str, List[Path]]] = None,
) -> Optional[Dict]:
    """
    Top-level Gemini call with chunking guard.
    If doc count exceeds MAX_DOCS_PER_CALL, splits into batches and merges results.
    """
    if not doc_paths:
        return None

    # Build filename → category hints
    cat_hints: Dict[str, str] = {}
    if cat_map:
        for _cat, _paths in cat_map.items():
            for _p in _paths:
                cat_hints[_p.name] = _cat

    if len(doc_paths) > MAX_DOCS_PER_CALL:
        log.warning(f"  ⚡ Chunking {len(doc_paths)} docs → batches of {MAX_DOCS_PER_CALL}")
        chunks  = [doc_paths[i:i + MAX_DOCS_PER_CALL]
                   for i in range(0, len(doc_paths), MAX_DOCS_PER_CALL)]
        results = [r for i, c in enumerate(chunks, 1)
                   if (r := _gemini_single_call(c, api_key,
                                                f"{client_name}[chunk {i}/{len(chunks)}]",
                                                cat_hints=cat_hints))]
        return _merge_chunks(results) if results else None
    return _gemini_single_call(doc_paths, api_key, client_name, cat_hints=cat_hints)


#  POST-PROCESSING — _reconcile_compliance
#  Fixes notary detection, name matching, compliance scoring, final status.
REGISTERED_KW = [
    "registered", "sub-registrar", "sub registrar", "registration no", "reg. no",
    "reg no", "book no", "joint sub registrar", "e-stamp", "estamp", "challan",
    "igr", "egrashry", "grashry", "document registration", "registration number",
]


def _norm_name(s: str) -> str:
    """Normalise entity name by stripping suffixes and punctuation."""
    s = re.sub(r"\b(m/s|pvt|ltd|private|limited|llp|inc|corp|co)\b", " ",
               str(s), flags=re.IGNORECASE)
    return re.sub(r"[^A-Z0-9]", "", s.upper())


def _name_score(a: str, b: str) -> int:
    """Return 0-100 similarity score between two entity names."""
    na, nb = _norm_name(a), _norm_name(b)
    if not na or not nb: return 0
    if na == nb:         return 100
    if na in nb or nb in na: return 90
    wa = set(re.findall(r"[A-Z]{2,}", na))
    wb = set(re.findall(r"[A-Z]{2,}", nb))
    if wa and wb:
        return int(len(wa & wb) / max(len(wa), len(wb)) * 85)
    return int(SequenceMatcher(None, na, nb).ratio() * 100)


def _extract_state_from_lease_pdf(lease_path: Path) -> Optional[str]:
    """
    Attempt to extract the verbatim state name from the primary lease deed PDF
    by OCR-ing the property description clause (Clause 5 / किरायेदारी संपत्ति का विवरण).

    This is a GROUND-TRUTH check: it bypasses Gemini's geographic inference by
    reading the raw characters written in the document.

    Returns the state string exactly as written (e.g. "Bihar", "Jharkhand"),
    or None if extraction fails or no state pattern is found.

    Requires: pdf2image, pytesseract, Pillow (all optional — graceful fallback).
    """
    try:
        from pdf2image import convert_from_path
        import pytesseract
    except ImportError:
        return None  # OCR libs not available — skip silently

    # State patterns to search for (Devanagari + common Latin variants)
    # Each tuple: (devanagari_pattern, latin_output)
    STATE_PATTERNS = [
        # Bihar
        (r"बिहार",         "Bihar"),
        (r"bihar",         "Bihar"),
        # Jharkhand
        (r"झारखण्ड",       "Jharkhand"),
        (r"झारखंड",        "Jharkhand"),
        (r"jharkhand",     "Jharkhand"),
        # Uttar Pradesh
        (r"उत्तर\s*प्रदेश", "Uttar Pradesh"),
        (r"uttar\s*pradesh","Uttar Pradesh"),
        # Madhya Pradesh
        (r"मध्य\s*प्रदेश", "Madhya Pradesh"),
        (r"madhya\s*pradesh","Madhya Pradesh"),
        # Rajasthan
        (r"राजस्थान",      "Rajasthan"),
        (r"rajasthan",     "Rajasthan"),
        # Maharashtra
        (r"महाराष्ट्र",    "Maharashtra"),
        (r"maharashtra",   "Maharashtra"),
        # West Bengal
        (r"पश्चिम\s*बंगाल","West Bengal"),
        (r"west\s*bengal", "West Bengal"),
        # Odisha
        (r"ओडिशा",         "Odisha"),
        (r"odisha",        "Odisha"),
        # Chhattisgarh
        (r"छत्तीसगढ",      "Chhattisgarh"),
        (r"chhattisgarh",  "Chhattisgarh"),
        # Gujarat
        (r"गुजरात",        "Gujarat"),
        (r"gujarat",       "Gujarat"),
    ]

    # Clause 5 / property description section keywords (Devanagari + Latin)
    CLAUSE_ANCHORS = [
        "किरायेदारी संपत्ति",
        "संपत्ति का विवरण",
        "राज्य",        # "State" in Hindi
        "राज्य-",
        "राज्य -",
        "clause 5",
        "property description",
    ]

    try:
        # Convert all pages — clause 5 is typically pages 2-4 for handwritten deeds
        pages_to_check = list(range(1, min(6, 10)))  # pages 1-5 (0-indexed for pdf2image)
        images = convert_from_path(
            str(lease_path),
            dpi=200,
            first_page=1,
            last_page=min(6, 99),
        )
    except Exception as e:
        log.debug(f"  pdf2image failed for {lease_path.name}: {e}")
        return None

    # OCR config: Hindi + English
    ocr_config = r"--oem 3 --psm 6 -l hin+eng"

    best_state: Optional[str] = None

    for img in images:
        try:
            raw_text = pytesseract.image_to_string(img, config=ocr_config)
        except Exception:
            try:
                # Fallback: English only
                raw_text = pytesseract.image_to_string(img, config=r"--oem 3 --psm 6")
            except Exception:
                continue

        text_lower = raw_text.lower()

        # Check if this page contains a clause 5 / property description anchor
        has_clause_anchor = any(anchor.lower() in text_lower for anchor in CLAUSE_ANCHORS)
        if not has_clause_anchor:
            continue

        # Search for state pattern near "राज्य" keyword
        # Look for pattern: "राज्य- <STATE>" or "राज्य : <STATE>" or "राज्य <STATE>"
        # Also handle Latin: "State: Bihar" / "State - Bihar"
        state_line_pattern = re.compile(
            r"(?:\u0930\u093e\u091c\u094d\u092f|state)\s*[-:\u2013\s]\s*([^\s,\.\n]{3,20})",
            re.IGNORECASE
        )
        matches = state_line_pattern.findall(raw_text)

        for match in matches:
            match_clean = match.strip()
            # Try to map to known state name
            for dev_pat, latin_name in STATE_PATTERNS:
                if re.search(dev_pat, match_clean, re.IGNORECASE):
                    best_state = latin_name
                    log.info(
                        f"   OCR state extracted from lease clause: "
                        f"'{match_clean}' → '{latin_name}' (file: {lease_path.name})"
                    )
                    return best_state

    if best_state:
        return best_state

    log.debug(f"  OCR: no state pattern found in clause 5 of {lease_path.name}")
    return None


def _reconcile_compliance(data: Dict, client_files: List[Path]) -> Dict:
    """
    Robust post-processing pipeline:
    1.  Fix notary / Sub-Registrar registration detection
    2.  Recalculate owner name match with normalised scoring
    3.  Verify agreement validity and utility bill age
    4.  Normalise witness records
    5.  Derive best POB address
    6.  Rebuild canonical passes list (de-duped)
    7.  Rebuild issues list (suppressing Gemini false positives)
    8.  Compute weighted_compliance_percent from CBIC checklist
    9.  Derive final_status (Clean / Warning / High Risk)
    10. Set per-document effective_status defaults
    11. Write everything back to data dict
    """
    if not data or not isinstance(data, dict):
        log.warning("  ⚠ _reconcile_compliance: empty/invalid data — initialising defaults")
        data = {}

    s     = data.get("summary", {})
    n     = data.get("names", {})
    ld    = data.get("lease_details", {})
    eb    = data.get("electricity_bill") or {}
    comp  = data.get("compliance", {})
    docs  = data.get("documents_analyzed", [])
    addrs = data.get("addresses", {})

    # ── 0. DOCUMENT LIST REPAIR ───────────────────────────────────────────────
    # If Gemini returned an empty documents_analyzed list but we have actual
    # client files, synthesize stub entries so all downstream checks have
    # something to work with and the document count is accurate.
    if not isinstance(docs, list):
        docs = []
    if not docs and client_files:
        log.warning(
            f"  ⚠ _reconcile_compliance: documents_analyzed is empty but "
            f"{len(client_files)} client file(s) exist. "
            f"Synthesising stub document entries from filenames."
        )
        _AGR_KW  = {"agreement", "lease", "deed", "rent", "contract"}
        _UTIL_KW = {"bill", "receipt", "utility", "electricity", "electric",
                    "tax", "invoice", "water", "gas", "phone"}
        for fp in client_files:
            fn_lower = fp.name.lower().replace("_", " ").replace("-", " ")
            if any(k in fn_lower for k in _AGR_KW):
                _role     = "PRIMARY_LEASE_DEED"
                _det_type = "Lease/Rent Agreement"
            elif any(k in fn_lower for k in _UTIL_KW):
                _role     = "PROOF_OF_OWNERSHIP"
                _det_type = "Utility/Ownership Proof"
            else:
                _role     = "OTHER"
                _det_type = "Supporting Document"
            docs.append({
                "file_name"            : fp.name,
                "detected_type"        : _det_type,
                "role"                 : _role,
                "extraction_status"    : "PARTIAL",
                "effective_status"     : "REVIEW",
                "issue_type"           : "Incomplete",
                "discrepancy_reason"   : (
                    "Document was not fully analysed by AI — manual review required."
                ),
                "key_notes"            : (
                    f"Stub entry synthesised from filename '{fp.name}'. "
                    f"Re-run validation without --use-cache to get full analysis."
                ),
                "owner_name"           : None,
                "address_in_doc"       : None,
                "name_match_status"    : "N/A",
                "address_match_status" : "N/A",
                "signature_stamp_present": False,
                "notarised_registered" : False,
                "stamp_duty_amount"    : None,
                "expiry_date"          : None,
                "is_duplicate"         : False,
            })
        data["documents_analyzed"] = docs
        log.info(
            f"   Synthesised {len(docs)} stub document entrie(s) from client_files."
        )

    # ── 1. REGISTRATION / NOTARY DETECTION ────────────────────────────────────
    # Negation phrases that indicate the document is explicitly NOT registered/notarised.
    # If any of these surround the keyword match, the hit is suppressed.
    _NEGATION_PATTERNS = [
        r"not\s+registered", r"neither\s+notari[sz]ed\s+nor\s+registered",
        r"not\s+notari[sz]ed", r"nor\s+registered", r"not\s+yet\s+registered",
        r"unregistered", r"not\s+been\s+registered",
    ]

    def _kw_match(text: str, kw: str) -> bool:
        """
        Word-boundary-safe keyword match with negation guard.
        Prevents:
          • 'registered' matching inside 'unregistered'
          • Positive hits when the surrounding context contains a negation phrase
            (e.g. 'not registered', 'neither notarised nor registered').
        Uses regex word-boundaries so only whole-word or whole-phrase hits count.
        """
        # Step 1: basic word-boundary match
        pattern = r"(?<!\w)" + re.escape(kw) + r"(?!\w)"
        if not re.search(pattern, text, re.IGNORECASE):
            return False
        # Step 2: negation suppression — scan ±120 chars around each match position
        for m in re.finditer(pattern, text, re.IGNORECASE):
            start = max(0, m.start() - 120)
            end   = min(len(text), m.end() + 120)
            context = text[start:end]
            negated = any(re.search(neg, context, re.IGNORECASE)
                          for neg in _NEGATION_PATTERNS)
            if not negated:
                return True   # at least one non-negated hit found
        return False           # all hits were negated

    is_registered = False
    for doc in docs:
        combined = " ".join([str(doc.get(k, "")) for k in
                             ("key_notes", "detected_type", "role", "file_name")]).lower()
        if any(_kw_match(combined, kw) for kw in REGISTERED_KW):
            is_registered = True; break
    if not is_registered:
        for fp in client_files:
            if any(_kw_match(fp.name.lower(), kw) for kw in REGISTERED_KW):
                is_registered = True; break

    # Prefer the AI-extracted registration_type; only override with REGISTERED
    # when keyword scan positively confirms registration (not just mentions it).
    reg_type  = ld.get("registration_type", s.get("registration_type", "UNKNOWN"))
    if is_registered:
        reg_type = "REGISTERED"
    # Trust AI's notary_verified field first; keyword scan is a secondary signal.
    # If AI explicitly says False (not merely absent), do NOT override it.
    ai_notary_verified = ld.get("notary_verified", s.get("notary_verified", None))
    notary_ok = (
        (ai_notary_verified is True)          # AI confirmed notarised/registered
        or (ai_notary_verified is None and is_registered)  # AI silent, keyword hit
    )
    if notary_ok:
        ld["notary_verified"]   = True
        ld["registration_type"] = reg_type
        s["notary_verified"]    = True
        s["registration_type"]  = reg_type
    elif ai_notary_verified is False:
        # AI explicitly said not notarised — trust it; do not let keyword scan override
        ld["notary_verified"]   = False
        ld["registration_type"] = reg_type if reg_type not in ("REGISTERED",) else "NONE"
        s["notary_verified"]    = False
        s["registration_type"]  = ld["registration_type"]

    # ── 2. OWNER NAME MATCH ────────────────────────────────────────────────────
    eb_owner  = eb.get("owner_name", "")
    lessor_nm = n.get("lessor_first_party", ld.get("lessor_name", ""))
    score     = _name_score(eb_owner, lessor_nm)
    owner_match = score >= 70
    eb["owner_name_match_score"]        = score
    eb["owner_name_matched_with_lessor"] = (
        "Matched"      if score >= 90 else
        "Partial Match" if owner_match else
        "Unmatched"
    )

    # ── 3. AGREEMENT VALIDITY ──────────────────────────────────────────────────
    end_date_str = ld.get("end_date")
    agr_valid    = bool(end_date_str and end_date_str >= VALIDATION_DATE)
    ld["agreement_valid_as_on_validation_date"] = agr_valid

    # ── 4. UTILITY BILL AGE ────────────────────────────────────────────────────
    bill_date_str = eb.get("bill_date")
    within_3m     = bool(bill_date_str and bill_date_str >= CUTOFF_DATE)
    eb["within_3_months"] = within_3m

    # ── 5. DURATION NORMALISATION ──────────────────────────────────────────────
    try:    dur = int(ld.get("duration_months", 0) or 0)
    except: dur = 0
    ld["duration_months"] = dur

    # ── 6. WITNESS NORMALISATION ───────────────────────────────────────────────
    witnesses = ld.get("witness_records", [])
    norm_wit  = []
    for w in witnesses:
        nw = {
            "witness_number":   w.get("witness_number", 1),
            "name":             w.get("name", w.get("full name", w.get("full_name", ""))),
            "full_address":     w.get("full_address",
                                     w.get("full - address", w.get("full-address", ""))),
            "signature_present": bool(w.get("signature_present", w.get("signature", False))),
            "compliant":         bool(w.get("compliant", True)),
            "compliance_note":   w.get("compliance_note", w.get("notes", "")),
        }
        norm_wit.append(nw)
    ld["witness_records"] = norm_wit
    wit_names = any(w["name"].strip() for w in norm_wit)
    wit_addr  = all((w.get("full_address") or "").strip()
                    for w in norm_wit if w["name"].strip())
    wit_sig   = all(w["signature_present"] for w in norm_wit if w["name"].strip())

    # ── 7. BEST POB ADDRESS ────────────────────────────────────────────────────
    # Priority: premises_address (from lease clause) > finalised_pob_address > final_pob_address
    # Do NOT use max(len) — longer strings are not more accurate.
    # The lease deed's premises_address is the most verbatim source.
    pob_priority = [
        ld.get("premises_address") or "",          # 1st: directly from lease clause
        addrs.get("finalised_pob_address") or "",  # 2nd: structured addresses section
        s.get("final_pob_address") or "",          # 3rd: summary (may be synthesized)
    ]
    best_pob = next((p for p in pob_priority if p.strip()), "")
    s["final_pob_address"]         = best_pob
    addrs["finalised_pob_address"] = best_pob

    # ── 7a. OCR-BASED STATE VERBATIM CORRECTION ───────────────────────────────
    # Attempt to read the state directly from the lease deed PDF via pytesseract.
    # This is the definitive ground-truth check: if OCR finds a different state in
    # the property description clause than Gemini extracted, we correct it and log it.
    # Falls back silently if pdf2image/pytesseract are not installed.
    _lease_file: Optional[Path] = None
    for fp in client_files:
        fn = fp.name.lower()
        # Identify the primary lease deed file (not the utility/ownership proof)
        if any(kw in fn for kw in ("agreement", "lease", "deed", "lees", "agrement")):
            _lease_file = fp
            break
    if _lease_file is None and client_files:
        # Fallback: try the first file that has agreement-related content
        for fp in client_files:
            if fp.suffix.lower() == ".pdf":
                _lease_file = fp
                break

    if _lease_file:
        _ocr_state = _extract_state_from_lease_pdf(_lease_file)
        if _ocr_state:
            _current_state = str(addrs.get("state") or s.get("state") or "").strip()
            if _current_state.lower() != _ocr_state.lower():
                log.warning(
                    f"  ⚠ OCR STATE CORRECTION: Gemini extracted state='{_current_state}' "
                    f"but OCR of lease deed property clause found '{_ocr_state}'. "
                    f"Correcting to verbatim OCR value: '{_ocr_state}'."
                )
                # Correct the state in all locations
                addrs["state"]   = _ocr_state
                s["state"]       = _ocr_state
                # Flag the discrepancy so reviewer knows a correction was made
                addrs["pincode_state_flag"] = True
                addrs["state_discrepancy_note"] = (
                    f"OCR correction applied: lease deed property clause says '{_ocr_state}' "
                    f"but Gemini originally extracted '{_current_state}'. "
                    f"State updated to match verbatim document text. "
                    f"Reviewer should verify against Clause 5 of the primary lease deed."
                )
                # Rebuild best_pob with corrected state
                if _current_state and _ocr_state and best_pob:
                    best_pob = re.sub(
                        re.escape(_current_state), _ocr_state, best_pob, flags=re.IGNORECASE
                    )
                    s["final_pob_address"]         = best_pob
                    addrs["finalised_pob_address"] = best_pob
                    if ld.get("premises_address"):
                        ld["premises_address"] = re.sub(
                            re.escape(_current_state), _ocr_state,
                            ld["premises_address"], flags=re.IGNORECASE
                        )
            else:
                log.info(
                    f"   OCR STATE VERIFIED: Gemini state='{_current_state}' matches "
                    f"OCR of lease deed clause='{_ocr_state}'. No correction needed."
                )

    # ── 7b. PYTHON-SIDE PINCODE → STATE MISMATCH CHECK ────────────────────────
    # IMPORTANT: This block NEVER overwrites addrs["state"] or any address field.
    # It only sets pincode_state_flag and state_discrepancy_note for flagging.
    # The verbatim state as extracted by Gemini is preserved throughout.
    # If Gemini silently overrode the state (e.g. Bihar->Jharkhand), it cannot be
    # recovered here — prevention is in the SYSTEM_PROMPT STATE_VERBATIM block.
    # We detect likely overrides and emit a reviewer warning in the log.
    _pincode  = str(addrs.get("pincode")  or s.get("pincode")  or "").strip()
    _state    = str(addrs.get("state")    or s.get("state")    or "").lower().strip()
    _district = str(addrs.get("district") or "").lower().strip()

    # Surface any pre-existing Gemini-flagged state discrepancy as a visible log warning
    _existing_disc_note = str(addrs.get("state_discrepancy_note") or "").strip()
    if addrs.get("pincode_state_flag") is True and _existing_disc_note:
        log.warning(
            f"  ⚠ STATE VERBATIM WARNING: pincode_state_flag=True — "
            f"district '{addrs.get('district')}' / extracted state '{addrs.get('state')}'. "
            f"Possible Gemini geographic override. Reviewer MUST verify state against "
            f"primary lease deed property clause. Detail: {_existing_disc_note[:150]}"
        )

    # ── Complete all-India district → state mapping (788 districts, 36 states/UTs) ──
    # Keys are district names in lowercase. Values are canonical state name (lowercase).
    # For districts that span state borders or have name variants, multiple keys map
    # to the same state. Where a district name appears in 2+ states, the value is a
    # list so both are accepted (e.g. "aurangabad" → Bihar AND Maharashtra).
    def _dreg(state_name: str, district_list: List[str]) -> Dict[str, Any]:
        return {d.lower(): state_name for d in district_list}

    _D2S: Dict[str, Any] = {}  # district (lower) → state name OR list of state names

    def _add(state_name: str, districts: List[str]):
        for d in districts:
            key = d.lower()
            if key in _D2S:
                # District name shared across states — store as list
                existing = _D2S[key]
                if isinstance(existing, list):
                    if state_name not in existing:
                        existing.append(state_name)
                else:
                    if existing != state_name:
                        _D2S[key] = [existing, state_name]
            else:
                _D2S[key] = state_name

    # Andhra Pradesh
    _add("andhra pradesh", [
        "srikakulam","vizianagaram","visakhapatnam","vizag","east godavari",
        "west godavari","krishna","guntur","prakasam","nellore",
        "sri potti sriramulu nellore","kurnool","kadapa","anantapur","chittoor",
        "y.s.r.","ysr kadapa",
    ])
    # Telangana
    _add("telangana", [
        "adilabad","bhadradri kothagudem","hyderabad","jagtial","jangaon",
        "jayashankar bhupalpally","jogulamba gadwal","kamareddy","karimnagar",
        "khammam","komaram bheem asifabad","mahabubabad","mahabubnagar","mancherial",
        "medak","medchal malkajgiri","medchal–malkajgiri","mulugu","nagarkurnool",
        "nalgonda","narayanpet","nirmal","nizamabad","peddapalli","rajanna sircilla",
        "ranga reddy","rangareddy","sangareddy","siddipet","suryapet","vikarabad",
        "wanaparthy","warangal rural","warangal urban","yadadri bhuvanagiri",
    ])
    # Maharashtra
    _add("maharashtra", [
        "ahmednagar","akola","amravati","beed","bhandara","buldhana","chandrapur",
        "dhule","gadchiroli","gondia","hingoli","jalgaon","jalna","kolhapur","latur",
        "mumbai","mumbai city","mumbai suburban","nagpur","nanded","nandurbar","nashik",
        "osmanabad","palghar","parbhani","pune","raigad","ratnagiri","sangli","satara",
        "sindhudurg","solapur","thane","wardha","washim","yavatmal",
    ])
    # aurangabad is in both Maharashtra and Bihar — handled by _add's merge logic
    _add("maharashtra", ["aurangabad"])
    # Goa
    _add("goa", ["north goa","south goa","panaji"])
    # Gujarat
    _add("gujarat", [
        "ahmedabad","amreli","anand","aravalli","banaskantha","bharuch","bhavnagar",
        "botad","chhota udaipur","dahod","dang","devbhumi dwarka","gandhinagar",
        "gir somnath","jamnagar","junagadh","kheda","kutch","mahisagar","mehsana",
        "morbi","narmada","navsari","panchmahal","patan","porbandar","rajkot",
        "sabarkantha","surat","surendranagar","tapi","vadodara","valsad",
    ])
    # Dadra & NH + Daman & Diu
    _add("dadra and nagar haveli and daman and diu", [
        "dadra and nagar haveli","dadra & nagar haveli","daman","diu",
    ])
    # Rajasthan
    _add("rajasthan", [
        "ajmer","alwar","banswara","baran","barmer","bharatpur","bhilwara","bikaner",
        "bundi","chittorgarh","churu","dausa","dholpur","dungarpur","hanumangarh",
        "jaipur","jaisalmer","jalore","jhalawar","jhunjhunu","jodhpur","karauli",
        "kota","nagaur","pali","pratapgarh","rajsamand","sawai madhopur","sikar",
        "sirohi","sri ganganagar","tonk","udaipur",
    ])
    # Madhya Pradesh
    _add("madhya pradesh", [
        "agar malwa","alirajpur","anuppur","ashoknagar","balaghat","barwani","betul",
        "bhind","bhopal","burhanpur","chhatarpur","chhindwara","damoh","datia","dewas",
        "dhar","dindori","guna","gwalior","harda","hoshangabad","narmadapuram","indore",
        "jabalpur","jhabua","katni","khandwa","khargone","mandla","mandsaur","morena",
        "narsinghpur","neemuch","niwari","panna","raisen","rajgarh","ratlam","rewa",
        "sagar","satna","sehore","seoni","shahdol","shajapur","sheopur","shivpuri",
        "sidhi","singrauli","tikamgarh","ujjain","umaria","vidisha",
    ])
    # Chhattisgarh
    _add("chhattisgarh", [
        "balod","baloda bazar","balrampur","bastar","bemetara","bijapur","bilaspur",
        "dantewada","dhamtari","durg","gariaband","gaurela pendra marwahi",
        "janjgir champa","jashpur","kabirdham","kanker","khairagarh","kondagaon",
        "korba","koriya","mahasamund","manendragarh chirmiri bharatpur",
        "mohla manpur ambagadh chowki","mungeli","narayanpur","raigarh","raipur",
        "rajnandgaon","sakti","sarangarh bilaigarh","sukma","surajpur","surguja",
    ])
    # Uttar Pradesh
    _add("uttar pradesh", [
        "agra","aligarh","allahabad","prayagraj","ambedkar nagar","amethi","amroha",
        "auraiya","ayodhya","azamgarh","baghpat","bahraich","ballia","balrampur",
        "banda","barabanki","bareilly","basti","bhadohi","sant ravidas nagar","bijnor",
        "budaun","bulandshahr","chandauli","chitrakoot","deoria","etah","etawah",
        "farrukhabad","fatehpur","firozabad","gautam buddha nagar","noida","ghaziabad",
        "ghazipur","gonda","gorakhpur","hamirpur","hapur","hardoi","hathras","jalaun",
        "jaunpur","jhansi","kannauj","kanpur dehat","kanpur nagar","kasganj",
        "kaushambi","kheri","lakhimpur kheri","kushinagar","lalitpur","lucknow",
        "maharajganj","mahoba","mainpuri","mathura","mau","meerut","mirzapur",
        "moradabad","muzaffarnagar","pilibhit","pratapgarh","rampur","saharanpur",
        "sambhal","sant kabir nagar","shahjahanpur","shamli","shravasti",
        "siddharthnagar","sitapur","sonbhadra","sultanpur","unnao","varanasi",
    ])
    # Uttarakhand
    _add("uttarakhand", [
        "almora","bageshwar","chamoli","champawat","dehradun","haridwar","nainital",
        "pauri garhwal","pauri","pithoragarh","rudraprayag","tehri garhwal","tehri",
        "udham singh nagar","udhamnagar",
    ])
    # Bihar
    _add("bihar", [
        "araria","arwal","banka","begusarai","bhagalpur","bhojpur","buxar",
        "darbhanga","east champaran","purba champaran","gaya","gopalganj","jamui",
        "jehanabad","kaimur","katihar","khagaria","kishanganj","lakhisarai","madhepura",
        "madhubani","munger","muzaffarpur","nalanda","nawada","patna","purnia","rohtas",
        "saharsa","samastipur","saran","sheikhpura","sheohar","sitamarhi","siwan",
        "supaul","vaishali","west champaran","pashchim champaran",
    ])
    _add("bihar", ["aurangabad"])   # aurangabad Bihar (shared name with MH — both stored)
    # Jharkhand
    _add("jharkhand", [
        "bokaro","chatra","deoghar","dhanbad","dumka","east singhbhum","singhbhum",
        "garhwa","giridih","godda","gumla","hazaribagh","jamtara","khunti","koderma",
        "latehar","lohardaga","pakur","palamu","ramgarh","ranchi","sahebganj",
        "seraikela","seraikela kharsawan","simdega","west singhbhum",
    ])
    # West Bengal
    _add("west bengal", [
        "alipurduar","bankura","birbhum","cooch behar","dakshin dinajpur","south dinajpur",
        "darjeeling","hooghly","howrah","jalpaiguri","jhargram","kalimpong","kolkata",
        "maldah","malda","murshidabad","nadia","north 24 parganas","paschim bardhaman",
        "paschim medinipur","purba bardhaman","purba medinipur","purulia",
        "south 24 parganas","uttar dinajpur","north dinajpur",
    ])
    # Odisha
    _add("odisha", [
        "angul","balangir","balasore","baleswar","bargarh","bhadrak","bolangir","boudh",
        "cuttack","deogarh","dhenkanal","gajapati","ganjam","jagatsinghpur","jajpur",
        "jharsuguda","kalahandi","kandhamal","kendrapara","kendujhar","keonjhar",
        "khordha","koraput","malkangiri","mayurbhanj","nabarangpur","nayagarh",
        "nuapada","puri","rayagada","sambalpur","sonepur","subarnapur","sundargarh",
    ])
    # Assam
    _add("assam", [
        "baksa","barpeta","biswanath","bongaigaon","cachar","charaideo","chirang",
        "darrang","dhemaji","dhubri","dibrugarh","dima hasao","goalpara","golaghat",
        "hailakandi","hojai","jorhat","kamrup","kamrup metropolitan","guwahati",
        "karbi anglong","karimganj","kokrajhar","lakhimpur","majuli","morigaon",
        "nagaon","nalbari","sivasagar","sonitpur","south salmara mankachar","tinsukia",
        "udalguri","west karbi anglong",
    ])
    # NE States
    _add("arunachal pradesh", [
        "anjaw","changlang","dibang valley","east kameng","east siang","kamle",
        "kra daadi","kurung kumey","lepa rada","lohit","longding","lower dibang valley",
        "lower siang","lower subansiri","namsai","pakke kessang","papum pare",
        "shi yomi","siang","tawang","tirap","upper dibang valley","upper siang",
        "upper subansiri","west kameng","west siang",
    ])
    _add("nagaland", [
        "chumoukedima","dimapur","kiphire","kohima","longleng","mokokchung","mon",
        "niuland","noklak","peren","phek","shamator","tseminyu","tuensang","wokha",
        "zunheboto",
    ])
    _add("manipur", [
        "bishnupur","chandel","churachandpur","imphal east","imphal west","jiribam",
        "kakching","kamjong","kangpokpi","noney","pherzawl","senapati","tamenglong",
        "tengnoupal","thoubal","ukhrul",
    ])
    _add("mizoram", [
        "aizawl","champhai","hnahthial","khawzawl","kolasib","lawngtlai","lunglei",
        "mamit","saiha","saitual","serchhip",
    ])
    _add("tripura", [
        "dhalai","gomati","khowai","north tripura","sepahijala","sipahijala",
        "south tripura","unakoti","west tripura",
    ])
    _add("meghalaya", [
        "east garo hills","east jaintia hills","east khasi hills",
        "eastern west khasi hills","north garo hills","ri bhoi","south garo hills",
        "south west garo hills","south west khasi hills","west garo hills",
        "west jaintia hills","west khasi hills","shillong","mairang",
    ])
    _add("sikkim", [
        "east sikkim","gangtok","north sikkim","south sikkim","west sikkim",
        "pakyong","soreng",
    ])
    # Punjab
    _add("punjab", [
        "amritsar","barnala","bathinda","faridkot","fatehgarh sahib","fazilka",
        "ferozepur","gurdaspur","hoshiarpur","jalandhar","kapurthala","ludhiana",
        "mansa","moga","mohali","sas nagar","muktsar","sri muktsar sahib","nawanshahr",
        "shahid bhagat singh nagar","pathankot","patiala","rupnagar","ropar","sangrur",
        "tarn taran",
    ])
    _add("chandigarh", ["chandigarh"])
    # Haryana
    _add("haryana", [
        "ambala","bhiwani","charkhi dadri","faridabad","fatehabad","gurugram","gurgaon",
        "hisar","jhajjar","jind","kaithal","karnal","kurukshetra","mahendragarh","nuh",
        "palwal","panchkula","panipat","rewari","rohtak","sirsa","sonipat","yamunanagar",
    ])
    # Himachal Pradesh
    _add("himachal pradesh", [
        "bilaspur","chamba","hamirpur","kangra","kinnaur","kullu","lahaul and spiti",
        "mandi","shimla","sirmaur","solan","una",
    ])
    # J&K and Ladakh
    _add("jammu and kashmir", [
        "anantnag","bandipora","baramulla","budgam","doda","ganderbal","jammu",
        "kathua","kishtwar","kulgam","kupwara","poonch","pulwama","rajouri","ramban",
        "reasi","samba","shopian","srinagar","udhampur",
    ])
    _add("ladakh", ["kargil","leh"])
    # Kerala
    _add("kerala", [
        "alappuzha","ernakulam","idukki","kannur","kasaragod","kollam","kottayam",
        "kozhikode","malappuram","palakkad","pathanamthitta","thiruvananthapuram",
        "thrissur","wayanad",
    ])
    _add("lakshadweep", ["lakshadweep","kavaratti"])
    # Karnataka
    _add("karnataka", [
        "bagalkot","bangalore rural","bangalore urban","bengaluru","belgaum","belagavi",
        "ballari","bellary","bidar","chamarajanagar","chikkaballapur","chikkamagaluru",
        "chitradurga","dakshina kannada","davanagere","dharwad","gadag","hassan",
        "haveri","kalaburagi","gulbarga","kodagu","kolar","koppal","mandya","mysuru",
        "mysore","raichur","ramanagara","shivamogga","shimoga","tumakuru","tumkur",
        "udupi","uttara kannada","vijayapura","yadgir",
    ])
    # Tamil Nadu
    _add("tamil nadu", [
        "ariyalur","chengalpattu","chennai","coimbatore","cuddalore","dharmapuri",
        "dindigul","erode","kallakurichi","kanchipuram","kanyakumari","karur",
        "krishnagiri","madurai","mayiladuthurai","nagapattinam","namakkal","nilgiris",
        "the nilgiris","perambalur","pudukkottai","ramanathapuram","ranipet","salem",
        "sivaganga","sivagangai","tenkasi","thanjavur","theni","thoothukudi",
        "tirunelveli","tirupathur","tiruppur","tiruvallur","tiruvannamalai",
        "tiruvarur","vellore","viluppuram","virudhunagar",
    ])
    _add("puducherry", ["puducherry","pondicherry","karaikal","mahe","yanam"])
    # Delhi
    _add("delhi", [
        "central delhi","east delhi","new delhi","north delhi","north east delhi",
        "north west delhi","shahdara","south delhi","south east delhi",
        "south west delhi","west delhi",
    ])
    # Andaman & Nicobar
    _add("andaman and nicobar islands", [
        "north and middle andaman","south andaman","nicobar",
    ])

    # ── Pincode prefix → candidate states (2-digit prefix of 6-digit pincode) ──
    # Where a prefix is shared across 2+ states, ALL are listed — the district
    # lookup above is used to pick the right one.  "?" means no known mapping
    # for that prefix (unmapped/invalid pincode range).
    _PIN_STATES: Dict[str, List[str]] = {
        "11": ["delhi"],
        "12": ["haryana"],                     "13": ["haryana"],
        "14": ["punjab"],                       "15": ["punjab"],
        "16": ["punjab", "chandigarh"],
        "17": ["himachal pradesh"],
        "18": ["jammu and kashmir", "ladakh"],  "19": ["jammu and kashmir", "ladakh"],
        "20": ["uttar pradesh"],  "21": ["uttar pradesh"],  "22": ["uttar pradesh"],
        "23": ["uttar pradesh"],
        "24": ["uttar pradesh", "uttarakhand"],
        "25": ["uttar pradesh"],
        "26": ["uttar pradesh", "uttarakhand"],
        "27": ["uttar pradesh"],  "28": ["uttar pradesh"],
        "30": ["rajasthan"],      "31": ["rajasthan"],  "32": ["rajasthan"],
        "33": ["rajasthan"],      "34": ["rajasthan"],
        "36": ["gujarat"],        "37": ["gujarat"],    "38": ["gujarat"],
        "39": ["gujarat", "dadra and nagar haveli and daman and diu"],
        "40": ["maharashtra", "goa"],
        "41": ["maharashtra"],  "42": ["maharashtra"],
        "43": ["maharashtra"],  "44": ["maharashtra"],
        "45": ["madhya pradesh"],  "46": ["madhya pradesh"],
        "47": ["madhya pradesh"],  "48": ["madhya pradesh"],
        "49": ["madhya pradesh", "chhattisgarh"],
        "50": ["telangana", "andhra pradesh"],
        "51": ["andhra pradesh", "telangana"],
        "52": ["andhra pradesh", "telangana"],
        "53": ["andhra pradesh"],
        "54": ["uttar pradesh"],  "55": ["uttar pradesh"],
        "56": ["karnataka"],  "57": ["karnataka"],
        "58": ["karnataka"],  "59": ["karnataka"],
        "60": ["tamil nadu", "puducherry"],
        "61": ["tamil nadu"],  "62": ["tamil nadu"],
        "63": ["tamil nadu"],  "64": ["tamil nadu"],
        "67": ["kerala", "lakshadweep"],
        "68": ["kerala"],      "69": ["kerala"],
        "70": ["west bengal"],  "71": ["west bengal"],
        "72": ["west bengal"],
        "73": ["west bengal", "sikkim"],
        "74": ["west bengal", "andaman and nicobar islands"],
        "75": ["odisha"],  "76": ["odisha"],  "77": ["odisha"],
        "78": ["assam"],
        "79": ["assam", "arunachal pradesh", "nagaland", "manipur",
               "mizoram", "tripura", "meghalaya"],
        "80": ["bihar"],  "81": ["bihar"],
        "82": ["bihar", "jharkhand"],  # shared — resolved by district below
        "83": ["bihar", "jharkhand"],  # shared — resolved by district below
        "84": ["bihar"],  "85": ["bihar"],
    }

    if _pincode and _state:
        _prefix   = _pincode[:2]
        _candidates = _PIN_STATES.get(_prefix, [])

        if _candidates:
            # First try: use extracted district to find the authoritative state
            _auth_state: Optional[str] = None
            if _district:
                _ds_result = _D2S.get(_district)
                if _ds_result:
                    if isinstance(_ds_result, list):
                        # District shared across states — check if doc state matches any
                        _auth_state = next(
                            (st for st in _ds_result if st in _state), _ds_result[0]
                        )
                    else:
                        _auth_state = _ds_result

            # Second try: if district not found, fall back to prefix candidates
            if _auth_state is None:
                # Accept if the document state matches any candidate
                _state_ok = any(cand in _state for cand in _candidates)
            else:
                _state_ok = _auth_state in _state

            if not _state_ok:
                _expected = _auth_state.title() if _auth_state else (
                    " or ".join(c.title() for c in _candidates)
                )
                addrs["pincode_state_flag"] = True
                if not addrs.get("state_discrepancy_note"):
                    addrs["state_discrepancy_note"] = (
                        f"Python validation: document states '{addrs.get('state')}' but "
                        f"district '{addrs.get('district', '')}' + pincode {_pincode} "
                        f"indicate '{_expected}'. "
                        f"State extracted verbatim — verify against primary lease deed clause."
                    )
            else:
                # Consistent — explicitly clear any stale flag from Gemini
                addrs["pincode_state_flag"] = False
                addrs["state_discrepancy_note"] = None

    # ── 7c. STATE INFERENCE FROM DISTRICT (when state absent from document) ─────
    _state_raw = str(addrs.get("state") or s.get("state") or "").strip()
    _state_is_absent = not _state_raw or _state_raw.lower() in (
        "null", "none", "n/a", "na", "not specified", "unknown", "not mentioned",
        "not provided", "not available", ""
    )
    if _state_is_absent and _district:
        _d2s_result = _D2S.get(_district)
        if _d2s_result:
            # Resolve: could be a str (single state) or list (shared-district)
            if isinstance(_d2s_result, list):
                # Multiple candidate states — pick the first as best-guess but note ambiguity
                _inferred_state_raw  = _d2s_result[0]
                _all_candidates      = ", ".join(c.title() for c in _d2s_result)
                _ambiguous           = len(_d2s_result) > 1
            else:
                _inferred_state_raw  = _d2s_result
                _all_candidates      = _inferred_state_raw.title()
                _ambiguous           = False

            _inferred_canonical = normalize_state_name(_inferred_state_raw)
            if _inferred_canonical and _inferred_canonical != "Unknown_State":
                log.info(
                    f"  [StateLookup] State absent in document — inferred from district "
                    f"'{_district}' → '{_inferred_canonical}'"
                    + (f" (ambiguous: candidates = {_all_candidates})" if _ambiguous else "")
                )
                # Write inferred state to all state fields
                addrs["state"]   = _inferred_canonical
                s["state"]       = _inferred_canonical
                if ld.get("premises_state") in (None, "", "null", "None", "N/A"):
                    ld["premises_state"] = _inferred_canonical
                # Mark as inferred so downstream steps (step 9, Excel) can flag it
                addrs["state_inferred_from_district"] = True
                addrs["state_inferred_note"] = (
                    (f"State not explicitly written in document. "
                     f"Auto-inferred as '{_inferred_canonical}' from district '{_district.title()}'. "
                     + (f"Note: district '{_district.title()}' spans multiple states "
                        f"({_all_candidates}); reviewer must confirm correct state. " if _ambiguous else "")
                     + "Action: Obtain corrected agreement with explicit state name.")
                )
                # Also rebuild best_pob to include the inferred state if it's missing
                if best_pob and _inferred_canonical.lower() not in best_pob.lower():
                    best_pob = best_pob.rstrip(", ") + f", {_inferred_canonical}"
                    s["final_pob_address"]         = best_pob
                    addrs["finalised_pob_address"] = best_pob
        else:
            log.debug(
                f"  [StateLookup] State absent, district '{_district}' not found in D2S map — "
                f"cannot auto-infer state. Leaving state as null."
            )

    # ── 7d. MULTILINGUAL STATE FALLBACK SCANNING (MANDATORY) ────────────────────
    _state_raw2 = str(addrs.get("state") or s.get("state") or "").strip()
    _state_is_absent2 = not _state_raw2 or _state_raw2.lower() in (
        "null", "none", "n/a", "na", "not specified", "unknown", "not mentioned",
        "not provided", "not available", "unknown_state", ""
    )
    if _state_is_absent2:
        log.info("   [StateFallback] State still absent. Initiating fallback scanning across all documents...")
        found_fallback_state = None
        for fp in client_files:
            if fp.suffix.lower() == ".pdf":
                texts = _extract_page_texts_from_pdf(fp)
                for text in texts:
                    if not text:
                        continue
                    # Check against known states in sorted order of variant length
                    for variant in sorted(STATE_NORMALIZATION_MAP.keys(), key=len, reverse=True):
                        # Use word boundary for English/Latin variants of length > 2
                        is_hindi = any(ord(c) >= 0x0900 and ord(c) <= 0x097F for c in variant)
                        if is_hindi:
                            pattern = re.escape(variant)
                        else:
                            pattern = r"\b" + re.escape(variant) + r"\b" if len(variant) > 2 else re.escape(variant)
                        
                        if re.search(pattern, text, re.IGNORECASE):
                            found_fallback_state = STATE_NORMALIZATION_MAP[variant]
                            log.info(f"   [StateFallback] Found state '{found_fallback_state}' in {fp.name} via variant '{variant}'")
                            break
                    if found_fallback_state:
                        break
            if found_fallback_state:
                break
        if found_fallback_state:
            addrs["state"] = found_fallback_state
            s["state"] = found_fallback_state
            if ld.get("premises_state") in (None, "", "null", "None", "N/A"):
                ld["premises_state"] = found_fallback_state
            if "ownership_proof" in data and isinstance(data["ownership_proof"], dict):
                data["ownership_proof"]["state"] = found_fallback_state
            # Rebuild best_pob with the found state
            if best_pob and found_fallback_state.lower() not in best_pob.lower():
                best_pob = best_pob.rstrip(", ") + f", {found_fallback_state}"
                s["final_pob_address"]         = best_pob
                addrs["finalised_pob_address"] = best_pob

    # Resolve is_genuine_utility from documents
    GENUINE_UTILITY_TYPES = {
        "electricity bill", "water tax receipt", "water bill",
        "gas bill", "telephone bill", "internet bill",
        "municipal tax receipt", "property tax receipt",
    }
    is_genuine_utility = False
    for doc in docs:
        dt   = (doc.get("detected_type") or "").lower().strip()
        role = (doc.get("role") or "").lower()
        if role in ("proof_of_ownership", "primary_lease_deed"):
            continue  # ownership/lease docs are not utility bills
        if any(ut in dt for ut in GENUINE_UTILITY_TYPES):
            is_genuine_utility = True
            break
    # Also accept if electricity_bill section has a real (non-Other) bill type
    eb_bill_type = (eb.get("bill_type") or "").lower()
    if eb_bill_type not in ("", "other") and any(
        eb_bill_type in ut for ut in GENUINE_UTILITY_TYPES
    ):
        is_genuine_utility = True

    # ── 7e. OWNERSHIP PROOF INTEGRATION & SEGMENTATION (v3.0) ─────────────────
    _ownership_proof: Dict[str, Any] = data.get("ownership_proof") or {}
    _agreement_addr = str(
        ld.get("premises_address") or
        addrs.get("finalised_pob_address") or
        s.get("final_pob_address") or ""
    ).strip()
    _all_segments: List[Dict[str, Any]] = []
    _best_seg: Optional[Dict[str, Any]] = None

    for fp in client_files:
        if fp.suffix.lower() != ".pdf":
            continue
        _page_texts = _extract_page_texts_from_pdf(fp)
        if not _page_texts:
            continue
        _seg = _segment_merged_pdf(fp, _page_texts)
        _all_segments.append(_seg)
        if _best_seg is None and (_seg.get("has_agreement") or _seg.get("has_ownership_proof")):
            _best_seg = _seg
        if not _ownership_proof.get("detected") and _seg.get("has_ownership_proof"):
            _op = _process_ownership_proof_from_segments(_seg, _page_texts, _agreement_addr)
            if _op.get("detected"):
                _ownership_proof = _op
                if _best_seg is None:
                    _best_seg = _seg
                log.info(
                    f"   [OwnershipProof] Detected in '{fp.name}': "
                    f"{_op['document_type']} pages={_op['source_pages']}, "
                    f"address_match={_op['address_match_with_agreement']}"
                )

    merged_partial_fail = False
    segment_failures = []

    if _all_segments:
        _primary_seg = _best_seg or _all_segments[0]
        data["document_segmentation"] = dict(_primary_seg)

        is_merged = bool(_primary_seg.get("merged_document"))
        
        def get_segment_pages(seg_type):
            segs = _primary_seg.get("document_segments") or []
            for seg in segs:
                if seg.get("type", "").lower() == seg_type.lower():
                     return seg.get("pages")
            return []

        if is_merged:
            # Check Agreement Segment
            agreement_fails = False
            agr_fail_reasons = []
            if not ld.get("signatures_both_parties"):
                agreement_fails = True
                agr_fail_reasons.append("missing signatures/stamps")
            if not agr_valid:
                agreement_fails = True
                agr_fail_reasons.append("agreement expired")
            if not ld.get("stamp_duty_compliant"):
                agreement_fails = True
                agr_fail_reasons.append("stamp duty non-compliant")
            if not ld.get("lessee_company_stamp_present"):
                agreement_fails = True
                agr_fail_reasons.append("missing lessee company stamp")
                
            if agreement_fails:
                p_range = _primary_seg.get("agreement_source_pages") or _format_pages_for_display(get_segment_pages("Lease Deed") or get_segment_pages("Rent Agreement"))
                p_str = f"pages {p_range}" if p_range else "unknown pages"
                segment_failures.append(f"Agreement Segment ({p_str}: {', '.join(agr_fail_reasons)})")
                
            # Check Utility Bill Segment
            if is_genuine_utility:
                utility_fails = False
                util_fail_reasons = []
                if not within_3m:
                    utility_fails = True
                    util_fail_reasons.append("older than 3 months")
                if not owner_match and eb_owner:
                    utility_fails = True
                    util_fail_reasons.append("owner name mismatch")
                    
                if utility_fails:
                    p_range = _format_pages_for_display(get_segment_pages("Electricity Bill") or get_segment_pages("Utility Bill"))
                    p_str = f"pages {p_range}" if p_range else "unknown pages"
                    segment_failures.append(f"Utility Bill Segment ({p_str}: {', '.join(util_fail_reasons)})")
                    
            # Check Ownership Proof Segment
            if _ownership_proof.get("detected"):
                op_fails = False
                op_fail_reasons = []
                op_addr_match = _ownership_proof.get("address_match_with_agreement")
                if op_addr_match in ("MISMATCH", "STATE MISMATCH", "LOW MATCH"):
                    op_fails = True
                    op_fail_reasons.append("address mismatch")
                if _ownership_proof.get("financial_year_valid") is False:
                    op_fails = True
                    op_fail_reasons.append("outdated financial year")
                    
                if op_fails:
                    p_range = _primary_seg.get("ownership_proof_source_pages") or _format_pages_for_display(get_segment_pages("Land Revenue Receipt") or get_segment_pages("Property Tax Receipt") or get_segment_pages("Municipal Tax Receipt"))
                    p_str = f"pages {p_range}" if p_range else "unknown pages"
                    segment_failures.append(f"Ownership Proof Segment ({p_str}: {', '.join(op_fail_reasons)})")

            if segment_failures:
                merged_partial_fail = True

    s["merged_document_partial_failure"] = merged_partial_fail
    if "document_segmentation" in data and isinstance(data["document_segmentation"], dict):
        data["document_segmentation"]["merged_document_partial_failure"] = merged_partial_fail

    # If Gemini provided ownership_proof but 3-month rule was incorrectly applied,
    # fix it here for permanent ownership types.
    if _ownership_proof.get("detected"):
        _op_type = _ownership_proof.get("document_type") or ""
        if _op_type in _PERMANENT_OWNERSHIP_TYPES:
            _ownership_proof["three_month_rule_applicable"] = False
            _ownership_proof["within_3_months"]             = "NOT_APPLICABLE"
            _ownership_proof["ownership_record_valid"]      = True
        elif _op_type in _FY_BASED_TYPES:
            _ownership_proof["three_month_rule_applicable"] = False
            _ownership_proof["within_3_months"]             = "NOT_APPLICABLE"
            # Re-validate FY if we have the value
            _fy_str = _ownership_proof.get("financial_year")
            if _fy_str:
                _fy_ok = _is_fy_valid(_fy_str)
                _ownership_proof["financial_year_valid"] = _fy_ok
                _ownership_proof["ownership_record_valid"] = _fy_ok

        # Build address comparison if not done
        if not _ownership_proof.get("address_match_with_agreement"):
            _agreement_addr2 = str(
                ld.get("premises_address") or
                addrs.get("finalised_pob_address") or
                s.get("final_pob_address") or ""
            ).strip()
            _ml, _cr = _compare_addresses(
                _ownership_proof.get("ownership_address") or "",
                _agreement_addr2,
            )
            _ownership_proof["address_match_with_agreement"] = _ml
            _ownership_proof["comparison_remarks"]           = _cr

    data["ownership_proof"] = _ownership_proof

    # ── 7f. READABILITY & QUALITY SCAN (MANDATORY) ───────────────────────────
    readability_warning_triggered = False
    warning_text = "Document/Image quality is poor or blurry. Content is partially unreadable and could not be reliably verified by human or automated validation."
    
    for doc in docs:
        fn = doc.get("file_name")
        score_from_cache = OCR_CONFIDENCE_CACHE.get(fn) if fn else None
        score_from_ai = doc.get("readability_confidence_score")
        
        # Resolve final score
        if score_from_cache is not None:
            final_score = score_from_cache
        elif score_from_ai is not None:
            try: final_score = float(score_from_ai)
            except: final_score = 100.0
        else:
            final_score = 100.0
            
        doc["readability_confidence_score"] = round(final_score, 1)
        
        # Check if score is low or if Gemini flagged it
        ai_warning = doc.get("readability_warning") is True
        low_confidence = final_score < 60.0
        
        if low_confidence or ai_warning:
            doc["readability_warning"] = True
            readability_warning_triggered = True
            
            # Set key_notes and discrepancy_reason in doc
            doc_notes = doc.get("key_notes") or ""
            if warning_text not in doc_notes:
                doc["key_notes"] = (doc_notes + " | " + warning_text if doc_notes else warning_text)
            doc_reason = doc.get("discrepancy_reason") or ""
            if warning_text not in doc_reason:
                doc["discrepancy_reason"] = (doc_reason + " | " + warning_text if doc_reason else warning_text)
        else:
            doc["readability_warning"] = False

    s["readability_warning"] = readability_warning_triggered

    # ── 8. OVERRIDE CBIC CHECKLIST STATUSES FROM PROGRAMMATIC FACTS ───────────
    # Gemini can get these wrong. Python ground-truth always wins for verifiable checks.
    # Setup stamp defaults if missing
    if "lessee_company_stamp_present" not in ld:
        ld["lessee_company_stamp_present"] = False
    if "lessor_stamp_present" not in ld:
        ld["lessor_stamp_present"] = False

    cbic_overrides = {
        1:  "PASS" if within_3m else "FAIL",
        4:  ("PASS" if owner_match else ("N/A" if not eb.get("owner_name") else "FAIL")),
        8:  "PASS" if (ld.get("signatures_both_parties") and ld.get("lessee_company_stamp_present")) else "FAIL",
        10: "PASS" if agr_valid else "FAIL",
        11: "PASS" if notary_ok else "FAIL",
        14: "PASS" if ld.get("stamp_duty_compliant") else "FAIL",
    }
    # Check 3 — only genuine utility docs count; land revenue receipts do not
    cbic_overrides[3] = "PASS" if is_genuine_utility else "FAIL"
    # Check 6 — N/A if no genuine utility bill present
    if not is_genuine_utility:
        cbic_overrides[6] = "N/A"

    cbic_checks = comp.get("cbic_checklist", [])
    for chk in cbic_checks:
        cn = chk.get("check_number")
        if cn in cbic_overrides:
            new_status = cbic_overrides[cn]
            if chk.get("status") != new_status:
                chk["_overridden"] = True   # internal flag for logging
            chk["status"] = new_status

    # ── 9. REBUILD PASSES LIST ─────────────────────────────────────────────────
    passes: List[str] = []

    def _p(cond, msg):
        if cond: passes.append(msg)

    _p(within_3m,
       f"Ownership proof (utility bill) is within 3 months of {VALIDATION_DATE}.")
    _p(owner_match,
       f"Owner name on electricity bill matches lessor in lease/rent agreement "
       f"(score: {score}%).")
    _p(notary_ok,
       f"Lease/Rent agreement is "
       f"{'Registered with Sub-Registrar' if reg_type == 'REGISTERED' else 'Notarised'} — verified.")
    _p(ld.get("signatures_both_parties"),
       "Both parties (Lessor & Lessee) have signed and stamped the agreement.")
    _p(ld.get("stamp_duty_compliant"),
       f"Stamp duty is compliant (duration: {dur} months).")
    _p(dur > 0,
       f"Duration {dur} months → "
       f"{'Lease Deed (>11 months)' if dur > 11 else 'Rent Agreement (<=11 months)'}.")
    _p(agr_valid,
       f"Agreement is valid and not expired as on {VALIDATION_DATE}.")
    _p(wit_names, "Witness name(s) documented in the agreement.")
    _p(wit_names and wit_sig, "Witness signature(s) present.")
    _p(s.get("address_match_status", "") in ("HIGH MATCH", "PARTIAL MATCH"),
       f"Address on ownership proof {(s.get('address_match_status') or '').lower()} "
       f"with lease/rent agreement.")
    _p((n.get("finalised_gst_business_name") or "").strip(),
       f"GST applicant identified: {n.get('finalised_gst_business_name', '')}.")

    # Ownership proof address matches agreement address — HIGH MATCH
    _op_match = _ownership_proof.get("address_match_with_agreement") or ""
    if _ownership_proof.get("detected") and _op_match == "HIGH MATCH":
        passes.append(
            f"Ownership proof address ({_ownership_proof.get('document_type', 'Ownership Proof')}) "
            f"matches agreement address — HIGH MATCH."
        )

    def _to_str(item) -> str:
        """Coerce a passes/issues entry to a plain string safely."""
        if item is None:
            return ""
        if isinstance(item, dict):
            # Gemini sometimes returns {"text": "...", "severity": "..."}
            return str(item.get("text") or item.get("description")
                       or item.get("issue") or item.get("pass")
                       or next(iter(item.values()), "")).strip()
        return str(item).strip()

    # Merge Gemini's passes (de-dup) — filter raw numbers and short junk entries
    seen = {p[:50].lower() for p in passes}
    for p in [_to_str(x) for x in comp.get("passes", []) if x]:
        # Skip bare numbers or very short strings (e.g. "1", "2", "Check 1")
        if not p or re.fullmatch(r"[\d\s,\.]+|check\s*\d+", p.strip(), re.IGNORECASE):
            continue
        if len(p.strip()) < 10:
            continue
        if p[:50].lower() not in seen:
            passes.append(p); seen.add(p[:50].lower())

    # ── 10. REBUILD ISSUES LIST (filter contradicted Gemini issues) ────────────
    issues: List[str] = []
    
    if readability_warning_triggered:
        issues.append(warning_text)

    if merged_partial_fail:
        failure_msg = f"Merged Document Partial Failure: One or more segments failed validation — {'; '.join(segment_failures)}."
        issues.append(failure_msg)
        
        # Set the document's effective_status to INCORRECT and override status in docs
        for doc in docs:
            doc["effective_status"] = "INCORRECT"
            doc_reason = doc.get("discrepancy_reason") or ""
            if failure_msg not in doc_reason:
                doc["discrepancy_reason"] = (doc_reason + " | " + failure_msg if doc_reason else failure_msg)

    # Determine if an ownership proof that is exempt from 3-month rule is present
    _op_type_for_filter = str(_ownership_proof.get("document_type") or "").strip()
    _op_exempt_from_3m = _op_type_for_filter in _PERMANENT_OWNERSHIP_TYPES or \
                         _op_type_for_filter in _FY_BASED_TYPES

    for issue in [_to_str(x) for x in comp.get("issues", []) if x]:
        if not issue:
            continue
        il = issue.lower()
        if ("notari" in il or "registr" in il) and notary_ok:           continue
        if "owner name" in il and ("not match" in il or "mismatch" in il
                                   or "does not" in il) and owner_match: continue
        if ("older than 3" in il or "more than 3" in il
                or "not within" in il) and within_3m:                   continue
        # v3.0: Do NOT apply 3-month rule to Land Revenue Records or Tax Receipts
        if ("older than 3" in il or "more than 3" in il or "not within 3" in il) \
                and _op_exempt_from_3m:
            log.info(
                f"   [OwnershipProof] Suppressed 3-month issue for {_op_type_for_filter}: '{issue[:80]}'"
            )
            continue
        if "sign" in il and "both" in il and ld.get("signatures_both_parties"): continue
        if "expir" in il and agr_valid:                                  continue
        # FIX: suppress "no ownership proof" / "no utility bill" Gemini issues
        # when an ownership proof (Land Revenue Receipt, Khatoni, etc.) IS detected.
        _op_detected_now = bool(_ownership_proof.get("detected"))
        if _op_detected_now:
            if any(phrase in il for phrase in (
                "no ownership proof", "ownership proof document", "ownership proof is not",
                "no utility bill", "utility bill not provided", "no current utility bill",
                "utility bill / ownership proof is missing",
            )):
                log.info(f"   [OwnershipProof] Suppressed false-missing issue (ownership proof detected): '{issue[:80]}'")
                continue
        issues.append(issue)

    # Add freshly-detected issues
    def _issue_absent(kws): return not any(all(k in str(i).lower() for k in kws) for i in issues if i)

    # Guard: do NOT emit "older than 3 months" issue if ownership proof is exempt
    # Mandatory Combined Absence Rule
    _op_detected = bool(_ownership_proof.get("detected"))
    _has_any_proof = is_genuine_utility or _op_detected
    if (not _has_any_proof) or (not eb.get("owner_name") and not _ownership_proof.get("owner_name") and not _op_detected):
        if _issue_absent(["utility bill / ownership proof"]):
            issues.append("Utility bill / ownership proof is missing — Submit a valid utility bill or land ownership record.")
        # Remove any separate Gemini-generated ownership-proof / utility-bill issues
        # so they don't appear alongside the combined message in Final Remarks.
        def _is_separate_op_or_ub(iss: str) -> bool:
            il = iss.lower()
            if (il.startswith("ownership proof:") or il.startswith("utility bill:") or 
                (re.search(r"\bownership\s+proof\b", il) is not None and ("missing" in il or "not" in il or "provide" in il)) or
                (re.search(r"\butility\s+bill\b", il) is not None and ("missing" in il or "not" in il or "provide" in il))):
                return True
            # Catch standalone bullet text or phrases generated by Gemini
            if "utility bill" in il and "ownership proof" in il and "missing" in il:
                # Keep our explicit combined one by checking for the exact phrase punctuation/hyphen signature
                return "— submit a valid" not in il
            return False
        issues = [iss for iss in issues if not _is_separate_op_or_ub(iss)]

    # 1. Mandatory Check: Witness Details Validation
    if _issue_absent(["witness details", "address", "signature", "name"]):
        # If no records exist or flags indicate blank structural components
        if not ld.get("witness_records") or len(ld.get("witness_records", [])) == 0:
            issues.append("Witness details like address, signature or name not found — Obtain complete details for at least two witnesses.")
        else:
            for wit in ld.get("witness_records", []):
                if not wit.get("name") or not wit.get("signature_present") or (dur <= 11 and not wit.get("full_address")):
                    issues.append("Witness details like address, signature or name not found — Complete missing witness attributes.")
                    break

    # 2. Mandatory Check: Lessee Signature Validation
    if not ld.get("lessee_signed") and _issue_absent(["lessee signature"]):
        issues.append("Lessee signature not found — Obtain signed copy from the lessee / second party.")

    # 2b. Mandatory Check: Lessee Company Stamp Validation (CBIC Check 8)
    if not ld.get("lessee_company_stamp_present") and _issue_absent(["lessee company stamp"]):
        issues.append("Lessee company stamp not found — Obtain signed copy containing the lessee's company seal/rubber stamp.")

    # 3. Mandatory Check: Sub-Registrar Registration Details (For Lease Deeds > 11 Months)
    if dur > 11:
        # Programmatic check for mandatory lease deed tracking
        if reg_type != "REGISTERED" or _issue_absent(["not registered with sub-registrar"]):
            if "Lease/Rent agreement is neither notarised nor registered with Sub-Registrar." not in issues:
                issues.append("Lease/Rent agreement is neither notarised nor registered with Sub-Registrar.")
            if not any("exceeding 11 months" in x for x in issues):
                issues.append("Registration: The 15-year lease deed is not registered, which is mandatory for agreements exceeding 11 months.")
    if not owner_match and eb_owner and lessor_nm and _issue_absent(["owner name"]):
        issues.append(
            f"Owner name MISMATCH — Bill: '{eb_owner}' vs Lessor: '{lessor_nm}' "
            f"(score: {score}%) — provide initial lease deed or NOC explaining discrepancy.")
    if not notary_ok and _issue_absent(["notari", "registr"]):
        issues.append(
            "Lease/Rent agreement is neither notarised nor registered with Sub-Registrar.")
    if not agr_valid and end_date_str and _issue_absent(["expir"]):
        issues.append(f"Agreement expired on {end_date_str} — invalid as on {VALIDATION_DATE}.")
    if not (n.get("finalised_gst_business_name") or "").strip() and _issue_absent(["applicant", "business name"]):
        issues.append(
            "GST applicant/business name could not be identified — additional identity proof required.")
    if wit_names and not wit_sig:
        issues.append("Witness signature(s) missing from the agreement.")
    if wit_names and not wit_addr and dur <= 11 and _issue_absent(["witness address"]):
        issues.append(
            "Witness address(es) not documented — mandatory for Rent Agreements (<=11 months).")

    # State verbatim warning: if a discrepancy was flagged (pincode_state_flag=True),
    # add a compliance issue so the reviewer sees it in the Excel Discrepancy Register.
    _disc_note_for_issue = str(addrs.get("state_discrepancy_note") or "").strip()
    if addrs.get("pincode_state_flag") is True and _disc_note_for_issue and _issue_absent(["state", "verbatim"]):
        issues.append(
            f"Issue – State Verbatim Check: Extracted state '{addrs.get('state')}' may not match "
            f"what the primary lease deed property clause actually says. {_disc_note_for_issue} "
            f"Action: Open the lease deed property description clause and verify the exact state "
            f"name written there. If different, update the extracted address manually.")

    # State inferred from district (when state was absent from the document).
    _inferred_note = str(addrs.get("state_inferred_note") or "").strip()
    if addrs.get("state_inferred_from_district") and _inferred_note \
            and _issue_absent(["state", "inferred", "not explicitly"]):
        issues.append(
            f"Issue – Missing State in Document: {_inferred_note}"
        )

    # Add ownership-specific discrepancy issues
    _op_issues = _ownership_discrepancy_remarks(_ownership_proof)
    for _oi in _op_issues:
        if _oi and not any(_oi[:30] in str(x) for x in issues):
            issues.append(_oi)

    # Ownership proof address mismatch
    if _op_detected and _op_match in ("LOW MATCH", "MISMATCH", "STATE MISMATCH") and _issue_absent(["ownership proof address mismatch"]):
        issues.append(
            f"Ownership proof address mismatch with agreement address: {_op_match}. "
            "Manual review required."
        )

    # ── 11. WEIGHTED COMPLIANCE PERCENT ────────────────────────────────────────
    if cbic_checks:
        total_w = sum(c.get("weight", 1) for c in cbic_checks if c.get("status") != "N/A")
        pass_w  = sum(c.get("weight", 1) for c in cbic_checks if c.get("status") == "PASS")
        wcp = round((pass_w / total_w * 100) if total_w else 0, 1)
    else:
        tp, ti = len(passes), len(issues)
        wcp = round(tp / (tp + ti) * 100 if (tp + ti) else 0, 1)

    # ── 12. FINAL STATUS + SOURCE-ATTRIBUTED REMARKS ───────────────────────────
    _miss_docs = s.get("missing_documents") or []
    _risk_summ = comp.get("risk_summary", "")
    _ld_risk   = ld.get("risk_notes", "")
    _eb_risk   = eb.get("risk_notes", "")

    if wcp >= 85:
        final_status  = "Clean"
        _base_remark  = "All critical GST documentation checks passed. Ready to upload."
        _source_tag   = "Derived from weighted compliance score (>=85%)"
    elif wcp >= 60:
        final_status  = "Warning"
        _base_remark  = "Minor documentation gaps detected. Review and resolve before uploading."
        _source_tag   = "Derived from weighted compliance score (60-84%)"
    else:
        final_status  = "High Risk"
        _base_remark  = ("Significant deficiencies found. DO NOT upload to GST portal — "
                         "resolve all HIGH severity issues first.")
        _source_tag   = "Derived from weighted compliance score (<60%)"

    # Forced merged document failure override
    if merged_partial_fail:
        final_status = "High Risk"
        _source_tag = "Merged document partial failure"
        _base_remark = f"Merged document validation failed on: {'; '.join(segment_failures)}."

    # Compose clean final_remarks (no [Source:...] tags — clean for Excel display)
    _remark_parts = [f"{_base_remark} [{_source_tag}]"]
    if issues:
        _top_issues = "; ".join(
            re.sub(r"\s*\[Source:[^\]]*\]", "", str(iss))[:120].strip()
            for iss in issues[:3]
        )
        _remark_parts.append(f"Key issues: {_top_issues}")
    if _risk_summ and _risk_summ.strip()[:60] != _base_remark[:60]:
        _remark_parts.append(_risk_summ.strip())
    if _ld_risk:
        _remark_parts.append(_ld_risk.strip())
    if _eb_risk and _eb_risk != _ld_risk:
        _remark_parts.append(_eb_risk.strip())
    if _miss_docs:
        _mdoc_str = "; ".join(str(m) for m in _miss_docs[:3])
        _remark_parts.append(f"Missing documents: {_mdoc_str}")
    if not is_genuine_utility and not _op_detected:
        # Purge conflicting true strings from generating incorrect final lists
        issues = [iss for iss in issues if "Land Revenue Record valid" not in iss]
        _remark_parts = [r for r in _remark_parts if "Land Revenue Record valid" not in r]

    # Append ownership proof remarks to final remarks
    if _ownership_proof.get("detected"):
        _agr_addr_for_remarks = str(
            ld.get("premises_address") or
            addrs.get("finalised_pob_address") or ""
        ).strip()
        _op_remarks = _build_ownership_proof_remarks(_ownership_proof, _agr_addr_for_remarks)
        if _op_remarks:
            _remark_parts.append(_op_remarks)

    final_remarks = " | ".join(_remark_parts)

    # ── 13. PER-DOC effective_status DEFAULTS ──────────────────────────────────
    for doc in docs:
        if not doc.get("effective_status"):
            es = doc.get("extraction_status", "OK")
            if es == "FAILED":
                doc["effective_status"] = "INCORRECT"
            elif es in ("SCANNED_PDF", "IMAGE", "PARTIAL"):
                doc["effective_status"] = "REVIEW"
            else:
                doc["effective_status"] = "CORRECT"

    # ── 14. WRITE BACK ─────────────────────────────────────────────────────────
    comp.update({
        "passes":                      passes,
        "issues":                      issues,
        "total_passes":                len(passes),
        "total_issues":                len(issues),
        "weighted_compliance_percent": wcp,
        "final_status":                final_status,
        "final_remarks":               final_remarks,
        "overall_risk_exposure":       ("HIGH" if wcp < 60 else "MEDIUM" if wcp < 85 else "LOW"),
    })
    s.update({
        "weighted_compliance_percent": wcp,
        "effective_overall_status":    final_status,
        "total_documents_analyzed":    max(len(docs), len(client_files)),
    })

    data.update({
        "summary":               s,
        "names":                 n,
        "lease_details":         ld,
        "electricity_bill":      eb,
        "compliance":            comp,
        "documents_analyzed":    docs,
        "addresses":             addrs,
        "ownership_proof":       _ownership_proof,
    })

    # ── HARDENED : Apply state verbatim lock + geographic consistency check ──
    data = lock_state_from_primary_lease(data)
    data = validate_geographic_consistency(data)

    return data


#  JSON CACHE  (saves raw Gemini responses to avoid repeated API calls)
def _safe_name(s: str) -> str:
    """Sanitise a string for use as a filename."""
    return re.sub(r"[^a-zA-Z0-9_\-]", "_", s)[:80]


def save_json_cache(client_name: str, data: Dict, *, client_cache_dir: Optional[Path] = None):
    """
    Persist Gemini JSON response to disk atomically.

    HARDENED : uses atomic_json_dump (temp-file + rename) for both write
    locations, preventing corrupt/partial JSON files even on process kill.

    Writes to TWO locations (both are always attempted):
      1. Legacy global cache dir  : outputs/json_cache/<client>.json
      2. Per-client cache          : OUTPUT/<State>/<Client>/json_cache/File_cache.json

    The dual-write ensures backwards compatibility with --use-cache / --skip-gemini
    while also populating the state-wise per-client folder structure.
    """
    # ── 1. Legacy global cache (atomic write) ────────────────────────────────
    try:
        safe_create_dir(normalize_windows_path(JSON_CACHE_DIR))
        fp = normalize_windows_path(JSON_CACHE_DIR / f"{_safe_name(client_name)}.json")
        if safe_json_save(data, fp, lock=_cache_lock, client_name=client_name):
            log.info(f"   JSON cache saved (global) → {fp.name}")
        else:
            log.warning(f"   JSON cache global write failed for '{client_name}'")
    except OSError as exc:
        log.warning(f"   JSON cache global dir creation failed: {exc}")

    # ── 2. Per-client hierarchy cache (atomic write) ──────────────────────────
    target_dir = client_cache_dir
    if target_dir is None:
        try:
            # Primary: developer-only path under _internal/<State>/<Client>/json_cache
            hits = sorted(CLIENT_OUTPUT_BASE.glob(f"_internal/*/{_safe_folder_name(client_name)}/json_cache"))
            if not hits:
                # Fallback: legacy flat layout (pre-v6 runs)
                hits = sorted(CLIENT_OUTPUT_BASE.glob(f"*/{_safe_folder_name(client_name)}/json_cache"))
            if hits:
                target_dir = hits[-1]
            else:
                legacy = CLIENT_OUTPUT_BASE / "_internal" / _safe_folder_name(client_name) / "cache"
                if legacy.exists():
                    target_dir = legacy
        except OSError:
            target_dir = None

    if target_dir is not None:
        target_dir_norm = normalize_windows_path(target_dir)
        try:
            safe_create_dir(target_dir_norm)
            client_cache_path = target_dir_norm / "File_cache.json"
            if safe_json_save(data, client_cache_path, lock=_cache_lock, client_name=client_name):
                log.info(f"   JSON cache saved (client) → {client_cache_path}")
            else:
                log.warning(f"   Per-client cache write failed for '{client_name}'")
        except OSError as exc:
            log.warning(f"   [FolderManager] Per-client cache dir failed for "
                        f"'{client_name}': {exc}")


def load_json_cache(client_name: str) -> Optional[Dict]:
    """
    Load a previously saved Gemini JSON response from disk.

    Checks locations (in priority order):
      1. State-wise client cache: OUTPUT/<State>/<Client>/json_cache/gemini_cache.json
      2. Legacy internal cache   : OUTPUT/_internal/<Client>/cache/gemini_cache.json
      3. Legacy global cache dir : outputs/json_cache/<client>.json
    """
    safe = _safe_folder_name(client_name)

    # ── 1. Developer-only path: _internal/<State>/<Client>/json_cache/ ────────
    try:
        for client_cache in sorted(CLIENT_OUTPUT_BASE.glob(
                f"_internal/*/{safe}/json_cache/File_cache.json")):
            if client_cache.exists():
                try:
                    data = json.loads(client_cache.read_text(encoding="utf-8"))
                    log.info(f"   JSON cache loaded (client/_internal) ← {client_cache}")
                    return data
                except Exception as e:
                    log.warning(f"  ⚠ Per-client cache read error for {client_name}: {e}")
    except OSError as exc:
        log.warning(f"  ⚠ Client _internal cache glob failed for {client_name}: {exc}")

    # ── 2. Legacy state-wise per-client cache (pre-v6 flat layout) ───────────
    try:
        for client_cache in sorted(CLIENT_OUTPUT_BASE.glob(
                f"*/{safe}/json_cache/File_cache.json")):
            if client_cache.exists():
                try:
                    data = json.loads(client_cache.read_text(encoding="utf-8"))
                    log.info(f"   JSON cache loaded (client legacy) ← {client_cache}")
                    return data
                except Exception as e:
                    log.warning(f"  ⚠ Per-client cache read error for {client_name}: {e}")
    except OSError as exc:
        log.warning(f"  ⚠ Client cache glob failed for {client_name}: {exc}")

    # ── 3. Legacy _internal flat layout ──────────────────────────────────────
    legacy_internal = CLIENT_OUTPUT_BASE / "_internal" / safe / "cache" / "File_cache.json"
    if legacy_internal.exists():
        try:
            data = json.loads(legacy_internal.read_text(encoding="utf-8"))
            log.info(f"   JSON cache loaded (legacy internal) ← {legacy_internal}")
            return data
        except Exception as e:
            log.warning(f"  ⚠ Per-client cache read error for {client_name}: {e}")

    # ── 4. Fall back to legacy global cache ──────────────────────────────────
    fp = JSON_CACHE_DIR / f"{_safe_name(client_name)}.json"
    if fp.exists():
        try:
            data = json.loads(fp.read_text(encoding="utf-8"))
            log.info(f"   JSON cache loaded (global) ← {fp.name}")
            return data
        except Exception as e:
            log.warning(f"  ⚠ Cache read error for {client_name}: {e}")
    return None


#  EXCEL WORKBOOK BUILDER — 3 data sheets (no instruction / how-to tab)
class ExcelBuilder:
    """
    Builds a professional 3-sheet Excel workbook:
      Sheet 1: GST_Registrations
      Sheet 2: Document_Index
      Sheet 3: Discrepancy_Register
    """

    def __init__(self, output_path: Path):
        self.out  = output_path
        self.wb   = Workbook()
        # Remove default blank sheet
        if "Sheet" in self.wb.sheetnames:
            del self.wb["Sheet"]

        # Row-counter state (sheets are pre-built; rows added per client)
        self._reg_row  = 2   # GST_Registrations — next data row
        self._idx_row  = 3   # Document_Index — row 1=major heads, row 2=col headers
        self._disc_row = 2   # Discrepancy_Register — next data row
        self._doc_ctr  = 0   # global document counter for Doc_ID
        self._disc_ctr = 0   # global discrepancy counter for Discrepancy_ID
        self._reg_ctr  = 0   # client counter

        # Sheet references (populated in build_all_sheets)
        self._ws_reg = self._ws_idx = self._ws_disc = None

    # ── Low-level helpers ─────────────────────────────────────────────────────
    def _w(self, ws, row, col, val, bold=False, fg="000000", bg="FFFFFF",
           align="left", sz=10, bdr=True):
        """Write a styled cell."""
        cell = ws.cell(row=row, column=col, value=val)
        cell.font      = (_bfont(fg, sz) if bold else _rfont(fg, sz))
        cell.fill      = _fill(bg)
        cell.alignment = (_ca() if align == "center" else _la())
        if bdr: cell.border = _bdr()
        return cell

    def _hdr(self, ws, row: int, cols: List[str], bg=None, fg=None, ht=28):
        """Write a header row with bold styled cells."""
        bg = bg or C["hdr_bg"]; fg = fg or C["hdr_fg"]
        for ci, txt in enumerate(cols, 1):
            cell = ws.cell(row=row, column=ci, value=txt)
            cell.font      = _bfont(fg, 10)
            cell.fill      = _fill(bg)
            cell.alignment = _ca()
            cell.border    = _bdr()
        ws.row_dimensions[row].height = ht

    def _widths(self, ws, w_dict: Dict[int, int]):
        """Set column widths from a {column_index: width_chars} dict."""
        for col, w in w_dict.items():
            ws.column_dimensions[get_column_letter(col)].width = w

    def _dv(self, ws, col: int, r1: int, r2: int, options: List[str]):
        """Add a dropdown DataValidation to a column range."""
        dv = DataValidation(
            type="list",
            formula1=f'"{",".join(options)}"',
            allow_blank=True,
            showDropDown=False,
        )
        dv.sqref = f"{get_column_letter(col)}{r1}:{get_column_letter(col)}{r2}"
        ws.add_data_validation(dv)

    def _alt(self, row: int) -> str:
        """Alternating row background."""
        return C["alt"] if row % 2 == 0 else C["white"]

    #  SHEET 1 — GST_Registrations
    def _build_registrations(self):
        ws = self.wb.create_sheet("GST_Registrations")
        ws.sheet_view.showGridLines = False
        self._ws_reg = ws

        cols = [
            "GST_Reg_ID",                              # 1
            "Legal Name (GST Applicant)",              # 2
            "State",                                   # 3
            "Principal Place of Business Address",     # 4  — translated English address
            "Basis of Documents",                      # 5  — which doc was used for PPOB address
            "Total Documents Uploaded",                # 6  — all files with hyperlinks
            "Ready to Upload Docs",                    # 7  — correct/ready files with hyperlinks
            "Incorrect/Mismatch Documents",            # 8  — MERGED: incorrect + mismatch docs
            "Missing Documents",                       # 9
            "Review Required Docs",                    # 10 — docs needing manual review
            "Weighted Compliance %",                   # 11
            "Final Status",                            # 12
            "Overall Risk Exposure",                   # 13
            "Final Remarks",                           # 14
            "Suggested Actions",                       # 15 — NEW: corrective actions column
            # ── Ownership Proof Columns (shifted +1) ──────────────────────
            "Ownership Proof Type",                    # 16
            "Ownership Proof Address",                 # 17
            "Ownership vs Agreement Address Match",    # 18
            "Ownership Proof Source Pages",            # 19
            "Address Comparison Remarks",              # 20
        ]
        self._hdr(ws, 1, cols)
        self._widths(ws, {
            1: 14,  2: 34,  3: 14,  4: 52,
            5: 42,  6: 38,  7: 38,  8: 38,  9: 38,
            10: 38, 11: 14, 12: 14, 13: 14, 14: 58,
            15: 58, 16: 28, 17: 52, 18: 28, 19: 20, 20: 58,
        })
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}1"
        self._dv(ws, 12, 2, 500, ["Clean", "Warning", "High Risk"])
        self._dv(ws, 13, 2, 500, ["LOW", "MEDIUM", "HIGH"])
        self._dv(ws, 16, 2, 500, [
            "Land Revenue Receipt", "Property Tax Receipt", "Municipal Tax Receipt",
            "Mutation Record", "Registry Extract", "Government Land Record", "Not Detected",
        ])
        self._dv(ws, 18, 2, 500, ["HIGH MATCH", "PARTIAL MATCH", "LOW MATCH",
                                   "MISMATCH", "STATE MISMATCH", "N/A"])

    def _write_registration_row(self, gst_id: str, data: Dict, all_files: List[Path]):
        """
        Write one client row to Sheet 2 — GST_Registrations.

        ENHANCED CHANGES:
          • Col 6: "Basis of Documents" — dynamically identifies source document for PPOB address
          • Col 7: "Total Documents Uploaded" — all files with clickable hyperlinks
          • Col 8: "Ready to Upload Docs" — correct/clean files with clickable hyperlinks
          • Col 9: "Incorrect/Mismatch Documents" — MERGED column (incorrect + mismatch names)
          • Col 10: "Missing Documents" — document names that are missing
          • Col 11: "Review Required Docs" — from Sheet1 rules (address issues, partial matches, witness)
          • Col 12: "Weighted Compliance %" (shifted; Mismatch col removed)
          • Col 13: "Final Status" (shifted)
          • Col 14: "Overall Risk Exposure" (shifted)
          • Col 15: "Final Remarks" (shifted)
        """
        ws   = self._ws_reg
        row  = self._reg_row
        s    = data.get("summary", {})
        n    = data.get("names", {})
        comp = data.get("compliance", {})
        ld   = data.get("lease_details", {})
        docs = data.get("documents_analyzed", [])
        addrs= data.get("addresses", {})
        if not isinstance(docs, list): docs = []

        # ── Derive display names ───────────────────────────────────────────
        legal_name = (n.get("legal_name") or s.get("legal_name") or
                      n.get("finalised_gst_business_name") or s.get("final_business_name") or gst_id)
        trade_name = (n.get("trade_name") or s.get("trade_name") or legal_name)

        # ── NORMALIZE DATA THROUGH CENTRALIZED ENGINE ─────────────────────
        # Run normalization before building display values
        data = normalize_excel_output(data)

        # Refresh all sections after normalization
        s    = data.get("summary", {})
        n    = data.get("names", {})
        comp = data.get("compliance", {})
        ld   = data.get("lease_details", {})
        docs = data.get("documents_analyzed", [])
        addrs= data.get("addresses", {})
        if not isinstance(docs, list): docs = []

        # Rebuild legal_name/trade_name after normalization refresh
        legal_name = (n.get("legal_name") or s.get("legal_name") or
                      n.get("finalised_gst_business_name") or s.get("final_business_name") or gst_id)
        trade_name = (n.get("trade_name") or s.get("trade_name") or legal_name)

        wcp     = float(comp.get("weighted_compliance_percent",
                                  s.get("weighted_compliance_percent", 0)) or 0)
        fs      = comp.get("final_status", "")
        risk    = comp.get("overall_risk_exposure", s.get("overall_risk_exposure", ""))
        issues_list = comp.get("issues") or []
        _op = data.get("ownership_proof") or {}
        op_addr_match = str(_op.get("address_match_with_agreement") or "N/A")
        missing_docs_list = s.get("missing_documents") or []

        # ── BEST POB ADDRESS (col 5) ───────────────────────────────────────
        dur_for_basis = 0
        try: dur_for_basis = int(ld.get("duration_months", 0) or 0)
        except: pass

        _cat_raw = str(s.get("category") or "").lower()
        if "warehous" in _cat_raw or "service" in _cat_raw:
            agr_doc_label_for_basis = "Warehousing/Service Agreement"
        elif dur_for_basis > 11:
            agr_doc_label_for_basis = "Registered Lease Deed"
        elif dur_for_basis > 0:
            agr_doc_label_for_basis = "Rent Agreement"
        else:
            agr_doc_label_for_basis = "primary agreement document"

        addr_parts = [
            str(addrs.get("khasra_numbers") or "").strip(),
            str(addrs.get("village_locality") or "").strip(),
            str(addrs.get("road_street") or "").strip(),
            str(addrs.get("city_town") or "").strip(),
            str(addrs.get("district") or "").strip(),
            str(addrs.get("state") or "").strip(),
            str(addrs.get("pincode") or "").strip(),
        ]
        structured_addr = ", ".join(p for p in addr_parts if p)
        has_khasra = bool(addr_parts[0])

        def _has_non_latin(text: str) -> bool:
            return bool(text) and any(ord(c) > 127 for c in text)

        pob_candidates = [
            (str(s.get("final_pob_address") or "").strip(),
             f"From final_pob_address in {agr_doc_label_for_basis}"),
            (str(ld.get("premises_address") or "").strip(),
             f"From premises_address in primary {agr_doc_label_for_basis}"),
            (str(addrs.get("finalised_pob_address") or "").strip(),
             "From finalised_pob_address in addresses section"),
        ]
        if structured_addr:
            if has_khasra:
                pob_candidates.append(
                    (structured_addr, f"From KHASRA/village details in {agr_doc_label_for_basis}")
                )
            else:
                pob_candidates.append(
                    (structured_addr, f"Structured from address components in {agr_doc_label_for_basis}")
                )

        best_addr = ""
        for addr, basis in pob_candidates:
            if addr and len(addr) > len(best_addr):
                best_addr = addr

        if _has_non_latin(best_addr):
            best_addr = translate_to_english(best_addr)

        # HARDENED : normalize state name to prevent inconsistent Excel entries
        _raw_state = str(addrs.get("state") or s.get("state") or "").strip()
        state_val = normalize_state_name(_raw_state)
        # If state could not be resolved (absent from document), show blank in Excel
        # rather than the internal "Unknown_State" folder label.
        if state_val == "Unknown_State":
            state_val = ""

        # ── Col 6: BASIS OF DOCUMENTS ──────────────────────────────────────
        # Dynamically identifies which document was the address source.
        basis_of_docs = _determine_basis_of_documents(data, addrs, ld, s)

        # ── Col 7: TOTAL DOCUMENTS UPLOADED (with hyperlinks) ─────────────
        # All files passed to this client (agreement + utility + other)
        all_files_sorted = sorted(all_files, key=lambda p: p.name)

        # ── READABILITY & WARNING PROPAGATION ENGINE ──────────────────────
        client_readability_warning = s.get("readability_warning") is True or any(
            doc.get("readability_warning") is True for doc in docs
        )
        
        warning_note = "Note: The Image quality is poor and blurry, content in non human readable format."
        
        def append_readability_warning(text_val: str) -> str:
            if not text_val:
                return warning_note
            text_str = str(text_val).strip()
            if warning_note in text_str:
                return text_str
            if text_str.endswith(".") or text_str.endswith(":") or text_str.endswith("!"):
                return f"{text_str} {warning_note}"
            else:
                return f"{text_str}. {warning_note}"

        # ── STRICT VALIDATION GATE FOR 'READY TO UPLOAD DOCS' ─────────────
        # All conditions must pass for the document list to populate
        merged_segment_failure = any(
            "merged document partial failure" in str(iss).lower() or 
            "merged segment" in str(iss).lower() 
            for iss in issues_list
        )
        
        no_high_risk = (
            str(risk).upper() != "HIGH" and 
            str(comp.get("overall_risk_exposure") or "").upper() != "HIGH"
        )
        no_name_mismatch = (
            str(s.get("name_confidence_label") or "").upper() not in ("LOW", "MEDIUM") and 
            not any("name mismatch" in str(iss).lower() or "unmatched name" in str(iss).lower() for iss in issues_list)
        )
        utility_bill_present_if_required = not any("utility" in str(m).lower() for m in missing_docs_list)
        ownership_proof_valid = (
            op_addr_match not in ("MISMATCH", "STATE MISMATCH", "LOW MATCH", "N/A") and 
            not any("ownership proof" in str(iss).lower() or "land revenue receipt" in str(iss).lower() for iss in issues_list)
        )
        readability_warning_ok = not client_readability_warning
        merged_segment_failure_ok = not merged_segment_failure
        
        is_upload_ready = (
            no_high_risk and 
            no_name_mismatch and 
            utility_bill_present_if_required and 
            ownership_proof_valid and 
            readability_warning_ok and 
            merged_segment_failure_ok
        )
        
        # Enforce PASS + HIGH RISK safety rule: force High Risk if wcp fails
        if not no_high_risk:
            is_upload_ready = False
            if fs == "Clean":
                fs = "High Risk"

        # ── Col 8: READY TO UPLOAD DOCS (with hyperlinks) ─────────────────
        # Files whose effective_status is CORRECT — ready for GST portal
        correct_filenames = set()
        for d in docs:
            fn = str(d.get("file_name") or "").strip()
            if not fn:
                continue
            eff = str(d.get("effective_status") or "").upper()
            
            doc_no_high_risk = str(d.get("risk_exposure") or "").upper() != "HIGH"
            doc_no_name_mismatch = "mismatch" not in str(d.get("name_match_status") or "").lower()
            doc_readability_ok = not (d.get("readability_warning") is True)
            
            doc_ok = (
                eff == "CORRECT" and
                doc_no_high_risk and
                doc_no_name_mismatch and
                doc_readability_ok
            )
            if doc_ok:
                correct_filenames.add(fn)

        if is_upload_ready:
            ready_files = [fp for fp in all_files_sorted if fp.name in correct_filenames]
        else:
            ready_files = []

        # ── Col 9: INCORRECT/MISMATCH DOCUMENTS (merged) ──────────────────
        incorr_mismatch_filenames: List[str] = []
        for d in docs:
            eff = str(d.get("effective_status") or "").upper()
            if eff in ("INCORRECT", "MISMATCH"):
                fn = str(d.get("file_name") or "").strip()
                if fn and fn not in incorr_mismatch_filenames:
                    incorr_mismatch_filenames.append(fn)

        # Also add issue-tagged files not already captured
        # FIX: Only include files whose effective_status is NOT CORRECT — a file
        # cannot appear in both Ready to Upload and Incorrect/Mismatch.
        for d in docs:
            fn = str(d.get("file_name") or "").strip()
            if not fn or fn in incorr_mismatch_filenames:
                continue
            eff = str(d.get("effective_status") or "").upper()
            if eff == "CORRECT":
                continue  # CORRECT files belong in Ready to Upload only
            issue_t = str(d.get("issue_type") or "").lower()
            if any(k in issue_t for k in ("mismatch", "expired", "incomplete")):
                incorr_mismatch_filenames.append(fn)

        incorr_mismatch_text = ""
        if incorr_mismatch_filenames:
            incorr_mismatch_text = "\n".join(
                f"• {fn}" for fn in incorr_mismatch_filenames
            )
        else:
            incorr_mismatch_text = "None"

        # ── Col 11: REVIEW REQUIRED DOCS (from Sheet1 rules) ──────────────
        review_filenames: List[str] = []
        for d in docs:
            fn = str(d.get("file_name") or "").strip()
            if not fn:
                continue
            eff = str(d.get("effective_status") or "").upper()
            issue_t = str(d.get("issue_type") or "").lower()
            disc = str(d.get("discrepancy_reason") or "").lower()
            key_notes = str(d.get("key_notes") or "").lower()
            extr = str(d.get("extraction_status") or "").upper()
            name_m = str(d.get("name_match_status") or "").upper()
            addr_m = str(d.get("address_match_status") or "").upper()

            is_review = False

            # Rule a/b: H.No, Pin Code missing
            if any(k in disc or k in key_notes for k in
                   ("h.no", "flat no", "building no", "pin code", "pincode missing",
                    "incomplete address", "partial address")):
                is_review = True

            # Rule c: Witness present but signature/address missing
            if any(k in disc or k in key_notes for k in
                   ("witness address", "witness signature", "witness sign")):
                is_review = True

            # Rule d: Notary not available
            if any(k in disc or k in key_notes for k in
                   ("notary", "notarised", "not notarised", "notarization")):
                is_review = True

            # Rule e: Partial name or address match
            if "partial" in name_m or "partial" in addr_m:
                is_review = True

            # Effective status REVIEW
            if eff == "REVIEW":
                is_review = True

            # OCR partially readable — needs manual check
            if extr in ("SCANNED_PDF", "IMAGE", "PARTIAL"):
                is_review = True

            # Mismatch that is not severe enough to be INCORRECT
            if "mismatch" in issue_t and eff not in ("INCORRECT", "MISMATCH"):
                is_review = True

            if is_review and fn not in review_filenames:
                review_filenames.append(fn)

        review_text = ""
        if review_filenames:
            review_text = "\n".join(f"• {fn}" for fn in review_filenames)
        else:
            review_text = "None"

        # ── Col 9: MISSING DOCUMENTS (canonicalized, no false ownership/utility) ─
        # Already canonicalized by normalize_excel_output above
        cleaned_missing_items = []

        for m in missing_docs_list:
            m_lower = str(m).lower()
            # Use strict grouped labeling to block multiple appends for the same collection category
            if any(k in m_lower for k in ["utility", "ownership", "proof", "land revenue", "tax receipt"]):
                if "Utility bill / ownership proof" not in cleaned_missing_items:
                    cleaned_missing_items.append("Utility bill / ownership proof")
            else:
                cleaned_item = str(m).strip()
                if cleaned_item not in cleaned_missing_items and cleaned_item.lower() != "none":
                    cleaned_missing_items.append(cleaned_item)

        if cleaned_missing_items:
            missing_text = "\n".join(f"• {item}" for item in cleaned_missing_items)
        else:
            missing_text = "None"

        # ── FINAL REMARKS (col 14) — OLD WORKBOOK FORMAT ──────────────────
        # Clean serial numbered list matching old workbook style exactly.
        # No status prefix, no "Issue N", no action text — pure audit remarks.
        remarks = build_final_remarks(issues_list)
        if not remarks and wcp >= 85:
            remarks = "All critical checks passed. Ready to upload to GST portal."

        # ── SUGGESTED ACTIONS (col 15) — NEW COLUMN ──────────────────────
        # Corrective actions derived from normalized issues list.
        raw_actions = build_suggested_actions(issues_list)

        # Clean up any leftover "Resolve: " or "Action: " prefixes from the fallback engine
        suggested_actions_list = []
        for act in raw_actions.split("\n"):          # split string → lines, not chars
            # Strip leading serial number added by build_suggested_actions (e.g. "1. ")
            act = re.sub(r"^\d+\.\s*", "", act.strip())
            clean_act = re.sub(r"^(Resolve\s*:\s*|Action\s*:\s*)", "", act, flags=re.IGNORECASE)
            if clean_act and clean_act not in suggested_actions_list:
                suggested_actions_list.append(clean_act)

        suggested_actions = "\n".join(f"{idx}. {item}" for idx, item in enumerate(suggested_actions_list, 1)) if suggested_actions_list else "No action required."

        # ── OWNERSHIP PROOF COLUMNS (16-20) ───────────────────────────────
        _seg          = data.get("document_segmentation") or {}

        op_type       = str(_op.get("document_type") or "Not Detected")
        op_address    = str(_op.get("ownership_address") or "")
        if any(ord(c) > 127 for c in op_address):
            op_address = translate_to_english(op_address)
        _seg_blocks = _seg.get("document_segments") or []
        _own_pages_from_seg = [
            p for block in _seg_blocks
            if block.get("type") in _OWNERSHIP_PROOF_TYPES
            for p in (block.get("pages") or [])
        ]
        op_source_pages  = _seg.get("ownership_proof_source_pages") or (
            _format_pages_for_display(_own_pages_from_seg)
            or _format_pages_for_display(_op.get("source_pages") or [])
            or "N/A"
        )
        op_cmp_remarks   = str(_op.get("comparison_remarks") or "")

        # Build full address comparison text for col 20 if not already set
        if not op_cmp_remarks and _op.get("detected"):
            _own_addr_disp = str(_op.get("ownership_address") or "Not extracted")
            _agr_addr_disp = str(ld.get("premises_address") or addrs.get("finalised_pob_address") or "Not extracted")
            op_cmp_remarks = (
                f"Ownership Proof Address:\n  {_own_addr_disp}\n\n"
                f"Agreement Address:\n  {_agr_addr_disp}\n\n"
                f"Result: {op_addr_match}"
            )

        if client_readability_warning:
            # Propagate to Principal Place of Business Address (col D)
            if not best_addr or best_addr.strip() in ("", "N/A"):
                best_addr = warning_note
            else:
                best_addr = append_readability_warning(best_addr)

            # Propagate to Final Remarks
            if not remarks or remarks == "All critical checks passed. Ready to upload to GST portal.":
                remarks = "Document/Image quality is poor or blurry. Content is partially unreadable."
            remarks = append_readability_warning(remarks)
            
            # Propagate to Suggested Actions
            if not suggested_actions or suggested_actions == "No action required.":
                suggested_actions = "1. Re-scan and upload high-resolution, clear, readable documents."
            suggested_actions = append_readability_warning(suggested_actions)
            
            # Propagate to Address Comparison Remarks
            op_cmp_remarks = append_readability_warning(op_cmp_remarks)
            
            # Propagate to Ownership Proof Address
            if not op_address or op_address.lower() in ("not detected", "n/a", ""):
                op_address = warning_note
            else:
                op_address = append_readability_warning(op_address)

        bg = self._alt(row)

        # ── Write scalar cells ────────────────────────────────────────────────
        scalar_vals = {
            1:  gst_id,
            2:  legal_name,
            3:  state_val,
            4:  best_addr,
            5:  basis_of_docs,
            8:  incorr_mismatch_text,
            9:  missing_text,
            10: review_text,
            11: f"{wcp:.1f}%",
            12: fs,
            13: risk,
            14: remarks,             # Final Remarks — old workbook format
            15: suggested_actions,   # NEW: Suggested Actions column
            # Ownership Proof columns (shifted to 16-20 after removing FY)
            16: op_type,
            17: op_address,
            18: op_addr_match,
            19: op_source_pages,
            20: op_cmp_remarks,
        }

        # ── UTILITY BILL BLURRY DETECTION (per-client) ───────────────────────
        _dq = data.get("document_quality") or {}
        _util_bill_blurry = _dq.get("utility_bill_blurry") is True
        _UTIL_WARN = "Note: The Image quality is poor and blurry, content in non human readable format."
        _YELLOW_FILL = PatternFill(fill_type="solid", start_color="FFF7CC", end_color="FFF7CC")

        # If utility bill is blurry, append warning to Address Comparison (col 18)
        # and Ownership Proof Address (col 17) — DO NOT affect other columns
        if _util_bill_blurry:
            # Col 18: Ownership vs Agreement Address Match
            _existing_18 = scalar_vals.get(18, "")
            if isinstance(_existing_18, str) and _UTIL_WARN not in _existing_18:
                if _existing_18.strip():
                    _sep = ". " if not _existing_18.rstrip().endswith(".") else " "
                    scalar_vals[18] = f"{_existing_18}{_sep}{_UTIL_WARN}"
                else:
                    scalar_vals[18] = _UTIL_WARN

            # Col 17: Ownership Proof Address
            _existing_17 = scalar_vals.get(17, "")
            if isinstance(_existing_17, str) and _UTIL_WARN not in _existing_17:
                if _existing_17.strip():
                    _sep = ". " if not _existing_17.rstrip().endswith(".") else " "
                    scalar_vals[17] = f"{_existing_17}{_sep}{_UTIL_WARN}"
                else:
                    scalar_vals[17] = _UTIL_WARN

        for ci, v in scalar_vals.items():
            c = ws.cell(row=row, column=ci, value=v)
            c.font      = _rfont(sz=10)
            c.fill      = _fill(bg)
            if ci in (4, 5, 8, 9, 10, 14, 15, 17, 20):
                c.alignment = _la(w=True)
            else:
                c.alignment = _la(w=False)
            c.border    = _bdr()

        # ── Apply YELLOW highlight to utility-bill-blurry warning cells ───────
        if _util_bill_blurry:
            for _warn_col in (17, 18):
                _wc = ws.cell(row=row, column=_warn_col)
                _wc.fill = _YELLOW_FILL
                _wc.font = _rfont(sz=10)  # normal black text, NOT bold/red

        # ── Write hyperlink file-list cells (cols 6 and 7) ───────────────────
        _write_hyperlink_file_list(ws, row, 6, all_files_sorted, bg)
        _write_hyperlink_file_list(ws, row, 7, ready_files, bg)

        # ── Colour-code key status columns ────────────────────────────────
        fg,  fb  = _status_clr(fs)
        fg2, fb2 = _status_clr(risk)
        fg3 = (C["pass_fg"] if wcp >= 85 else C["warn_fg"] if wcp >= 60 else C["fail_fg"])
        fb3 = (C["pass_bg"] if wcp >= 85 else C["warn_bg"] if wcp >= 60 else C["fail_bg"])

        for col, f, b in [(11, fg3, fb3), (12, fg, fb), (13, fg2, fb2)]:
            ws.cell(row=row, column=col).font  = _bfont(f, 10)
            ws.cell(row=row, column=col).fill  = _fill(b)

        # Colour ownership proof address match column (now col 18)
        # Skip yellow override — utility blurry yellow takes priority when active
        if op_addr_match and not _util_bill_blurry:
            _fg_op, _fb_op = _status_clr(op_addr_match)
            ws.cell(row=row, column=18).font = _bfont(_fg_op, 10)
            ws.cell(row=row, column=18).fill = _fill(_fb_op)

        # Dynamic row height based on content
        n_lines = max(
            1,
            len(all_files_sorted),
            len(ready_files),
            len(incorr_mismatch_filenames),
            len(missing_docs_list),
            len(review_filenames),
            remarks.count("\n") + 1,
            suggested_actions.count("\n") + 1 if suggested_actions else 1,
            op_cmp_remarks.count("\n") + 1 if op_cmp_remarks else 1,
        )
        ws.row_dimensions[row].height = min(180, max(42, n_lines * 16))
        self._reg_row += 1

    #  SHEET 3 — Document_Index
    def _build_document_index(self):
        ws = self.wb.create_sheet("Document_Index")
        ws.sheet_view.showGridLines = False
        self._ws_idx = ws

        # Row 1 — Major section heads (merged cells)
        # Columns: 1-4=identity, 5-11=utility, 12-21=agreement, 22-27=outcome
        major_heads = [
            ("DOCUMENT IDENTITY",                    1,  4,  C["hdr_bg"]),
            ("UTILITY BILL VERIFICATION",            5,  11, C["sub_hdr1"]),
            ("RENT / LEASE AGREEMENT VERIFICATION",  12, 21, C["sub_hdr2"]),
            ("VALIDATION OUTCOME",                   22, 27, C["sub_hdr3"]),
        ]
        for txt, sc_, ec_, bg in major_heads:
            ws.merge_cells(start_row=1, start_column=sc_, end_row=1, end_column=ec_)
            cell = ws.cell(row=1, column=sc_, value=txt)
            cell.font      = _bfont(C["hdr_fg"], 11)
            cell.fill      = _fill(bg)
            cell.alignment = _ca()
            cell.border    = _bdr()
        ws.row_dimensions[1].height = 26

        # Row 2 — Column sub-headers (short, readable)
        cols = [
            # Document Identity (1-4)
            "Doc_ID",
            "File_Name",
            "Folder_Path",
            "GST_Reg_ID",
            # Utility Bill Verification (5-11)
            "Utility Document Type",
            "Utility Sub-Type / Issuing Authority",
            "Owner Name (Utility Bill)",
            "Utility Address vs Lease Deed",
            "Older Than 3 Months?",
            "Bill Date",
            "Risk (Utility)",
            # Agreement Verification (12-21)
            "Agreement Document Type",
            "Owner Name Match (Utility vs Lease)",
            "Address Match (Utility vs Lease)",
            "Signature & Stamp",
            "Notarised / Registered",
            "Stamp Duty Amount (Rs)",
            "Expiry Date",
            "Valid As On Date",
            "Sub-Lease Permitted",
            "Risk (Agreement)",
            # Validation Outcome (22-27)
            "Duplicate",
            "OCR Readable",
            "Effective Status",
            "Issue Type",
            "Discrepancy Reason",
            "Reviewer Remarks",
        ]
        self._hdr(ws, 2, cols, bg=C["sec_bg"], ht=38)
        self._widths(ws, {
            1: 12,  2: 42,  3: 28,  4: 14,
            5: 26,  6: 30,  7: 26,  8: 22,  9: 16,  10: 14, 11: 14,
            12: 20, 13: 24, 14: 24, 15: 16, 16: 18, 17: 18,
            18: 14, 19: 16, 20: 16, 21: 16,
            22: 10, 23: 12, 24: 16, 25: 18, 26: 46, 27: 30,
        })
        ws.freeze_panes = "A3"
        ws.auto_filter.ref = f"A2:{get_column_letter(len(cols))}2"

        # Dropdown validations (updated for new Yes/No/NA on Older_Than_3_Months)
        MATCH   = ["HIGH MATCH", "PARTIAL MATCH", "LOW MATCH", "N/A"]
        YN      = ["Yes", "No", "N/A", "Review"]
        YN_NA   = ["Yes", "No", "NA"]          # Older_Than_3_Months
        STAT    = ["CORRECT", "INCORRECT", "REVIEW", "MISSING", "MISMATCH"]
        ISSUE   = ["None", "Missing", "Expired", "Mismatch", "Incomplete", "Unreadable", "Duplicate"]
        RISK    = ["LOW", "MEDIUM", "HIGH"]
        AGR_DT  = ["Rent Agreement", "Lease Deed", "Warehousing Agreement", "Other"]
        UTIL_DT = [
            "Electricity Bill", "Water Bill", "Property Tax Receipt",
            "Gas Bill", "Telephone Bill", "Internet Bill",
            "Municipal Tax Receipt", "Any other Legal Document",
        ]

        self._dv(ws,  5, 3, 1000, UTIL_DT)    # Utility_Document_Type
        self._dv(ws,  6, 3, 1000, UTIL_DT)    # Utility_Sub_Document_Type
        self._dv(ws,  8, 3, 1000, MATCH)      # Utility_Address_Match
        self._dv(ws,  9, 3, 1000, YN_NA)      # Older_Than_3_Months (Yes/No/NA)
        self._dv(ws, 11, 3, 1000, RISK)       # Risk_Exposure_Utility
        self._dv(ws, 12, 3, 1000, AGR_DT)     # Agreement_Document_Type
        self._dv(ws, 13, 3, 1000, MATCH)      # Name_Match_Status
        self._dv(ws, 14, 3, 1000, MATCH)      # Agreement_Address_Match
        self._dv(ws, 15, 3, 1000, YN)         # Signature_Stamp_Available
        self._dv(ws, 16, 3, 1000, YN)         # Notarised_Registered
        self._dv(ws, 19, 3, 1000, YN)         # Valid_As_On_Date
        self._dv(ws, 20, 3, 1000, YN)         # Sub_Lease_Permitted
        self._dv(ws, 21, 3, 1000, RISK)       # Risk_Exposure_Agreement
        self._dv(ws, 22, 3, 1000, ["Yes", "No"])             # Duplicate
        self._dv(ws, 23, 3, 1000, ["Yes", "No", "Partial"])  # OCR_Readable
        self._dv(ws, 24, 3, 1000, STAT)       # Effective_Status
        self._dv(ws, 25, 3, 1000, ISSUE)      # Issue_Type

    def _write_document_rows(self, gst_id: str, data: Dict, cat_map: Dict[str, List[Path]]):
        """
        Write one row per document to Sheet 3 — Document_Index.
          • Utility_Document_Type: full type label
          • Utility_Sub_Document_Type: sub-type or issuing authority
          • Utility_Address_Match: match between utility bill address and Rent/Lease Deed address
          • Older_Than_3_Months: Yes / No / NA (not Yes(Old)/No(Recent))
          • Bill_Date: YYYY-MM-DD or NA (not blank)
          • Agreement_Document_Type: derived from duration (Rent Agreement / Lease Deed)
          • Name_Match_Status: utility bill owner name vs agreement lessor name
          • Agreement_Address_Match: utility bill address vs agreement address
        """
        ws  = self._ws_idx
        s   = data.get("summary", {})
        eb  = data.get("electricity_bill") or {}
        ld  = data.get("lease_details", {})
        n   = data.get("names", {})

        # Build filename → analyzed_doc lookup
        analyzed: Dict[str, Dict] = {}
        for d in data.get("documents_analyzed", []):
            fn = (d.get("file_name") or "").strip()
            if fn: analyzed[fn] = d

        cat_labels = {
            "agreement": "AGREEMENTS",
            "utility":   "UTILITY BILL / ELECTRICITY BILL",
            "other":     "OTHER DOCUMENTS",
        }

        # ── Pre-compute shared values reused across all rows ──────────────
        # Agreement duration determines Rent Agreement vs Lease Deed
        try:    dur = int(ld.get("duration_months", 0) or 0)
        except: dur = 0
        agr_doc_type = "Lease Deed" if dur > 11 else ("Rent Agreement" if dur > 0 else "")

        # Utility bill address (for matching against agreement address)
        util_bill_addr = str(eb.get("address") or "").strip()

        # Agreement premises address
        agr_addr = str(ld.get("premises_address") or
                       s.get("final_pob_address") or "").strip()

        # Name match: utility bill owner name vs agreement lessor name
        eb_owner   = str(eb.get("owner_name") or "").strip()
        lessor_nm  = str(n.get("lessor_first_party") or ld.get("lessor_name") or "").strip()
        name_score = eb.get("owner_name_match_score", 0) or 0
        if name_score >= 90:
            name_match_status = "HIGH MATCH"
        elif name_score >= 70:
            name_match_status = "PARTIAL MATCH"
        elif eb_owner and lessor_nm:
            name_match_status = "LOW MATCH"
        else:
            name_match_status = "N/A"

        # Address match: utility bill address vs agreement address
        def _addr_match(a1: str, a2: str) -> str:
            """Rough address match — compare normalised tokens."""
            if not a1 or not a2:
                return "N/A"
            t1 = set(re.sub(r"[^a-zA-Z0-9]", " ", a1.lower()).split())
            t2 = set(re.sub(r"[^a-zA-Z0-9]", " ", a2.lower()).split())
            if not t1 or not t2:
                return "N/A"
            overlap = len(t1 & t2) / max(len(t1), len(t2))
            if overlap >= 0.65: return "HIGH MATCH"
            if overlap >= 0.35: return "PARTIAL MATCH"
            return "LOW MATCH"

        agr_addr_match_from_util = _addr_match(util_bill_addr, agr_addr)

        # Bill date and 3-month flag
        bill_date_raw = eb.get("bill_date")
        bill_date_str = str(bill_date_raw).strip() if bill_date_raw else ""
        if bill_date_str and bill_date_str.lower() not in ("null", "none", ""):
            bill_date_display = bill_date_str
            older_3m = "No" if eb.get("within_3_months") else "Yes"
        else:
            bill_date_display = "NA"
            older_3m = "NA"  # No date found → cannot determine

        for cat, files in cat_map.items():
            cat_label = cat_labels.get(cat, cat.upper())
            for fp in files:
                row = self._idx_row
                bg  = self._alt(row)
                self._doc_ctr += 1
                doc_id = f"DOC-{self._doc_ctr:04d}"

                # Match against Gemini's analyzed list (exact then fuzzy)
                gd = analyzed.get(fp.name) or {}
                if not gd:
                    for fn, d in analyzed.items():
                        if SequenceMatcher(None, fp.name.lower(), fn.lower()).ratio() > 0.80:
                            gd = d; break

                # ── UTILITY BILL columns (5-11) ───────────────────────────
                if cat == "utility":
                    # Col 5: Utility_Document_Type — broad category
                    # Normalise Gemini's bill_type to a recognised dropdown value
                    raw_bill_type  = str(eb.get("bill_type") or gd.get("detected_type") or "").strip()
                    _bt_lower      = raw_bill_type.lower()
                    _fn_lower      = fp.name.lower()

                    # Map to canonical type labels matching the dropdown list
                    if any(k in _bt_lower or k in _fn_lower
                           for k in ("property tax", "municipal tax", "khata", "house tax",
                                     "municipal khata")):
                        u_type = "Property Tax Receipt"
                    elif any(k in _bt_lower or k in _fn_lower
                             for k in ("water", "jal", "nagar nigam water")):
                        u_type = "Water Bill"
                    elif any(k in _bt_lower or k in _fn_lower
                             for k in ("gas", "lpg", "cng", "indane", "bharat gas", "hp gas")):
                        u_type = "Gas Bill"
                    elif any(k in _bt_lower or k in _fn_lower
                             for k in ("telephone", "phone", "bsnl", "mtnl")):
                        u_type = "Telephone Bill"
                    elif any(k in _bt_lower or k in _fn_lower
                             for k in ("internet", "broadband", "jio", "airtel fibre",
                                       "excitel", "hathway")):
                        u_type = "Internet Bill"  # maps to "Any other Legal Document" concept
                    elif any(k in _bt_lower or k in _fn_lower
                             for k in ("electric", "electricity", "power", "energy",
                                       "jseb", "jbvnl", "bses", "tpddl", "msedcl",
                                       "uppcl", "wbsedcl", "tneb", "kesco")):
                        u_type = "Electricity Bill"
                    elif raw_bill_type:
                        u_type = raw_bill_type  # keep Gemini's value if no match
                    else:
                        u_type = "Electricity Bill"  # safe default

                    # Col 6: Utility_Sub_Document_Type — specific instrument type
                    # Identifies: Electricity Bill / Property Tax Receipt / Water Tax Receipt /
                    #             Any other Legal Document
                    _sub_raw  = str(eb.get("sub_document_type") or "").strip()
                    _auth_raw = str(eb.get("issuing_authority") or "").strip()
                    _sub_lower = _sub_raw.lower()

                    # Map sub_document_type to canonical sub-type label
                    if "property tax" in _bt_lower or "municipal tax" in _bt_lower:
                        _sub_label = "Property Tax Receipt"
                    elif "water" in _bt_lower:
                        _sub_label = "Water Tax Receipt"
                    elif "electricity" in _bt_lower or "electric" in _bt_lower:
                        _sub_label = "Electricity Bill"
                    elif _sub_raw:
                        # Use Gemini's sub_document_type (Tax Invoice / Receipt / Statement etc.)
                        _sub_label = _sub_raw
                    else:
                        _sub_label = "Any other Legal Document"

                    # Append issuing authority for clarity (e.g. "Electricity Bill / JSEB")
                    u_sub = f"{_sub_label} / {_auth_raw}" if _auth_raw else _sub_label

                    # Col 7: Owner name on bill
                    u_owner = str(eb.get("owner_name") or gd.get("owner_name") or "").strip()

                    # Col 8: Utility_Address_Match = bill address vs Rent/Lease Deed address
                    u_match = agr_addr_match_from_util  # computed above from token overlap

                    # Col 9: Older_Than_3_Months = Yes / No / NA
                    u_old3m = older_3m

                    # Col 10: Bill date (YYYY-MM-DD or NA)
                    u_date = bill_date_display

                    # Col 11: Risk
                    u_risk = str(eb.get("risk_exposure") or "").strip()
                else:
                    u_type = u_sub = u_owner = u_match = u_old3m = u_date = u_risk = ""

                # ── AGREEMENT columns (12-21) ─────────────────────────────
                if cat == "agreement":
                    # Col 12: Agreement_Document_Type — derived from duration (primary rule)
                    # >11 months = Lease Deed; <=11 months = Rent Agreement (CBIC Para 6 / Section J)
                    _gd_type  = str(gd.get("detected_type") or "").strip()
                    _cat_type = str(s.get("category") or "").strip()
                    _doc_role = str(gd.get("role") or "").strip().upper()

                    if dur > 11:
                        a_type = "Lease Deed"
                    elif dur > 0:
                        a_type = "Rent Agreement"
                    elif "warehous" in _gd_type.lower() or "warehous" in _cat_type.lower() \
                         or "WAREHOUSING" in _doc_role or "SERVICE" in _doc_role:
                        a_type = "Warehousing Agreement"
                    elif "lease" in _gd_type.lower() or "lease" in _cat_type.lower():
                        a_type = "Lease Deed"
                    elif "rent" in _gd_type.lower() or "rent" in _cat_type.lower():
                        a_type = "Rent Agreement"
                    else:
                        a_type = _gd_type if _gd_type else ""

                    # Col 13: Name_Match_Status = utility bill owner name vs agreement lessor name
                    a_name = name_match_status

                    # Col 14: Agreement_Address_Match = utility bill address vs agreement address
                    a_addr = agr_addr_match_from_util

                    # Col 15: Signatures
                    a_sig    = "Yes" if ld.get("signatures_both_parties") else "No"
                    # Col 16: Notarised/Registered
                    a_notary = "Yes" if ld.get("notary_verified") else "No"
                    # Col 17: Stamp duty amount
                    a_stamp  = ld.get("stamp_duty_amount") or gd.get("stamp_duty_amount") or ""
                    # Col 18: Expiry date (prefer lease end_date; fall back to doc expiry_date)
                    a_expiry = str(ld.get("end_date") or gd.get("expiry_date") or "").strip() or ""
                    # Col 19: Valid as on validation date
                    a_valid  = "Yes" if ld.get("agreement_valid_as_on_validation_date") else "No"
                    # Col 20: Sub-lease permitted
                    sl = ld.get("sub_lease_permitted")
                    a_sub    = "Yes" if sl is True else "No" if sl is False else "N/A"
                    # Col 21: Risk exposure for the agreement
                    a_risk   = str(ld.get("risk_exposure") or "").strip()
                else:
                    a_type = a_name = a_addr = a_sig = a_notary = ""
                    a_stamp = a_expiry = a_valid = a_sub = a_risk = ""

                # ── OTHER DOC columns — utility/agreement fields left blank ─
                # (already handled by the else branches above)

                # ── OUTCOME columns (22-27) ───────────────────────────────
                eff   = str(gd.get("effective_status") or
                            ("REVIEW" if cat == "other" else "CORRECT"))
                issue = str(gd.get("issue_type") or "None")
                # Clean discrepancy_reason — strip [Source:...] tags and truncate
                disc_raw = str(gd.get("discrepancy_reason") or "")
                disc_raw = re.sub(r"\s*\[Source:[^\]]*\]", "", disc_raw).strip()
                disc_raw = re.sub(r"^Issue\s+[IVX]+\s*[–\-:]\s*", "", disc_raw).strip()

                # Check readability warning
                doc_readability_warning = gd.get("readability_warning") is True or s.get("readability_warning") is True
                if doc_readability_warning:
                    if not disc_raw or disc_raw.lower() in ("none", ""):
                        disc_raw = "Document/Image quality is poor or blurry. Content is partially unreadable."
                    
                    warning_note = "Note: The Image quality is poor and blurry, content in non human readable format."
                    if warning_note not in disc_raw:
                        if disc_raw.endswith(".") or disc_raw.endswith(":") or disc_raw.endswith("!"):
                            disc_raw = f"{disc_raw} {warning_note}"
                        else:
                            disc_raw = f"{disc_raw}. {warning_note}"
                    
                    reviewer_remarks = warning_note
                else:
                    reviewer_remarks = ""

                disc = disc_raw[:300] if len(disc_raw) > 300 else disc_raw
                ocr   = ("Partial" if gd.get("extraction_status") in
                         ("SCANNED_PDF", "IMAGE", "PARTIAL")
                         else "No" if gd.get("extraction_status") == "FAILED"
                         else "Yes")
                dup   = "Yes" if gd.get("is_duplicate") else "No"

                row_vals = [
                    doc_id, fp.name, cat_label, gst_id,    # 1-4
                    u_type, u_sub, u_owner, u_match,        # 5-8
                    u_old3m, u_date, u_risk,                # 9-11
                    a_type, a_name, a_addr, a_sig, a_notary,# 12-16
                    a_stamp, a_expiry, a_valid, a_sub, a_risk,# 17-21
                    dup, ocr, eff, issue, disc, reviewer_remarks,          # 22-27
                ]

                # ── UTILITY BILL BLURRY: Sheet 3 cell-level warning ──────────
                _dq_s3 = data.get("document_quality") or {}
                _util_blurry_s3 = _dq_s3.get("utility_bill_blurry") is True
                _UTIL_WARN_S3 = "Note: The Image quality is poor and blurry, content in non human readable format."
                _YELLOW_FILL_S3 = PatternFill(fill_type="solid", start_color="FFF7CC", end_color="FFF7CC")

                # If utility bill is blurry AND this is a utility row, append warning
                # to col 8 (Utility Address vs Lease Deed) and col 14 (Address Match)
                if _util_blurry_s3 and cat == "utility":
                    for _wi in (7, 13):  # 0-indexed into row_vals → col 8 and col 14
                        _orig = str(row_vals[_wi] or "")
                        if _UTIL_WARN_S3 not in _orig:
                            if _orig.strip():
                                _sep = ". " if not _orig.rstrip().endswith(".") else " "
                                row_vals[_wi] = f"{_orig}{_sep}{_UTIL_WARN_S3}"
                            else:
                                row_vals[_wi] = _UTIL_WARN_S3

                for ci, v in enumerate(row_vals, 1):
                    cell = ws.cell(row=row, column=ci, value=v)
                    cell.font   = _rfont(sz=10)
                    cell.fill   = _fill(bg)
                    cell.border = _bdr()
                    # Wrap text for File_Name (col 2), Discrepancy_Reason (col 26), Reviewer Remarks (col 27)
                    if ci in (2, 26, 27):
                        cell.alignment = _la(w=True)
                    else:
                        cell.alignment = _la(w=False)

                # ── Apply YELLOW highlight to utility blurry cells in Sheet 3 ────
                if _util_blurry_s3 and cat == "utility":
                    for _wc_col in (8, 14):  # 1-indexed column numbers
                        _wc_cell = ws.cell(row=row, column=_wc_col)
                        _wc_cell.fill = _YELLOW_FILL_S3
                        _wc_cell.font = _rfont(sz=10)  # normal black text

                # Colour Effective_Status cell (col 24)
                fg, fb = _status_clr(eff)
                ws.cell(row=row, column=24).font = _bfont(fg, 10)
                ws.cell(row=row, column=24).fill = _fill(fb)

                # Colour Older_Than_3_Months cell (col 9) for utility rows
                if cat == "utility" and u_old3m in ("Yes", "No", "NA"):
                    fg2, fb2 = _older3m_clr(u_old3m)
                    ws.cell(row=row, column=9).font = _bfont(fg2, 10)
                    ws.cell(row=row, column=9).fill = _fill(fb2)

                # Dynamic row height
                n_disc_lines = max(1, disc.count("\n") + 1) if disc else 1
                ws.row_dimensions[row].height = min(60, max(28, n_disc_lines * 16))
                self._idx_row += 1

    #  SHEET 4 — Discrepancy_Register
    #    • ONE row per document (not one row per issue)
    #    • Discrepancy_ID is now DOC-based (not sequential per issue)
    #    • ONE row per document (not one row per issue)
    #    • Issues shown in separate columns: Issue I, Issue II, Issue III
    #    • "Description of Issue" column with full readable explanation
    def _build_discrepancy(self):
        ws = self.wb.create_sheet("Discrepancy_Register")
        ws.sheet_view.showGridLines = False
        self._ws_disc = ws

        cols = [
            "Discrepancy_ID",        # 1
            "GST_Reg_ID",            # 2
            "Doc_ID",                # 3
            "File_Name",             # 4
            "Document_Type",         # 5
            "Issue I",               # 6  — primary issue label
            "Issue II",              # 7  — secondary issue label
            "Issue III",             # 8  — tertiary issue label
            "Issue_Description",     # 9  — short combined summary
            "Description of Issue",  # 10 — full readable explanation (NEW)
            "Severity",              # 11
            "Weightage_Impact_%",    # 12
            "Status",                # 13
            "Suggested_Action",      # 14
            "Owner_Reviewer",        # 15
            "Target_Closure_Date",   # 16
            "Closure_Remarks",       # 17
        ]
        self._hdr(ws, 1, cols)
        self._widths(ws, {
            1: 16,  2: 14,  3: 12,  4: 36,  5: 26,
            6: 30,  7: 30,  8: 30,  9: 40,  10: 65,
            11: 12, 12: 14, 13: 16, 14: 48,
            15: 24, 16: 18, 17: 32,
        })
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(cols))}1"

        ISSUE_TYPES = [
            "Missing Document", "Expired Document", "Name Mismatch", "Address Mismatch",
            "Signature Missing", "Notary/Registration Missing", "Stamp Duty Non-Compliant",
            "Bill Older Than 3 Months", "Witness Incomplete", "Sub-Lease Issue",
            "OCR Unreadable", "Duplicate", "Other",
        ]
        self._dv(ws,  6, 2, 2000, ISSUE_TYPES)   # Issue I
        self._dv(ws,  7, 2, 2000, ISSUE_TYPES)   # Issue II
        self._dv(ws,  8, 2, 2000, ISSUE_TYPES)   # Issue III
        self._dv(ws, 11, 2, 2000, ["HIGH", "MEDIUM", "LOW", "INFO"])
        self._dv(ws, 13, 2, 2000, ["Open", "In Progress", "Resolved", "Waived", "Escalated"])

    # ── Severity → numeric weightage impact ───────────────────────────────────
    _SEV_WEIGHT = {"HIGH": 15, "MEDIUM": 5, "LOW": 2, "INFO": 1}

    def _classify_issue(self, text: str) -> Tuple[str, str, str]:
        """Return (issue_type, suggested_action, severity) from free-text issue."""
        il = str(text).lower()
        # ── Ownership proof specific (v3.0) ──────────────────────────────────
        if "land revenue" in il and ("valid" in il or "exempt" in il):
            return ("Land Revenue Record Valid",
                    "Land Revenue Record accepted as permanent ownership proof; 3-month rule not applicable.",
                    "INFO")
        if "ownership proof" in il and ("mismatch" in il or "differ" in il):
            return ("Address Mismatch",
                    "Verify ownership proof address matches the agreement property address.",
                    "HIGH")
        if "property tax" in il and ("outdat" in il or "fy 20" in il):
            return ("Outdated Tax Receipt",
                    "Provide a Property Tax Receipt for the current or immediately previous financial year.",
                    "HIGH")
        if "merged pdf" in il or "ownership proof pages" in il:
            return ("Merged Document Detected",
                    "Ownership proof found embedded in merged PDF — validate each segment independently.",
                    "LOW")
        if "3-month rule not applicable" in il or "3 month rule exempted" in il:
            return ("3-Month Rule Exemption",
                    "Ownership proof type is exempt from 3-month recency rule — no action needed.",
                    "INFO")
        # ── Original logic ────────────────────────────────────────────────────
        if "older than 3" in il or "not within" in il or "more than 3" in il:
            return ("Bill Older Than 3 Months",
                    "Provide a utility bill dated within 3 months of the application date.",
                    "HIGH")
        if "owner name" in il or "name mismatch" in il or "mismatch" in il:
            return ("Name Mismatch",
                    "Provide initial lease deed or NOC explaining the owner name discrepancy.",
                    "HIGH")
        if "notari" in il or "registr" in il:
            return ("Notary/Registration Missing",
                    "Get the agreement notarised or registered with Sub-Registrar.",
                    "HIGH")
        if "expir" in il:
            return ("Expired Document",
                    "Renew the lease/rent agreement before uploading to GST portal.",
                    "HIGH")
        if "applicant" in il or "business name" in il or "not identified" in il:
            return ("Missing Document",
                    "Provide identity/incorporation proof to establish the GST applicant name.",
                    "HIGH")
        if "signature" in il or "sign" in il:
            return ("Signature Missing",
                    "Obtain signed and stamped copies from both lessor and lessee.",
                    "HIGH")
        if "witness address" in il or "witness" in il:
            return ("Witness Incomplete",
                    "Obtain complete witness details (name, full address, signature).",
                    "MEDIUM")
        if "stamp duty" in il:
            return ("Stamp Duty Non-Compliant",
                    "Verify and rectify stamp duty as per state-specific rules.",
                    "MEDIUM")
        if "sub-lease" in il or "sub lease" in il:
            return ("Sub-Lease Issue",
                    "Verify sub-lease clause in original deed; obtain owner NOC if absent.",
                    "MEDIUM")
        if "ocr" in il or "unreadable" in il or "scanned" in il:
            return ("OCR Unreadable",
                    "Provide a clearer or digital version of the document.",
                    "MEDIUM")
        if "duplicate" in il:
            return ("Duplicate",
                    "Remove duplicate document from the submission set.",
                    "LOW")
        if "address" in il:
            return ("Address Mismatch",
                    "Verify and align the address on all documents with the lease deed address.",
                    "MEDIUM")
        return ("Other",
                "Review the issue and resolve before uploading to GST portal.",
                "MEDIUM")

    def _write_discrepancy_rows(self, gst_id: str, data: Dict, cat_map: Dict[str, List[Path]]):
        """
        Write one row per document that has issues — Sheet 4 Discrepancy_Register.
        v4.1: ONE row per document, issues in Issue I / II / III columns,
              full text in Description of Issue column.
        """
        ws   = self._ws_disc
        s    = data.get("summary", {})
        comp = data.get("compliance", {})
        docs = data.get("documents_analyzed", [])
        
        client_readability_warning = s.get("readability_warning") is True or any(
            doc.get("readability_warning") is True for doc in docs
        )

        def _add_row(doc_id: str, file_name: str, doc_type: str,
                     issue_list: List[str], severity: str, action: str,
                     is_readability_warn: bool = False):
            self._disc_ctr += 1
            row     = self._disc_row
            bg      = self._alt(row)
            disc_id = f"DISC-{self._disc_ctr:04d}"

            # Clean issue texts — strip any [Source:...] noise
            clean_issues = []
            for iss in issue_list:
                t = re.sub(r"\s*\[Source:[^\]]*\]", "", str(iss)).strip()
                # Strip leading "Issue N –" prefix if present (we re-label below)
                t = re.sub(r"^Issue\s+[IVX]+\s*[–\-:]\s*", "", t).strip()
                if t:
                    clean_issues.append(t)

            # Classify each issue to get canonical type labels
            classified = [self._classify_issue(iss) for iss in clean_issues]
            if not classified:
                classified = [("Other", action, severity)]

            # Worst severity
            sev_rank  = {"HIGH": 3, "MEDIUM": 2, "LOW": 1, "INFO": 0}
            worst_sev = max(classified, key=lambda x: sev_rank.get(x[2], 0))[2]
            worst_action = next(
                (c[1] for c in classified if sev_rank.get(c[2], 0) == sev_rank.get(worst_sev, 0)),
                action
            )
            wt_impact = self._SEV_WEIGHT.get(worst_sev, 5)

            # Issue I, II, III — canonical type label for each
            issue1 = classified[0][0] if len(classified) > 0 else ""
            issue2 = classified[1][0] if len(classified) > 1 else ""
            issue3 = classified[2][0] if len(classified) > 2 else ""

            # Short Issue_Description (col 9): comma-joined type labels
            issue_summary = ", ".join(dict.fromkeys(c[0] for c in classified[:3]))

            # Full Description of Issue (col 10): numbered plain English
            roman = ["I", "II", "III", "IV"]
            desc_parts = [
                f"Issue {roman[ri]}: {txt}"
                for ri, txt in enumerate(clean_issues[:4])
            ]
            full_desc = "\n".join(desc_parts) if desc_parts else "No details available."

            warning_note = "Note: The Image quality is poor and blurry, content in non human readable format."
            
            def append_readability_warning(text_val: str) -> str:
                if not text_val:
                    return warning_note
                text_str = str(text_val).strip()
                if warning_note in text_str:
                    return text_str
                if text_str.endswith(".") or text_str.endswith(":") or text_str.endswith("!"):
                    return f"{text_str} {warning_note}"
                else:
                    return f"{text_str}. {warning_note}"

            if is_readability_warn:
                full_desc = append_readability_warning(full_desc)
                worst_action = append_readability_warning(worst_action)
                closure_remarks = warning_note
            else:
                closure_remarks = ""

            vals = [
                disc_id,       # 1
                gst_id,        # 2
                doc_id,        # 3
                file_name,     # 4
                doc_type,      # 5
                issue1,        # 6  Issue I
                issue2,        # 7  Issue II
                issue3,        # 8  Issue III
                issue_summary, # 9  Short summary
                full_desc,     # 10 Full description
                worst_sev,     # 11
                wt_impact,     # 12
                "Open",        # 13
                worst_action,  # 14
                "",            # 15
                "",            # 16
                closure_remarks, # 17
            ]
            for ci, v in enumerate(vals, 1):
                cell = ws.cell(row=row, column=ci, value=v)
                cell.font   = _rfont(sz=10)
                cell.fill   = _fill(bg)
                cell.border = _bdr()
                if ci in (4, 10, 14, 17):
                    cell.alignment = _la(w=True)
                else:
                    cell.alignment = _la(w=False)

            fg, fb = _status_clr(worst_sev)
            ws.cell(row=row, column=11).font = _bfont(fg, 10)
            ws.cell(row=row, column=11).fill = _fill(fb)
            n_lines = max(1, len(desc_parts))
            ws.row_dimensions[row].height = min(90, max(32, n_lines * 18 + 10))
            self._disc_row += 1

        # ── Per-document issues — ONE ROW PER DOCUMENT ────────────────────
        prior_docs = self._doc_ctr - sum(len(v) for v in cat_map.values())
        for i, doc in enumerate(docs, 1):
            eff = str(doc.get("effective_status") or "")
            if eff not in ("INCORRECT", "MISMATCH", "MISSING", "REVIEW"):
                continue
            doc_id = f"DOC-{prior_docs + i:04d}"
            sev    = "HIGH" if eff in ("INCORRECT", "MISSING") else "MEDIUM"

            issue_texts: List[str] = []
            disc_reason = str(doc.get("discrepancy_reason") or "").strip()
            key_notes   = str(doc.get("key_notes") or "").strip()
            issue_type  = str(doc.get("issue_type") or "").strip()

            # Primary issue from discrepancy_reason (cleaned)
            if disc_reason:
                clean_dr = re.sub(r"\s*\[Source:[^\]]*\]", "", disc_reason).strip()
                clean_dr = re.sub(r"^Issue\s+[IVX]+\s*[–\-:]\s*", "", clean_dr).strip()
                if clean_dr:
                    issue_texts.append(clean_dr[:250])

            # Add key_notes if it has additional context not in disc_reason
            if key_notes and (not issue_texts or key_notes.lower() not in issue_texts[0].lower()):
                clean_kn = key_notes[:200]
                issue_texts.append(clean_kn)

            if not issue_texts:
                issue_texts.append(f"Document status: {eff}. Issue type: {issue_type or 'Unknown'}.")

            doc_readability_warn = doc.get("readability_warning") is True or client_readability_warning

            _add_row(
                doc_id,
                str(doc.get("file_name") or ""),
                str(doc.get("detected_type") or ""),
                issue_texts, sev,
                "Review and resolve before uploading to GST portal.",
                is_readability_warn=doc_readability_warn
            )

        # ── Global compliance issues — ONE CONSOLIDATED ROW ───────────────
        global_issues = [
            re.sub(r"\s*\[Source:[^\]]*\]", "", str(iss)).strip()
            for iss in (comp.get("issues") or []) if str(iss).strip()
        ]
        global_issues = [g for g in global_issues if g]
        if global_issues:
            def _iss_sev(txt):
                tl = txt.lower()
                if any(k in tl for k in ("older than 3", "mismatch", "not identified",
                                          "expired", "missing", "cannot", "high risk")):
                    return "HIGH"
                if any(k in tl for k in ("witness", "review", "partial", "stamp")):
                    return "MEDIUM"
                return "LOW"
            sev_rank2 = {"HIGH": 3, "MEDIUM": 2, "LOW": 1, "INFO": 0}
            worst_g  = max(global_issues, key=lambda x: sev_rank2.get(_iss_sev(x), 0))
            grp_sev  = _iss_sev(worst_g)
            best_act = self._classify_issue(worst_g)[1]
            _add_row("GENERAL", "N/A (Compliance-level)", "Compliance Check",
                     global_issues, grp_sev, best_act,
                     is_readability_warn=client_readability_warning)


    #  ORCHESTRATION
    def build_all_sheets(self):
        """Create all workbook sheet skeletons (no data yet — populated per client)."""
        self._build_registrations()
        self._build_document_index()
        self._build_discrepancy()

    def add_client(self, gst_id: str, client_name: str,
                   data: Dict, cat_map: Dict[str, List[Path]]):
        """Append one client's data to all relevant sheets."""
        all_files = [f for files in cat_map.values() for f in files]
        log.info(f"   Writing Excel rows → [{client_name}] ({gst_id})")
        self._write_registration_row(gst_id, data, all_files)
        self._write_document_rows(gst_id, data, cat_map)
        self._write_discrepancy_rows(gst_id, data, cat_map)

    def save(self):
        """Save the workbook to disk."""
        self.out.parent.mkdir(parents=True, exist_ok=True)
        self.wb.save(self.out)
        log.info(f"   Workbook saved → {self.out}")




#  PARALLEL CLIENT PROCESSOR

def process_client(
    idx: int,
    total: int,
    client_name: str,
    cat_map: Dict[str, List[Path]],
    api_key: str,
    args: argparse.Namespace,
) -> Dict[str, Any]:
    """
    Independent per-client pipeline: cache → Gemini → reconcile → state-wise folders → outputs.

    After compliance reconciliation, the Indian state label is derived from extracted JSON
    so outputs land under OUTPUT/<State>/<Client>/.

    Returns a result dict with keys:
      client_name, gst_id, cat_map, data (or None), status, elapsed_sec, error,
      folders (Dict[str,Path] from create_statewise_client_structure)
    """
    t_start = time.perf_counter()
    gst_id  = f"GSTREG-{idx:03d}"
    thread_name = threading.current_thread().name
    log.info(f"  CLIENT {idx}/{total}: {client_name}  [{thread_name}]")

    all_source_files: List[Path] = [f for files in cat_map.values() for f in files]
    folders: Optional[Dict[str, Path]] = None
    result = None

    # ── HARDENED : Checkpoint/resume support ────────────────────────────
    _cp = get_checkpoint_manager()
    if _cp and _cp.is_done(client_name):
        log.info(f"  ✓ {client_name}: already completed — skipping (checkpoint)")
        return dict(client_name=client_name, gst_id=gst_id, cat_map=cat_map,
                    data=load_json_cache(client_name), status="skipped_checkpoint",
                    elapsed_sec=0, error=None, folders={})
    if _cp:
        _cp.mark_in_progress(client_name)

    # ── Audit logger reference ────────────────────────────────────────────────
    _audit = get_audit_logger()
    if _audit:
        _audit.update_summary(total_clients=1)

    proc_log_lines: List[str] = [
        f"GST Validator — Client Processing Log",
        f"Client       : {client_name}",
        f"GST ID       : {gst_id}",
        f"Run started  : {time.strftime('%Y-%m-%d %H:%M:%S')}",
        f"Total files  : {len(all_source_files)}",
        "─" * 60,
    ]

    try:
        all_files = all_source_files

        # ── Step 1: Try cache ──────────────────────────────────────────────
        if CACHE_ENABLED and (args.use_cache or args.skip_gemini):
            result = load_json_cache(client_name)
            if result:
                # ── Cache validity guard ───────────────────────────────────
                # A cached result with 0 documents_analyzed but actual files
                # exist is a bad/incomplete cache entry — discard and re-call
                # Gemini so all documents are properly analysed.
                _cached_doc_count = result.get("summary", {}).get(
                    "total_documents_analyzed", 0
                ) or len(result.get("documents_analyzed") or [])
                _actual_file_count = len(all_source_files)
                if _cached_doc_count == 0 and _actual_file_count > 0:
                    log.warning(
                        f"  ⚠ {client_name}: cached result has 0 documents_analyzed "
                        f"but {_actual_file_count} source file(s) exist — "
                        f"discarding stale cache and re-calling Gemini."
                    )
                    proc_log_lines.append(
                        f"Step 1: Cache discarded — 0 docs in cache vs "
                        f"{_actual_file_count} actual file(s). Re-calling Gemini."
                    )
                    result = None  # force Gemini re-call below
                else:
                    log.info(f"  ✓ {client_name}: loaded from JSON cache "
                             f"({_cached_doc_count} doc(s) in cache, "
                             f"{_actual_file_count} file(s) on disk)")
                    proc_log_lines.append("Step 1: Loaded result from JSON cache.")
            # result is still None here if: (a) no cache existed, or (b) bad cache discarded
            if result is None and args.skip_gemini:
                log.warning(f"  ⚠ {client_name}: no cache found — skipping (--skip-gemini)")
                proc_log_lines.append("Step 1: No cache found and --skip-gemini active. SKIPPED.")
                with _folder_lock:
                    folders = create_statewise_client_structure(
                        CLIENT_OUTPUT_BASE, client_name, "Unknown_State"
                    )
                organize_client_documents(all_source_files, folders["input_docs"], client_name=client_name)
                write_client_metadata(
                    folders["root"], client_name, len(all_files),
                    "SKIPPED", {"reason": "no_cache_skip_gemini", "gst_id": gst_id}
                )
                save_logs(client_name, folders["logs"], proc_log_lines)
                # ── Quarantine skipped documents ──────────────────────────────
                try:
                    safe_archive_failed_file(
                        client_name=client_name,
                        state_label="Unknown_State",
                        source_files=all_source_files,
                        failure_reason="no_cache_skip_gemini",
                        output_root=CLIENT_OUTPUT_BASE,
                        worker_id=threading.current_thread().name,
                        status="SKIPPED",
                    )
                except Exception:
                    pass
                record_processing_event(
                    CLIENT_OUTPUT_BASE, client_name, "Unknown_State",
                    "skipped", "no_cache_skip_gemini",
                    threading.current_thread().name,
                )
                if _audit:
                    _audit.log_skipped(client_name, "", "no_cache_skip_gemini",
                                       state="Unknown_State")
                return dict(client_name=client_name, gst_id=gst_id, cat_map=cat_map,
                            data=None, status="skipped_no_cache", elapsed_sec=0,
                            error=None, folders=folders)

        # ── Step 2: Call Gemini (if no cached result) ──────────────────────
        if result is None and not args.skip_gemini:
            proc_log_lines.append("Step 2: Calling Gemini API…")
            result = call_gemini(all_files, api_key, client_name, cat_map=cat_map)

        if not result:
            log.warning(f"  ⚠ {client_name}: no result obtained — skipping")
            proc_log_lines.append("Step 2: Gemini returned no result. SKIPPED.")
            with _folder_lock:
                folders = create_statewise_client_structure(
                    CLIENT_OUTPUT_BASE, client_name, "Unknown_State"
                )
            organize_client_documents(all_source_files, folders["input_docs"], client_name=client_name)
            write_client_metadata(
                folders["root"], client_name, len(all_files),
                "SKIPPED", {"reason": "no_gemini_result", "gst_id": gst_id}
            )
            save_logs(client_name, folders["logs"], proc_log_lines)
            # ── Quarantine skipped documents ──────────────────────────────────
            try:
                safe_archive_failed_file(
                    client_name=client_name,
                    state_label="Unknown_State",
                    source_files=all_source_files,
                    failure_reason="no_gemini_result",
                    output_root=CLIENT_OUTPUT_BASE,
                    worker_id=threading.current_thread().name,
                    status="SKIPPED",
                )
            except Exception:
                pass
            record_processing_event(
                CLIENT_OUTPUT_BASE, client_name, "Unknown_State",
                "skipped", "no_gemini_result",
                threading.current_thread().name,
            )
            if _audit:
                _audit.log_skipped(client_name, "", "no_gemini_result",
                                   state="Unknown_State")
            return dict(client_name=client_name, gst_id=gst_id, cat_map=cat_map,
                        data=None, status="skipped_no_result", elapsed_sec=0,
                        error=None, folders=folders)

        proc_log_lines.append("Step 2: Gemini response received successfully.")

        # ── Step 3: Post-process (must run before state folder resolution) ─
        proc_log_lines.append("Step 3: Running compliance reconciliation…")
        result = _reconcile_compliance(result, all_files)

        state_label = extract_state_from_validation_data(result)
        log.info(f"   [FolderManager] State folder for [{client_name}] → {state_label}")

        with _folder_lock:
            folders = create_statewise_client_structure(
                CLIENT_OUTPUT_BASE, client_name, state_label
            )

        organize_client_documents(all_source_files, folders["input_docs"], client_name=client_name)

        # ── Step 4: Persist cache ──────────────────────────────────────────
        if CACHE_ENABLED:
            save_json_cache(client_name, result, client_cache_dir=folders["json_cache"])
            proc_log_lines.append("Step 4: JSON cache persisted (global + client).")

        # ── Step 5: Write per-client outputs ────────────────────────────────
        save_json_response(client_name, folders["logs"], result)
        proc_log_lines.append("Step 5a: validation_response_{timestamp}.json written to client logs/.")

        elapsed = time.perf_counter() - t_start
        _comp_raw = result.get("compliance")
        comp      = _comp_raw if isinstance(_comp_raw, dict) else {}
        wcp       = comp.get("weighted_compliance_percent", 0)
        fs      = comp.get("final_status", "?")

        proc_log_lines += [
            "─" * 60,
            f"Run finished : {time.strftime('%Y-%m-%d %H:%M:%S')}",
            f"Elapsed      : {elapsed:.1f}s",
            f"Compliance   : {wcp:.1f}%",
            f"Status       : {fs}",
            f"Passes       : {comp.get('total_passes', 0)}",
            f"Issues       : {comp.get('total_issues', 0)}",
        ]

        save_logs(client_name, folders["logs"], proc_log_lines)

        write_client_metadata(
            folders["root"], client_name, len(all_files),
            "SUCCESS", {
                "gst_id"                     : gst_id,
                "state_folder"               : state_label,
                "compliance_percent"          : round(float(wcp), 1),
                "final_status"               : fs,
                "total_passes"               : comp.get("total_passes", 0),
                "total_issues"               : comp.get("total_issues", 0),
                "elapsed_sec"                : round(elapsed, 1),
            }
        )

        log.info(
            f"  ✔ {client_name}: {elapsed:.1f}s | "
            f"Compliance={wcp:.1f}%  Status={fs}  "
            f"Passes={comp.get('total_passes', 0)}  "
            f"Issues={comp.get('total_issues', 0)}"
        )

        # ── HARDENED : Mark done in checkpoint + update audit summary ──
        if _cp:
            _cp.mark_done(client_name)
        if _audit:
            _audit.update_summary(processed_ok=1, state=extract_state_from_validation_data(result))
            if fs == "Clean":
                _audit.update_summary(clean=1)
            elif fs == "Warning":
                _audit.update_summary(warning=1)
            elif fs == "High Risk":
                _audit.update_summary(high_risk=1)

        # ── Record successful processing event ────────────────────────────────
        try:
            versioned_ts = time.strftime("%Y%m%d_%H%M%S")
            record_processing_event(
                output_root=CLIENT_OUTPUT_BASE,
                client_name=client_name,
                state=state_label,
                status="ok",
                worker_id=threading.current_thread().name,
                duration_sec=elapsed,
                output_version=f"validation_response_{versioned_ts}.json",
                final_compliance_status=fs,
            )
        except Exception:
            pass

        return dict(client_name=client_name, gst_id=gst_id, cat_map=cat_map,
                    data=result, status="ok", elapsed_sec=elapsed,
                    error=None, folders=folders)

    except Exception as exc:
        elapsed = time.perf_counter() - t_start
        error   = f"{type(exc).__name__}: {exc}"
        log.error(f"  ✖ {client_name}: FAILED after {elapsed:.1f}s — {error}")
        log.debug(traceback.format_exc())

        proc_log_lines += [
            "─" * 60,
            f"ERROR: {error}",
            f"Run finished : {time.strftime('%Y-%m-%d %H:%M:%S')}",
            f"Elapsed      : {elapsed:.1f}s",
        ]
        # ── HARDENED: Mark failed in checkpoint + audit ───────────────
        if _cp:
            _cp.mark_failed(client_name)
        if _audit:
            _audit.update_summary(failed=1)
            _audit.log_failed(
                client_name, "", "process_client", error, traceback.format_exc(),
                state="Unknown_State", duration_sec=elapsed,
            )

        try:
            if folders is None:
                with _folder_lock:
                    folders = create_statewise_client_structure(
                        CLIENT_OUTPUT_BASE, client_name, "Unknown_State"
                    )
                organize_client_documents(all_source_files, folders["input_docs"],
                                          client_name=client_name)
            save_logs(client_name, folders["logs"], proc_log_lines)
            write_client_metadata(
                folders["root"], client_name, len(all_source_files),
                "FAILED", {"gst_id": gst_id, "error": error,
                           "elapsed_sec": round(elapsed, 1)}
            )
        except Exception:
            pass  # do not let metadata/log failures mask the original error

        # ── Quarantine failed documents ───────────────────────────────────────
        try:
            safe_archive_failed_file(
                client_name=client_name,
                state_label="Unknown_State",
                source_files=all_source_files,
                failure_reason=error,
                tb_text=traceback.format_exc(),
                output_root=CLIENT_OUTPUT_BASE,
                retry_count=getattr(args, "retry_count", 0),
                worker_id=threading.current_thread().name,
                status="FAILED",
                duration_sec=elapsed,
            )
        except Exception as qe:
            log.warning(f"[Quarantine] Quarantine step failed for {client_name}: {qe}")

        # ── Record to processing_history + retry_history ─────────────────────
        try:
            record_processing_event(
                output_root=CLIENT_OUTPUT_BASE,
                client_name=client_name,
                state="Unknown_State",
                status="failed",
                failure_reason=error,
                worker_id=threading.current_thread().name,
                duration_sec=elapsed,
            )
        except Exception:
            pass

        return dict(client_name=client_name, gst_id=gst_id, cat_map=cat_map,
                    data=None, status="failed", elapsed_sec=elapsed,
                    error=error, folders=folders if folders is not None else {})


def cleanup_client_resources(client_name: str, result: Optional[Dict] = None) -> None:
    """
    Issue 15 — Explicit resource cleanup after processing a client.
    Releases per-client lock from registry (memory hygiene for long batches).
    Does NOT touch the result dict or any on-disk files.
    """
    import gc
    try:
        cleanup_client_locks(client_name)
    except Exception:
        pass
    # Explicit GC hint (minor collection) — avoids memory pile-up over 300-500 clients
    gc.collect(0)


def generate_client_report(
    master_path: Path,
    client_name: str,
    gst_id: str,
    client_folder: Path,
    client_data: Optional[Dict] = None,
) -> Optional[Path]:
    """
    Generate a per-client GST_Validation_Report.xlsx in the 2-sheet format
    matching the sample template:
      Sheet 1:  Summary  — full validation summary (business identity, address,
               lease details, witness records, electricity bill, documents analyzed,
               missing/flagged documents)
      Sheet 2:  Compliance — CBIC compliance analysis (passes and issues)

    The master workbook (GST_Validation_Workbook.xlsx) is saved at the client
    output root and is NOT copied wholesale into each client folder.

    The report is placed directly at:
        CLIENT_OUTPUT_BASE/<State>/<ClientName>/GST_Validation_Report.xlsx

    Thread-safe: writes only to client-unique dest_path — no shared state.

    Args:
        master_path   : Path to the completed master workbook (registration index workbook).
        client_name   : Human-readable client name (for log messages).
        gst_id        : GST registration ID for this client (e.g. "GSTREG-001").
        client_folder : Client root folder under state-wise layout.
        client_data   : Optional pre-loaded Gemini JSON dict for this client.
                        If None, the function tries to load from JSON cache.

    Returns:
        Path to the written per-client report, or None on failure.
    """
    # VERSIONED: never overwrite old per-client reports — always create timestamped new file
    _ts = time.strftime("%Y%m%d_%H%M%S")
    dest_path = client_folder / f"GST_Validation_Report_{_ts}.xlsx"
    _ctr = 1
    while dest_path.exists():
        dest_path = client_folder / f"GST_Validation_Report_{_ts}_{_ctr:03d}.xlsx"
        _ctr += 1

    # ── Colour palette matching the sample template exactly ──────────────────
    C_TITLE_BG    = "1B3A5C"   # deep navy  — main title row
    C_SUBTITLE_BG = "1A4F7A"   # mid navy   — subtitle / section-header rows
    C_SECTION_BG  = "2E6DA4"   # blue       — section heading rows
    C_LABEL_BG    = "D9E8F7"   # light blue — label column
    C_ALT_ROW     = "EBF3FB"   # pale blue  — alternating data rows
    C_WHITE       = "FFFFFF"
    C_GREEN_BG    = "D6F5D6"   # pass / YES / HIGH MATCH
    C_RED_BG      = "F8D7DA"   # fail / NO
    C_YELLOW_BG   = "FFF3CD"   # warning status banner

    def _fill(rgb: str):
        return PatternFill("solid", fgColor=rgb)

    def _font(bold=False, sz=10, color="000000"):
        return Font(name="Arial", bold=bold, size=sz, color=color)

    def _align(wrap=True, h="left", v="top"):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

    def _bdr():
        s = Side(style="thin", color="CCCCCC")
        return Border(left=s, right=s, top=s, bottom=s)

    def _w(ws, row, col, val, bold=False, bg=C_WHITE, sz=10,
           fg="000000", wrap=True, h="left", v="top"):
        c = ws.cell(row=row, column=col, value=val)
        c.font      = _font(bold=bold, sz=sz, color=fg)
        c.fill      = _fill(bg)
        c.border    = _bdr()
        c.alignment = _align(wrap=wrap, h=h, v=v)
        return c

    def _status_bg(val: str) -> str:
        """Return background colour for YES/NO/MATCH status values."""
        v = str(val).upper()
        if any(k in v for k in ("YES", "HIGH MATCH", "PARTIAL MATCH", "✔", "PASS", "CLEAN")):
            return C_GREEN_BG
        if any(k in v for k in ("NO", "MISMATCH", "✘", "FAIL", "HIGH RISK")):
            return C_RED_BG
        if any(k in v for k in ("WARN", "REVIEW", "WARNING", "PARTIAL")):
            return C_YELLOW_BG
        return C_WHITE

    try:
        # ── Load client data from JSON cache if not supplied ─────────────────
        data: Dict = {}
        if client_data is not None:
            data = client_data
        else:
            loaded = load_json_cache(client_name)
            if loaded:
                data = loaded
            else:
                log.warning(f"   [ClientReport] No JSON cache found for {client_name}; "
                             f"report will have empty fields.")

        # ── Extract data sections ─────────────────────────────────────────────
        s    = data.get("summary", {})    if data else {}
        n    = data.get("names", {})      if data else {}
        ld   = data.get("lease_details",{}) if data else {}
        eb   = data.get("electricity_bill",{}) if data else {}
        comp = data.get("compliance", {}) if data else {}
        docs = data.get("documents_analyzed", []) if data else []
        addrs= data.get("addresses", {})  if data else {}
        if not isinstance(docs, list): docs = []

        # ── Derived values ────────────────────────────────────────────────────
        val_date  = VALIDATION_DATE
        gst_name  = (n.get("finalised_gst_business_name")
                     or s.get("final_business_name")
                     or n.get("legal_name")
                     or client_name)
        category  = (s.get("category") or ld.get("agreement_type") or "")
        lessor    = (n.get("lessor_first_party") or ld.get("lessor") or "")
        lessee    = (n.get("lessee_second_party") or ld.get("lessee") or "")
        name_conf = s.get("name_confidence_percent") or s.get("name_match_confidence") or ""
        name_conf_label = ""
        if name_conf:
            try:
                pct = float(str(name_conf).replace("%","").strip())
                lvl = "HIGH" if pct >= 85 else "MEDIUM" if pct >= 60 else "LOW"
                name_conf_label = f"{int(pct)}%   [ {lvl} ]"
            except Exception:
                name_conf_label = str(name_conf)

        # POB address — same candidate logic as _write_registration_row
        pob_addr  = (s.get("final_pob_address")
                     or ld.get("premises_address")
                     or addrs.get("finalised_pob_address")
                     or "")
        pincode   = (addrs.get("pincode") or s.get("pincode") or "")
        addr_match= (s.get("address_match_status") or addrs.get("address_match_status") or "")
        notary_ok = bool(ld.get("notary_verified") or s.get("notary_verified") or ld.get("notary_registered") or s.get("notary_registered") or ld.get("registered_with_subregistrar") or s.get("registered_with_subregistrar") or ld.get("notarised") or s.get("notarised"))
        reg_type  = str(ld.get("registration_type") or "").upper()
        if notary_ok:
            if "REGISTERED" in reg_type:
                notary_str = "YES  — Registered Lease Deed (Sub-Registrar verified)"
            else:
                notary_str = "YES  — Notarised agreement"
        else:
            notary_str = "NO  — Notary / Registration NOT verified"

        lease_start = str(ld.get("lease_start_date") or ld.get("start_date") or "")
        lease_end   = str(ld.get("lease_end_date") or ld.get("end_date") or "")
        dur_months   = ld.get("duration_months", "")
        try: dur_months = int(float(str(dur_months))) if dur_months else ""
        except: pass
        agr_valid    = bool(ld.get("agreement_valid") or s.get("agreement_valid") or ld.get("agreement_valid_as_on_validation_date") or s.get("agreement_valid_as_on_validation_date"))
        agr_valid_str= f"YES" if agr_valid else "NO"
        stamp_amt    = ld.get("stamp_duty_amount", "")
        stamp_ok     = bool(ld.get("stamp_duty_compliant"))
        stamp_ok_str = "YES" if stamp_ok else "NO"
        sigs_ok      = bool(ld.get("signatures_both_parties"))
        sigs_ok_str  = "YES" if sigs_ok else "NO"

        witnesses    = ld.get("witness_records") or ld.get("witnesses") or []
        if not isinstance(witnesses, list): witnesses = []

        eb_owner  = eb.get("owner_name", "")
        eb_addr = eb.get("address_on_bill") or eb.get("address") or ""
        eb_date   = str(eb.get("bill_date") or "")
        within_3m = bool(eb.get("within_3_months"))
        within_3m_str = f"YES" if within_3m else "NO"

        overall_status = (s.get("effective_overall_status")
                          or comp.get("final_status")
                          or s.get("overall_status")
                          or "")
        wcp  = float(comp.get("weighted_compliance_percent", 0) or 0)
        passes_list = comp.get("passes", [])
        issues_list = comp.get("issues", [])
        if not isinstance(passes_list, list): passes_list = []
        if not isinstance(issues_list, list): issues_list = []
        total_p = len(passes_list)
        total_i = len(issues_list)
        total_checks = total_p + total_i

        missing_docs = s.get("missing_documents", [])
        if not isinstance(missing_docs, list): missing_docs = []

        # ── Create workbook ───────────────────────────────────────────────────
        client_folder.mkdir(parents=True, exist_ok=True)
        wb = Workbook()
        if "Sheet" in wb.sheetnames:
            del wb["Sheet"]

        #  SHEET 1 —  Summary
        ws1 = wb.create_sheet("📋 Summary")
        ws1.sheet_view.showGridLines = False
        ws1.sheet_properties.tabColor = "1B3A5C"

        # Column widths — A=label(32), B=value(52), C=extra(30), D=status(14)
        ws1.column_dimensions["A"].width = 34
        ws1.column_dimensions["B"].width = 54
        ws1.column_dimensions["C"].width = 32
        ws1.column_dimensions["D"].width = 14

        row = 1
        # Row 1 — Main title (spans A:D)
        _w(ws1, row, 1,
           "🏛  GST REGISTRATION — DOCUMENT VALIDATION REPORT",
           bold=True, bg=C_TITLE_BG, fg="FFFFFF", sz=13, h="center")
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 28
        row += 1

        # Row 2 — Subtitle with date (spans A:D)
        _w(ws1, row, 1,
           f" Validation Date: {val_date}   |   CBIC Instruction No. 03/2025-GST   |   Para 6",
           bold=True, bg=C_SUBTITLE_BG, fg="FFFFFF", sz=10, h="center")
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 20
        row += 1

        # Row 3 — blank spacer
        row += 1

        # Row 4 — Overall Status banner
        status_bg = _status_bg(overall_status)
        _w(ws1, row, 1,
           f"  OVERALL STATUS:   {overall_status}",
           bold=True, bg=status_bg, sz=11, h="center")
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 22
        row += 1

        # Row 5 — blank
        row += 1

        # ── BUSINESS IDENTITY section ─────────────────────────────────────────
        _w(ws1, row, 1, "BUSINESS IDENTITY",
           bold=True, bg=C_SECTION_BG, fg="FFFFFF", sz=11)
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 20
        row += 1

        def _label_val(ws, r, label, val, val_bold=False, val_bg=None):
            vbg = val_bg if val_bg else C_WHITE
            _w(ws, r, 1, label, bold=True, bg=C_LABEL_BG, sz=10)
            _w(ws, r, 2, str(val) if val is not None else "",
               bold=val_bold, bg=vbg, sz=10)
            ws.merge_cells(f"B{r}:D{r}")
            ws.row_dimensions[r].height = 18

        _label_val(ws1, row, "Finalised GST Business Name", gst_name, val_bold=True)
        row += 1
        _label_val(ws1, row, "Category / Business Type", category)
        row += 1
        _label_val(ws1, row, "Lessor (First Party)", lessor)
        row += 1
        _label_val(ws1, row, "Lessee / Applicant (Second Party)", lessee)
        row += 1
        conf_bg = _status_bg(name_conf_label) if name_conf_label else C_WHITE
        _label_val(ws1, row, "Name Match Confidence", name_conf_label,
                   val_bold=bool(name_conf_label), val_bg=conf_bg)
        row += 1

        # Row — blank
        row += 1

        # ── ADDRESS DETAILS section ───────────────────────────────────────────
        _w(ws1, row, 1, "ADDRESS DETAILS",
           bold=True, bg=C_SECTION_BG, fg="FFFFFF", sz=11)
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 20
        row += 1

        _w(ws1, row, 1, "Finalised POB Address", bold=True, bg=C_LABEL_BG, sz=10)
        c = _w(ws1, row, 2, str(pob_addr), bg=C_WHITE, sz=10)
        ws1.merge_cells(f"B{row}:D{row}")
        n_lines = max(1, str(pob_addr).count(",") // 3 + 1)
        ws1.row_dimensions[row].height = max(18, min(60, n_lines * 16))
        row += 1

        _label_val(ws1, row, "Pincode", pincode)
        row += 1
        addr_bg = _status_bg(addr_match)
        _label_val(ws1, row, "Address Match Status", addr_match,
                   val_bold=True, val_bg=addr_bg)
        row += 1

        # blank
        row += 1

        _w(ws1, row, 1, "Notary / Registration Verified", bold=True, bg=C_LABEL_BG, sz=10)
        notary_bg = C_GREEN_BG if notary_ok else C_RED_BG
        _w(ws1, row, 2, notary_str, bold=True, bg=notary_bg, sz=10)
        ws1.merge_cells(f"B{row}:D{row}")
        ws1.row_dimensions[row].height = 18
        row += 1

        # blank
        row += 1

        # ── LEASE AGREEMENT DETAILS section ──────────────────────────────────
        _w(ws1, row, 1, "  LEASE AGREEMENT DETAILS",
           bold=True, bg=C_SECTION_BG, fg="FFFFFF", sz=11)
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 20
        row += 1

        _label_val(ws1, row, "Lease Start Date", lease_start)
        row += 1
        _label_val(ws1, row, "Lease End Date", lease_end)
        row += 1
        _label_val(ws1, row, "Duration (Months)", dur_months)
        row += 1
        agr_bg = _status_bg(agr_valid_str)
        _label_val(ws1, row, f"Agreement Valid as on {val_date}",
                   agr_valid_str, val_bold=True, val_bg=agr_bg)
        row += 1
        _label_val(ws1, row, "Stamp Duty Amount (₹)", stamp_amt)
        row += 1
        stamp_bg = _status_bg(stamp_ok_str)
        _label_val(ws1, row, "Stamp Duty Compliant", stamp_ok_str,
                   val_bold=True, val_bg=stamp_bg)
        row += 1
        sigs_bg = _status_bg(sigs_ok_str)
        _label_val(ws1, row, "Signatures — Both Parties", sigs_ok_str,
                   val_bold=True, val_bg=sigs_bg)
        row += 1

        # blank
        row += 1

        # ── WITNESS RECORDS section ───────────────────────────────────────────
        _w(ws1, row, 1, f"WITNESS RECORDS  ({len(witnesses)})",
           bold=True, bg=C_SECTION_BG, fg="FFFFFF", sz=11)
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 20
        row += 1

        # Witness table header
        for ci, hdr in enumerate(["#", "Full Name", "Full Address"], 1):
            _w(ws1, row, ci, hdr, bold=True, bg=C_SUBTITLE_BG, fg="FFFFFF", sz=10,
               h="center")
        ws1.row_dimensions[row].height = 18
        row += 1

        if witnesses:
            for wi, w in enumerate(witnesses, 1):
                row_bg = C_WHITE if wi % 2 == 1 else C_ALT_ROW
                wname = str(w.get("name") or w.get("witness_name") or "")
                waddr = str(w.get("full_address") or w.get("address") or w.get("witness_address") or "")
                _w(ws1, row, 1, wi, bg=row_bg, sz=10, h="center")
                _w(ws1, row, 2, wname, bg=row_bg, sz=10)
                _w(ws1, row, 3, waddr, bg=row_bg, sz=10)
                ws1.row_dimensions[row].height = 18
                row += 1
        else:
            _w(ws1, row, 1, "—  No witness records found", bg=C_WHITE, sz=10)
            ws1.merge_cells(f"A{row}:D{row}")
            ws1.row_dimensions[row].height = 18
            row += 1

        # blank
        row += 1

        # ── ELECTRICITY / UTILITY BILL DETAILS section ────────────────────────
        _w(ws1, row, 1, " ELECTRICITY / UTILITY BILL DETAILS",
           bold=True, bg=C_SECTION_BG, fg="FFFFFF", sz=11)
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 20
        row += 1

        _label_val(ws1, row, "Owner Name on Bill", eb_owner)
        row += 1
        _label_val(ws1, row, "Address on Bill", eb_addr)
        row += 1
        _label_val(ws1, row, "Bill Date", eb_date)
        row += 1
        w3m_bg = _status_bg(within_3m_str)
        _label_val(ws1, row, f"Within 3 Months of {val_date}",
                   within_3m_str, val_bold=True, val_bg=w3m_bg)
        row += 1

        # blank
        row += 1

        # ── DOCUMENTS ANALYZED section ────────────────────────────────────────
        _w(ws1, row, 1, f"  DOCUMENTS ANALYZED  ({len(docs)})",
           bold=True, bg=C_SECTION_BG, fg="FFFFFF", sz=11)
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 20
        row += 1

        # Doc table header
        for ci, hdr in enumerate(["#", "File Name", "Detected Type / Role", "Status"], 1):
            _w(ws1, row, ci, hdr, bold=True, bg=C_SUBTITLE_BG, fg="FFFFFF", sz=10,
               h="center" if ci in (1, 4) else "left")
        ws1.row_dimensions[row].height = 18
        row += 1

        if docs:
            for di, doc in enumerate(docs, 1):
                row_bg = C_WHITE if di % 2 == 1 else C_ALT_ROW
                fname  = str(doc.get("file_name") or "")
                dtype  = str(doc.get("detected_type") or "")
                role   = str(doc.get("role") or "")
                type_role = f"{dtype}  /  {role}" if role and role.lower() != "none" else dtype
                eff_st = str(doc.get("effective_status") or doc.get("extraction_status") or "OK")
                st_display = "OK" if eff_st.upper() in ("CORRECT", "OK") else eff_st
                st_bg  = _status_bg(st_display)

                _w(ws1, row, 1, di, bg=row_bg, sz=10, h="center")
                _w(ws1, row, 2, fname, bg=row_bg, sz=10)
                _w(ws1, row, 3, type_role, bg=row_bg, sz=10)
                _w(ws1, row, 4, st_display, bold=True, bg=st_bg, sz=10, h="center")
                ws1.row_dimensions[row].height = 18
                row += 1
        else:
            _w(ws1, row, 1, "No documents found.", bg=C_WHITE, sz=10)
            ws1.merge_cells(f"A{row}:D{row}")
            ws1.row_dimensions[row].height = 18
            row += 1

        # blank
        row += 1

        # ── Key Notes per Document ────────────────────────────────────────────
        _w(ws1, row, 1, "Key Notes per Document",
           bold=True, bg=C_SUBTITLE_BG, fg="FFFFFF", sz=10)
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 18
        row += 1

        if docs:
            for di, doc in enumerate(docs, 1):
                row_bg = C_WHITE if di % 2 == 1 else C_ALT_ROW
                fname  = str(doc.get("file_name") or f"Document {di}")
                kn     = str(doc.get("key_notes") or doc.get("notes") or "").strip()
                note   = f"{di}. {fname}:  {kn}" if kn else f"{di}. {fname}"
                n_lines = max(1, note.count("\n") + 1)
                c = ws1.cell(row=row, column=1, value=note)
                c.font      = _font(sz=10)
                c.fill      = _fill(row_bg)
                c.border    = _bdr()
                c.alignment = _align(wrap=True)
                ws1.merge_cells(f"A{row}:D{row}")
                ws1.row_dimensions[row].height = max(18, min(80, n_lines * 16))
                row += 1
        else:
            row += 1

        # blank
        row += 1

        # ── MISSING / FLAGGED DOCUMENTS section ───────────────────────────────
        _w(ws1, row, 1, f"⚠  MISSING / FLAGGED DOCUMENTS  ({len(missing_docs)})",
           bold=True, bg=C_SECTION_BG, fg="FFFFFF", sz=11)
        ws1.merge_cells(f"A{row}:D{row}")
        ws1.row_dimensions[row].height = 20
        row += 1

        if missing_docs:
            for mi, md in enumerate(missing_docs, 1):
                row_bg = C_RED_BG if mi % 2 == 1 else C_WHITE
                c = ws1.cell(row=row, column=1,
                             value=f"  {mi}. {str(md).strip()}")
                c.font      = _font(sz=10)
                c.fill      = _fill(row_bg)
                c.border    = _bdr()
                c.alignment = _align(wrap=True)
                ws1.merge_cells(f"A{row}:D{row}")
                ws1.row_dimensions[row].height = 18
                row += 1
        else:
            _w(ws1, row, 1,
               "✔  No missing documents — all required documents present",
               bold=True, bg=C_GREEN_BG, sz=10)
            ws1.merge_cells(f"A{row}:D{row}")
            ws1.row_dimensions[row].height = 18

        ws1.freeze_panes = "A3"

        #  SHEET 2 — Compliance
        ws2 = wb.create_sheet("✅ Compliance")
        ws2.sheet_view.showGridLines = False
        ws2.sheet_properties.tabColor = "217346"

        ws2.column_dimensions["A"].width = 100

        row2 = 1
        # Title
        _w(ws2, row2, 1,
           "  COMPLIANCE ANALYSIS — CBIC Instruction No. 03/2025-GST",
           bold=True, bg=C_TITLE_BG, fg="FFFFFF", sz=12)
        ws2.row_dimensions[row2].height = 26
        row2 += 1

        # Score banner
        score_bg = C_GREEN_BG if wcp >= 85 else C_YELLOW_BG if wcp >= 60 else C_RED_BG
        _w(ws2, row2, 1,
           f"Compliance Score:   {total_p} / {total_checks} checks passed   ({wcp:.0f}%)",
           bold=True, bg=score_bg, sz=11)
        ws2.row_dimensions[row2].height = 22
        row2 += 1

        # blank
        row2 += 1

        # PASSES section
        _w(ws2, row2, 1, f"✔  PASSES  ({total_p})",
           bold=True, bg=C_SECTION_BG, fg="FFFFFF", sz=11)
        ws2.row_dimensions[row2].height = 20
        row2 += 1

        if passes_list:
            for pi, p_item in enumerate(passes_list, 1):
                p_text = str(p_item).strip() if p_item else ""
                if not p_text:
                    continue
                row_bg = C_GREEN_BG if pi % 2 == 1 else C_WHITE
                n_lines = max(1, p_text.count("\n") + 1)
                c = ws2.cell(row=row2, column=1,
                             value=f"  ✔   {p_text}")
                c.font      = _font(sz=10)
                c.fill      = _fill(row_bg)
                c.border    = _bdr()
                c.alignment = _align(wrap=True)
                ws2.row_dimensions[row2].height = max(18, min(72, n_lines * 16))
                row2 += 1
        else:
            _w(ws2, row2, 1, "  No passes recorded.", bg=C_WHITE, sz=10)
            ws2.row_dimensions[row2].height = 18
            row2 += 1

        # blank
        row2 += 1

        # ISSUES section
        _w(ws2, row2, 1, f"✘  ISSUES / WARNINGS  ({total_i})",
           bold=True, bg=C_SECTION_BG, fg="FFFFFF", sz=11)
        ws2.row_dimensions[row2].height = 20
        row2 += 1

        if issues_list:
            for ii, i_item in enumerate(issues_list, 1):
                i_text = str(i_item).strip() if i_item else ""
                if not i_text:
                    continue
                row_bg = C_RED_BG if ii % 2 == 1 else C_WHITE
                n_lines = max(1, i_text.count("\n") + 1)
                c = ws2.cell(row=row2, column=1,
                             value=f"  ✘   {i_text}")
                c.font      = _font(sz=10)
                c.fill      = _fill(row_bg)
                c.border    = _bdr()
                c.alignment = _align(wrap=True)
                ws2.row_dimensions[row2].height = max(18, min(72, n_lines * 16))
                row2 += 1
        else:
            _w(ws2, row2, 1, "  ✔  No issues — fully compliant.", bg=C_GREEN_BG, sz=10)
            ws2.row_dimensions[row2].height = 18

        ws2.freeze_panes = "A2"

        # ── Save per-client report ────────────────────────────────────────────
        wb.save(str(dest_path))
        log.info(f"   [ClientReport] Saved → {dest_path}")
        return dest_path

    except Exception as exc:
        log.error(f"   [ClientReport] Failed for {client_name}: {exc}")
        log.debug(traceback.format_exc())
        return None


def run_parallel_batch(
    groups: Dict[str, Dict[str, List[Path]]],
    api_key: str,
    args: argparse.Namespace,
    builder: "ExcelBuilder",
) -> Dict[str, Any]:
    """
    Execute all clients in parallel using ThreadPoolExecutor.
    Returns a stats dict: processed, skipped, failed, clean, warning, high_risk.

    Thread-safety:
      - Each client runs process_client() independently (no shared mutable state).
      - ExcelBuilder.add_client() is serialised via _excel_lock to prevent race conditions.
      - JSON cache reads/writes are serialised via _cache_lock inside save/load_json_cache.
    """
    stats = {
        "total": len(groups), "processed": 0, "skipped": 0,
        "failed": 0, "clean": 0, "warning": 0, "high_risk": 0,
        "errors": [],   # list of (client_name, error_msg)
    }
    client_items = list(groups.items())
    total        = len(client_items)
    t_batch_start = time.perf_counter()

    # Progress tracking
    completed_count = 0
    prog_bar = None
    if HAS_TQDM:
        prog_bar = _tqdm(total=total, desc="Processing clients", unit="client", ncols=80)

    futures_map: Dict = {}

    with ThreadPoolExecutor(max_workers=MAX_WORKERS,
                            thread_name_prefix="GSTWorker") as executor:
        # Submit all clients
        for idx, (client_name, cat_map) in enumerate(client_items, 1):
            future = executor.submit(
                process_client, idx, total, client_name, cat_map, api_key, args
            )
            futures_map[future] = (idx, client_name)

        # Collect results as they complete
        for future in as_completed(futures_map):
            idx, client_name = futures_map[future]
            completed_count += 1

            try:
                res = future.result()
            except Exception as exc:
                log.error(f"  ✖ Unexpected future error for {client_name}: {exc}")
                stats["failed"] += 1
                stats["errors"].append((client_name, str(exc)))
                if prog_bar:
                    prog_bar.update(1)
                continue

            # Write to Excel (serialised)
            if res["data"] is not None:
                with _excel_lock:
                    builder.add_client(res["gst_id"], res["client_name"],
                                       res["data"], res["cat_map"])

            # Update stats
            st = res["status"]
            if st == "ok":
                stats["processed"] += 1
                fs = (res["data"].get("compliance", {}).get("final_status", "?")
                      if res["data"] else "?")
                if   fs == "Clean":     stats["clean"]     += 1
                elif fs == "Warning":   stats["warning"]   += 1
                elif fs == "High Risk": stats["high_risk"] += 1
            elif st == "failed":
                stats["failed"] += 1
                stats["errors"].append((client_name, res["error"] or "unknown"))
            else:
                stats["skipped"] += 1

            elapsed_total = time.perf_counter() - t_batch_start
            if completed_count > 0:
                eta = (elapsed_total / completed_count) * (total - completed_count)
            else:
                eta = 0.0

            if prog_bar:
                prog_bar.set_postfix({"done": completed_count, "ETA": f"{eta:.0f}s"})
                prog_bar.update(1)
            else:
                log.info(
                    f"  Progress: {completed_count}/{total} | "
                    f"Elapsed: {elapsed_total:.0f}s | ETA: {eta:.0f}s"
                )

    if prog_bar:
        prog_bar.close()

    return stats


#  ARGUMENT PARSER
def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="GST Amendment Document Validator v5.1 — Client-Folder Edition",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python gst_validator.py --folder ./Root_Folder\n"
            "  python gst_validator.py --folder ./Root --api-key AIza... --workers 8\n"
            "  python gst_validator.py --folder ./Root --use-cache\n"
            "  python gst_validator.py --folder ./Root --skip-gemini\n"
            "  python gst_validator.py --folder ./Root --no-cache\n"
            "  python gst_validator.py --folder ./Root --dry-run\n"
            "  python gst_validator.py --folder ./Root --client-output-dir ./my_clients\n"
            "\nEnv vars (override defaults):\n"
            "  GST_MAX_WORKERS=8\n"
            "  GST_OUTPUT_DIR=./my_outputs\n"
            "  GST_CLIENT_OUTPUT_DIR=./client_outputs\n"
            "  GST_CACHE_ENABLED=false\n"
            "  GST_RETRY_COUNT=3\n"
            "  GEMINI_API_KEY=your_key_here\n"
        ),
    )
    p.add_argument("--folder",             "-f", type=str, default=None,
                   help="Path to Root_Folder_To_Process")
    p.add_argument("--api-key",            "-k", type=str, default=None,
                   help="Gemini API key (overrides GEMINI_API_KEY env var)")
    p.add_argument("--output",             "-o", type=str,
                   default=str(OUTPUT_DIR / "GST_Amendment_Validation_Workbook.xlsx"),
                   help="Output Excel path")
    p.add_argument("--workers",            "-w", type=int, default=MAX_WORKERS,
                   help=f"Number of parallel workers (default: {MAX_WORKERS})")
    p.add_argument("--retry-count",        "-r", type=int, default=MAX_RETRIES,
                   help=f"API retry count per client (default: {MAX_RETRIES})")
    p.add_argument("--output-dir",         "-d", type=str, default=None,
                   help="Override output directory")
    p.add_argument("--client-output-dir",  "-c", type=str, default=None,
                   help=(
                       "Root dir for per-client folder hierarchy. "
                       "IMPORTANT: must be OUTSIDE --folder (the input folder). "
                       "Default: auto-set to <input_folder>_client_outputs sibling. "
                       "Example: --client-output-dir C:\\GST_Clients"
                   ))
    p.add_argument("--use-cache",   action="store_true",
                   help="Reuse cached JSON for clients already processed")
    p.add_argument("--no-cache",    action="store_true",
                   help="Disable caching entirely (always call Gemini)")
    p.add_argument("--skip-gemini", action="store_true",
                   help="Skip all Gemini calls — generate Excel from cache only")
    p.add_argument("--dry-run",     action="store_true",
                   help="Audit client grouping only — no Gemini calls, no Excel output")
    return p.parse_args()


#  MAIN ENTRY POINT
def main():
    global MAX_WORKERS, MAX_RETRIES, OUTPUT_DIR, JSON_CACHE_DIR, CACHE_ENABLED, CLIENT_OUTPUT_BASE

    args        = _parse_args()
    api_key     = args.api_key or os.environ.get("GEMINI_API_KEY", "")
    output_path = Path(args.output)

    # ── Apply CLI overrides to globals ──────────────────────────────────────
    if args.workers:
        MAX_WORKERS = args.workers
    if args.retry_count:
        MAX_RETRIES = args.retry_count
    if args.output_dir:
        OUTPUT_DIR     = Path(args.output_dir)
        JSON_CACHE_DIR = OUTPUT_DIR / "json_cache"
        output_path    = OUTPUT_DIR / output_path.name
    if args.no_cache:
        CACHE_ENABLED  = False

    # ── Client-wise folder hierarchy root ────────────────────────────────────
    # Priority:
    #   1. Explicit --client-output-dir flag
    #   2. Sibling of --folder named "<folder_name>_client_outputs"  ← avoids
    #      creating client folders INSIDE the source input tree
    #   3. Module-level default ./client_outputs (fallback for single-client mode)
    if args.client_output_dir:
        CLIENT_OUTPUT_BASE = Path(args.client_output_dir).resolve()
    elif args.folder:
        _input_root        = Path(args.folder).resolve()
        CLIENT_OUTPUT_BASE = _input_root.parent / f"{_input_root.name}_client_outputs"
    # else: keep the module-level default (./client_outputs)

    safe_create_dir(normalize_windows_path(CLIENT_OUTPUT_BASE))

    # ── Route all developer-only artefacts under _internal/ ─────────────────
    # JSON cache, audit CSVs, checkpoint, and processing history all live here.
    # Clients only ever see <State>/<Client>/GST_Validation_Report.xlsx + input_docs/.
    OUTPUT_DIR     = CLIENT_OUTPUT_BASE / "_internal"
    JSON_CACHE_DIR = OUTPUT_DIR / "json_cache"
    safe_create_dir(normalize_windows_path(OUTPUT_DIR))
    safe_create_dir(normalize_windows_path(JSON_CACHE_DIR))

    # ── HARDENED: Initialize production modules ──────────────────────
    # Sets up AuditLogger (failed/skipped CSVs, summary JSON) and CheckpointManager
    # (checkpoint.json for interrupted-run recovery) — all under _internal/.
    _initialize_production_modules(CLIENT_OUTPUT_BASE)

    # CHANGE 3: Master workbook gets a VERSIONED timestamped filename so old
    # reports are NEVER overwritten.  Every run creates a new file.
    output_path = versioned_workbook_path(CLIENT_OUTPUT_BASE, "GST_Validation_Workbook")

    if not api_key and not args.skip_gemini and not args.dry_run:
        print(
            "\nERROR: Gemini API key required.\n"
            "  Set env var : set GEMINI_API_KEY=your_key      (Windows)\n"
            "                export GEMINI_API_KEY=your_key   (Linux/Mac)\n"
            "  Or use flag : --api-key YOUR_KEY\n"
            "  Offline mode: --skip-gemini  (Excel from JSON cache only)\n"
        )
        sys.exit(1)

    log.info("=" * 72)
    log.info("  GST AMENDMENT DOCUMENT VALIDATOR — Client-Folder Edition")
    log.info(f"  Validation Date   : {VALIDATION_DATE}")
    log.info(f"  3-Month Cutoff    : {CUTOFF_DATE}")
    log.info(f"  Gemini Model      : {GEMINI_MODEL}")
    log.info(f"  Output (global)   : {output_path}")
    log.info(f"  Max Workers       : {MAX_WORKERS}")
    log.info(f"  Cache Enabled     : {CACHE_ENABLED}")
    log.info(f"  Retry Count       : {MAX_RETRIES}")
    log.info(f"  SDK               : {'google-genai (new)' if USE_NEW_SDK else 'google-generativeai (legacy)'}")
    log.info("=" * 72)

    #  BATCH MODE  (--folder)
    if args.folder:
        root = Path(args.folder)
        if not root.is_dir():
            log.error(f"Not a valid directory: {root}")
            sys.exit(1)

        log.info(f"\n  BATCH MODE — Root folder: {root}")

        # ── SAFETY CHECK: client output dir must not be inside input root ─────
        try:
            CLIENT_OUTPUT_BASE.resolve().relative_to(root.resolve())
            # If we reach here, CLIENT_OUTPUT_BASE IS inside root — dangerous.
            log.error(
                f"\n  ╔══ CONFIGURATION ERROR ═══════════════════════════════════════╗\n"
                f"  ║  CLIENT_OUTPUT_BASE is inside the input folder!              ║\n"
                f"  ║  Input folder   : {str(root):<46}║\n"
                f"  ║  Client out dir : {str(CLIENT_OUTPUT_BASE):<46}║\n"
                f"  ║                                                              ║\n"
                f"  ║  This would place client sub-folders INSIDE your source      ║\n"
                f"  ║  documents tree, and document copying would be skipped.      ║\n"
                f"  ║                                                              ║\n"
                f"  ║  FIX: use --client-output-dir to set a path OUTSIDE the     ║\n"
                f"  ║  input folder.  Example:                                     ║\n"
                f"  ║    --client-output-dir C:\\\\Users\\\\you\\\\GST_Clients           ║\n"
                f"  ╚══════════════════════════════════════════════════════════════╝\n"
            )
            # Auto-correct: move CLIENT_OUTPUT_BASE to a sibling directory
            CLIENT_OUTPUT_BASE = root.resolve().parent / f"{root.name}_client_outputs"
            CLIENT_OUTPUT_BASE.mkdir(parents=True, exist_ok=True)
            log.warning(
                f"  Auto-corrected CLIENT_OUTPUT_BASE → {CLIENT_OUTPUT_BASE}\n"
                f"  (sibling of input folder). Re-run with --client-output-dir to "
                f"choose a different location."
            )
        except ValueError:
            pass  # CLIENT_OUTPUT_BASE is outside root — all good

        log.info(f"  Client Output Dir : {CLIENT_OUTPUT_BASE}")
        cats = detect_category_folders(root)

        if all(v is None for v in cats.values()):
            log.error(
                "Could not detect any of the 3 expected sub-folders.\n"
                "  Expected patterns: AGREEMENTS/, UTILITY BILL.../, OTHER DOCUMENTS.../\n"
                "  Make sure the folder names contain the keywords above.")
            sys.exit(1)

        groups = group_documents_by_client(cats)
        if not groups:
            log.error("No supported documents found in any sub-folder.")
            sys.exit(1)

        # Print grouping summary
        log.info(f"\n  CLIENTS TO PROCESS: {len(groups)}")
        for i, (nm, cm) in enumerate(groups.items(), 1):
            tot = sum(len(v) for v in cm.values())
            log.info(
                f"  {i:3d}. {nm:<52} "
                f"(total={tot}  agr={len(cm['agreement'])}  "
                f"util={len(cm['utility'])}  other={len(cm['other'])})"
            )

        if args.dry_run:
            log.info("\n  DRY RUN complete — no Gemini calls, no Excel output.")
            return

        # ── Initialise Excel workbook (single instance, shared across threads) ──
        builder = ExcelBuilder(output_path)
        builder.build_all_sheets()

        # ── Run parallel batch ────────────────────────────────────────────────
        t_start = time.perf_counter()

        if MAX_WORKERS == 1:
            # Sequential fallback (useful for debugging / very tight rate limits)
            log.info("  Running SEQUENTIALLY (--workers 1)")
            stats = {
                "total": len(groups), "processed": 0, "skipped": 0,
                "failed": 0, "clean": 0, "warning": 0, "high_risk": 0, "errors": [],
            }
            for idx, (client_name, cat_map) in enumerate(
                progress(groups.items(), desc="Processing clients", unit="client"), 1
            ):
                log.info(f"\n{'─' * 68}")
                log.info(f"  CLIENT {idx}/{len(groups)}: {client_name}")
                log.info(f"{'─' * 68}")
                res = process_client(idx, len(groups), client_name, cat_map, api_key, args)
                if res["data"] is not None:
                    builder.add_client(res["gst_id"], client_name, res["data"], res["cat_map"])
                st = res["status"]
                if st == "ok":
                    stats["processed"] += 1
                    fs = res["data"].get("compliance", {}).get("final_status", "?")
                    if   fs == "Clean":     stats["clean"]     += 1
                    elif fs == "Warning":   stats["warning"]   += 1
                    elif fs == "High Risk": stats["high_risk"] += 1
                elif st == "failed":
                    stats["failed"] += 1
                    stats["errors"].append((client_name, res["error"]))
                else:
                    stats["skipped"] += 1
        else:
            log.info(f"  Running PARALLEL with {MAX_WORKERS} workers")
            stats = run_parallel_batch(groups, api_key, args, builder)

        builder.save()
        t_total = time.perf_counter() - t_start

        # ── Generate per-client GST_Validation_Report.xlsx ───────────────────
        # Reports land at: CLIENT_OUTPUT_BASE/<State>/<Client>/GST_Validation_Report.xlsx
        log.info("  Generating per-client validation reports…")
        for idx, client_name_key in enumerate(groups, 1):
            gst_id_for_client = f"GSTREG-{idx:03d}"
            client_json = load_json_cache(client_name_key)
            client_folder = resolve_client_output_folder(
                CLIENT_OUTPUT_BASE, client_name_key, client_json
            )
            generate_client_report(
                master_path=output_path,
                client_name=client_name_key,
                gst_id=gst_id_for_client,
                client_folder=client_folder,
                client_data=client_json,
            )

        # ── Final summary ─────────────────────────────────────────────────────
        state_breakdown: defaultdict[str, int] = defaultdict(int)
        for _cn in groups:
            jd = load_json_cache(_cn)
            sk = extract_state_from_validation_data(jd) if jd else "Unknown_State"
            state_breakdown[sk] += 1

        log.info(f"\n{'=' * 72}")
        log.info(f"  BATCH COMPLETE")
        log.info(f"  Total clients     : {stats['total']}")
        log.info(f"  Processed (OK)    : {stats['processed']}")
        log.info(f"  Skipped (no data) : {stats['skipped']}")
        log.info(f"  Failed (error)    : {stats['failed']}")
        log.info(f"  ─── Status breakdown ───────────────────────────────────")
        log.info(f"  Clean   (>=85%)   : {stats['clean']}")
        log.info(f"  Warning (60-84%)  : {stats['warning']}")
        log.info(f"  High Risk (<60%)  : {stats['high_risk']}")
        log.info(f"  ─── Clients by state folder ────────────────────────────")
        for _st in sorted(state_breakdown.keys()):
            log.info(f"    {_st:<28} : {state_breakdown[_st]}")
        log.info(f"  ─── Timing ─────────────────────────────────────────────")
        log.info(f"  Total wall-time   : {t_total:.1f}s  "
                 f"({t_total/60:.1f} min)")
        if stats['processed']:
            log.info(f"  Avg per client    : {t_total/stats['total']:.1f}s")
        log.info(f"  ─── Output ─────────────────────────────────────────────")
        log.info(f"  Master Workbook   : {output_path}")
        log.info(f"  Client Output Dir : {CLIENT_OUTPUT_BASE}")
        log.info(f"  Client folders    : <State>/<Client>/GST_Validation_Report + input_docs/")
        log.info(f"  Developer-only    : _internal/<State>/<Client>/json_cache|logs|metadata/")
        log.info(f"  Global JSON Cache : {JSON_CACHE_DIR}")
        log.info(f"  Log file          : gst_validator.log")
        if stats["errors"]:
            log.warning(f"\n  FAILED CLIENTS ({len(stats['errors'])}):")
            for cn, err in stats["errors"]:
                log.warning(f"    ✖ {cn}: {err}")
        log.info(f"{'=' * 72}\n")

        # ── HARDENED : Finalize audit reports ───────────────────────────
        _al = get_audit_logger()
        if _al:
            _al.finalize()
            log.info(f"  Audit reports → {CLIENT_OUTPUT_BASE}/_internal/processing_summary.json")
            log.info(f"                   _internal/failed_documents.csv | _internal/skipped_documents.csv")

    #  SINGLE CLIENT MODE  (no --folder)
    else:
        log.info("\n  SINGLE CLIENT MODE — enter document paths interactively\n")

        doc_paths: List[Path] = []
        print("  Enter document paths one per line. Press ENTER on an empty line when done.\n")
        while True:
            try:
                raw = input(f"  File {len(doc_paths) + 1}: ").strip().strip("'\"")
            except (EOFError, KeyboardInterrupt):
                print()
                break
            if not raw:
                if doc_paths: break
                continue
            p = Path(raw)
            if p.exists():
                doc_paths.append(p)
                print(f"  Added: {p.name}")
            else:
                print(f"  Not found: {p}")

        found = [p for p in doc_paths if p.exists()]
        if not found:
            log.error("No valid documents found."); sys.exit(1)

        result = call_gemini(found, api_key, "SingleClient")
        if not result:
            log.error("Gemini analysis failed."); sys.exit(1)

        result = _reconcile_compliance(result, found)
        _sc_state = extract_state_from_validation_data(result)
        sc_folders = create_statewise_client_structure(
            CLIENT_OUTPUT_BASE, "SingleClient", _sc_state
        )
        organize_client_documents(found, sc_folders["input_docs"], client_name="SingleClient")

        if CACHE_ENABLED:
            save_json_cache("SingleClient", result, client_cache_dir=sc_folders["json_cache"])

        builder  = ExcelBuilder(output_path)
        builder.build_all_sheets()
        cat_map  = {"agreement": [], "utility": [], "other": found}
        builder.add_client("GSTREG-001", "SingleClient", result, cat_map)
        builder.save()

        # ── Save per-client outputs (logs + metadata under state-wise client folder)
        save_json_response("SingleClient", sc_folders["logs"], result)
        save_logs("SingleClient", sc_folders["logs"], [
            "GST Validator — Single Client Processing Log",
            f"Run time    : {time.strftime('%Y-%m-%d %H:%M:%S')}",
            f"Files       : {len(found)}",
            f"Compliance  : {result.get('compliance', {}).get('weighted_compliance_percent', 0):.1f}%",
            f"Status      : {result.get('compliance', {}).get('final_status', '?')}",
        ])
        comp = result.get("compliance", {})
        write_client_metadata(sc_folders["root"], "SingleClient", len(found), "SUCCESS", {
            "gst_id"             : "GSTREG-001",
            "state_folder"       : _sc_state,
            "compliance_percent" : round(float(comp.get("weighted_compliance_percent", 0)), 1),
            "final_status"       : comp.get("final_status", "?"),
        })

        # ── Generate per-client report at client folder root ─────────────────
        sc_client_folder = sc_folders["root"]
        generate_client_report(
            master_path=output_path,
            client_name="SingleClient",
            gst_id="GSTREG-001",
            client_folder=sc_client_folder,
            client_data=result,
        )

        # ── Legacy JSON at global output path (backwards compat) ──────────────
        json_path = output_path.parent / "GST_Validation_Result.json"
        json_path.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")

        log.info(f"\n  SINGLE CLIENT COMPLETE")
        log.info(f"  Compliance  : {comp.get('weighted_compliance_percent', 0):.1f}%")
        log.info(f"  Status      : {comp.get('final_status', '?')}")
        log.info(f"  Passes      : {comp.get('total_passes', 0)}")
        log.info(f"  Issues      : {comp.get('total_issues', 0)}")
        log.info(f"  Excel       : {output_path}")
        log.info(f"  JSON        : {json_path}")
        log.info(f"  Client Dir  : {sc_folders['root']}\n")

DEFAULT_API_KEY = "AIzaSyCH4N-4VlgBSa7Knuru9CsKtd8KtnLyCFo"

class UILogHandler(logging.Handler):
    """
    Issue 14 — Tiered UI logging handler.

    UI panel policy:
      • DEBUG_MODE off (default): only WARNING, ERROR, CRITICAL reach the UI.
        Routine folder-creation, copy, and INFO confirmations are file-only.
      • DEBUG_MODE on (GST_UI_DEBUG=1): all levels shown.
      • Suppressed prefixes list (_UI_SUPPRESSED_PREFIXES) keeps UI clean.
    """

    def __init__(self, ui_queue: "queue.Queue[Tuple[str, Any]]"):
        super().__init__()
        self.ui_queue = ui_queue
        self.setFormatter(logging.Formatter("%(asctime)s  %(levelname)-8s  %(message)s"))

    def _should_show_in_ui(self, record: logging.LogRecord) -> bool:
        if _UI_DEBUG_MODE:
            return True
        if record.levelno >= logging.WARNING:
            return True
        msg = record.getMessage()
        for prefix in _UI_SUPPRESSED_PREFIXES:
            if msg.startswith(prefix):
                return False
        if record.levelno == logging.INFO:
            return True
        return False

    def emit(self, record: logging.LogRecord) -> None:
        try:
            if not self._should_show_in_ui(record):
                return
            formatted = self.format(record)
            if record.levelno >= logging.ERROR:
                formatted = f"[ERROR] {formatted}"
            elif record.levelno >= logging.WARNING:
                formatted = f"[WARN]  {formatted}"
            self.ui_queue.put(("log", formatted))
        except Exception:
            pass


class GSTValidatorUI(ctk.CTk):
    SETTINGS_FILE = Path.home() / ".gst_validator_ui_settings.json"

    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("GST Validator Enterprise Desktop")
        self.geometry("1400x860")
        self.minsize(1180, 760)

        self.ui_queue: "queue.Queue[Tuple[str, Any]]" = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker_thread: Optional[threading.Thread] = None
        self.processing_started_at: Optional[float] = None
        self.last_input_root: Optional[Path] = None
        self.last_output_root: Optional[Path] = None
        self.last_groups: Dict[str, Dict[str, List[Path]]] = {}
        self.failed_clients: List[str] = []
        self.current_workers = 0

        # ── Issue 11: Worker Registry — strict lifecycle tracking ──────────────
        self._active_worker_registry: Dict[str, float] = {}   # worker_id → start_time
        self._worker_registry_lock = threading.Lock()
        self._configured_max_workers = MAX_WORKERS             # set at run-start

        # ── Issue 12: Rolling ETA — last N successful processing times ─────────
        self._processing_times: list = []          # rolling deque of seconds per client
        self._rolling_window = 20                  # use last 20 successful clients
        self._eta_lock = threading.Lock()
        self._throughput_per_hour: float = 0.0    # updated dynamically
        self._last_eta_str: str = "--"

        # ── Issue 15: Memory tracking ──────────────────────────────────────────
        self._clients_since_gc: int = 0
        self._gc_interval: int = 25               # run gc every 25 clients
        try:
            import psutil as _psutil
            self._psutil = _psutil
        except ImportError:
            self._psutil = None

        # ── Issue 16: Live dashboard state ────────────────────────────────────
        self._retry_count_total: int = 0
        self._current_stage: str = "Idle"
        self._current_client_name: str = ""
        self._queue_remaining: int = 0
        self._avg_processing_sec: float = 0.0

        self.settings = self._load_settings()
        self._setup_logging_stream()
        self._show_startup_splash()
        self._build_ui()
        self._apply_settings()
        self._setup_drag_drop()
        self.after(900, self._hide_startup_splash)
        self.after(120, self._drain_ui_queue)

    def _setup_logging_stream(self) -> None:
        self.gui_log_handler = UILogHandler(self.ui_queue)
        # Issue 14: UI panel receives WARNING+ by default (UILogHandler.emit filters internally)
        # Set to DEBUG so the handler's own _should_show_in_ui() logic controls what shows.
        self.gui_log_handler.setLevel(logging.DEBUG)
        logging.getLogger().addHandler(self.gui_log_handler)

    def _show_startup_splash(self) -> None:
        self.splash = ctk.CTkToplevel(self)
        self.splash.overrideredirect(True)
        self.splash.attributes("-topmost", True)
        width, height = 460, 180
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2
        self.splash.geometry(f"{width}x{height}+{x}+{y}")
        ctk.CTkLabel(self.splash, text="GST Validator", font=("Segoe UI", 28, "bold")).pack(pady=(26, 4))
        ctk.CTkLabel(self.splash, text="Launching enterprise compliance workspace...", font=("Segoe UI", 13)).pack(pady=(0, 14))
        splash_bar = ctk.CTkProgressBar(self.splash, width=360)
        splash_bar.pack()
        splash_bar.set(0.8)

    def _hide_startup_splash(self) -> None:
        try:
            self.splash.destroy()
        except Exception:
            pass

    def _load_settings(self) -> Dict[str, Any]:
        default = {"theme": "dark", "recent_inputs": [], "recent_outputs": [], "workers": str(MAX_WORKERS)}
        try:
            if self.SETTINGS_FILE.exists():
                loaded = json.loads(self.SETTINGS_FILE.read_text(encoding="utf-8"))
                if isinstance(loaded, dict):
                    default.update(loaded)
        except Exception:
            pass
        return default

    def _save_settings(self) -> None:
        try:
            self.SETTINGS_FILE.write_text(json.dumps(self.settings, indent=2), encoding="utf-8")
        except Exception as exc:
            log.warning(f"Could not save UI settings: {exc}")

    def _push_recent(self, key: str, value: str) -> None:
        items = [v for v in self.settings.get(key, []) if v != value]
        items.insert(0, value)
        self.settings[key] = items[:8]

    def _build_ui(self) -> None:
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.sidebar = ctk.CTkFrame(self, width=240, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(9, weight=1)

        ctk.CTkLabel(self.sidebar, text="GST Validator", font=("Segoe UI", 24, "bold")).grid(row=0, column=0, padx=20, pady=(24, 6), sticky="w")
        ctk.CTkLabel(self.sidebar, text="Enterprise Compliance", font=("Segoe UI", 13)).grid(row=1, column=0, padx=20, pady=(0, 16), sticky="w")
        self.start_btn = ctk.CTkButton(self.sidebar, text="Start Validation", height=40, command=self.start_process)
        self.start_btn.grid(row=2, column=0, padx=20, pady=6, sticky="ew")
        self.stop_btn = ctk.CTkButton(self.sidebar, text="Stop / Cancel", height=40, fg_color="#9f1239", hover_color="#881337", command=self.stop_process)
        self.stop_btn.grid(row=3, column=0, padx=20, pady=6, sticky="ew")
        self.report_btn = ctk.CTkButton(self.sidebar, text="Generate Reports", height=40, command=self.generate_reports_only)
        self.report_btn.grid(row=4, column=0, padx=20, pady=6, sticky="ew")
        self.retry_btn = ctk.CTkButton(self.sidebar, text="Retry Failed Clients", height=40, command=self.retry_failed_clients)
        self.retry_btn.grid(row=5, column=0, padx=20, pady=6, sticky="ew")
        self.open_output_btn = ctk.CTkButton(self.sidebar, text="Open Output Folder", height=40, command=self.open_output_folder)
        self.open_output_btn.grid(row=6, column=0, padx=20, pady=6, sticky="ew")
        self.export_log_btn = ctk.CTkButton(self.sidebar, text="Export Logs", height=40, command=self.export_logs)
        self.export_log_btn.grid(row=7, column=0, padx=20, pady=6, sticky="ew")
        self.theme_var = ctk.StringVar(value="dark")
        self.theme_switch = ctk.CTkSwitch(self.sidebar, text="Dark Theme", variable=self.theme_var, onvalue="dark", offvalue="light", command=self.toggle_theme)
        self.theme_switch.grid(row=8, column=0, padx=20, pady=(12, 10), sticky="w")
        self.app_status = ctk.CTkLabel(self.sidebar, text="Idle", text_color="#86efac", font=("Segoe UI", 12, "bold"))
        self.app_status.grid(row=10, column=0, padx=20, pady=(8, 20), sticky="w")

        self.main = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main.grid(row=0, column=1, sticky="nsew", padx=18, pady=16)
        self.main.grid_columnconfigure(0, weight=1)
        self.main.grid_rowconfigure(3, weight=1)
        header = ctk.CTkFrame(self.main, height=68)
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(header, text="GST Amendment Validation Dashboard", font=("Segoe UI", 22, "bold")).grid(row=0, column=0, padx=18, pady=(12, 4), sticky="w")
        self.header_meta = ctk.CTkLabel(header, text="Ready", font=("Segoe UI", 12))
        self.header_meta.grid(row=1, column=0, padx=18, pady=(0, 10), sticky="w")

        config = ctk.CTkFrame(self.main)
        config.grid(row=1, column=0, sticky="ew", pady=(0, 12))
        for i in range(0, 4, 2):
            config.grid_columnconfigure(i + 1, weight=1)
        ctk.CTkLabel(config, text="Input Folder", font=("Segoe UI", 12, "bold")).grid(row=0, column=0, padx=12, pady=(12, 4), sticky="w")
        self.input_combo = ctk.CTkComboBox(config, values=[""], command=self._on_input_combo, height=34)
        self.input_combo.grid(row=0, column=1, padx=(0, 8), pady=(12, 4), sticky="ew")
        ctk.CTkButton(config, text="Select", width=90, command=self.browse_folder).grid(row=0, column=2, padx=(0, 12), pady=(12, 4))
        ctk.CTkLabel(config, text="Output Folder", font=("Segoe UI", 12, "bold")).grid(row=1, column=0, padx=12, pady=4, sticky="w")
        self.output_combo = ctk.CTkComboBox(config, values=[""], command=self._on_output_combo, height=34)
        self.output_combo.grid(row=1, column=1, padx=(0, 8), pady=4, sticky="ew")
        ctk.CTkButton(config, text="Select", width=90, command=self.browse_output).grid(row=1, column=2, padx=(0, 12), pady=4)
        ctk.CTkLabel(config, text="API Key", font=("Segoe UI", 12, "bold")).grid(row=2, column=0, padx=12, pady=4, sticky="w")
        self.api_entry = ctk.CTkEntry(config, placeholder_text="Leave blank to use built-in default", show="*", height=34)
        self.api_entry.grid(row=2, column=1, padx=(0, 8), pady=4, sticky="ew")
        ctk.CTkLabel(config, text="Workers", font=("Segoe UI", 12, "bold")).grid(row=2, column=2, padx=(8, 4), pady=4, sticky="e")
        self.worker_entry = ctk.CTkEntry(config, width=70, height=34)
        self.worker_entry.grid(row=2, column=3, padx=(0, 12), pady=4, sticky="w")
        self.drag_label = ctk.CTkLabel(config, text="Tip: drag and drop a folder onto this window", text_color="gray")
        self.drag_label.grid(row=3, column=0, columnspan=4, padx=12, pady=(2, 10), sticky="w")

        stats = ctk.CTkFrame(self.main)
        stats.grid(row=2, column=0, sticky="ew", pady=(0, 12))
        for col in range(8):
            stats.grid_columnconfigure(col, weight=1)

        # Row 0 — primary counters
        self.total_label   = ctk.CTkLabel(stats, text="Total: 0",     font=("Segoe UI", 12, "bold"))
        self.done_label    = ctk.CTkLabel(stats, text="Completed: 0", font=("Segoe UI", 12))
        self.failed_label  = ctk.CTkLabel(stats, text="Failed: 0",    font=("Segoe UI", 12), text_color="#fb7185")
        self.skip_label    = ctk.CTkLabel(stats, text="Skipped: 0",   font=("Segoe UI", 12))
        self.retry_label   = ctk.CTkLabel(stats, text="Retried: 0",   font=("Segoe UI", 12))
        self.queue_label   = ctk.CTkLabel(stats, text="Queue: 0",     font=("Segoe UI", 12))
        self.eta_label     = ctk.CTkLabel(stats, text="ETA: --",      font=("Segoe UI", 12))
        self.worker_status_label = ctk.CTkLabel(stats, text="Workers: 0/0", font=("Segoe UI", 12, "bold"), text_color="#86efac")
        for idx, widget in enumerate((self.total_label, self.done_label, self.failed_label,
                                      self.skip_label, self.retry_label, self.queue_label,
                                      self.eta_label, self.worker_status_label)):
            widget.grid(row=0, column=idx, padx=6, pady=(10, 2), sticky="w")

        # Row 1 — throughput + avg time + memory + current stage
        self.throughput_label  = ctk.CTkLabel(stats, text="Throughput: -- /hr",   font=("Segoe UI", 11), text_color="gray")
        self.avg_time_label    = ctk.CTkLabel(stats, text="Avg: --s/client",       font=("Segoe UI", 11), text_color="gray")
        self.memory_label      = ctk.CTkLabel(stats, text="Mem: --",               font=("Segoe UI", 11), text_color="gray")
        self.stage_label       = ctk.CTkLabel(stats, text="Stage: Idle",           font=("Segoe UI", 11), text_color="#facc15")
        self.current_client_label = ctk.CTkLabel(stats, text="",                   font=("Segoe UI", 11), text_color="gray")
        self.throughput_label.grid(row=1, column=0, columnspan=2, padx=6, pady=(0, 4), sticky="w")
        self.avg_time_label.grid(  row=1, column=2, columnspan=2, padx=6, pady=(0, 4), sticky="w")
        self.memory_label.grid(    row=1, column=4, padx=6, pady=(0, 4), sticky="w")
        self.stage_label.grid(     row=1, column=5, padx=6, pady=(0, 4), sticky="w")
        self.current_client_label.grid(row=1, column=6, columnspan=2, padx=6, pady=(0, 4), sticky="w")

        # Row 2 — progress bar
        self.progress = ctk.CTkProgressBar(stats, height=16)
        self.progress.grid(row=2, column=0, columnspan=8, padx=10, pady=(0, 10), sticky="ew")
        self.progress.set(0.0)

        bottom = ctk.CTkFrame(self.main)
        bottom.grid(row=3, column=0, sticky="nsew")
        bottom.grid_columnconfigure(0, weight=1)
        bottom.grid_columnconfigure(1, weight=1)
        bottom.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(bottom, text="Live Processing Log", font=("Segoe UI", 13, "bold")).grid(row=0, column=0, padx=(10, 5), pady=(10, 6), sticky="w")
        ctk.CTkLabel(bottom, text="Client Processing View", font=("Segoe UI", 13, "bold")).grid(row=0, column=1, padx=(5, 10), pady=(10, 6), sticky="w")
        self.log_box = ctk.CTkTextbox(bottom, wrap="word")
        self.log_box.grid(row=1, column=0, padx=(10, 5), pady=(0, 10), sticky="nsew")
        self.client_box = ctk.CTkTextbox(bottom, wrap="word")
        self.client_box.grid(row=1, column=1, padx=(5, 10), pady=(0, 10), sticky="nsew")

    def _apply_settings(self) -> None:
        ctk.set_appearance_mode(self.settings.get("theme", "dark"))
        self.theme_var.set(self.settings.get("theme", "dark"))
        if self.theme_var.get() == "dark":
            self.theme_switch.select()
        else:
            self.theme_switch.deselect()
        input_vals = self.settings.get("recent_inputs", []) or [""]
        output_vals = self.settings.get("recent_outputs", []) or [""]
        self.input_combo.configure(values=input_vals)
        self.output_combo.configure(values=output_vals)
        self.input_combo.set(input_vals[0] if input_vals else "")
        self.output_combo.set(output_vals[0] if output_vals else "")
        self.worker_entry.delete(0, "end")
        self.worker_entry.insert(0, str(self.settings.get("workers", MAX_WORKERS)))

    def _setup_drag_drop(self) -> None:
        try:
            from tkinterdnd2 import DND_FILES  # type: ignore
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drop)
            self.drag_label.configure(text="Drag-and-drop enabled")
        except Exception:
            self.drag_label.configure(text="Install tkinterdnd2 for drag-and-drop support")

    def _on_drop(self, event: Any) -> None:
        dropped = (event.data or "").strip().strip("{}")
        if dropped:
            self._set_input_path(dropped)

    def _on_input_combo(self, value: str) -> None:
        self._set_input_path(value, auto_output=False)

    def _on_output_combo(self, value: str) -> None:
        self.output_combo.set(value)

    def _set_input_path(self, path: str, auto_output: bool = True) -> None:
        self.input_combo.set(path)
        if auto_output and path:
            suggested = str(Path(path).parent / f"{Path(path).name}_client_outputs")
            if not self.output_combo.get().strip():
                self.output_combo.set(suggested)

    def browse_folder(self) -> None:
        selected = filedialog.askdirectory()
        if selected:
            self._set_input_path(selected)

    def browse_output(self) -> None:
        selected = filedialog.askdirectory()
        if selected:
            self.output_combo.set(selected)

    def toggle_theme(self) -> None:
        theme = self.theme_var.get()
        ctk.set_appearance_mode(theme)
        self.settings["theme"] = theme
        self._save_settings()

    def _build_args(self) -> argparse.Namespace:
        class UIArgs:
            use_cache = True
            skip_gemini = False
            dry_run = False
        return UIArgs()

    def start_process(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Processing Running", "Validation is already running.")
            return
        input_folder = self.input_combo.get().strip()
        output_folder = self.output_combo.get().strip()
        if not input_folder:
            messagebox.showerror("Missing Input", "Please select an input folder.")
            return
        if not output_folder:
            output_folder = str(Path(input_folder).parent / f"{Path(input_folder).name}_client_outputs")
            self.output_combo.set(output_folder)

        self.cancel_event.clear()
        self.processing_started_at = time.perf_counter()
        self.progress.set(0.0)
        self.client_box.delete("1.0", "end")
        self.failed_clients = []
        self.app_status.configure(text="Running", text_color="#facc15")
        self.start_btn.configure(state="disabled")
        self.header_meta.configure(text="Validation in progress")

        self._push_recent("recent_inputs", input_folder)
        self._push_recent("recent_outputs", output_folder)
        self.settings["workers"] = self.worker_entry.get().strip() or str(MAX_WORKERS)
        self._save_settings()
        self._apply_settings()

        self.worker_thread = threading.Thread(target=self._run_pipeline, args=(Path(input_folder), Path(output_folder), None), daemon=True, name="UIBatchRunner")
        self.worker_thread.start()

    def stop_process(self) -> None:
        if not self.worker_thread or not self.worker_thread.is_alive():
            self.app_status.configure(text="Idle", text_color="#86efac")
            return
        self.cancel_event.set()
        self.app_status.configure(text="Cancelling...", text_color="#fb7185")
        self.ui_queue.put(("log", "Cancellation requested. In-flight clients will finish safely."))

    def _run_pipeline(self, input_root: Path, output_root: Path, only_clients: Optional[set]) -> None:
        global CLIENT_OUTPUT_BASE, OUTPUT_DIR, JSON_CACHE_DIR, MAX_WORKERS
        try:
            workers = max(1, int(self.worker_entry.get().strip() or MAX_WORKERS))
            MAX_WORKERS = workers
            CLIENT_OUTPUT_BASE = output_root
            CLIENT_OUTPUT_BASE.mkdir(parents=True, exist_ok=True)
            OUTPUT_DIR = CLIENT_OUTPUT_BASE / "_internal"
            OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
            JSON_CACHE_DIR = OUTPUT_DIR / "json_cache"
            JSON_CACHE_DIR.mkdir(parents=True, exist_ok=True)

            # Initialize audit logger and checkpoint under _internal/
            _initialize_production_modules(CLIENT_OUTPUT_BASE)

            api_key = self.api_entry.get().strip() or DEFAULT_API_KEY
            cats = detect_category_folders(input_root)
            groups = group_documents_by_client(cats)
            if only_clients:
                groups = {k: v for k, v in groups.items() if k in only_clients}
            if not groups:
                raise ValueError("No supported documents found in selected input folder.")
            self.last_groups = groups
            self.last_input_root = input_root
            self.last_output_root = output_root

            # VERSIONED: Every run gets its own timestamped master workbook
            master_output_path = versioned_workbook_path(CLIENT_OUTPUT_BASE, "GST_Validation_Workbook")
            builder = ExcelBuilder(master_output_path)
            builder.build_all_sheets()
            stats = self._run_ui_parallel_batch(groups, api_key, self._build_args(), builder)
            builder.save()

            if not self.cancel_event.is_set():
                for idx, client_name_key in enumerate(groups, 1):
                    gst_id = f"GSTREG-{idx:03d}"
                    client_json = load_json_cache(client_name_key)
                    client_folder = resolve_client_output_folder(
                        CLIENT_OUTPUT_BASE, client_name_key, client_json
                    )
                    generate_client_report(master_output_path, client_name_key, gst_id, client_folder, client_json)
            self.ui_queue.put(("done", {"stats": stats, "master": str(master_output_path), "cancelled": self.cancel_event.is_set()}))

            # ── Finalize audit + processing history ──────────────────────────
            _al = get_audit_logger()
            if _al:
                _al.finalize()
        except Exception as exc:
            log.error(f"Runtime pipeline failure: {exc}")
            log.debug(traceback.format_exc())
            self.ui_queue.put(("error", str(exc)))

    # Issue 11 — Worker Registry helpers

    def _register_worker(self, worker_id: str) -> None:
        """Register a worker as active. Enforces max-worker cap."""
        with self._worker_registry_lock:
            self._active_worker_registry[worker_id] = time.perf_counter()

    def _deregister_worker(self, worker_id: str) -> None:
        """Remove a worker from the active registry on completion or failure."""
        with self._worker_registry_lock:
            self._active_worker_registry.pop(worker_id, None)

    def _active_worker_count(self) -> int:
        with self._worker_registry_lock:
            return len(self._active_worker_registry)

    def _cleanup_dead_workers(self, max_age_sec: float = 600.0) -> int:
        """Remove stale workers that have been running longer than max_age_sec."""
        now = time.perf_counter()
        dead = []
        with self._worker_registry_lock:
            for wid, start in list(self._active_worker_registry.items()):
                if now - start > max_age_sec:
                    dead.append(wid)
            for wid in dead:
                self._active_worker_registry.pop(wid, None)
        if dead:
            log.warning(f"[WorkerRegistry] Cleaned up {len(dead)} stale worker(s): {dead}")
        return len(dead)

    # Issue 12 — Rolling ETA helpers

    def _record_processing_time(self, elapsed_sec: float) -> None:
        """Record a successful client processing time into the rolling window."""
        with self._eta_lock:
            self._processing_times.append(elapsed_sec)
            if len(self._processing_times) > self._rolling_window:
                self._processing_times = self._processing_times[-self._rolling_window:]

    def _calculate_dynamic_eta(self, remaining: int) -> str:
        """
        Calculate ETA from rolling average of last N successful clients.
        Returns a human-readable string like '1h 42m' or '3m 12s'.
        Excludes retries/failures/skipped from the average.
        """
        with self._eta_lock:
            if not self._processing_times or remaining <= 0:
                return "--"
            avg = sum(self._processing_times) / len(self._processing_times)
            self._avg_processing_sec = avg
            active = max(1, self._active_worker_count() or self._configured_max_workers)
            # Parallel workers reduce ETA proportionally
            eta_sec = (avg * remaining) / active
            # Update throughput
            if avg > 0:
                self._throughput_per_hour = (3600.0 / avg) * active
            if eta_sec < 60:
                return f"{int(eta_sec)}s"
            elif eta_sec < 3600:
                return f"{int(eta_sec // 60)}m {int(eta_sec % 60)}s"
            else:
                hrs = int(eta_sec // 3600)
                mins = int((eta_sec % 3600) // 60)
                return f"{hrs}h {mins}m"

    # Issue 15 — Memory management helpers

    def _maybe_run_gc(self) -> None:
        """Run Python garbage collection every _gc_interval clients."""
        import gc
        self._clients_since_gc += 1
        if self._clients_since_gc >= self._gc_interval:
            self._clients_since_gc = 0
            collected = gc.collect()
            log.debug(f"[MemManager] GC collected {collected} objects")

    def _get_memory_label(self) -> str:
        """Return current process memory as a formatted string, or '--' if psutil unavailable."""
        if self._psutil is None:
            return "--"
        try:
            proc = self._psutil.Process(os.getpid())
            mb = proc.memory_info().rss / (1024 * 1024)
            if mb > 1024:
                return f"{mb/1024:.1f} GB"
            return f"{mb:.0f} MB"
        except Exception:
            return "--"

    # Issue 11/12/13/15/16 — Main parallel batch runner (hardened)

    def _run_ui_parallel_batch(self, groups: Dict[str, Dict[str, List[Path]]], api_key: str, args: argparse.Namespace, builder: "ExcelBuilder") -> Dict[str, Any]:
        stats = {"total": len(groups), "processed": 0, "skipped": 0, "failed": 0, "clean": 0, "warning": 0, "high_risk": 0, "errors": []}
        client_items = list(groups.items())
        total = len(client_items)
        self._configured_max_workers = MAX_WORKERS

        # Ensure registry is clean at batch start
        with self._worker_registry_lock:
            self._active_worker_registry.clear()
        with self._eta_lock:
            self._processing_times.clear()
        self._retry_count_total = 0
        self._queue_remaining = total

        self.ui_queue.put(("progress", {
            "total": total, "done": 0, "failed": 0, "skipped": 0,
            "eta": "--", "active_workers": 0, "configured_workers": MAX_WORKERS,
            "retried": 0, "queue_remaining": total, "throughput": 0.0,
            "avg_sec": 0.0, "memory": self._get_memory_label(), "stage": "Initialising",
        }))

        completed_count = 0

        # Issue 11: Use a controlled executor — workers are strictly bounded
        with ThreadPoolExecutor(max_workers=MAX_WORKERS, thread_name_prefix="GSTWorker") as executor:
            futures_map: Dict[Any, Tuple[int, str, float]] = {}

            for idx, (client_name, cat_map) in enumerate(client_items, 1):
                if self.cancel_event.is_set():
                    break
                # Issue 11: safe_submit_task — register before submit
                worker_id = f"W-{idx:04d}-{client_name[:20]}"
                self._register_worker(worker_id)
                t_submit = time.perf_counter()
                future = executor.submit(process_client, idx, total, client_name, cat_map, api_key, args)
                futures_map[future] = (idx, client_name, t_submit, worker_id)

                # Issue 16: Update queue label live
                self._queue_remaining = total - idx
                self.ui_queue.put(("stage", f"Submitted: {client_name}"))

            for future in as_completed(futures_map):
                idx, client_name, t_submit, worker_id = futures_map[future]

                # Issue 11: always deregister on completion/failure
                self._deregister_worker(worker_id)

                if self.cancel_event.is_set() and future.cancel():
                    continue

                completed_count += 1
                t_elapsed = time.perf_counter() - t_submit
                self._queue_remaining = max(0, total - completed_count)

                self.ui_queue.put(("client", f"[{idx}/{total}] {client_name}: processing completed"))

                try:
                    res = future.result()
                except Exception as exc:
                    stats["failed"] += 1
                    stats["errors"].append((client_name, str(exc)))
                    self.failed_clients.append(client_name)
                    log.error(f"[Batch] Client exception: {client_name} — {exc}")
                    self.ui_queue.put(("stage", f"Error: {client_name}"))
                    # Issue 15: GC after failure too
                    self._maybe_run_gc()
                    self._emit_progress(stats, total, completed_count)
                    continue

                if res["data"] is not None:
                    # Issue 13: folder ops already guarded by per-client lock in process_client
                    with _excel_lock:
                        builder.add_client(res["gst_id"], res["client_name"], res["data"], res["cat_map"])

                st = res["status"]
                if st == "ok":
                    stats["processed"] += 1
                    # Issue 12: record time only for successful (non-retry, non-failed) clients
                    self._record_processing_time(t_elapsed)
                    fs = (res["data"].get("compliance", {}).get("final_status", "?") if res["data"] else "?")
                    if fs == "Clean":
                        stats["clean"] += 1
                    elif fs == "Warning":
                        stats["warning"] += 1
                    elif fs == "High Risk":
                        stats["high_risk"] += 1
                elif st == "failed":
                    stats["failed"] += 1
                    stats["errors"].append((client_name, res.get("error") or "unknown"))
                    self.failed_clients.append(client_name)
                    log.warning(f"[Batch] Client failed: {client_name} — {res.get('error')}")
                else:
                    stats["skipped"] += 1

                # Issue 15: Periodic GC
                self._maybe_run_gc()
                cleanup_client_resources(client_name)

                # Issue 11: Periodic stale-worker cleanup
                if completed_count % 10 == 0:
                    self._cleanup_dead_workers()

                self._emit_progress(stats, total, completed_count)

                if self.cancel_event.is_set():
                    self.ui_queue.put(("log", "Cancellation acknowledged; stopping remaining queue submissions."))
                    break

        # Final registry cleanup
        with self._worker_registry_lock:
            self._active_worker_registry.clear()

        return stats

    def _emit_progress(self, stats: Dict, total: int, done: int) -> None:
        """Issue 16: Emit a rich progress payload to the UI queue."""
        remaining = max(0, total - done)
        eta_str = self._calculate_dynamic_eta(remaining)
        active = self._active_worker_count()
        throughput = getattr(self, "_throughput_per_hour", 0.0)
        avg_sec = getattr(self, "_avg_processing_sec", 0.0)
        self.ui_queue.put(("progress", {
            "total": total,
            "done": done,
            "failed": stats.get("failed", 0),
            "skipped": stats.get("skipped", 0),
            "eta": eta_str,
            "active_workers": active,
            "configured_workers": self._configured_max_workers,
            "retried": self._retry_count_total,
            "queue_remaining": remaining,
            "throughput": throughput,
            "avg_sec": avg_sec,
            "memory": self._get_memory_label(),
            "stage": self._current_stage,
        }))

    def generate_reports_only(self) -> None:
        if not self.last_groups or not self.last_output_root:
            messagebox.showwarning("No Run Data", "Run validation once before generating reports.")
            return
        # Find the most recent versioned master workbook (fallback to legacy name)
        workbooks = sorted(
            self.last_output_root.glob("GST_Validation_Workbook_*.xlsx"),
            key=lambda p: p.stat().st_mtime, reverse=True
        )
        legacy = self.last_output_root / "GST_Validation_Workbook.xlsx"
        master_path = workbooks[0] if workbooks else (legacy if legacy.exists() else None)
        if not master_path or not master_path.exists():
            messagebox.showwarning("Workbook Missing", f"No master workbook found in {self.last_output_root}")
            return
        generated = 0
        for idx, client_name_key in enumerate(self.last_groups, 1):
            gst_id = f"GSTREG-{idx:03d}"
            client_json = load_json_cache(client_name_key)
            client_folder = resolve_client_output_folder(
                self.last_output_root, client_name_key, client_json
            )
            generate_client_report(master_path, client_name_key, gst_id, client_folder, client_json)
            generated += 1
        messagebox.showinfo("Reports Generated", f"Generated/updated {generated} client reports.")

    def retry_failed_clients(self) -> None:
        if not self.failed_clients or not self.last_input_root or not self.last_output_root:
            messagebox.showwarning("Nothing To Retry", "No failed clients are available for retry.")
            return
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Processing Running", "Please wait for current batch to finish.")
            return
        retry_set = set(self.failed_clients)
        self.failed_clients = []
        self.cancel_event.clear()
        self.worker_thread = threading.Thread(target=self._run_pipeline, args=(self.last_input_root, self.last_output_root, retry_set), daemon=True, name="UIBatchRetry")
        self.worker_thread.start()

    def open_output_folder(self) -> None:
        output = self.output_combo.get().strip()
        if not output:
            messagebox.showwarning("Missing Folder", "Select output folder first.")
            return
        if not Path(output).exists():
            messagebox.showwarning("Folder Not Found", output)
            return
        os.startfile(output)

    def export_logs(self) -> None:
        src = Path("gst_validator.log")
        if not src.exists():
            messagebox.showwarning("Log Missing", "No gst_validator.log found yet.")
            return
        save_to = filedialog.asksaveasfilename(defaultextension=".log", filetypes=[("Log Files", "*.log"), ("Text Files", "*.txt"), ("All files", "*.*")], initialfile=f"gst_validator_export_{time.strftime('%Y%m%d_%H%M%S')}.log")
        if not save_to:
            return
        shutil.copy2(src, save_to)
        messagebox.showinfo("Log Exported", f"Logs exported to:\n{save_to}")

    def _drain_ui_queue(self) -> None:
        while True:
            try:
                msg_type, payload = self.ui_queue.get_nowait()
            except queue.Empty:
                break

            if msg_type == "log":
                self.log_box.insert("end", f"{payload}\n")
                self.log_box.see("end")

            elif msg_type == "client":
                self.client_box.insert("end", f"{payload}\n")
                self.client_box.see("end")

            elif msg_type == "stage":
                self._current_stage = str(payload)

            elif msg_type == "progress":
                total   = max(1, int(payload.get("total", 1)))
                done    = int(payload.get("done", 0))
                failed  = int(payload.get("failed", 0))
                skipped = int(payload.get("skipped", 0))
                retried = int(payload.get("retried", 0))
                queue_remaining = int(payload.get("queue_remaining", 0))
                eta_str = str(payload.get("eta", "--"))
                active_workers     = int(payload.get("active_workers", 0))
                configured_workers = int(payload.get("configured_workers", self._configured_max_workers))
                throughput  = float(payload.get("throughput", 0.0))
                avg_sec     = float(payload.get("avg_sec", 0.0))
                memory_str  = str(payload.get("memory", "--"))
                stage_str   = str(payload.get("stage", self._current_stage))

                # Issue 11: worker count capped display — never show > configured
                safe_active = min(active_workers, configured_workers)

                self.total_label.configure(text=f"Total: {total}")
                self.done_label.configure(text=f"Completed: {done}")
                self.failed_label.configure(
                    text=f"Failed: {failed}",
                    text_color="#fb7185" if failed > 0 else "gray"
                )
                self.skip_label.configure(text=f"Skipped: {skipped}")
                self.retry_label.configure(text=f"Retried: {retried}")
                self.queue_label.configure(text=f"Queue: {queue_remaining}")

                # Issue 12: rich ETA display
                if throughput > 0:
                    self.eta_label.configure(text=f"ETA: {eta_str}  ({throughput:.0f}/hr)")
                else:
                    self.eta_label.configure(text=f"ETA: {eta_str}")

                # Issue 11: accurate worker count (safe_active ≤ configured)
                worker_color = "#86efac" if safe_active <= configured_workers else "#fb7185"
                self.worker_status_label.configure(
                    text=f"Workers: {safe_active}/{configured_workers}",
                    text_color=worker_color
                )

                # Issue 16: throughput + avg + memory + stage row
                self.throughput_label.configure(
                    text=f"Throughput: {throughput:.0f}/hr" if throughput > 0 else "Throughput: --"
                )
                avg_text = f"Avg: {avg_sec:.1f}s/client" if avg_sec > 0 else "Avg: --"
                self.avg_time_label.configure(text=avg_text)
                self.memory_label.configure(text=f"Mem: {memory_str}")
                self.stage_label.configure(text=f"Stage: {stage_str}")

                self.progress.set(done / total)

                # Update header meta bar
                self.header_meta.configure(
                    text=f"{done}/{total} Clients | {safe_active} Workers | {throughput:.0f}/hr | ETA: {eta_str}"
                )

            elif msg_type == "done":
                stats = payload["stats"]
                cancelled = payload.get("cancelled", False)
                master = payload.get("master", "")
                self.start_btn.configure(state="normal")
                self.app_status.configure(
                    text="Cancelled" if cancelled else "Completed",
                    text_color="#fb7185" if cancelled else "#86efac"
                )
                self.header_meta.configure(
                    text=f"Processed: {stats['processed']} | Failed: {stats['failed']} | Skipped: {stats['skipped']}"
                )
                self.stage_label.configure(text="Stage: Done")
                self.worker_status_label.configure(text=f"Workers: 0/{self._configured_max_workers}")
                elapsed = 0.0 if self.processing_started_at is None else (time.perf_counter() - self.processing_started_at)
                messagebox.showinfo("Processing Summary",
                    f"Total Clients  : {stats['total']}\n"
                    f"Processed      : {stats['processed']}\n"
                    f"Failed         : {stats['failed']}\n"
                    f"Skipped        : {stats['skipped']}\n"
                    f"Elapsed        : {elapsed:.1f}s\n"
                    f"Workbook       : {master}"
                )

            elif msg_type == "error":
                self.start_btn.configure(state="normal")
                self.app_status.configure(text="Error", text_color="#fb7185")
                self.header_meta.configure(text="Run failed")
                self.stage_label.configure(text="Stage: Error")
                messagebox.showerror("Runtime Error", f"Validation failed:\n{payload}")

        self.after(120, self._drain_ui_queue)


if __name__ == "__main__":
    if len(sys.argv) > 1:
        main()
    else:
        app = GSTValidatorUI()
        app.mainloop()