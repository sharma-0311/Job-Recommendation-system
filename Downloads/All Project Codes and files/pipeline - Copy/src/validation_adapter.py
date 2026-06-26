import os
import re
import sys
import subprocess
import threading
from pathlib import Path
from typing import Callable, Optional

class ValidationAdapter:
    """
    Executes gst_validator_v2.py as a clean background subprocess.
    Streams output in real-time and parses progress metrics using regular expressions.
    """
    def __init__(
        self,
        input_folder: Path,
        output_folder: Path,
        api_key: str,
        workers: int = 4,
        use_cache: bool = True,
        skip_gemini: bool = False
    ):
        self.input_folder = input_folder
        self.output_folder = output_folder
        self.api_key = api_key
        self.workers = workers
        self.use_cache = use_cache
        self.skip_gemini = skip_gemini
        
        self._process: Optional[subprocess.Popen] = None
        self._thread: Optional[threading.Thread] = None
        self._cancel_event = threading.Event()

    def start(
        self,
        log_callback: Callable[[str], None],
        progress_callback: Callable[[int, int, str], None],
        completion_callback: Callable[[bool, Optional[str], Optional[Path]], None]
    ) -> None:
        """Starts the validation subprocess inside a background thread."""
        self._cancel_event.clear()
        self._thread = threading.Thread(
            target=self._run_process,
            args=(log_callback, progress_callback, completion_callback),
            daemon=True,
            name="ValidatorAdapterThread"
        )
        self._thread.start()

    def stop(self) -> None:
        """Cancels and terminates the validation subprocess safely."""
        self._cancel_event.set()
        if self._process:
            try:
                # Windows taskkill bypasses process group issues
                if sys.platform == "win32":
                    subprocess.run(
                        ["taskkill", "/F", "/T", "/PID", str(self._process.pid)],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL
                    )
                else:
                    self._process.terminate()
            except Exception:
                pass

    def _run_process(
        self,
        log_callback: Callable[[str], None],
        progress_callback: Callable[[int, int, str], None],
        completion_callback: Callable[[bool, Optional[str], Optional[Path]], None]
    ) -> None:
        # Determine the Python interpreter executable and target command under PyInstaller
        is_frozen = getattr(sys, 'frozen', False)
        if is_frozen:
            python_exe = sys.executable
            cmd = [
                python_exe,
                "--folder", str(self.input_folder),
                "--client-output-dir", str(self.output_folder),
                "--workers", str(self.workers)
            ]
        else:
            python_exe = sys.executable
            gst_script_path = "gst_validator_v2.py"
            cmd = [
                python_exe,
                gst_script_path,
                "--folder", str(self.input_folder),
                "--client-output-dir", str(self.output_folder),
                "--workers", str(self.workers)
            ]
        
        if self.api_key:
            cmd.extend(["--api-key", self.api_key])
        if self.skip_gemini:
            cmd.append("--skip-gemini")
        if self.use_cache:
            cmd.append("--use-cache")
        else:
            cmd.append("--no-cache")

        # Regex patterns to parse progress updates
        progress_pattern_1 = re.compile(r"Progress:\s*(\d+)/(\d+)\s*(?:\|\s*Elapsed:.*?ETA:\s*(\d+s))?")
        progress_pattern_2 = re.compile(r"(\d+)/(\d+)\s+\[.*?ETA:\s*(\w+)\]")

        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"

        try:
            self._process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
                bufsize=1,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            )

            detected_master_path: Optional[Path] = None
            total_clients = 0

            while True:
                if self._cancel_event.is_set():
                    break
                    
                line = self._process.stdout.readline()
                if not line:
                    break
                    
                # Clean and send log line
                clean_line = line.rstrip()
                log_callback(clean_line)

                # Search for master Excel path saved message
                if "Workbook saved →" in clean_line:
                    try:
                        p_str = clean_line.split("→")[-1].strip()
                        detected_master_path = Path(p_str)
                    except Exception:
                        pass
                elif "Output (global)   :" in clean_line:
                    try:
                        p_str = clean_line.split(":")[-1].strip()
                        detected_master_path = Path(p_str)
                    except Exception:
                        pass
                
                # Check for client grouping counts
                if "CLIENTS TO PROCESS:" in clean_line:
                    try:
                        total_clients = int(clean_line.split(":")[-1].strip())
                    except Exception:
                        pass

                # Parse progress markers
                m1 = progress_pattern_1.search(clean_line)
                if m1:
                    done = int(m1.group(1))
                    total = int(m1.group(2))
                    eta = m1.group(3) or "--"
                    progress_callback(done, total, eta)
                    continue

                m2 = progress_pattern_2.search(clean_line)
                if m2:
                    done = int(m2.group(1))
                    total = int(m2.group(2))
                    eta = m2.group(3) or "--"
                    progress_callback(done, total, eta)
                    continue

            self._process.wait()
            success = (self._process.returncode == 0) and not self._cancel_event.is_set()
            
            # Fallback path scan if not detected in logs
            if not detected_master_path and success:
                # Check for newest master workbook in client output base
                internal_dir = self.output_folder / "_internal"
                workbooks = sorted(
                    self.output_folder.glob("GST_Validation_Workbook_*.xlsx"),
                    key=lambda p: p.stat().st_mtime,
                    reverse=True
                )
                if workbooks:
                    detected_master_path = workbooks[0]
                else:
                    legacy = self.output_folder / "GST_Validation_Workbook.xlsx"
                    if legacy.exists():
                        detected_master_path = legacy

            completion_callback(success, None if success else f"Exit code {self._process.returncode}", detected_master_path)

        except Exception as e:
            completion_callback(False, str(e), None)
        finally:
            if self._process:
                try:
                    self._process.stdout.close()
                except Exception:
                    pass
                self._process = None
