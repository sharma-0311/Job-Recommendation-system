import os
import re
import sys
import subprocess
import threading
from pathlib import Path
from typing import Callable, Optional

class RegistrationAdapter:
    """
    Executes amend_req.py as a background subprocess.
    Streams console output in real-time and parses dynamic GST Portal execution markers.
    """
    def __init__(
        self,
        excel_path: Path,
        checkpoint_path: Path,
        username: str,
        password: str,
        timeout: int = 15,
        headless: bool = False,
        trocr_model: Optional[str] = None,
        workers: int = 1,
        engine: str = "selenium"
    ):
        self.excel_path = excel_path
        self.checkpoint_path = checkpoint_path
        self.username = username
        self.password = password
        self.timeout = timeout
        self.headless = headless
        self.trocr_model = trocr_model
        self.workers = workers
        self.engine = engine
        
        self._process: Optional[subprocess.Popen] = None
        self._thread: Optional[threading.Thread] = None
        self._cancel_event = threading.Event()

    def start(
        self,
        log_callback: Callable[[str], None],
        status_callback: Callable[[str, str, int, int], None], # (current_gstin, current_stage, done_count, failed_count)
        completion_callback: Callable[[bool, Optional[str]], None],
        captcha_callback: Callable[[str, str], None]
    ) -> None:
        """Starts the Selenium bulk APoB runner subprocess."""
        self._cancel_event.clear()
        self._thread = threading.Thread(
            target=self._run_process,
            args=(log_callback, status_callback, completion_callback, captcha_callback),
            daemon=True,
            name="RegistrationAdapterThread"
        )
        self._thread.start()

    def submit_captcha(self, value: str) -> None:
        """Writes the manual CAPTCHA solution to the subprocess stdin."""
        if self._process and self._process.stdin:
            try:
                self._process.stdin.write(f"{value}\n")
                self._process.stdin.flush()
            except Exception:
                pass

    def stop(self) -> None:
        """Kills the active Selenium process chain."""
        self._cancel_event.set()
        if self._process:
            try:
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
        status_callback: Callable[[str, str, int, int], None],
        completion_callback: Callable[[bool, Optional[str]], None],
        captcha_callback: Callable[[str], None]
    ) -> None:
        # Determine the Python interpreter executable and target command under PyInstaller
        is_frozen = getattr(sys, 'frozen', False)
        if is_frozen:
            python_exe = sys.executable
            cmd = [
                python_exe,
                "--excel", str(self.excel_path),
                "--checkpoint", str(self.checkpoint_path),
                "--timeout", str(self.timeout),
                "--no-wait"
            ]
        else:
            python_exe = sys.executable
            amend_script_path = "amend_req.py"
            cmd = [
                python_exe,
                amend_script_path,
                "--excel", str(self.excel_path),
                "--checkpoint", str(self.checkpoint_path),
                "--timeout", str(self.timeout),
                "--no-wait"
            ]
        
        if self.username:
            cmd.extend(["--username", self.username])
        if self.password:
            cmd.extend(["--password", self.password])
        if self.headless:
            cmd.append("--headless")
        if self.trocr_model:
            cmd.extend(["--model", self.trocr_model])
        if self.workers > 1:
            cmd.extend(["--workers", str(self.workers)])
        if self.engine:
            cmd.extend(["--engine", self.engine])

        env = os.environ.copy()
        env["PYTHONUNBUFFERED"] = "1"
        # Force Python's stdout/stderr standard streams to UTF-8
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUTF8"] = "1"

        try:
            self._process = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                env=env,
                bufsize=1,
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
            )

            current_gstin = "Initializing"
            current_stage = "Logging In"
            done_count = 0
            failed_count = 0
            self._last_captcha_path = None

            # Regex parser keys
            user_session_pattern = re.compile(r"\[USER SESSION: (.*?)\]")
            gstin_group_pattern = re.compile(r"\[GSTIN: (.*?)\]")
            progress_pattern = re.compile(r"Registration (\d+) of (\d+)")
            success_pattern = re.compile(r"successfully marked completed|successfully finalized and completed")
            failed_pattern = re.compile(r"GSTIN Processing Error|business_fail|failed_clients")

            while True:
                if self._cancel_event.is_set():
                    break
                    
                line = self._process.stdout.readline()
                if not line:
                    break
                    
                clean_line = line.rstrip()
                log_callback(clean_line)

                # Dynamically parse current stage and progress
                m_session = user_session_pattern.search(clean_line)
                if m_session:
                    current_stage = f"Active Session: {m_session.group(1)}"
                    status_callback(current_gstin, current_stage, done_count, failed_count)
                    continue

                m_gstin = gstin_group_pattern.search(clean_line)
                if m_gstin:
                    current_gstin = m_gstin.group(1)
                    current_stage = "Processing APoB"
                    status_callback(current_gstin, current_stage, done_count, failed_count)
                    continue

                m_prog = progress_pattern.search(clean_line)
                if m_prog:
                    current_stage = f"APoB record {m_prog.group(1)}/{m_prog.group(2)}"
                    status_callback(current_gstin, current_stage, done_count, failed_count)
                    continue

                if success_pattern.search(clean_line):
                    done_count += 1
                    current_stage = "APoB Row Success"
                    status_callback(current_gstin, current_stage, done_count, failed_count)
                    continue

                if failed_pattern.search(clean_line):
                    failed_count += 1
                    current_stage = "APoB Row Failed"
                    status_callback(current_gstin, current_stage, done_count, failed_count)
                    continue

                # Capture last captcha image path and hash
                if "[CAPTCHA] Saved -> " in clean_line:
                    self._last_captcha_path = clean_line.split("[CAPTCHA] Saved -> ")[-1].strip()
                
                if "[CAPTCHA] Portal Captcha Hash: " in clean_line:
                    self._portal_captcha_hash = clean_line.split("[CAPTCHA] Portal Captcha Hash: ")[-1].strip()

                # Trigger manual captcha callback if auto-solve failed
                if "Auto-solve failed. Switched to user manual" in clean_line:
                    if self._last_captcha_path and os.path.exists(self._last_captcha_path):
                        portal_h = getattr(self, "_portal_captcha_hash", "N/A")
                        captcha_callback(self._last_captcha_path, portal_h)

                # Auto-detect Captcha OCR solves
                if "[CAPTCHA] OCR result:" in clean_line:
                    current_stage = "CAPTCHA Solving"
                    status_callback(current_gstin, current_stage, done_count, failed_count)
                    continue
                
                # Check for profile details fetch
                if "Auto-fetching contact details" in clean_line:
                    current_stage = "Fetching Contact Details"
                    status_callback(current_gstin, current_stage, done_count, failed_count)
                    continue

            self._process.wait()
            success = (self._process.returncode == 0) and not self._cancel_event.is_set()
            completion_callback(success, None if success else f"Exit code {self._process.returncode}")

        except Exception as e:
            completion_callback(False, str(e))
        finally:
            if self._process:
                try:
                    self._process.stdout.close()
                except Exception:
                    pass
                self._process = None
