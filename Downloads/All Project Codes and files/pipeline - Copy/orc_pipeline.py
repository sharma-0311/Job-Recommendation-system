import os
import re
import sys

# Globally configure standard streams to UTF-8 on Windows to prevent UnicodeEncodeError
for stream in (sys.stdout, sys.stderr):
    if stream is not None and hasattr(stream, "reconfigure"):
        try:
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
        except Exception:
            pass

# ── PyInstaller Subprocess Dependencies Verification Layer ────────────────────
# These module-level imports ensure PyInstaller packages the dependencies
# required by the background subprocesses (amend_req.py & gst_validator_v2.py).
try:
    import webdriver_manager
    import webdriver_manager.chrome
    import selenium
    import selenium.webdriver
    import openpyxl
    import pandas
    import cv2
    import numpy
    import PIL
    import torch
    import transformers
    import tqdm
    import google.genai
    import google.generativeai
    import urllib3
except ImportError:
    pass

import time
import json
import threading
import queue
import webbrowser
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional

import openpyxl
import customtkinter as ctk
from tkinter import filedialog, messagebox

# Import orchestration modules
from src.security import SecureVault
from src.checkpoint_manager import CheckpointManager
from src.batch_builder import BatchBuilder
from src.validation_adapter import ValidationAdapter
from src.registration_adapter import RegistrationAdapter
from src.reporting_engine import ReportingEngine

class GSTSmartHubApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Color Token System (Premium SaaS Palette - Dual Theme)
        self.C_MAIN_BG = ("#F8FAFC", "#020617")        # Slate 50 / Slate 955
        self.C_SIDEBAR_BG = ("#F1F5F9", "#0B0F19")     # Slate 100 / Slate 900
        self.C_CARD_BG = ("#FFFFFF", "#151F32")        # White / Slate 800 Card
        self.C_CARD_DARK = ("#F8FAFC", "#0F172A")      # Slate 50 / Slate 900 alternative
        self.C_BORDER = ("#E2E8F0", "#1E293B")         # Slate 200 / Slate 800 subtle borders
        self.C_TEXT_PRIMARY = ("#0F172A", "#F8FAFC")   # Slate 900 / Slate 50
        self.C_TEXT_SECONDARY = ("#475569", "#94A3B8") # Slate 600 / Slate 400
        self.C_TEXT_MUTED = ("#94A3B8", "#475569")     # Slate 400 / Slate 600
        
        self.C_ACCENT = ("#1D4ED8", "#60A5FA")         # Deep Blue / Light Blue
        self.C_ACCENT_HOVER = ("#1E40AF", "#3B82F6")   # Even Deeper / Medium Blue
        self.C_SUCCESS = ("#047857", "#34D399")        # Deep Emerald / Light Emerald
        self.C_SUCCESS_BG = ("#D1FAE5", "#064E3B")     # Light Emerald / Dark Emerald
        self.C_DANGER = ("#B91C1C", "#F87171")         # Deep Red / Light Red
        self.C_DANGER_BG = ("#FEE2E2", "#7F1D1D")       # Light Red / Dark Red
        self.C_WARNING = ("#B45309", "#FBBF24")        # Deep Amber / Light Amber
        self.C_WARNING_BG = ("#FEF3C7", "#78350F")      # Light Amber / Dark Amber
        self.C_WARNING_HOVER = ("#92400E", "#D97706")   # Amber hover
        self.validation_failed_count = 0

        # Style Definitions
        ctk.set_appearance_mode("light") # Defaulting to light theme
        ctk.set_default_color_theme("blue")
        
        self.title("GST Smart Automation Hub")
        self.geometry("1400x900")
        self.minsize(1180, 780)
        self.configure(fg_color=self.C_MAIN_BG)

        # Runtime States
        self.input_folder_path: Optional[Path] = None
        self.output_folder_path: Optional[Path] = None
        self.master_workbook_path: Optional[Path] = None
        self.active_batch_id: Optional[str] = None
        
        self.selected_records: List[Dict[str, Any]] = []
        self.grid_rows: List[Dict[str, Any]] = [] # Grid row widget map
        self.execution_durations: Dict[str, float] = {} # Record tracking for final report

        # Thread Safety Communication Queue
        self.ui_queue: queue.Queue[tuple[str, Any]] = queue.Queue()
        
        # Adapters references
        self.validator_adapter: Optional[ValidationAdapter] = None
        self.registration_adapter: Optional[RegistrationAdapter] = None
        self.checkpoint_mgr: Optional[CheckpointManager] = None

        # Build UI layout
        self._build_sidebar()
        self._build_content_area()
        
        # Dynamic telemetry updates loop
        self.after(100, self._drain_ui_queue)
        
        # Setup Drag and Drop
        self._setup_drag_and_drop()

    def _build_sidebar(self) -> None:
        """Constructs the left-side fixed navigation sidebar console."""
        self.sidebar = ctk.CTkFrame(self, width=280, corner_radius=0, fg_color=self.C_SIDEBAR_BG, border_width=0)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        # Dynamic Logo Loader
        logo_path = Path(__file__).parent / "logo.png"
        logo_loaded = False
        if logo_path.exists():
            try:
                from PIL import Image
                pil_img = Image.open(logo_path)
                w, h = pil_img.size
                target_w = 200
                target_h = int((target_w / w) * h)
                
                logo_img = ctk.CTkImage(
                    light_image=pil_img,
                    dark_image=pil_img,
                    size=(target_w, target_h)
                )
                self.logo_lbl = ctk.CTkLabel(self.sidebar, image=logo_img, text="")
                self.logo_lbl.pack(pady=(24, 2), padx=28, anchor="w")
                logo_loaded = True
            except Exception:
                pass

        if not logo_loaded:
            # Fallback text logo
            ctk.CTkLabel(
                self.sidebar,
                text="GST SMART HUB",
                font=("Segoe UI", 22, "bold"),
                text_color="#60A5FA"
            ).pack(pady=(32, 2), padx=28, anchor="w")

        ctk.CTkLabel(
            self.sidebar,
            text="Enterprise Orchestrator",
            font=("Segoe UI", 11, "italic"),
            text_color=self.C_TEXT_MUTED
        ).pack(pady=(0, 24), padx=28, anchor="w")

        # 9 Tab Navigation Setup
        self.nav_items: Dict[str, Dict[str, Any]] = {}
        tabs = [
            ("dashboard", "📊 Dashboard"),
            ("upload", "📂 Upload & Validate"),
            ("results", "📝 Review & Select"),
            ("credentials", "🔑 Credentials"),
            ("processing", "⚙️ Processing"),
            ("reports", "📈 Reports"),
            ("logs", "📜 Live Logs"),
            ("settings", "🔧 Settings"),
            ("about", "ℹ️ About")
        ]
        
        for name, text in tabs:
            row_frm = ctk.CTkFrame(self.sidebar, fg_color="transparent")
            row_frm.pack(fill="x", padx=16, pady=3)
            
            # Active indicator colored strip
            indicator = ctk.CTkFrame(row_frm, width=4, height=36, corner_radius=2, fg_color="transparent")
            indicator.pack(side="left", fill="y", padx=(0, 8))
            
            btn = ctk.CTkButton(
                row_frm,
                text=text,
                anchor="w",
                height=36,
                corner_radius=6,
                font=("Segoe UI", 12, "bold"),
                fg_color="transparent",
                text_color=self.C_TEXT_SECONDARY,
                hover_color=self.C_CARD_BG,
                command=lambda n=name: self._switch_screen(n)
            )
            btn.pack(side="left", fill="x", expand=True)
            
            self.nav_items[name] = {"indicator": indicator, "button": btn}

        # Separator Line
        ctk.CTkFrame(self.sidebar, height=1, fg_color=self.C_BORDER).pack(fill="x", padx=20, pady=20)

        # Active Monitor Status Panel
        ctk.CTkLabel(
            self.sidebar,
            text="SYSTEM COMPLIANCE STATE",
            font=("Segoe UI", 10, "bold"),
            text_color=self.C_TEXT_MUTED
        ).pack(padx=28, anchor="w", pady=(4, 4))

        self.lbl_status_indicator = ctk.CTkLabel(
            self.sidebar,
            text="● Hub Standby",
            font=("Segoe UI", 13, "bold"),
            text_color=self.C_SUCCESS
        )
        self.lbl_status_indicator.pack(padx=28, anchor="w")

        # Theme Selector Segmented Button in sidebar footer
        ctk.CTkLabel(
            self.sidebar,
            text="THEME MODE",
            font=("Segoe UI", 9, "bold"),
            text_color=self.C_TEXT_MUTED
        ).pack(side="bottom", padx=20, pady=(12, 2), anchor="w")
        
        self.seg_theme = ctk.CTkSegmentedButton(
            self.sidebar,
            values=["Light", "Dark", "System"],
            command=self._on_theme_changed,
            height=28,
            font=("Segoe UI", 11, "bold")
        )
        self.seg_theme.pack(side="bottom", fill="x", padx=20, pady=(2, 10))
        self.seg_theme.set("Light")

        # Browse Output folders button in sidebar
        self.btn_open_outputs = ctk.CTkButton(
            self.sidebar,
            text="Browse Output Folders",
            fg_color="transparent",
            border_width=1,
            border_color=self.C_BORDER,
            text_color=self.C_TEXT_SECONDARY,
            hover_color=self.C_CARD_BG,
            height=34,
            corner_radius=6,
            command=self._open_output_dir
        )
        self.btn_open_outputs.pack(side="bottom", fill="x", padx=20, pady=28)

    def _build_content_area(self) -> None:
        """Builds the main container with a stepper header and the screen cards below it."""
        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.pack(side="right", fill="both", expand=True, padx=40, pady=28)

        # 1. Header and Workflow Stepper
        self.header_area = ctk.CTkFrame(self.content, fg_color="transparent")
        self.header_area.pack(fill="x", pady=(0, 16))
        self._build_stepper()

        # Divider
        ctk.CTkFrame(self.content, height=1, fg_color=self.C_BORDER).pack(fill="x", pady=(0, 24))

        # 2. Main Scrollable/Switchable Screen container
        self.screen_container = ctk.CTkFrame(self.content, fg_color="transparent")
        self.screen_container.pack(fill="both", expand=True)

        # Register Screens Map
        self.screens: Dict[str, ctk.CTkFrame] = {}

        # Build each screen pane
        self._build_dashboard_screen()
        self._build_upload_screen()
        self._build_results_screen()
        self._build_credentials_screen()
        self._build_processing_screen()
        self._build_reports_screen()
        self._build_logs_screen()
        self._build_settings_screen()
        self._build_about_screen()

        # Switch to first tab initially
        self._switch_screen("dashboard")

    def _build_stepper(self) -> None:
        """Constructs a modern horizontal workflow stepper with connection indicators."""
        self.stepper_frame = ctk.CTkFrame(
            self.header_area, 
            fg_color=self.C_SIDEBAR_BG, 
            height=82, 
            corner_radius=10, 
            border_width=1, 
            border_color=self.C_BORDER
        )
        self.stepper_frame.pack(fill="x", ipady=6)
        self.stepper_frame.pack_propagate(False)

        inner_stepper = ctk.CTkFrame(self.stepper_frame, fg_color="transparent")
        inner_stepper.pack(expand=True, fill="both", padx=32)

        # Stepper Columns (Grid configuration)
        for col_idx in range(9):
            if col_idx % 2 == 0:
                inner_stepper.grid_columnconfigure(col_idx, weight=0)
            else:
                inner_stepper.grid_columnconfigure(col_idx, weight=1)

        self.steps_info = [
            ("Upload & Validate", "1"),
            ("Review & Select", "2"),
            ("Credentials", "3"),
            ("Processing", "4"),
            ("Completed", "5")
        ]

        self.step_widgets = []
        self.step_connectors = []

        for idx, (label_text, num) in enumerate(self.steps_info):
            col = idx * 2
            
            # Step frame
            step_unit = ctk.CTkFrame(inner_stepper, fg_color="transparent")
            step_unit.grid(row=0, column=col, padx=8, pady=4)
            
            circle = ctk.CTkFrame(step_unit, width=32, height=32, corner_radius=16, border_width=2, border_color=self.C_BORDER, fg_color=self.C_MAIN_BG)
            circle.pack()
            circle.pack_propagate(False)
            
            num_lbl = ctk.CTkLabel(circle, text=num, font=("Segoe UI", 12, "bold"), text_color=self.C_TEXT_SECONDARY)
            num_lbl.pack(expand=True)
            
            lbl = ctk.CTkLabel(step_unit, text=label_text, font=("Segoe UI", 11, "bold"), text_color=self.C_TEXT_MUTED)
            lbl.pack(pady=(4, 0))
            
            self.step_widgets.append({
                "circle": circle,
                "label": lbl,
                "num_lbl": num_lbl
            })
            
            # Connection connector
            if idx < 4:
                conn_col = col + 1
                connector = ctk.CTkFrame(inner_stepper, height=2, fg_color=self.C_BORDER)
                connector.grid(row=0, column=conn_col, sticky="ew", pady=(0, 24))
                self.step_connectors.append(connector)

    def _update_stepper(self, active_step: int) -> None:
        """Updates the visual stepper highlight (1-5 step logic)."""
        for idx, w in enumerate(self.step_widgets):
            step_num = idx + 1
            if step_num < active_step:
                w["circle"].configure(border_color=self.C_SUCCESS, fg_color=self.C_SUCCESS_BG)
                w["num_lbl"].configure(text="✓", text_color=self.C_SUCCESS)
                w["label"].configure(text_color=self.C_TEXT_PRIMARY)
            elif step_num == active_step:
                w["circle"].configure(border_color=self.C_ACCENT, fg_color=self.C_MAIN_BG)
                w["num_lbl"].configure(text=str(step_num), text_color=self.C_ACCENT)
                w["label"].configure(text_color=self.C_TEXT_PRIMARY)
            else:
                w["circle"].configure(border_color=self.C_BORDER, fg_color=self.C_MAIN_BG)
                w["num_lbl"].configure(text=str(step_num), text_color=self.C_TEXT_MUTED)
                w["label"].configure(text_color=self.C_TEXT_MUTED)

        for idx, conn in enumerate(self.step_connectors):
            if idx + 1 < active_step:
                conn.configure(fg_color=self.C_SUCCESS)
            elif idx + 1 == active_step:
                conn.configure(fg_color=self.C_ACCENT)
            else:
                conn.configure(fg_color=self.C_BORDER)

    def _switch_screen(self, name: str) -> None:
        """Packs and shows the selected screen frame, updating sidebar buttons and syncs stepper."""
        for scr_name, scr_frame in self.screens.items():
            scr_frame.pack_forget()

        self.screens[name].pack(fill="both", expand=True)

        # Highlight active sidebar element
        for n, item in self.nav_items.items():
            if n == name:
                item["indicator"].configure(fg_color=self.C_ACCENT)
                item["button"].configure(fg_color=self.C_BORDER, text_color=self.C_ACCENT)
            else:
                item["indicator"].configure(fg_color="transparent")
                item["button"].configure(fg_color="transparent", text_color=self.C_TEXT_SECONDARY)

        # Workflow stepper sync
        step_mapping = {
            "upload": 1,
            "results": 2,
            "credentials": 3,
            "processing": 4,
            "reports": 5
        }
        if name in step_mapping:
            self._update_stepper(step_mapping[name])

    # ── SCREEN 1: EXECUTIVE DASHBOARD ────────────────────────────────────────
    def _build_dashboard_screen(self) -> None:
        scr = ctk.CTkFrame(self.screen_container, fg_color="transparent")
        self.screens["dashboard"] = scr

        # Header Info Card
        header_card = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=10)
        header_card.pack(fill="x", pady=(0, 20), ipady=12, ipadx=20)
        
        ctk.CTkLabel(
            header_card,
            text="Smart Automation Hub Workspace",
            font=("Segoe UI", 24, "bold"),
            text_color=self.C_TEXT_PRIMARY
        ).pack(anchor="w")
        
        ctk.CTkLabel(
            header_card,
            text="Enterprise-grade GST compliance & Place of Business (APoB) registration automation orchestrator.",
            font=("Segoe UI", 12),
            text_color=self.C_TEXT_SECONDARY
        ).pack(anchor="w", pady=(4, 0))

        # 5-Column Dynamic Telemetry Stats Grid
        stats_frame = ctk.CTkFrame(scr, fg_color="transparent")
        stats_frame.pack(fill="x", pady=(0, 20))
        for col in range(5):
            stats_frame.grid_columnconfigure(col, weight=1)

        # Metric 1: Total Batch Card
        self.card_total = ctk.CTkFrame(stats_frame, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, height=100, corner_radius=8)
        self.card_total.grid(row=0, column=0, padx=(0, 8), sticky="nsew")
        self.card_total.pack_propagate(False)
        self.val_stat_total = ctk.CTkLabel(self.card_total, text="0", font=("Segoe UI", 32, "bold"), text_color=self.C_TEXT_PRIMARY)
        self.val_stat_total.pack(pady=(16, 2))
        ctk.CTkLabel(self.card_total, text="TOTAL BATCHED", font=("Segoe UI", 10, "bold"), text_color=self.C_TEXT_MUTED).pack()

        # Metric 2: Pending Card
        self.card_pending = ctk.CTkFrame(stats_frame, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, height=100, corner_radius=8)
        self.card_pending.grid(row=0, column=1, padx=8, sticky="nsew")
        self.card_pending.pack_propagate(False)
        self.val_stat_pending = ctk.CTkLabel(self.card_pending, text="0", font=("Segoe UI", 32, "bold"), text_color=self.C_TEXT_SECONDARY)
        self.val_stat_pending.pack(pady=(16, 2))
        ctk.CTkLabel(self.card_pending, text="PENDING QUEUE", font=("Segoe UI", 10, "bold"), text_color=self.C_TEXT_MUTED).pack()

        # Metric 3: Success Card
        self.card_completed = ctk.CTkFrame(stats_frame, fg_color=self.C_SUCCESS_BG, border_width=1, border_color=self.C_SUCCESS, height=100, corner_radius=8)
        self.card_completed.grid(row=0, column=2, padx=8, sticky="nsew")
        self.card_completed.pack_propagate(False)
        self.val_stat_completed = ctk.CTkLabel(self.card_completed, text="0", font=("Segoe UI", 32, "bold"), text_color="#10B981")
        self.val_stat_completed.pack(pady=(16, 2))
        ctk.CTkLabel(self.card_completed, text="SUCCESS DRAFTS", font=("Segoe UI", 10, "bold"), text_color="#34D399").pack()

        # Metric 4: Failed Card
        self.card_failed = ctk.CTkFrame(stats_frame, fg_color=self.C_DANGER_BG, border_width=1, border_color=self.C_DANGER, height=100, corner_radius=8)
        self.card_failed.grid(row=0, column=3, padx=8, sticky="nsew")
        self.card_failed.pack_propagate(False)
        self.val_stat_failed = ctk.CTkLabel(self.card_failed, text="0", font=("Segoe UI", 32, "bold"), text_color="#F87171")
        self.val_stat_failed.pack(pady=(16, 2))
        ctk.CTkLabel(self.card_failed, text="FAILED RUNS", font=("Segoe UI", 10, "bold"), text_color="#FCA5A5").pack()

        # Metric 5: ETA Card
        self.card_eta = ctk.CTkFrame(stats_frame, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, height=100, corner_radius=8)
        self.card_eta.grid(row=0, column=4, padx=(8, 0), sticky="nsew")
        self.card_eta.pack_propagate(False)
        self.val_stat_eta = ctk.CTkLabel(self.card_eta, text="--", font=("Segoe UI", 30, "bold"), text_color=self.C_WARNING)
        self.val_stat_eta.pack(pady=(18, 2))
        ctk.CTkLabel(self.card_eta, text="ESTIMATED ETA", font=("Segoe UI", 10, "bold"), text_color=self.C_TEXT_MUTED).pack()

        # Active System Telemetry Overview Card
        telemetry_card = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=10)
        telemetry_card.pack(fill="both", expand=True, pady=10, ipady=16, ipadx=24)

        ctk.CTkLabel(
            telemetry_card,
            text="LIVE ENGINE TELEMETRY",
            font=("Segoe UI", 12, "bold"),
            text_color="#60A5FA"
        ).pack(anchor="w", padx=16, pady=(10, 16))

        # Client Active Information
        self.lbl_tel_client = ctk.CTkLabel(
            telemetry_card,
            text="Current Active Record: Hub Standby",
            font=("Segoe UI", 15, "bold"),
            text_color=self.C_TEXT_PRIMARY
        )
        self.lbl_tel_client.pack(anchor="w", padx=16, pady=4)

        self.lbl_tel_stage = ctk.CTkLabel(
            telemetry_card,
            text="Current Pipeline Stage: Idle",
            font=("Segoe UI", 13),
            text_color=self.C_WARNING
        )
        self.lbl_tel_stage.pack(anchor="w", padx=16, pady=4)

        # Operational Instructions / Guidelines
        ctk.CTkFrame(telemetry_card, height=1, fg_color=self.C_BORDER).pack(fill="x", padx=16, pady=16)

        ctk.CTkLabel(
            telemetry_card,
            text="Quick Workflow Checklist:",
            font=("Segoe UI", 12, "bold"),
            text_color=self.C_TEXT_SECONDARY
        ).pack(anchor="w", padx=16, pady=2)

        checklist_items = [
            "✔ 1. Document Auditing: Scan, cluster files, and perform Gemini validations on client root directory.",
            "✔ 2. Interactive Results: Review parsed addresses, edit GSTIN keys, verify status validations.",
            "✔ 3. Credentials & settings: Supply secure authentication tokens and configure webdriver parameters.",
            "✔ 4. Bulk Processing: Run selenium orchestration to verify place of business amendments."
        ]
        for item in checklist_items:
            ctk.CTkLabel(
                telemetry_card,
                text=item,
                font=("Segoe UI", 11),
                text_color=self.C_TEXT_SECONDARY
            ).pack(anchor="w", padx=32, pady=2)

    # ── SCREEN 2: UPLOAD & VALIDATION WORKSPACE ──────────────────────────────
    def _build_upload_screen(self) -> None:
        scr = ctk.CTkFrame(self.screen_container, fg_color="transparent")
        self.screens["upload"] = scr

        # Header Info Card
        header = ctk.CTkFrame(scr, fg_color="transparent")
        header.pack(fill="x", pady=(0, 16))
        ctk.CTkLabel(
            header,
            text="Client Documents Validation Workspace",
            font=("Segoe UI", 24, "bold"),
            text_color=self.C_TEXT_PRIMARY
        ).pack(anchor="w")
        ctk.CTkLabel(
            header,
            text="Perform Gemini-based document validation, category clustering, and compliance auditing.",
            font=("Segoe UI", 12),
            text_color=self.C_TEXT_SECONDARY
        ).pack(anchor="w", pady=(2, 0))

        # Drag and Drop visual card
        self.card_drop = ctk.CTkFrame(
            scr, 
            height=160, 
            fg_color=self.C_SIDEBAR_BG, 
            border_width=1, 
            border_color=self.C_BORDER,
            corner_radius=8
        )
        self.card_drop.pack(fill="x", pady=10)
        self.card_drop.pack_propagate(False)

        ctk.CTkLabel(
            self.card_drop,
            text="📂",
            font=("Segoe UI", 36)
        ).pack(pady=(28, 4))

        self.lbl_drop_title = ctk.CTkLabel(
            self.card_drop,
            text="Drag & Drop Client Root Folder Here",
            font=("Segoe UI", 15, "bold"),
            text_color="#60A5FA"
        )
        self.lbl_drop_title.pack()

        self.lbl_drop_subtitle = ctk.CTkLabel(
            self.card_drop,
            text="or select directory manually using browse selector below",
            font=("Segoe UI", 11),
            text_color=self.C_TEXT_MUTED
        )
        self.lbl_drop_subtitle.pack(pady=(2, 0))

        # Input Path selection row card
        card_path = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=6)
        card_path.pack(fill="x", pady=10, ipady=8, ipadx=10)

        ctk.CTkLabel(
            card_path,
            text="Client Root Folder Path:",
            font=("Segoe UI", 12, "bold"),
            text_color=self.C_TEXT_SECONDARY
        ).pack(side="left", padx=(15, 10))

        self.entry_input_path = ctk.CTkEntry(
            card_path,
            placeholder_text="No folder loaded yet...",
            height=34,
            border_width=1,
            border_color=self.C_BORDER,
            fg_color=self.C_MAIN_BG,
            text_color=self.C_TEXT_PRIMARY,
            font=("Segoe UI", 12)
        )
        self.entry_input_path.pack(side="left", fill="x", expand=True, padx=(0, 10))

        btn_browse = ctk.CTkButton(
            card_path,
            text="Browse Folder",
            fg_color=self.C_ACCENT,
            hover_color=self.C_ACCENT_HOVER,
            text_color="white",
            width=110,
            height=34,
            corner_radius=4,
            font=("Segoe UI", 12, "bold"),
            command=self._browse_input_folder
        )
        btn_browse.pack(side="right", padx=(0, 10))

        # Gemini API Key & Offline mode configurations card
        card_gemini = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=6)
        card_gemini.pack(fill="x", pady=10, ipady=8, ipadx=10)

        ctk.CTkLabel(
            card_gemini,
            text="Gemini API Key (Optional):",
            font=("Segoe UI", 12, "bold"),
            text_color=self.C_TEXT_SECONDARY
        ).pack(side="left", padx=(15, 10))

        self.entry_gemini_key = ctk.CTkEntry(
            card_gemini,
            placeholder_text="Enter Gemini API key here... (Falls back to Env Var if empty)",
            show="*",
            height=34,
            border_width=1,
            border_color=self.C_BORDER,
            fg_color=self.C_MAIN_BG,
            text_color=self.C_TEXT_PRIMARY,
            font=("Segoe UI", 12)
        )
        self.entry_gemini_key.pack(side="left", fill="x", expand=True, padx=(0, 15))

        self.chk_skip_gemini = ctk.CTkCheckBox(
            card_gemini,
            text="Offline Mode (Skip Gemini / Cache Only)",
            font=("Segoe UI", 12, "bold"),
            fg_color=self.C_ACCENT,
            hover_color=self.C_ACCENT_HOVER,
            text_color=self.C_TEXT_PRIMARY
        )
        self.chk_skip_gemini.pack(side="right", padx=(0, 15))

        # Action Execution Area
        row_actions = ctk.CTkFrame(scr, fg_color="transparent")
        row_actions.pack(fill="x", pady=(10, 10))

        self.btn_validate = ctk.CTkButton(
            row_actions,
            text="Start Validation Audit",
            fg_color=self.C_SUCCESS,
            hover_color="#047857",
            text_color="white",
            width=180,
            height=42,
            font=("Segoe UI", 13, "bold"),
            command=self._start_validation
        )
        self.btn_validate.pack(side="left")

        self.btn_stop_validation = ctk.CTkButton(
            row_actions,
            text="Abort Process",
            fg_color=self.C_DANGER,
            hover_color="#B91C1C",
            text_color="white",
            width=130,
            height=42,
            font=("Segoe UI", 13, "bold"),
            state="disabled",
            command=self._stop_validation
        )
        self.btn_stop_validation.pack(side="left", padx=10)

        self.btn_import_excel = ctk.CTkButton(
            row_actions,
            text="Import Processed Excel",
            fg_color="transparent",
            border_width=1,
            border_color=self.C_BORDER,
            hover_color=self.C_CARD_BG,
            text_color=self.C_TEXT_PRIMARY,
            width=180,
            height=42,
            font=("Segoe UI", 13, "bold"),
            command=self._import_processed_excel
        )
        self.btn_import_excel.pack(side="left", padx=10)

        # Progress indicators
        row_prog_meta = ctk.CTkFrame(scr, fg_color="transparent")
        row_prog_meta.pack(fill="x", pady=(10, 2))
        
        ctk.CTkLabel(
            row_prog_meta,
            text="Audit validation progress:",
            font=("Segoe UI", 11, "bold"),
            text_color=self.C_TEXT_SECONDARY
        ).pack(side="left")

        self.lbl_validation_eta = ctk.CTkLabel(
            row_prog_meta,
            text="ETA: Standby  | Completed: 0/0",
            font=("Segoe UI", 11),
            text_color=self.C_TEXT_MUTED
        )
        self.lbl_validation_eta.pack(side="right")

        self.prog_validation = ctk.CTkProgressBar(scr, height=10, progress_color=self.C_SUCCESS, fg_color=self.C_SIDEBAR_BG)
        self.prog_validation.pack(fill="x", pady=(0, 15))
        self.prog_validation.set(0.0)

    # ── SCREEN 3: HIGH-FIDELITY RESULTS GRID ──────────────────────────────────
    def _build_results_screen(self) -> None:
        scr = ctk.CTkFrame(self.screen_container, fg_color="transparent")
        self.screens["results"] = scr

        # Header Row
        header = ctk.CTkFrame(scr, fg_color="transparent")
        header.pack(fill="x", pady=(0, 12))
        
        ctk.CTkLabel(
            header,
            text="Validation Results Grid",
            font=("Segoe UI", 24, "bold"),
            text_color=self.C_TEXT_PRIMARY
        ).pack(side="left")

        self.btn_load_last = ctk.CTkButton(
            header,
            text="🔄 Refresh Grid",
            fg_color="transparent",
            border_width=1,
            border_color=self.C_BORDER,
            text_color=self.C_TEXT_PRIMARY,
            hover_color=self.C_CARD_BG,
            width=120,
            height=32,
            font=("Segoe UI", 12, "bold"),
            command=self._load_validation_results
        )
        self.btn_load_last.pack(side="right")

        # Interactive Controls toolbar panel Card
        card_controls = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=6)
        card_controls.pack(fill="x", pady=(0, 10), ipady=6, ipadx=10)

        ctk.CTkButton(
            card_controls,
            text="Select All",
            width=90,
            height=28,
            fg_color="transparent",
            border_width=1,
            border_color=self.C_BORDER,
            text_color=self.C_TEXT_PRIMARY,
            hover_color=self.C_CARD_BG,
            font=("Segoe UI", 11, "bold"),
            command=self._grid_select_all
        ).pack(side="left", padx=(15, 5))

        ctk.CTkButton(
            card_controls,
            text="Deselect All",
            width=100,
            height=28,
            fg_color="transparent",
            border_width=1,
            border_color=self.C_BORDER,
            text_color=self.C_TEXT_PRIMARY,
            hover_color=self.C_CARD_BG,
            font=("Segoe UI", 11, "bold"),
            command=self._grid_deselect_all
        ).pack(side="left", padx=5)

        ctk.CTkFrame(card_controls, width=1, height=20, fg_color=self.C_BORDER).pack(side="left", padx=10)

        self.entry_grid_search = ctk.CTkEntry(
            card_controls,
            placeholder_text="🔍 Filter records by name, state, or ID...",
            width=320,
            height=30,
            border_width=1,
            border_color=self.C_BORDER,
            fg_color=self.C_MAIN_BG,
            font=("Segoe UI", 11)
        )
        self.entry_grid_search.pack(side="left", padx=5)
        self.entry_grid_search.bind("<KeyRelease>", self._grid_search_filter)

        # Sticky Table Grid Columns Definition
        self.grid_cols = [
            ("Select", 50),
            ("Client Legal Name", 180),
            ("GSTIN (Editable)", 160),
            ("GST_Reg_ID", 100),
            ("Status Code", 110),
            ("Verification Remarks", 240),
            ("State", 110),
            ("POB Extracted Address", 340),
            ("Utility Bill", 120)
        ]

        # Sticky Column Header Container (Stays locked on top)
        self.grid_header_frame = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, corner_radius=6, border_width=1, border_color=self.C_BORDER)
        self.grid_header_frame.pack(fill="x", padx=(12, 28), pady=(5, 0), ipady=4)
        
        for col_num, (col_name, col_width) in enumerate(self.grid_cols):
            self.grid_header_frame.grid_columnconfigure(col_num, minsize=col_width)
            lbl = ctk.CTkLabel(
                self.grid_header_frame,
                text=col_name.upper(),
                font=("Segoe UI", 10, "bold"),
                text_color=self.C_TEXT_MUTED,
                anchor="w"
            )
            lbl.grid(row=0, column=col_num, padx=8, pady=8, sticky="w")

        # Scrollable Data table container
        self.grid_scroll_container = ctk.CTkScrollableFrame(
            scr, 
            fg_color=self.C_MAIN_BG, 
            border_width=1, 
            border_color=self.C_BORDER,
            corner_radius=8
        )
        self.grid_scroll_container.pack(fill="both", expand=True, pady=(5, 5))
        for col_num, (_, col_width) in enumerate(self.grid_cols):
            self.grid_scroll_container.grid_columnconfigure(col_num, minsize=col_width)

        # Action Footer Row
        card_footer = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=6)
        card_footer.pack(fill="x", pady=(10, 0), ipady=8, ipadx=15)

        self.lbl_selected_counter = ctk.CTkLabel(
            card_footer,
            text="Selected: 0 rows",
            font=("Segoe UI", 14, "bold"),
            text_color="#93C5FD"
        )
        self.lbl_selected_counter.pack(side="left", padx=15)

        self.btn_proceed_to_reg = ctk.CTkButton(
            card_footer,
            text="Proceed to Credentials Settings",
            fg_color=self.C_SUCCESS,
            hover_color="#047857",
            text_color="white",
            font=("Segoe UI", 13, "bold"),
            width=230,
            height=36,
            command=self._proceed_to_settings
        )
        self.btn_proceed_to_reg.pack(side="right", padx=15)

    # ── SCREEN 4: CREDENTIALS WORKSPACE ──────────────────────────────────────
    def _build_credentials_screen(self) -> None:
        scr = ctk.CTkFrame(self.screen_container, fg_color="transparent")
        self.screens["credentials"] = scr

        ctk.CTkLabel(
            scr,
            text="GST Portal Credentials Workspace",
            font=("Segoe UI", 24, "bold"),
            text_color=self.C_TEXT_PRIMARY
        ).pack(anchor="w", pady=(0, 6))
        
        ctk.CTkLabel(
            scr,
            text="Ensure authentication credentials are set prior to executing selenium-based bulk portal runs.",
            font=("Segoe UI", 12),
            text_color=self.C_TEXT_SECONDARY
        ).pack(anchor="w", pady=(0, 16))

        card_auth = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=8)
        card_auth.pack(fill="x", pady=10, ipady=20, ipadx=24)

        ctk.CTkLabel(
            card_auth,
            text="SECURE CREDENTIALS VAULT",
            font=("Segoe UI", 11, "bold"),
            text_color="#60A5FA"
        ).pack(anchor="w", padx=20, pady=(10, 16))

        # Username entry
        ctk.CTkLabel(card_auth, text="GST Portal Username:", font=("Segoe UI", 12), text_color=self.C_TEXT_SECONDARY).pack(anchor="w", padx=20)
        self.entry_gst_user = ctk.CTkEntry(
            card_auth, 
            placeholder_text="Enter Portal username...", 
            height=36, 
            border_width=1, 
            border_color=self.C_BORDER, 
            fg_color=self.C_MAIN_BG,
            font=("Segoe UI", 12)
        )
        self.entry_gst_user.pack(fill="x", padx=20, pady=(4, 16))

        # Password entry
        ctk.CTkLabel(card_auth, text="GST Portal Password:", font=("Segoe UI", 12), text_color=self.C_TEXT_SECONDARY).pack(anchor="w", padx=20)
        
        pass_row = ctk.CTkFrame(card_auth, fg_color="transparent")
        pass_row.pack(fill="x", padx=20, pady=(4, 20))
        
        self.entry_gst_pass = ctk.CTkEntry(
            pass_row, 
            placeholder_text="Enter Portal password...", 
            show="*", 
            height=36, 
            border_width=1, 
            border_color=self.C_BORDER, 
            fg_color=self.C_MAIN_BG,
            font=("Segoe UI", 12)
        )
        self.entry_gst_pass.pack(side="left", fill="x", expand=True, padx=(0, 8))
        
        self.btn_show_pass = ctk.CTkButton(
            pass_row,
            text="👁",
            width=36,
            height=36,
            fg_color="transparent",
            border_width=1,
            border_color=self.C_BORDER,
            text_color=self.C_TEXT_PRIMARY,
            hover_color=self.C_CARD_BG,
            font=("Segoe UI", 14),
            command=self._toggle_password_visibility
        )
        self.btn_show_pass.pack(side="right")

        # Warning information card
        info_card = ctk.CTkFrame(scr, fg_color=self.C_WARNING_BG, border_width=1, border_color=self.C_WARNING, corner_radius=6)
        info_card.pack(fill="x", pady=10, ipady=10, ipadx=15)
        
        ctk.CTkLabel(
            info_card,
            text="🔐 Security Guidelines:\n• Credentials are stored temporarily in volatile secure memory and are never persisted to disk.\n• Active session elements are wiped automatically immediately after orchestration completion.",
            font=("Segoe UI", 11),
            text_color=self.C_TEXT_PRIMARY,
            justify="left"
        ).pack(anchor="w", padx=15)

        # Action layout
        btn_next = ctk.CTkButton(
            scr,
            text="Proceed to Bulk Run Automation",
            fg_color=self.C_ACCENT,
            hover_color=self.C_ACCENT_HOVER,
            text_color="white",
            height=40,
            font=("Segoe UI", 13, "bold"),
            command=self._proceed_to_automation
        )
        btn_next.pack(side="right", pady=20)

    # ── SCREEN 5: PROCESSING & AUTOMATION CONSOLE ────────────────────────────
    def _build_processing_screen(self) -> None:
        scr = ctk.CTkFrame(self.screen_container, fg_color="transparent")
        self.screens["processing"] = scr

        ctk.CTkLabel(
            scr,
            text="Portal Automation live Run Workspace",
            font=("Segoe UI", 24, "bold"),
            text_color=self.C_TEXT_PRIMARY
        ).pack(anchor="w", pady=(0, 6))

        ctk.CTkLabel(
            scr,
            text="Execute, monitor, and abort place of business selenium compliance sequences.",
            font=("Segoe UI", 12),
            text_color=self.C_TEXT_SECONDARY
        ).pack(anchor="w", pady=(0, 16))

        # Execution Controls Card
        card_exec = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=8)
        card_exec.pack(fill="x", pady=10, ipady=12, ipadx=20)

        ctk.CTkLabel(
            card_exec,
            text="PORTAL AUTOMATION ENGINE CONTROLS",
            font=("Segoe UI", 11, "bold"),
            text_color="#60A5FA"
        ).pack(anchor="w", padx=20, pady=(10, 8))

        btn_row = ctk.CTkFrame(card_exec, fg_color="transparent")
        btn_row.pack(fill="x", padx=20, pady=10)

        self.btn_start_reg = ctk.CTkButton(
            btn_row,
            text="Start Portal Amendment Run",
            fg_color=self.C_SUCCESS,
            hover_color="#047857",
            text_color="white",
            height=42,
            font=("Segoe UI", 13, "bold"),
            command=self._start_registration
        )
        self.btn_start_reg.pack(side="left", padx=(0, 10))

        self.btn_stop_reg = ctk.CTkButton(
            btn_row,
            text="Abort Registration Run",
            fg_color=self.C_DANGER,
            hover_color="#B91C1C",
            text_color="white",
            height=42,
            font=("Segoe UI", 13, "bold"),
            state="disabled",
            command=self._stop_registration
        )
        self.btn_stop_reg.pack(side="left")

        # Cyber-cyan Live log box
        ctk.CTkLabel(
            scr,
            text="Portal Automation Live CLI Streams",
            font=("Segoe UI", 12, "bold"),
            text_color=self.C_TEXT_SECONDARY
        ).pack(anchor="w", pady=(10, 2))

        self.reg_log_box = ctk.CTkTextbox(
            scr, 
            wrap="word", 
            font=("Consolas", 11), 
            fg_color="#020617", 
            border_width=1, 
            border_color=self.C_BORDER,
            text_color="#06B6D4" # Cyberpunk cyan
        )
        self.reg_log_box.pack(fill="both", expand=True, pady=5)

        # Post Completion Actions panel Frame
        self.row_completion_actions = ctk.CTkFrame(scr, fg_color="transparent")
        self.row_completion_actions.pack(fill="x", pady=(10, 0))
        self.row_completion_actions.pack_forget()

        self.btn_open_report = ctk.CTkButton(
            self.row_completion_actions,
            text="📊 Open Summary Report Excel",
            fg_color=self.C_ACCENT,
            hover_color=self.C_ACCENT_HOVER,
            text_color="white",
            width=220,
            height=36,
            font=("Segoe UI", 12, "bold"),
            command=self._open_final_report
        )
        self.btn_open_report.pack(side="left")

    # ── SCREEN 6: REPORTS WORKSPACE ──────────────────────────────────────────
    def _build_reports_screen(self) -> None:
        scr = ctk.CTkFrame(self.screen_container, fg_color="transparent")
        self.screens["reports"] = scr

        ctk.CTkLabel(
            scr,
            text="Compliance Reports & Document Repository",
            font=("Segoe UI", 24, "bold"),
            text_color=self.C_TEXT_PRIMARY
        ).pack(anchor="w", pady=(0, 6))

        ctk.CTkLabel(
            scr,
            text="Manage, access, and generate outputs of audited records and automated actions.",
            font=("Segoe UI", 12),
            text_color=self.C_TEXT_SECONDARY
        ).pack(anchor="w", pady=(0, 16))

        card_reports = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=8)
        card_reports.pack(fill="both", expand=True, pady=10, ipady=16, ipadx=20)

        ctk.CTkLabel(
            card_reports,
            text="FINALIZED ARTEFACTS & ACTIONS",
            font=("Segoe UI", 12, "bold"),
            text_color="#60A5FA"
        ).pack(anchor="w", padx=20, pady=(10, 16))

        # Excel Report trigger
        btn_open_rep_scr = ctk.CTkButton(
            card_reports,
            text="📊 Open Summary Report Excel",
            fg_color=self.C_ACCENT,
            hover_color=self.C_ACCENT_HOVER,
            text_color="white",
            height=40,
            font=("Segoe UI", 13, "bold"),
            command=self._open_final_report
        )
        btn_open_rep_scr.pack(anchor="w", padx=20, pady=10)

        # Browse Output folders trigger
        btn_open_dirs = ctk.CTkButton(
            card_reports,
            text="📂 Open Output Directories",
            fg_color="transparent",
            border_width=1,
            border_color=self.C_BORDER,
            hover_color=self.C_CARD_BG,
            text_color=self.C_TEXT_PRIMARY,
            height=40,
            font=("Segoe UI", 13, "bold"),
            command=self._open_output_dir
        )
        btn_open_dirs.pack(anchor="w", padx=20, pady=10)

        ctk.CTkLabel(
            card_reports,
            text="Reports land under individual state folders in compliance directories.",
            font=("Segoe UI", 11),
            text_color=self.C_TEXT_MUTED
        ).pack(anchor="w", padx=20, pady=(16, 0))

    # ── SCREEN 7: LIVE AUDITOR TEXT LOGS ─────────────────────────────────────
    def _build_logs_screen(self) -> None:
        scr = ctk.CTkFrame(self.screen_container, fg_color="transparent")
        self.screens["logs"] = scr

        # Header Row (Title, Subtitle, Abort and Retry buttons)
        header_row = ctk.CTkFrame(scr, fg_color="transparent")
        header_row.pack(fill="x", pady=(0, 12))
        
        left_header = ctk.CTkFrame(header_row, fg_color="transparent")
        left_header.pack(side="left", fill="y")
        
        ctk.CTkLabel(
            left_header,
            text="Live Auditor Console Streams",
            font=("Segoe UI", 24, "bold"),
            text_color=self.C_TEXT_PRIMARY
        ).pack(anchor="w")

        ctk.CTkLabel(
            left_header,
            text="Review chronological diagnostic details, file analysis steps, and warning pipelines.",
            font=("Segoe UI", 12),
            text_color=self.C_TEXT_SECONDARY
        ).pack(anchor="w", pady=(2, 0))

        # Right side Buttons
        self.btn_stop_validation_logs = ctk.CTkButton(
            header_row,
            text="Abort Process",
            fg_color=self.C_DANGER,
            hover_color="#B91C1C",
            text_color="white",
            width=130,
            height=36,
            font=("Segoe UI", 13, "bold"),
            state="disabled",
            command=self._stop_validation
        )
        self.btn_stop_validation_logs.pack(side="right", padx=(10, 0))

        self.btn_retry_validation = ctk.CTkButton(
            header_row,
            text="Retry Failed/Skipped",
            fg_color=self.C_WARNING,
            hover_color=self.C_WARNING_HOVER,
            text_color="white",
            width=170,
            height=36,
            font=("Segoe UI", 13, "bold"),
            state="disabled",
            command=self._retry_validation
        )
        self.btn_retry_validation.pack(side="right", padx=(10, 0))

        # 1. Clean Validation Dashboard Live Stats Card (moved to Logs screen)
        self.card_val_dashboard = ctk.CTkFrame(
            scr, 
            fg_color=self.C_SIDEBAR_BG, 
            border_width=1, 
            border_color=self.C_BORDER,
            corner_radius=8
        )
        self.card_val_dashboard.pack(fill="x", pady=10)
        
        self.card_val_dashboard.grid_columnconfigure(0, weight=1)
        self.card_val_dashboard.grid_columnconfigure(1, weight=1)
        self.card_val_dashboard.grid_columnconfigure(2, weight=1)
        
        # Stat 1: Total Supplied
        f_supplied = ctk.CTkFrame(self.card_val_dashboard, fg_color="transparent")
        f_supplied.grid(row=0, column=0, pady=12, sticky="nsew")
        self.lbl_val_supplied = ctk.CTkLabel(f_supplied, text="0", font=("Segoe UI", 24, "bold"), text_color=self.C_TEXT_PRIMARY)
        self.lbl_val_supplied.pack()
        ctk.CTkLabel(f_supplied, text="FILES SUPPLIED", font=("Segoe UI", 9, "bold"), text_color=self.C_TEXT_MUTED).pack()
        
        # Stat 2: Completed (Success)
        f_completed = ctk.CTkFrame(self.card_val_dashboard, fg_color="transparent")
        f_completed.grid(row=0, column=1, pady=12, sticky="nsew")
        self.lbl_val_completed = ctk.CTkLabel(f_completed, text="0", font=("Segoe UI", 24, "bold"), text_color=self.C_SUCCESS)
        self.lbl_val_completed.pack()
        ctk.CTkLabel(f_completed, text="COMPLETED", font=("Segoe UI", 9, "bold"), text_color=self.C_TEXT_MUTED).pack()
        
        # Stat 3: Failed / Skipped
        f_failed = ctk.CTkFrame(self.card_val_dashboard, fg_color="transparent")
        f_failed.grid(row=0, column=2, pady=12, sticky="nsew")
        self.lbl_val_failed = ctk.CTkLabel(f_failed, text="0", font=("Segoe UI", 24, "bold"), text_color=self.C_DANGER)
        self.lbl_val_failed.pack()
        ctk.CTkLabel(f_failed, text="FAILED / SKIPPED", font=("Segoe UI", 9, "bold"), text_color=self.C_TEXT_MUTED).pack()

        # Green Console Log stream textbox
        self.log_txt_box = ctk.CTkTextbox(
            scr, 
            wrap="word", 
            font=("Consolas", 11), 
            fg_color="#020617", 
            border_width=1, 
            border_color=self.C_BORDER,
            text_color="#10B981" # Emerald/Green
        )
        self.log_txt_box.pack(fill="both", expand=True, pady=5)

    # ── SCREEN 8: AUTOMATION ENGINE OPTIONS ──────────────────────────────────
    def _build_settings_screen(self) -> None:
        scr = ctk.CTkFrame(self.screen_container, fg_color="transparent")
        self.screens["settings"] = scr

        ctk.CTkLabel(
            scr,
            text="System Settings & Configurations",
            font=("Segoe UI", 24, "bold"),
            text_color=self.C_TEXT_PRIMARY
        ).pack(anchor="w", pady=(0, 6))

        ctk.CTkLabel(
            scr,
            text="Configure engine limits, webdriver run options, and system thresholds.",
            font=("Segoe UI", 12),
            text_color=self.C_TEXT_SECONDARY
        ).pack(anchor="w", pady=(0, 16))

        card_options = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=8)
        card_options.pack(fill="both", expand=True, pady=10, ipady=16, ipadx=20)

        ctk.CTkLabel(
            card_options,
            text="AUTOMATION ENGINE OPTIONS",
            font=("Segoe UI", 12, "bold"),
            text_color="#60A5FA"
        ).pack(anchor="w", padx=20, pady=(10, 12))

        # Headless Mode
        self.chk_headless = ctk.CTkCheckBox(
            card_options, 
            text="Headless Mode (Background Selenium Run)", 
            font=("Segoe UI", 12),
            fg_color=self.C_ACCENT,
            hover_color=self.C_ACCENT_HOVER
        )
        self.chk_headless.pack(anchor="w", padx=20, pady=10)
        self.chk_headless.select()

        # Retry Attempts and Timings Container
        row_opts = ctk.CTkFrame(card_options, fg_color="transparent")
        row_opts.pack(fill="x", padx=20, pady=16)

        ctk.CTkLabel(row_opts, text="Retry Attempts Limit:", font=("Segoe UI", 12), text_color=self.C_TEXT_SECONDARY).grid(row=0, column=0, sticky="w", pady=8)
        self.entry_retry_limit = ctk.CTkEntry(row_opts, width=80, height=30, border_width=1, border_color=self.C_BORDER, fg_color=self.C_MAIN_BG, justify="center")
        self.entry_retry_limit.grid(row=0, column=1, sticky="w", padx=12, pady=8)
        self.entry_retry_limit.insert(0, "2")

        ctk.CTkLabel(row_opts, text="Selenium Timeout Threshold (sec):", font=("Segoe UI", 12), text_color=self.C_TEXT_SECONDARY).grid(row=1, column=0, sticky="w", pady=8)
        self.entry_timeout = ctk.CTkEntry(row_opts, width=80, height=30, border_width=1, border_color=self.C_BORDER, fg_color=self.C_MAIN_BG, justify="center")
        self.entry_timeout.grid(row=1, column=1, sticky="w", padx=12, pady=8)
        self.entry_timeout.insert(0, "15")

    # ── SCREEN 9: ABOUT & INSTRUCTION ALIGNMENT ──────────────────────────────
    def _build_about_screen(self) -> None:
        scr = ctk.CTkFrame(self.screen_container, fg_color="transparent")
        self.screens["about"] = scr

        ctk.CTkLabel(
            scr,
            text="About GST Smart Automation Hub",
            font=("Segoe UI", 24, "bold"),
            text_color=self.C_TEXT_PRIMARY
        ).pack(anchor="w", pady=(0, 6))

        ctk.CTkLabel(
            scr,
            text="Compliance details and engine specifications.",
            font=("Segoe UI", 12),
            text_color=self.C_TEXT_SECONDARY
        ).pack(anchor="w", pady=(0, 16))

        card_info = ctk.CTkFrame(scr, fg_color=self.C_SIDEBAR_BG, border_width=1, border_color=self.C_BORDER, corner_radius=8)
        card_info.pack(fill="both", expand=True, pady=10, ipady=16, ipadx=20)

        ctk.CTkLabel(
            card_info,
            text="COMPLIANCE ORCHESTRATION LAYER",
            font=("Segoe UI", 12, "bold"),
            text_color="#60A5FA"
        ).pack(anchor="w", padx=20, pady=(10, 12))

        desc = (
            "GST Smart Automation Hub is a premium enterprise-grade desktop compliance workspace.\n\n"
            "• Configured in alignment with CBIC Instruction No. 03/2025-GST (Para 6) for APoB Verification.\n"
            "• Integrates automated document audits with structural optical OCR clustering.\n"
            "• Employs dual-tier Gemini generative AI validation and selenium browser automation."
        )
        ctk.CTkLabel(
            card_info,
            text=desc,
            font=("Segoe UI", 12),
            text_color=self.C_TEXT_PRIMARY,
            justify="left"
        ).pack(anchor="w", padx=20, pady=10)

    def _open_output_dir(self) -> None:
        """Opens the outputs folder."""
        out = self.entry_input_path.get().strip()
        if out:
            p = Path(out).parent / f"{Path(out).name}_client_outputs"
            if p.exists():
                os.startfile(p)
                return
        messagebox.showwarning("Open Folder", "No outputs directory mapped yet. Execute validation first.")

    def _open_final_report(self) -> None:
        """Opens the finalized registrations summary workbook."""
        if self.output_folder_path:
            p = self.output_folder_path / "Final_Registration_Report.xlsx"
            if p.exists():
                os.startfile(p)
                return
        messagebox.showwarning("Open Report", "Final Registration Report not generated yet.")

    def _browse_input_folder(self) -> None:
        """Triggers Tkinter filedialog directories selectors."""
        selected = filedialog.askdirectory(title="Select Client root Documents Folder")
        if selected:
            self.input_folder_path = Path(selected)
            self.entry_input_path.delete(0, "end")
            self.entry_input_path.insert(0, selected)
            
            # Auto-infer outputs base sibling
            suggested = self.input_folder_path.parent / f"{self.input_folder_path.name}_client_outputs"
            self.output_folder_path = suggested

    def _setup_drag_and_drop(self) -> None:
        """Registers and binds drag and drop target components using tkinterdnd2 if available."""
        try:
            from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
            from tkinterdnd2.TkinterDnD import _require  # type: ignore
            # Dynamically initialize Tcl DnD package in the tk interpreter
            self.TkdndVersion = _require(self)
            
            # Dynamically patch DnDWrapper methods and attributes if they are missing
            if not hasattr(self, "drop_target_register"):
                for attr in dir(TkinterDnD.DnDWrapper):
                    if not attr.startswith("__"):
                        val = getattr(TkinterDnD.DnDWrapper, attr)
                        if callable(val):
                            setattr(self, attr, val.__get__(self, self.__class__))
                        else:
                            setattr(self, attr, val)
            
            self.drop_target_register(DND_FILES)
            self.dnd_bind("<<Drop>>", self._on_drag_drop)
            self.lbl_drop_title.configure(text="Drag & Drop Client Root Folder Here", text_color="#60A5FA")
            self.lbl_drop_subtitle.configure(text="Drag-and-drop enabled.")
        except Exception as e:
            self.lbl_drop_subtitle.configure(text=f"DND support disabled. Reason: {e}")

    def _on_drag_drop(self, event: Any) -> None:
        """Callback processing files dropped onto the Tkinter window."""
        dropped = (event.data or "").strip().strip("{}")
        if dropped:
            p = Path(dropped)
            if p.exists():
                if p.is_dir():
                    self.input_folder_path = p
                    self.entry_input_path.delete(0, "end")
                    self.entry_input_path.insert(0, str(p))
                    suggested = p.parent / f"{p.name}_client_outputs"
                    self.output_folder_path = suggested
                    self.lbl_drop_title.configure(text="Drag & Drop Client Root Folder Here", text_color="#60A5FA")
                    self.lbl_drop_subtitle.configure(text=f"Loaded folder: {p.name}")
                elif p.suffix.lower() == ".xlsx":
                    self.master_workbook_path = p
                    self.output_folder_path = p.parent
                    
                    # Try to auto-infer matching input folder if possible
                    name_str = p.parent.name
                    if name_str.endswith("_client_outputs"):
                        in_name = name_str[:-15]
                        in_path = p.parent.parent / in_name
                        if in_path.exists() and in_path.is_dir():
                            self.input_folder_path = in_path
                            self.entry_input_path.delete(0, "end")
                            self.entry_input_path.insert(0, str(in_path))
                    
                    self.log_txt_box.insert("end", f"\n[HUB] Loaded master validation Excel file via Drag-and-Drop:\n{p}\n")
                    self.log_txt_box.see("end")
                    
                    # Switch to results screen and load the validation results
                    self._switch_screen("results")
                    self._load_validation_results()
                else:
                    self.lbl_drop_subtitle.configure(text=f"Unsupported file type: {p.suffix}")

    # ── VALIDATION CONTROL DISPATCHERS ────────────────────────────────────────
    def _start_validation(self) -> None:
        """Spawns validation adapter subprocess in background thread."""
        in_path = self.entry_input_path.get().strip()
        if not in_path or not os.path.isdir(in_path):
            messagebox.showerror("Validation Error", "Please specify a valid client root directory.")
            return

        self.input_folder_path = Path(in_path)
        if not self.output_folder_path:
            self.output_folder_path = self.input_folder_path.parent / f"{self.input_folder_path.name}_client_outputs"

        # Reset Live Stats Dashboard Counters
        self.validation_failed_count = 0
        self.lbl_val_supplied.configure(text="0")
        self.lbl_val_completed.configure(text="0")
        self.lbl_val_failed.configure(text="0")

        self.btn_validate.configure(state="disabled")
        self.btn_retry_validation.configure(state="disabled")
        self.btn_stop_validation.configure(state="normal")
        self.btn_stop_validation_logs.configure(state="normal")
        self.log_txt_box.delete("1.0", "end")
        self.prog_validation.set(0.0)
        
        # Switch immediately to live logs screen
        self._switch_screen("logs")

        # Spawn validation subprocess
        api_key = self.entry_gemini_key.get().strip()
        if not api_key:
            try:
                from gst_validator_v2 import DEFAULT_API_KEY
                api_key = DEFAULT_API_KEY
            except ImportError:
                api_key = "AIzaSyCH4N-4VlgBSa7Knuru9CsKtd8KtnLyCFo"
        skip_gemini = self.chk_skip_gemini.get() == 1

        self.validator_adapter = ValidationAdapter(
            input_folder=self.input_folder_path,
            output_folder=self.output_folder_path,
            api_key=api_key,
            workers=4,
            skip_gemini=skip_gemini
        )

        def log_cb(line: str) -> None:
            self.ui_queue.put(("val_log", line))

        def prog_cb(done: int, total: int, eta: str) -> None:
            self.ui_queue.put(("val_prog", (done, total, eta)))

        def comp_cb(success: bool, err_msg: Optional[str], master_path: Optional[Path]) -> None:
            self.ui_queue.put(("val_comp", (success, err_msg, master_path)))

        self.validator_adapter.start(log_cb, prog_cb, comp_cb)

    def _stop_validation(self) -> None:
        """Terminates active validator thread."""
        if self.validator_adapter:
            self.validator_adapter.stop()
            self.log_txt_box.insert("end", "\n[HUB] Auditor validation aborted by operator.\n")
            self.btn_validate.configure(state="normal")
            self.btn_stop_validation.configure(state="disabled")
            self.btn_stop_validation_logs.configure(state="disabled")
            if self.validation_failed_count > 0:
                self.btn_retry_validation.configure(state="normal")
            else:
                self.btn_retry_validation.configure(state="disabled")

    def _retry_validation(self) -> None:
        """Triggers validation audit again to retry failed/skipped items."""
        self.log_txt_box.insert("end", "\n[HUB] Retrying failed/skipped validation audits (re-running with cache check)...\n")
        self._start_validation()

    def _import_processed_excel(self) -> None:
        """Opens file selector to select a pre-existing master validation workbook Excel."""
        selected_file = filedialog.askopenfilename(
            title="Select Pre-existing Master Validation Excel",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not selected_file:
            return

        p = Path(selected_file)
        if not p.exists():
            messagebox.showerror("Error", f"Selected file does not exist: {selected_file}")
            return

        self.master_workbook_path = p
        self.output_folder_path = p.parent

        name_str = p.parent.name
        if name_str.endswith("_client_outputs"):
            in_name = name_str[:-15]
            in_path = p.parent.parent / in_name
            if in_path.exists() and in_path.is_dir():
                self.input_folder_path = in_path
                self.entry_input_path.delete(0, "end")
                self.entry_input_path.insert(0, str(in_path))

        self.log_txt_box.insert("end", f"\n[HUB] Loaded master validation Excel file explicitly:\n{p}\n")
        self.log_txt_box.see("end")

        # Switch to results screen and load the validation results
        self._switch_screen("results")
        self._load_validation_results()

    # ── REGISTRATION AUTOMATION DISPATCHERS ────────────────────────────────────
    def _proceed_to_settings(self) -> None:
        """Transition from grid selection to credentials settings workspace."""
        self._switch_screen("credentials")

    def _proceed_to_automation(self) -> None:
        """Transitions from credentials settings to processing workspace after validation."""
        username = self.entry_gst_user.get().strip()
        password = self.entry_gst_pass.get().strip()
        if not username or not password:
            messagebox.showerror("Credentials Required", "Please configure your GST Portal Username and Password before proceeding to automation.")
            return
        self._switch_screen("processing")

    def _start_registration(self) -> None:
        """Triggers bulk portal automation with validation records."""
        # Grab grid selections
        self.selected_records = []
        for r_widgets in self.grid_rows:
            if r_widgets["chk"].get() == 1:
                self.selected_records.append(r_widgets["data"])

        if not self.selected_records:
            messagebox.showerror("Registration Error", "Please select at least one record in the Grid before proceeding.")
            self._switch_screen("results")
            return

        username = self.entry_gst_user.get().strip()
        password = self.entry_gst_pass.get().strip()
        if not username or not password:
            messagebox.showerror("Credentials Error", "Please enter valid GST Portal Username and Password.")
            return

        # Securely store in-memory
        SecureVault.store_credentials(username, password)
        
        # Clear entries from screen for security
        self.entry_gst_user.delete(0, "end")
        self.entry_gst_pass.delete(0, "end")

        self.btn_start_reg.configure(state="disabled")
        self.btn_stop_reg.configure(state="normal")
        self.reg_log_box.delete("1.0", "end")
        self.row_completion_actions.pack_forget()

        # Initialize folders & checkpoint managers
        internal_dir = self.output_folder_path / "_internal"
        self.checkpoint_mgr = CheckpointManager(internal_dir)

        # Setup dynamic Batch ID
        self.active_batch_id = BatchBuilder.generate_batch_id()
        os.environ["GST_RUN_TIME"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Initialize recovery checkpoints
        self.checkpoint_mgr.initialize_batch(self.active_batch_id, self.selected_records, operator="SmartHubOperator")

        # Compile Excel batch file for amend_req.py
        temp_batch_excel = internal_dir / "runtime_batch.xlsx"
        sel_ids = [r.get("GST_Reg_ID") for r in self.selected_records if r.get("GST_Reg_ID")]
        
        parsed_batch = BatchBuilder.build_runtime_excel(
            master_excel_path=self.master_workbook_path,
            selected_reg_ids=sel_ids,
            batch_id=self.active_batch_id,
            output_excel_path=temp_batch_excel,
            operator="SmartHubOperator"
        )

        # Trigger adapter execution
        headless = self.chk_headless.get() == 1
        timeout = int(self.entry_timeout.get().strip() or "15")

        self.val_stat_total.configure(text=f"{len(parsed_batch)}")
        self.val_stat_pending.configure(text=f"{len(parsed_batch)}")
        self.val_stat_completed.configure(text="0")
        self.val_stat_failed.configure(text="0")
        self.val_stat_eta.configure(text="Calc...")

        self.execution_durations.clear()
        
        # Run subprocess
        self.registration_adapter = RegistrationAdapter(
            excel_path=temp_batch_excel,
            checkpoint_path=internal_dir / "checkpoint.json",
            username=username,
            password=password,
            timeout=timeout,
            headless=headless
        )

        def log_cb(line: str) -> None:
            self.ui_queue.put(("reg_log", line))

        def status_cb(gstin: str, stage: str, done: int, failed: int) -> None:
            self.ui_queue.put(("reg_status", (gstin, stage, done, failed)))

        def comp_cb(success: bool, err_msg: Optional[str]) -> None:
            self.ui_queue.put(("reg_comp", (success, err_msg)))

        def captcha_cb(img_path: str, portal_hash: str = "N/A") -> None:
            self.ui_queue.put(("reg_captcha", (img_path, portal_hash)))

        self.registration_adapter.start(log_cb, status_cb, comp_cb, captcha_cb)

    def _stop_registration(self) -> None:
        """Kills registration adapter process thread."""
        if self.registration_adapter:
            self.registration_adapter.stop()
            self.reg_log_box.insert("end", "\n[HUB] Bulk amendment registration sequence terminated by operator.\n")
            self.btn_start_reg.configure(state="normal")
            self.btn_stop_reg.configure(state="disabled")
            SecureVault.clear()

    # ── INTERACTIVE TABLE GRID GENERATION ─────────────────────────────────────
    def _load_validation_results(self) -> None:
        """Scans directories, opens master workbook, and builds the visual list rows."""
        if not self.output_folder_path:
            messagebox.showwarning("Grid Update", "No client validation outputs found. Execute validation first.")
            return

        if self.master_workbook_path and self.master_workbook_path.exists():
            pass
        else:
            workbooks = sorted(
                self.output_folder_path.glob("GST_Validation_Workbook_*.xlsx"),
                key=lambda p: p.stat().st_mtime,
                reverse=True
            )
            if workbooks:
                self.master_workbook_path = workbooks[0]
            else:
                legacy = self.output_folder_path / "GST_Validation_Workbook.xlsx"
                if legacy.exists():
                    self.master_workbook_path = legacy
                else:
                    messagebox.showerror("Grid Error", "Master validation workbook missing. Validation might have failed.")
                    return

        # Parse master workbook
        try:
            wb = openpyxl.load_workbook(self.master_workbook_path, data_only=True)
            if "GST_Registrations" not in wb.sheetnames:
                wb.close()
                messagebox.showerror("Grid Error", "Target sheet 'GST_Registrations' not found in workbook.")
                return
            
            ws = wb["GST_Registrations"]
            rows = list(ws.iter_rows(values_only=True))
            wb.close()
        except Exception as e:
            messagebox.showerror("Grid Parse Error", f"Could not read spreadsheet details: {e}")
            return

        if len(rows) < 2:
            messagebox.showinfo("Grid Warning", "Spreadsheet contains no records.")
            return

        # Clear old rows from container
        for widget in self.grid_scroll_container.winfo_children():
            widget.destroy()

        self.grid_rows = []
        headers = [str(cell).strip() if cell else "" for cell in rows[0]]

        # Mappings indices
        col_indices = {
            "reg_id": -1, "name": -1, "state": -1, "address": -1, "status": -1, "remarks": -1,
            "util_status": -1, "compliance_percent": -1
        }
        for idx, h in enumerate(headers):
            hl = h.lower()
            if hl == "gst_reg_id": col_indices["reg_id"] = idx
            elif "legal name" in hl: col_indices["name"] = idx
            elif "state" in hl: col_indices["state"] = idx
            elif "address" in hl and "ppob" in hl: col_indices["address"] = idx
            elif "final status" in hl: col_indices["status"] = idx
            elif "remarks" in hl: col_indices["remarks"] = idx
            elif "utility" in hl or "bill" in hl: col_indices["util_status"] = idx
            elif "compliance" in hl or "wcp" in hl or "%" in hl: col_indices["compliance_percent"] = idx

        # GSTIN Indian regex search pattern for folder/file scan
        gstin_regex = re.compile(r"\b([0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1})\b")

        row_counter = 0
        for row_data in rows[1:]:
            if all(cell is None for cell in row_data):
                continue

            reg_id = str(row_data[col_indices["reg_id"]] or f"REG-{row_counter:03d}").strip()
            client_name = str(row_data[col_indices["name"]] or "Client").strip()
            state = str(row_data[col_indices["state"]] or "").strip()
            address = str(row_data[col_indices["address"]] or "").strip()
            status = str(row_data[col_indices["status"]] or "Warning").strip()
            remarks = str(row_data[col_indices["remarks"]] or "").strip()
            util_status = "Not Detected" if col_indices["util_status"] == -1 else str(row_data[col_indices["util_status"]] or "Not Detected")

            # Try to auto-infer GSTIN from files
            gstin_inferred = ""
            if self.input_folder_path:
                for path in self.input_folder_path.rglob("*"):
                    m = gstin_regex.search(path.name)
                    if m and client_name.lower() in path.name.lower():
                        gstin_inferred = m.group(1)
                        break

            # Alternate row background fill
            bg_row = self.C_CARD_BG if row_counter % 2 == 0 else self.C_CARD_DARK

            # Row container cell frame for absolute styling
            row_frame = ctk.CTkFrame(self.grid_scroll_container, fg_color=bg_row, corner_radius=4)
            row_frame.grid(row=row_counter, column=0, columnspan=len(self.grid_cols), sticky="ew", pady=2, ipady=4)
            for c_idx, (_, col_width) in enumerate(self.grid_cols):
                row_frame.grid_columnconfigure(c_idx, minsize=col_width)

            chk_var = ctk.IntVar(value=1)
            chk = ctk.CTkCheckBox(row_frame, text="", variable=chk_var, width=20, command=self._grid_on_check_changed)
            chk.grid(row=0, column=0, padx=8, pady=4)

            display_name = client_name[:22] + "..." if len(client_name) > 22 else client_name
            lbl_name = ctk.CTkLabel(row_frame, text=display_name, font=("Segoe UI", 12, "bold"), text_color=self.C_TEXT_PRIMARY, width=self.grid_cols[1][1], anchor="w")
            lbl_name.grid(row=0, column=1, padx=8, pady=4, sticky="w")

            entry_gstin = ctk.CTkEntry(
                row_frame, 
                width=self.grid_cols[2][1] - 16, 
                height=26, 
                border_width=1, 
                border_color=self.C_BORDER,
                fg_color=self.C_MAIN_BG,
                font=("Consolas", 11)
            )
            entry_gstin.grid(row=0, column=2, padx=8, pady=4, sticky="w")
            if gstin_inferred:
                entry_gstin.insert(0, gstin_inferred)

            lbl_reg_id = ctk.CTkLabel(row_frame, text=reg_id, font=("Consolas", 11), text_color=self.C_TEXT_SECONDARY, width=self.grid_cols[3][1], anchor="w")
            lbl_reg_id.grid(row=0, column=3, padx=8, pady=4, sticky="w")

            # Soft-badge rounded pill colors for status metrics
            badge_fg = "#94A3B8"
            badge_bg = "#334155"
            badge_border = "#475569"
            if status in ("Clean", "Valid"):
                badge_fg = "#34D399"
                badge_bg = "#064E3B"
                badge_border = "#047857"
            elif status in ("Warning", "Attention"):
                badge_fg = "#FBBF24"
                badge_bg = "#78350F"
                badge_border = "#D97706"
            elif status in ("High Risk", "Invalid"):
                badge_fg = "#F87171"
                badge_bg = "#7F1D1D"
                badge_border = "#B91C1C"

            # Calculate actual risk percentage from Weighted Compliance % column
            display_status = status.upper()
            if col_indices["compliance_percent"] != -1:
                comp_val = row_data[col_indices["compliance_percent"]]
                if comp_val is not None:
                    try:
                        # Clean and convert to float
                        val_num = float(str(comp_val).replace("%", "").strip())
                        if val_num <= 1.0:
                            val_num = val_num * 100
                        risk_pct = 100 - val_num
                        display_status = f"{int(risk_pct)}% RISK"
                    except Exception:
                        pass

            lbl_status_badge = ctk.CTkLabel(
                row_frame,
                text=display_status,
                font=("Segoe UI", 10, "bold"),
                text_color=badge_fg,
                fg_color=badge_bg,
                corner_radius=12,
                width=self.grid_cols[4][1] - 16,
                height=22
            )
            # visual border enhancement on status label
            lbl_status_badge.grid(row=0, column=4, padx=8, pady=4)

            lbl_remarks = ctk.CTkLabel(row_frame, text=remarks[:28] + "..." if len(remarks)>28 else remarks, font=("Segoe UI", 11), text_color=self.C_TEXT_SECONDARY, width=self.grid_cols[5][1], anchor="w")
            lbl_remarks.grid(row=0, column=5, padx=8, pady=4, sticky="w")

            lbl_state = ctk.CTkLabel(row_frame, text=state, font=("Segoe UI", 11), text_color=self.C_TEXT_PRIMARY, width=self.grid_cols[6][1], anchor="w")
            lbl_state.grid(row=0, column=6, padx=8, pady=4, sticky="w")

            lbl_addr = ctk.CTkLabel(row_frame, text=address[:42] + "..." if len(address)>42 else address, font=("Segoe UI", 11), text_color=self.C_TEXT_SECONDARY, width=self.grid_cols[7][1], anchor="w")
            lbl_addr.grid(row=0, column=7, padx=8, pady=4, sticky="w")

            lbl_util = ctk.CTkLabel(row_frame, text=util_status, font=("Segoe UI", 11), text_color=self.C_TEXT_SECONDARY, width=self.grid_cols[8][1], anchor="w")
            lbl_util.grid(row=0, column=8, padx=8, pady=4, sticky="w")

            raw_record = {}
            for col_idx, h in enumerate(headers):
                raw_record[h] = row_data[col_idx]

            row_widgets = {
                "chk": chk_var,
                "widgets": [row_frame, chk, lbl_name, entry_gstin, lbl_reg_id, lbl_status_badge, lbl_remarks, lbl_state, lbl_addr, lbl_util],
                "entry_gstin": entry_gstin,
                "client_name": client_name,
                "state": state,
                "gst_reg_id": reg_id,
                "data": raw_record
            }
            self.grid_rows.append(row_widgets)
            row_counter += 1

        self._grid_on_check_changed()
        messagebox.showinfo("Grid Load Success", f"Audited Excel validation loaded successfully: {len(self.grid_rows)} records added to grid.")

    def _grid_on_check_changed(self) -> None:
        """Callback to count selections."""
        total_selected = sum(1 for r in self.grid_rows if r["chk"].get() == 1)
        self.lbl_selected_counter.configure(text=f"Selected: {total_selected} / {len(self.grid_rows)} rows queued")

    def _grid_select_all(self) -> None:
        for r in self.grid_rows:
            r["chk"].set(1)
        self._grid_on_check_changed()

    def _grid_deselect_all(self) -> None:
        for r in self.grid_rows:
            r["chk"].set(0)
        self._grid_on_check_changed()

    def _grid_search_filter(self, event: Any) -> None:
        """Filters grid row visibilities based on keys."""
        query = self.entry_grid_search.get().strip().lower()
        for r in self.grid_rows:
            match = (
                query in r["client_name"].lower() or
                query in r["state"].lower() or
                query in r["gst_reg_id"].lower() or
                query in r["entry_gstin"].get().strip().lower()
            )
            # hide or show container frame widget
            if match:
                r["widgets"][0].grid()
            else:
                r["widgets"][0].grid_remove()

    # ── THE DRAIN LOOP FOR REAL-TIME CONSOLE PIPES ───────────────────────────
    def _drain_ui_queue(self) -> None:
        """Thread-safe UI updates draining queues."""
        while True:
            try:
                msg_type, payload = self.ui_queue.get_nowait()
            except queue.Empty:
                break

            if msg_type == "val_log":
                self.log_txt_box.insert("end", f"{payload}\n")
                self.log_txt_box.see("end")
                
                # Highly robust, independent parser blocks for live statistics card
                payload_lower = payload.lower()
                
                # 1. Total clients to process
                if "clients to process:" in payload_lower or "total clients" in payload_lower:
                    try:
                        total = int(payload.split(":")[-1].strip())
                        self.lbl_val_supplied.configure(text=str(total))
                    except Exception:
                        pass
                
                # 2. Success (completed) client audit (triggered by standard text markers or the checkmark)
                if "compliance=" in payload_lower or "status=" in payload_lower or "✔" in payload:
                    try:
                        # Increment completed count live (only if we haven't completed everything yet to avoid double counting summary)
                        current = int(self.lbl_val_completed.cget("text"))
                        supplied = int(self.lbl_val_supplied.cget("text"))
                        if supplied == 0 or current < supplied:
                            self.lbl_val_completed.configure(text=str(current + 1))
                    except Exception:
                        pass
                
                # 3. Failed client audit (triggered by standard text markers or the error cross)
                if "❌" in payload or "business_fail" in payload_lower:
                    try:
                        self.validation_failed_count += 1
                        self.lbl_val_failed.configure(text=str(self.validation_failed_count))
                    except Exception:
                        pass
                
                # 4. Final summary parser fallback/corrections
                if "processed (ok)" in payload_lower:
                    try:
                        ok = int(payload.split(":")[-1].strip())
                        self.lbl_val_completed.configure(text=str(ok))
                    except Exception:
                        pass
                if "failed (error)" in payload_lower:
                    try:
                        errs = int(payload.split(":")[-1].strip())
                        self.validation_failed_count = errs
                        self.lbl_val_failed.configure(text=str(errs))
                    except Exception:
                        pass

                # 5. Live progress bar and ETA label updates
                try:
                    supplied = int(self.lbl_val_supplied.cget("text"))
                    completed = int(self.lbl_val_completed.cget("text"))
                    if supplied > 0:
                        self.prog_validation.set(completed / supplied)
                        self.lbl_validation_eta.configure(text=f"Processed: {completed}/{supplied}")
                except Exception:
                    pass

            elif msg_type == "val_prog":
                done, total, eta = payload
                self.prog_validation.set(done / max(1, total))
                self.lbl_validation_eta.configure(text=f"ETA: {eta}  | Processed: {done}/{total}")
                self.lbl_val_supplied.configure(text=str(total))
                self.lbl_val_completed.configure(text=str(done))

            elif msg_type == "val_comp":
                success, err_msg, master_path = payload
                self.btn_validate.configure(state="normal")
                self.btn_stop_validation.configure(state="disabled")
                self.btn_stop_validation_logs.configure(state="disabled")
                if self.validation_failed_count > 0:
                    self.btn_retry_validation.configure(state="normal")
                else:
                    self.btn_retry_validation.configure(state="disabled")
                
                if success:
                    self.master_workbook_path = master_path
                    self.log_txt_box.insert("end", f"\n[HUB] Validation Completed Successfully!\nExcel saved at: {master_path}\n")
                    self._switch_screen("results")
                    self._load_validation_results()
                else:
                    self.log_txt_box.insert("end", f"\n[HUB] Validation audit execution failed:\n{err_msg}\n")
                    messagebox.showerror("Validation Failed", f"auditing run failed:\n{err_msg}")

            elif msg_type == "reg_log":
                self.reg_log_box.insert("end", f"{payload}\n")
                self.reg_log_box.see("end")

            elif msg_type == "reg_status":
                gstin, stage, done, failed = payload
                self.lbl_tel_client.configure(text=f"Current Active Record: {gstin}")
                self.lbl_tel_stage.configure(text=f"Current Pipeline Stage: {stage}")
                
                total_sel = len(self.selected_records)
                pending = max(0, total_sel - (done + failed))
                
                self.val_stat_pending.configure(text=f"{pending}")
                self.val_stat_completed.configure(text=f"{done}")
                self.val_stat_failed.configure(text=f"{failed}")

            elif msg_type == "reg_captcha":
                img_path, portal_hash = payload
                self._show_captcha_popup(img_path, portal_hash)

            elif msg_type == "reg_comp":
                success, err_msg = payload
                self.btn_start_reg.configure(state="normal")
                self.btn_stop_reg.configure(state="disabled")
                
                username, password = SecureVault.get_credentials()
                
                if self.output_folder_path:
                    completed_excel = Path("completed_clients.xlsx")
                    failed_excel = Path("failed_clients.xlsx")
                    final_report = self.output_folder_path / "Final_Registration_Report.xlsx"

                    self.reg_log_box.insert("end", "\n[HUB] Compiling Finalized Registration Excel Report...\n")
                    
                    for r_widgets in self.grid_rows:
                        gst_id = r_widgets["gst_reg_id"]
                        edited_gstin = r_widgets["entry_gstin"].get().strip().upper()
                        
                        for record in self.selected_records:
                            if record.get("GST_Reg_ID") == gst_id:
                                record["GSTIN"] = edited_gstin

                    ReportingEngine.generate_final_report(
                        batch_records=self.selected_records,
                        batch_id=self.active_batch_id or "BATCH_RUN",
                        completed_excel_path=completed_excel,
                        failed_excel_path=failed_excel,
                        output_report_path=final_report,
                        execution_durations=self.execution_durations
                    )

                    self.reg_log_box.insert("end", f"[HUB] Final excel summary report written to:\n{final_report}\n")
                    self.row_completion_actions.pack(fill="x", pady=(10, 0))

                SecureVault.clear()
                self._update_stepper(5) # Set stepper to completed state

                if success:
                    messagebox.showinfo("Bulk Run Complete", "Portal APoB registrations batch completed successfully!")
                else:
                    messagebox.showerror("Bulk Run Interrupt", f"Portal execution interrupted:\n{err_msg}")

        self.after(100, self._drain_ui_queue)

    def _on_theme_changed(self, value: str) -> None:
        """Dynamically switches application appearance mode."""
        ctk.set_appearance_mode(value.lower())

    def _toggle_password_visibility(self) -> None:
        """Toggles show/hide state of the portal password input."""
        current_show = self.entry_gst_pass.cget("show")
        if current_show == "*":
            self.entry_gst_pass.configure(show="")
            self.btn_show_pass.configure(text="🔒")
        else:
            self.entry_gst_pass.configure(show="*")
            self.btn_show_pass.configure(text="👁")

    def _show_captcha_popup(self, img_path: str, portal_hash: str = "N/A") -> None:
        """Displays a stylish GUI modal popup prompting the operator to manually solve the CAPTCHA."""
        from PIL import Image
        import hashlib
        import time
        
        # Check if the image exists
        if not os.path.exists(img_path):
            messagebox.showerror("CAPTCHA Error", f"CAPTCHA image file not found at: {img_path}")
            return
            
        # Compute and compare hashes, retrying in case of filesystem write lag
        popup_hash = "N/A"
        for attempt in range(3):
            try:
                if os.path.exists(img_path):
                    with open(img_path, "rb") as f:
                        popup_hash = hashlib.sha256(f.read()).hexdigest()
                    if portal_hash != "N/A" and popup_hash == portal_hash:
                        break
            except Exception:
                pass
            time.sleep(0.2)

        match_status = "MATCHED" if (portal_hash == "N/A" or popup_hash == portal_hash) else "MISMATCHED"

        # Log details in exact required format
        log_block = (
            f"\n[CAPTCHA]\n"
            f"Manual Mode Activated\n\n"
            f"Portal Captcha Hash:\n{portal_hash}\n\n"
            f"Popup Captcha Hash:\n{popup_hash}\n\n"
            f"Status:\n{match_status}\n"
        )
        print(log_block)
        try:
            import logging
            logging.getLogger("GST-Amendment.Hub").info(log_block)
        except Exception:
            pass
            
        # Create a Toplevel popup window
        popup = ctk.CTkToplevel(self)
        popup.title("Manual CAPTCHA Verification")
        
        # Calculate responsive size based on screen dimensions
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        popup_w = int(screen_width * 0.3)
        popup_h = int(screen_height * 0.4)
        
        # Clamp dimensions
        if popup_w < 400: popup_w = 400
        if popup_w > 550: popup_w = 550
        if popup_h < 320: popup_h = 320
        if popup_h > 420: popup_h = 420
        
        popup.geometry(f"{popup_w}x{popup_h}")
        popup.resizable(False, False)
        popup.configure(fg_color=self.C_MAIN_BG)
        popup.grab_set() # Make it modal
        
        # Center the popup relative to self
        popup.update_idletasks()
        x = self.winfo_x() + (self.winfo_width() // 2) - (popup.winfo_width() // 2)
        y = self.winfo_y() + (self.winfo_height() // 2) - (popup.winfo_height() // 2)
        popup.geometry(f"+{x}+{y}")
        
        # Title Label
        ctk.CTkLabel(
            popup,
            text="🔐 GST CAPTCHA Verification Required",
            font=("Segoe UI", 16, "bold"),
            text_color=self.C_TEXT_PRIMARY
        ).pack(pady=(20, 10))
        
        ctk.CTkLabel(
            popup,
            text="Headless auto-solve failed. Please enter the characters shown below:",
            font=("Segoe UI", 11),
            text_color=self.C_TEXT_SECONDARY
        ).pack(pady=(0, 15))
        
        # Load and display CAPTCHA Image
        try:
            pil_img = Image.open(img_path)
            orig_w, orig_h = pil_img.size
            
            # Scale image dynamically based on popup width
            target_w = int(popup_w * 0.5)
            scale = target_w / orig_w
            target_h = int(orig_h * scale)
            
            # Clamp image height to 25% of popup height
            if target_h > int(popup_h * 0.25):
                target_h = int(popup_h * 0.25)
                target_w = int(target_h * (orig_w / orig_h))
                
            ctk_img = ctk.CTkImage(light_image=pil_img, dark_image=pil_img, size=(target_w, target_h))
            img_label = ctk.CTkLabel(popup, image=ctk_img, text="")
            img_label.pack(pady=5)
        except Exception as e:
            ctk.CTkLabel(popup, text="[Error loading image]", text_color=self.C_DANGER).pack(pady=5)
            
        # Text Entry for CAPTCHA characters
        entry_captcha = ctk.CTkEntry(
            popup,
            placeholder_text="Enter CAPTCHA here...",
            width=200,
            height=36,
            font=("Segoe UI", 14, "bold"),
            justify="center",
            border_color=self.C_BORDER,
            fg_color=self.C_SIDEBAR_BG,
            text_color=self.C_TEXT_PRIMARY
        )
        entry_captcha.pack(pady=15)
        entry_captcha.focus_set()
        
        def on_submit(event=None):
            captcha_val = entry_captcha.get().strip().upper()
            if not captcha_val or len(captcha_val) != 6:
                messagebox.showerror("Verification Failed", "CAPTCHA must be exactly 6 characters.", parent=popup)
                return
            # Submit to automation engine
            if self.registration_adapter:
                self.registration_adapter.submit_captcha(captcha_val)
            popup.destroy()
            
        # Bind Return key to submit
        entry_captcha.bind("<Return>", on_submit)
        
        # Submit Button
        btn_submit = ctk.CTkButton(
            popup,
            text="Verify & Resume",
            fg_color=self.C_ACCENT,
            hover_color=self.C_ACCENT_HOVER,
            text_color="white",
            height=36,
            width=150,
            font=("Segoe UI", 12, "bold"),
            command=on_submit
        )
        btn_submit.pack(pady=(5, 20))

if __name__ == "__main__":
    import sys
    # PyInstaller Subprocess Redirection Layer:
    # If the packaged EXE is spawned by the adapter with arguments,
    # bypass the GUI entirely and run the automation engine / validator directly.
    if len(sys.argv) > 1 and ("--excel" in sys.argv or "--username" in sys.argv):
        try:
            from amend_req import execute_bulk_apob_pipeline
            execute_bulk_apob_pipeline()
        except Exception as e:
            print(f"Subprocess Execution Error: {e}", file=sys.stderr)
            sys.exit(1)
        sys.exit(0)
    
    elif len(sys.argv) > 1 and ("--folder" in sys.argv or "--client-output-dir" in sys.argv):
        try:
            from gst_validator_v2 import main as validator_main
            validator_main()
        except Exception as e:
            print(f"Subprocess Validation Error: {e}", file=sys.stderr)
            sys.exit(1)
        sys.exit(0)

    app = GSTSmartHubApp()
    app.mainloop()
