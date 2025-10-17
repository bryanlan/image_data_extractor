# -*- coding: utf-8 -*-
"""
Windows app for image->text extraction and summarization using local Ollama.
- Mode A: Multi-file picker. Sort by ctime, extract per image with extraction model + prompt,
          then prepend a single summary using summary model + prompt. Save as Word-compatible HTML.
- Mode B (updated): Global hotkey (default: ctrl+alt+p) or "Snap Now" captures the ACTIVE WINDOW ONCE.
          Visual indication on snap. No background polling.
- Deferred processing (default ON): queue frames; only process + summarize on "Complete & Save…".
  If OFF: extract immediately per snap; summary still waits for "Complete & Save…".
- Stores prompts, selected models, hotkey, hotkey-enabled, defer toggle to config.json in the install directory.

Dependencies:
  pip install ollama PySide6 pillow keyboard pywin32

Notes:
  * Requires Ollama daemon to be running locally (http://localhost:11434). We do NOT start/stop it.
  * No OCR fallback; strictly VLM extraction as requested.
  * If you pick a non-vision model for extraction, it will error (as requested).
  * If context length is exceeded, request fails (as requested).
  * Windows 11 only target is fine.
"""

import os
import sys
import json
import base64
import traceback
from pathlib import Path
from datetime import datetime

# UI
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QComboBox, QTextEdit, QSpinBox,
    QLineEdit, QMessageBox, QSystemTrayIcon, QMenu, QStyle, QCheckBox
)
from PySide6.QtGui import QAction

# Imaging
from PIL import ImageGrab

# Global hotkey
import keyboard  # global hook (Windows); may require admin depending on system

# Active window capture (Windows)
import win32gui

# Ollama
import ollama
from ollama import Client, ResponseError


APP_NAME = "Ollama Image Transcriber"

# Word-friendly HTML output by default
DEFAULT_EXTRACTION_PROMPT = """You are producing a clean, Word-compatible HTML fragment for a single image.
Requirements:
- Do NOT include <html>, <head>, or <body>. Output only the inner HTML fragment.
- Structure the content using semantic HTML: <h2>, <h3>, <p>, <ul>, <li>, <table>, <thead>, <tbody>, <tr>, <th>, <td>, <code>, <pre>.
- Start with <h3>“Facts”</h3> followed by a <ul> of all key facts you can infer from the image.
- If there are numbers, dates, currencies, years, or measurements and you are confident you understand them include a <h3>“Data”</h3> section with a small <table> (two columns: Label, Value). 
- If applicable, Add <h3>“Narrative”</h3> with a short paragraph summarizing what’s in the image.

"""
DEFAULT_SUMMARY_PROMPT = """You are producing a clean, Word-compatible HTML fragment that is an executive summary of various data.
Requirements:
- Do NOT include <html>, <head>, or <body>. Output only the inner HTML fragment.
- Start with <h1>Executive Summary</h1>
- Provide a short <p> overview, then a <ul> of crisp bullet points (3–8 bullets).
- If there are recurring entities, dates, or numeric themes across images, include a compact <h2>Highlights</h2> section with a small <table> (three columns: Theme, Evidence, Notes).
- No markdown; HTML only. Keep it concise and executive-ready.
"""

DEFAULT_HOTKEY = "ctrl+alt+p"
CONFIG_FILE = Path(__file__).resolve().parent / "config.json"
IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff")
OLLAMA_HOST = "http://localhost:11434"  # local daemon only

# --- HTML helpers for Word-compatible output ---
HTML_CSS = """
<style>
  body { font-family: Segoe UI, Arial, sans-serif; line-height: 1.35; font-size: 11pt; color: #222; }
  h1, h2, h3 { margin: 0.4em 0 0.2em; }
  h1 { font-size: 20pt; }
  h2 { font-size: 16pt; border-bottom: 1px solid #ddd; padding-bottom: 4px; }
  h3 { font-size: 13pt; color: #333; }
  .meta { color: #666; font-size: 9pt; margin: 0.2em 0 0.8em; }
  ul { margin: 0.2em 0 0.8em 1.2em; }
  table { border-collapse: collapse; margin: 0.3em 0 0.8em; }
  th, td { border: 1px solid #ddd; padding: 6px 8px; }
  th { background: #f7f7f7; }
  pre { background: #fbfbfb; border: 1px solid #eee; padding: 8px; overflow-x: auto; }
  .section { margin: 1.0em 0; }
</style>
"""

def build_html_doc(title: str, summary_html: str, sections_html: list[str]) -> str:
    body = []
    if summary_html:
        body.append(summary_html)
    if sections_html:
        body.append("\n".join(sections_html))
    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>{title}</title>
{HTML_CSS}
</head>
<body>
{''.join(body)}
</body>
</html>"""


def load_config():
    cfg = {
        "extraction_prompt": DEFAULT_EXTRACTION_PROMPT,
        "summary_prompt": DEFAULT_SUMMARY_PROMPT,
        "extraction_model": "",
        "summary_model": "",
        "hotkey": DEFAULT_HOTKEY,
        "hotkey_enabled": True,
        "defer_processing": True,
        "interim_save_folder": str(Path.home()),  # Default to user's home directory
        "include_summary": True,
    }
    if CONFIG_FILE.exists():
        try:
            cfg.update(json.loads(CONFIG_FILE.read_text(encoding="utf-8")))
        except Exception:
            # Ignore corrupt config, recreate later
            pass
    return cfg


def save_config(cfg: dict):
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception:
        pass


def human_ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def get_active_window_bbox():
    # Returns (left, top, right, bottom) for the foreground window
    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        raise RuntimeError("No active window handle.")
    rect = win32gui.GetWindowRect(hwnd)
    return rect, hwnd


class OllamaClientWrapper:
    def __init__(self, host=OLLAMA_HOST):
        self.client = Client(host=host)  # explicit host

    def ensure_daemon(self):
        try:
            _ = ollama.list()  # health check
        except Exception as e:
            raise RuntimeError(
                "Cannot connect to the Ollama daemon at http://localhost:11434. "
                "Start Ollama first and try again."
            ) from e

    def list_models(self):
        resp = ollama.list()
        names = []
        for m in resp.get("models", []):
            n = m.get("name") or m.get("model")
            if n:
                names.append(n)
        return sorted(names, key=str.lower)

    def extract_from_image_path(self, model: str, user_prompt: str, image_path: str, temperature: float = 0.0) -> str:
        with open(image_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode("utf-8")
        return self._extract_from_b64(model, user_prompt, b64, temperature)

    def extract_from_image_bytes(self, model: str, user_prompt: str, image_bytes: bytes, temperature: float = 0.0) -> str:
        b64 = base64.b64encode(image_bytes).decode("utf-8")
        return self._extract_from_b64(model, user_prompt, b64, temperature)

    def _extract_from_b64(self, model: str, user_prompt: str, b64: str, temperature: float) -> str:
        messages = [{
            "role": "user",
            "content": user_prompt,
            "images": [b64],
        }]
        resp = self.client.chat(model=model, messages=messages, options={"temperature": temperature}, stream=False)
        return resp["message"]["content"]

    def summarize_text(self, model: str, summary_prompt: str, text: str, temperature: float = 0.0) -> str:
        content = f"{summary_prompt}\n\n=== CONTENT START ===\n{text}\n=== CONTENT END ==="
        messages = [{"role": "user", "content": content}]
        resp = self.client.chat(model=model, messages=messages, options={"temperature": temperature}, stream=False)
        return resp["message"]["content"]


class HotkeyListener(QThread):
    pressed = Signal()  # emitted on hotkey press

    def __init__(self, hotkey: str, parent=None):
        super().__init__(parent)
        self.hotkey = hotkey
        self._hook_id = None

    def run(self):
        try:
            self._hook_id = keyboard.add_hotkey(self.hotkey, lambda: self.pressed.emit())
            keyboard.wait()  # blocks in this thread
        except Exception:
            pass

    def rebind(self, hotkey: str):
        try:
            if self._hook_id is not None:
                keyboard.remove_hotkey(self._hook_id)
        except Exception:
            pass
        self.hotkey = hotkey
        try:
            self._hook_id = keyboard.add_hotkey(self.hotkey, lambda: self.pressed.emit())
        except Exception:
            pass

    def stop(self):
        try:
            if self._hook_id is not None:
                keyboard.remove_hotkey(self._hook_id)
        except Exception:
            pass


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.setMinimumWidth(760)

        self.cfg = load_config()
        self.ollama = OllamaClientWrapper()

        # Fail fast if daemon is not running
        try:
            self.ollama.ensure_daemon()
        except Exception as e:
            QMessageBox.critical(self, "Ollama not running", str(e))
            sys.exit(2)

        # ---------- State for Mode B ----------
        # If defer_processing=True: store raw PNG bytes to process upon "Complete".
        # If defer_processing=False: process immediately and store HTML sections here.
        self.deferred_queue: list[tuple[str, bytes]] = []       # (timestamp, png_bytes)
        self.processed_sections_html: list[str] = []            # HTML fragments already extracted

        # ---------- UI ----------
        self.model_extract = QComboBox()
        self.model_summary = QComboBox()
        self.btn_refresh_models = QPushButton("Refresh models")
        self.btn_refresh_models.clicked.connect(self.refresh_models)

        self.prompt_extract = QTextEdit()
        self.prompt_extract.setPlainText(self.cfg.get("extraction_prompt", DEFAULT_EXTRACTION_PROMPT))
        self.btn_save_extract_prompt = QPushButton("Save Extraction Prompt")
        self.btn_save_extract_prompt.clicked.connect(self.save_extraction_prompt)
        self.btn_restore_extract_prompt = QPushButton("Restore Default Prompt")
        self.btn_restore_extract_prompt.clicked.connect(self.restore_default_extraction_prompt)

        self.prompt_summary = QTextEdit()
        self.prompt_summary.setPlainText(self.cfg.get("summary_prompt", DEFAULT_SUMMARY_PROMPT))
        self.btn_save_summary_prompt = QPushButton("Save Summary Prompt")
        self.btn_save_summary_prompt.clicked.connect(self.save_summary_prompt)
        self.btn_restore_summary_prompt = QPushButton("Restore Default Prompt")
        self.btn_restore_summary_prompt.clicked.connect(self.restore_default_summary_prompt)

        # New toggles
        self.defer_checkbox = QCheckBox("Defer processing until Complete")
        self.defer_checkbox.setChecked(bool(self.cfg.get("defer_processing", True)))
        self.defer_checkbox.stateChanged.connect(self.on_defer_changed)

        self.hotkey_enabled_checkbox = QCheckBox("Hotkey enabled")
        self.hotkey_enabled_checkbox.setChecked(bool(self.cfg.get("hotkey_enabled", True)))
        self.hotkey_enabled_checkbox.stateChanged.connect(self.on_hotkey_enabled_changed)

        self.include_summary_checkbox = QCheckBox("Include Summary")
        self.include_summary_checkbox.setChecked(bool(self.cfg.get("include_summary", True)))
        self.include_summary_checkbox.stateChanged.connect(self.on_include_summary_changed)

        self.hotkey_edit = QLineEdit(self.cfg.get("hotkey", DEFAULT_HOTKEY))
        self.hotkey_edit.setMaximumWidth(120)
        self.btn_rebind_hotkey = QPushButton("Bind hotkey")
        self.btn_rebind_hotkey.clicked.connect(self.rebind_hotkey)

        # Actions
        self.btn_mode_a = QPushButton("Process Saved Images")
        self.btn_mode_a.clicked.connect(self.process_images_mode)

        self.btn_complete = QPushButton("Complete & Save…")
        self.btn_complete.setVisible(False)  # only appears when there is at least one frame
        self.btn_complete.clicked.connect(self.complete_and_save)

        self.btn_set_interim_folder = QPushButton("Set Interim Save Folder")
        self.btn_set_interim_folder.clicked.connect(self.set_interim_save_folder)

        # Status labels
        self.status_label = QLabel("Ready.")
        self.status_label.setWordWrap(True)
        self.queue_label = QLabel("No outstanding frames to process")
        self.queue_label.setStyleSheet("color:#555;")
        self.flash_label = QLabel("")  # transient visual ping on snap
        self.flash_label.setStyleSheet("padding:4px 8px; color:#fff; background:#2e7d32; border-radius:6px;")
        self.flash_label.setVisible(False)

        # Layout
        root = QWidget()
        v = QVBoxLayout(root)

        # Models row
        mrow = QHBoxLayout()
        mrow.addWidget(QLabel("Extraction model:"))
        mrow.addWidget(self.model_extract, 1)
        mrow.addWidget(QLabel("Summary model:"))
        mrow.addWidget(self.model_summary, 1)
        mrow.addWidget(self.btn_refresh_models)
        v.addLayout(mrow)

        # Prompts
        h_extract = QHBoxLayout()
        h_extract.addWidget(self.prompt_extract, 1)
        h_extract.addWidget(self.btn_save_extract_prompt)
        h_extract.addWidget(self.btn_restore_extract_prompt)
        v.addWidget(QLabel("Extraction prompt:"))
        v.addLayout(h_extract)

        h_summary = QHBoxLayout()
        h_summary.addWidget(self.prompt_summary, 1)
        h_summary.addWidget(self.btn_save_summary_prompt)
        h_summary.addWidget(self.btn_restore_summary_prompt)
        v.addWidget(QLabel("Summary prompt:"))
        v.addLayout(h_summary)

        # Controls row: deferral + hotkey enable + edit/rebind
        crow = QHBoxLayout()
        crow.addWidget(self.defer_checkbox)
        crow.addSpacing(16)
        crow.addWidget(self.hotkey_enabled_checkbox)
        crow.addSpacing(16)
        crow.addWidget(QLabel("Global hotkey:"))
        crow.addWidget(self.hotkey_edit)
        crow.addWidget(self.btn_rebind_hotkey)
        crow.addWidget(self.include_summary_checkbox)
        v.addLayout(crow)

        # Actions row
        arow = QHBoxLayout()
        arow.addWidget(self.btn_mode_a)
        arow.addWidget(self.btn_set_interim_folder)
        arow.addWidget(self.btn_complete)
        v.addLayout(arow)

        v.addWidget(self.queue_label)
        v.addWidget(self.status_label)
        v.addWidget(self.flash_label, alignment=Qt.AlignLeft)

        self.setCentralWidget(root)
        self.add_discard_button()

        # Tray icon
        self.tray = QSystemTrayIcon(self)
        icon = self.style().standardIcon(QStyle.SP_ComputerIcon)
        self.tray.setIcon(icon)
        self.tray.setToolTip(APP_NAME)
        tray_menu = QMenu()
        act_show = QAction("Show", self)
        act_show.triggered.connect(lambda: (self.showNormal(), self.raise_(), self.activateWindow()))
        act_snap = QAction("Snap Frame (active window)", self)
        act_snap.triggered.connect(self.on_hotkey_pressed)
        act_complete = QAction("Complete & Save…", self)
        act_complete.triggered.connect(self.complete_and_save)
        act_quit = QAction("Quit", self)
        act_quit.triggered.connect(QApplication.instance().quit)
        tray_menu.addAction(act_show)
        tray_menu.addSeparator()
        tray_menu.addAction(act_snap)
        tray_menu.addAction(act_complete)
        tray_menu.addSeparator()
        tray_menu.addAction(act_quit)
        self.tray.setContextMenu(tray_menu)
        self.tray.show()

        # Hotkey listener (optional, based on toggle)
        self.hotkey_listener: HotkeyListener | None = None
        if self.hotkey_enabled_checkbox.isChecked():
            self.start_hotkey_listener()

        # Populate models
        self.refresh_models()

        # Apply last used models if they exist
        ex = self.cfg.get("extraction_model", "")
        sm = self.cfg.get("summary_model", "")
        if ex:
            idx = self.model_extract.findText(ex)
            if idx >= 0: self.model_extract.setCurrentIndex(idx)
        if sm:
            idx = self.model_summary.findText(sm)
            if idx >= 0: self.model_summary.setCurrentIndex(idx)

        self.status_update(f"Ready. (Hotkey: {self.cfg.get('hotkey', DEFAULT_HOTKEY)})")
        self.update_queue_state()

    # --------------------------- Utility & UI helpers ---------------------------

    def status_update(self, msg: str):
        self.status_label.setText(msg)
        self.tray.setToolTip(f"{APP_NAME}\n{msg}")

    def flash_ping(self, text="Captured frame"):
        self.flash_label.setText(text)
        self.flash_label.setVisible(True)
        self.tray.showMessage(APP_NAME, text, QSystemTrayIcon.Information, 1200)
        QTimer.singleShot(900, lambda: self.flash_label.setVisible(False))

    def save_current_config(self):
        self.cfg["extraction_prompt"] = self.prompt_extract.toPlainText().strip() or DEFAULT_EXTRACTION_PROMPT
        self.cfg["summary_prompt"] = self.prompt_summary.toPlainText().strip() or DEFAULT_SUMMARY_PROMPT
        self.cfg["hotkey"] = self.hotkey_edit.text().strip() or DEFAULT_HOTKEY
        self.cfg["hotkey_enabled"] = self.hotkey_enabled_checkbox.isChecked()
        self.cfg["defer_processing"] = self.defer_checkbox.isChecked()
        self.cfg["extraction_model"] = self.model_extract.currentText().strip()
        self.cfg["summary_model"] = self.model_summary.currentText().strip()
        self.cfg["include_summary"] = self.include_summary_checkbox.isChecked()
        save_config(self.cfg)

    def update_queue_state(self):
        if self.defer_checkbox.isChecked():
            n = len(self.deferred_queue)
        else:
            n = len(self.processed_sections_html)
        if n <= 0:
            self.queue_label.setText("No outstanding frames to process")
            self.btn_complete.setVisible(False)
            self.btn_discard.setVisible(False)
        else:
            self.queue_label.setText(f"Frames ready for completion ({n})")
            self.btn_complete.setVisible(True)
            self.btn_discard.setVisible(True)

    # --------------------------- Models ---------------------------

    def refresh_models(self):
        try:
            models = self.ollama.list_models()
        except Exception as e:
            QMessageBox.critical(self, "Ollama error", f"Failed to list models.\n\n{e}")
            return
        old_extract = self.model_extract.currentText()
        old_summary = self.model_summary.currentText()
        self.model_extract.clear()
        self.model_summary.clear()
        self.model_extract.addItems(models)
        self.model_summary.addItems(models)
        if old_extract:
            i = self.model_extract.findText(old_extract)
            if i >= 0: self.model_extract.setCurrentIndex(i)
        if old_summary:
            i = self.model_summary.findText(old_summary)
            if i >= 0: self.model_summary.setCurrentIndex(i)

    # --------------------------- Mode A ---------------------------

    def process_images_mode(self):
        self.save_current_config()
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Select image files",
            "",
            "Images (*.png *.jpg *.jpeg *.webp *.bmp *.tif *.tiff);;All Files (*.*)"
        )
        if not files:
            return

        files = [f for f in files if Path(f).suffix.lower() in IMAGE_EXTS]
        if not files:
            QMessageBox.warning(self, "No images", "No supported image files selected.")
            return

        files.sort(key=lambda p: os.path.getctime(p))

        extract_model = self.model_extract.currentText().strip()
        summary_model = self.model_summary.currentText().strip()
        if not extract_model or not summary_model:
            QMessageBox.warning(self, "Models required", "Pick both extraction and summary models.")
            return

        extraction_prompt = self.prompt_extract.toPlainText().strip() or DEFAULT_EXTRACTION_PROMPT
        summary_prompt = self.prompt_summary.toPlainText().strip() or DEFAULT_SUMMARY_PROMPT

        self.status_update(f"Processing {len(files)} images…")
        QApplication.processEvents()

        sections = []
        all_extractions = []
        try:
            for p in files:
                try:
                    fragment = self.ollama.extract_from_image_path(extract_model, extraction_prompt, p, temperature=0.0).strip()
                except Exception as e:
                    raise
                fname = Path(p).name
                cts = datetime.fromtimestamp(os.path.getctime(p))
                section_html = f"""
<section class=\"section">
  <h2>{fname}</h2>
  <p class=\"meta\">Created: {cts}</p>
  {fragment}
</section>
"""
                sections.append(section_html)
                all_extractions.append(fragment)
                self.status_update(f"Processed: {fname}")
                QApplication.processEvents()

            summary_html = ""
            if self.include_summary_checkbox.isChecked():
                summary_input = "\n\n".join(all_extractions)
                summary_html = self.ollama.summarize_text(summary_model, summary_prompt, summary_input, temperature=0.0).strip()
                if not summary_html or len(summary_html) < 100:
                    summary_html = '<div class="section"><h1>Executive Summary</h1><p>LLM did not return any summary data.</p></div>'
                elif not (summary_html.startswith('<') and ('<h1' in summary_html or '<section' in summary_html or '<div' in summary_html)):
                    summary_html = f'<div class="section"><h1>Executive Summary</h1><p>{summary_html}</p></div>'
            final_html = build_html_doc("Image Extractions", summary_html, sections)
            self.save_word_html_via_dialog(final_html)

        except ResponseError as re:
            QMessageBox.critical(self, "Ollama error", f"{re}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Processing failed:\n{e}\n\n{traceback.format_exc()}")
        finally:
            self.status_update("Ready.")

    def save_word_html_via_dialog(self, html_text: str):
        default_name = datetime.now().strftime("%Y%m%d_%H%M_run-summary.doc")
        out_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save Word-compatible HTML",
            default_name,
            "Word-compatible HTML (*.doc *.html *.htm);;All Files (*.*)"
        )
        if out_path:
            Path(out_path).write_text(html_text, encoding="utf-8")
            self.status_update(f"Saved: {out_path}")
        else:
            self.status_update("Save canceled.")

    # --------------------------- Mode B (updated: snap once per hotkey) ---------------------------

    def on_hotkey_pressed(self):
        """Capture active window once. Depending on 'defer', either queue raw image or extract now."""
        try:
            (l, t, r, b), hwnd = get_active_window_bbox()
            bbox = (l, t, r, b)
            img = ImageGrab.grab(bbox=bbox)
            from io import BytesIO
            buf = BytesIO()
            img.save(buf, format="PNG")
            png_bytes = buf.getvalue()
            when = human_ts()
        except Exception as e:
            QMessageBox.critical(self, "Capture error", f"Failed to capture active window:\n{e}")
            return

        self.flash_ping("Captured frame")

        extract_model = self.model_extract.currentText().strip()
        if not extract_model:
            QMessageBox.warning(self, "Extraction model needed", "Pick an extraction model first.")
            return

        extraction_prompt = self.prompt_extract.toPlainText().strip() or DEFAULT_EXTRACTION_PROMPT

        if self.defer_checkbox.isChecked():
            # Queue image for later processing
            self.deferred_queue.append((when, png_bytes))
            self.status_update(f"Queued frame at {when} (deferred)")
        else:
            # Process immediately; store the section HTML for later summary
            try:
                fragment = self.ollama.extract_from_image_bytes(extract_model, extraction_prompt, png_bytes, temperature=0.0).strip()
                section_html = f"""
<section class="section">
  <h2>Active window</h2>
  <p class="meta">Captured: {when}</p>
  {fragment}
</section>
"""
                self.processed_sections_html.append(section_html)
                self.status_update(f"Processed frame at {when} (immediate)")
            except ResponseError as re:
                QMessageBox.critical(self, "Ollama error", f"{re}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Extraction failed:\n{e}\n\n{traceback.format_exc()}")

        self.update_queue_state()

    def complete_and_save(self):
        """Process any queued frames (if deferred), then summarize ALL processed sections and save."""
        summary_model = self.model_summary.currentText().strip()
        if not summary_model:
            QMessageBox.warning(self, "Summary model needed", "Pick a summary model to finalize.")
            return

        summary_prompt = self.prompt_summary.toPlainText().strip() or DEFAULT_SUMMARY_PROMPT
        extract_model = self.model_extract.currentText().strip()
        extraction_prompt = self.prompt_extract.toPlainText().strip() or DEFAULT_EXTRACTION_PROMPT

        # If deferred, process queued images now
        if self.defer_checkbox.isChecked() and self.deferred_queue:
            self.status_update(f"Processing {len(self.deferred_queue)} queued frames…")
            QApplication.processEvents()
            try:
                for when, png_bytes in self.deferred_queue:
                    fragment = self.ollama.extract_from_image_bytes(extract_model, extraction_prompt, png_bytes, temperature=0.0).strip()
                    section_html = f"""
<section class="section">
  <h2>Active window</h2>
  <p class="meta">Captured: {when}</p>
  {fragment}
</section>
"""
                    self.processed_sections_html.append(section_html)
                # Clear the queue once processed
                self.deferred_queue.clear()
            except ResponseError as re:
                QMessageBox.critical(self, "Ollama error", f"{re}")
                return
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Processing queued frames failed:\n{e}\n\n{traceback.format_exc()}")
                return

        if not self.processed_sections_html:
            QMessageBox.information(self, "Nothing to do", "No frames have been processed yet.")
            self.update_queue_state()
            return

        try:
            joined = "\n".join(self.processed_sections_html)
            summary_html = self.ollama.summarize_text(summary_model, summary_prompt, joined, temperature=0.0).strip()
            final_html = build_html_doc("Active Window Extractions", summary_html, self.processed_sections_html)
            self.save_word_html_via_dialog(final_html)
        except ResponseError as re:
            QMessageBox.critical(self, "Ollama error", f"{re}")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Summarization failed:\n{e}\n\n{traceback.format_exc()}")
        finally:
            # Reset state after completion
            self.processed_sections_html.clear()
            self.deferred_queue.clear()
            self.update_queue_state()
            self.status_update("Ready.")

    # --------------------------- Toggles & bindings ---------------------------

    def rebind_hotkey(self):
        from PySide6.QtWidgets import QInputDialog
        hk, ok = QInputDialog.getText(self, "Bind Hotkey", "Enter new hotkey (e.g. ctrl+alt+p) or press Esc to cancel:", text=self.hotkey_edit.text())
        if ok and hk.strip():
            self.hotkey_edit.setText(hk.strip())
            self.hotkey_listener.rebind(hk.strip())
            self.cfg["hotkey"] = hk.strip()
            save_config(self.cfg)
            self.status_update(f"Global hotkey bound to: {hk.strip()}")
        else:
            self.status_update("Hotkey binding canceled.")

    def start_hotkey_listener(self):
        try:
            if self.hotkey_listener:
                self.hotkey_listener.stop()
            self.hotkey_listener = HotkeyListener(self.cfg.get("hotkey", DEFAULT_HOTKEY))
            self.hotkey_listener.pressed.connect(self.on_hotkey_pressed)
            self.hotkey_listener.start()
            self.status_update(f"Hotkey enabled ({self.cfg.get('hotkey', DEFAULT_HOTKEY)})")
        except Exception:
            # non-fatal; user can still click "Snap Now"
            self.status_update("Hotkey could not be enabled (permissions?).")

    def stop_hotkey_listener(self):
        try:
            if self.hotkey_listener:
                self.hotkey_listener.stop()
                self.hotkey_listener = None
        except Exception:
            pass
        self.status_update("Hotkey disabled")

    def on_hotkey_enabled_changed(self, state):
        enabled = state == Qt.Checked
        if enabled:
            self.start_hotkey_listener()
        else:
            self.stop_hotkey_listener()
        self.save_current_config()

    def on_defer_changed(self, state):
        # Just persist. Queue semantics handled at snap/complete time.
        self.save_current_config()
        self.update_queue_state()

    def on_include_summary_changed(self, state):
        self.cfg["include_summary"] = bool(state == Qt.Checked)
        save_config(self.cfg)

    # --------------------------- Lifecycle ---------------------------

    def closeEvent(self, event):
        self.save_current_config()
        try:
            self.stop_hotkey_listener()
        except Exception:
            pass
        return super().closeEvent(event)

    def save_extraction_prompt(self):
        """Persist the current extraction prompt to config.json and update status."""
        self.cfg["extraction_prompt"] = self.prompt_extract.toPlainText().strip() or DEFAULT_EXTRACTION_PROMPT
        save_config(self.cfg)
        self.status_update("Extraction prompt saved.")

    def save_summary_prompt(self):
        """Persist the current summary prompt to config.json and update status."""
        self.cfg["summary_prompt"] = self.prompt_summary.toPlainText().strip() or DEFAULT_SUMMARY_PROMPT
        save_config(self.cfg)
        self.status_update("Summary prompt saved.")

    def restore_default_extraction_prompt(self):
        self.prompt_extract.setPlainText(DEFAULT_EXTRACTION_PROMPT)
        self.cfg["extraction_prompt"] = DEFAULT_EXTRACTION_PROMPT
        save_config(self.cfg)
        self.status_update("Extraction prompt restored to default.")

    def restore_default_summary_prompt(self):
        self.prompt_summary.setPlainText(DEFAULT_SUMMARY_PROMPT)
        self.cfg["summary_prompt"] = DEFAULT_SUMMARY_PROMPT
        save_config(self.cfg)
        self.status_update("Summary prompt restored to default.")

    def add_discard_button(self):
        self.btn_discard = QPushButton("Discard Selected Frame")
        self.btn_discard.setVisible(False)
        self.btn_discard.clicked.connect(self.discard_selected_frame)
        # Add to layout after queue_label
        self.centralWidget().layout().insertWidget(self.centralWidget().layout().indexOf(self.queue_label) + 1, self.btn_discard)

    def discard_selected_frame(self):
        # For demo: remove last frame in deferred_queue
        if self.deferred_queue:
            self.deferred_queue.pop()
            self.update_queue_state()
            self.status_update("Last frame discarded.")
        else:
            self.status_update("No frames to discard.")

    def set_interim_save_folder(self):
        """Set the folder where interim HTML files will be saved."""
        folder = QFileDialog.getExistingDirectory(self, "Select Interim Save Folder", self.cfg.get("interim_save_folder", ""))
        if folder:
            self.cfg["interim_save_folder"] = folder
            save_config(self.cfg)
            self.status_update(f"Interim save folder set: {folder}")
        else:
            self.status_update("Interim save folder not changed.")

def main():
    app = QApplication(sys.argv)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
