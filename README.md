Ollama Image Transcriber
=========================

A small Windows utility that captures screenshots (active window) and uses a local Ollama VLM to extract structured information from images and produce an executive summary. The app saves Word-compatible HTML (.doc/.html) so the output opens nicely in Microsoft Word.

Key features
------------
- Mode A: "Process Saved Images" — open a set of image files and create a styled HTML report with per-image extracted HTML fragments and an optional executive summary.
- Mode B (hotkey-driven): press the global hotkey (default: `ctrl+alt+p`) to capture the active window. Captured frames are queued and can be processed later. Toggle hotkey and defer behavior from the UI.
- Prompts: editable extraction and summary prompts with Save and Restore Default buttons. Prompts are stored in `config.json`.
- Interim save folder: set a folder where interim captures may be saved.
- Discard: discard captured frames before finalizing.
- Output: Word-compatible HTML documents (.doc/.html) styled with light CSS for good Word rendering.

Requirements
------------
- Windows (app uses active-window capture APIs)
- Python 3.10+
- Ollama daemon running locally (http://localhost:11434)

Python dependencies
-------------------
Install via pip:

```powershell
pip install -r requirements.txt
```

(If `requirements.txt` is not present, install these packages manually: `ollama`, `PySide6`, `pillow`, `imagehash`, `keyboard`, `pywin32`.)

How to run
----------
1. Make sure the Ollama daemon is running locally.
2. From the project folder, run:

```powershell
python ollama_extractor.py
```

3. The main window appears. Pick extraction and summary models from the dropdowns. Edit prompts or restore defaults. Configure hotkey and whether summaries are included.

Usage notes
-----------
- Hotkey behavior: when the hotkey is enabled, pressing the hotkey captures the active window and queues frames (if "Defer processing" is checked). Use "Complete & Save…" to summarize (if enabled) and save the final HTML. Use "Discard Selected Frame" to remove frames from the queue.
- Process Saved Images: choose image files (Mode A) to batch-process and save a single HTML report.
- Interim save folder: choose a folder to store interim captures if needed by your workflow.

Configuration (config.json)
---------------------------
The app persists settings to `config.json` in the same folder as the script. Keys include:
- `extraction_prompt`: the extraction prompt text
- `summary_prompt`: the summary prompt text
- `extraction_model`: last selected extraction model
- `summary_model`: last selected summary model
- `hotkey`: the global hotkey string
- `hotkey_enabled`: true/false
- `defer_processing`: true/false
- `interim_save_folder`: path string
- `include_summary`: true/false

Troubleshooting
---------------
- If the app exits immediately, ensure Python can import required packages and that Ollama daemon is reachable at `http://localhost:11434`.
- If hotkey binding requires elevated privileges on Windows, run the app as Administrator.
- If the summary seems missing, check the "Include Summary" checkbox and inspect prompts.

License & notes
---------------
This is a small utility intended for local use with a local Ollama installation. It produces HTML output for Word—if you need `.docx` output later, consider integrating `python-docx`.
