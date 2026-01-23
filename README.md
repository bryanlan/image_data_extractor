# Visual Data Extractor

A Windows desktop application for capturing presentation slides and extracting structured data using AI vision models. Features automatic slide change detection, smart cropping, and flexible AI endpoints (local Ollama or cloud Azure/OpenAI).

![Visual Data Extractor UI](app_visual.png)

## Key Features

### Session-Based Workflow

The app uses a two-step workflow for captured slides:

1. **Step 1: Create PPT** - Crop images and generate a PowerPoint presentation
2. **Step 2: Extract with AI** - Send images to AI for text extraction → HTML document

### Capture Methods

**Manual Capture (Hotkey)**
- Press `Ctrl+Alt+P` (configurable) to capture the active window
- Instant feedback with visual flash notification
- Frames saved to session folder

**Auto-Capture (Slide Detection)**
- Press `Ctrl+Alt+R` (configurable) to toggle auto-capture on/off
- Red "RECORDING" overlay appears when active
- Automatically detects slide changes using perceptual hashing
- Configurable sensitivity (1-50, default 15) and polling interval (500-10000ms, default 5000ms)
- Only captures when significant visual change detected

### Slide Review & Duplicate Removal

Before cropping, review all captured slides:
- Thumbnail grid shows all images
- Click to select, Ctrl+click for multi-select
- Delete duplicates before proceeding
- Files remain on disk (only removed from processing list)

### Smart Cropping

- Automatic Teams meeting detection (removes toolbars/panels)
- Manual adjustment with live preview
- Crop all four edges: top, bottom, left, right
- Three options: No Cropping / Apply to All / Adjust Each Image

### AI Endpoint Flexibility

**Local & Cloud Support**
- **Local**: Ollama models running on your machine
- **Cloud**: Azure AI, Azure OpenAI, or OpenAI-compatible endpoints
- Choose Local or Cloud independently for extraction and summary
- Test button to verify configuration

### Output Formats

- **PowerPoint**: Slides with cropped images
- **HTML Report**: Word-compatible document with:
  - Executive Summary (optional)
  - Per-image sections: Facts, Data tables, Narrative
  - Embedded CSS for clean rendering

## User Interface

### Capture Settings
- **Capture hotkey**: Manual single-frame capture (default: `Ctrl+Alt+P`)
- **Auto-capture hotkey**: Toggle slide detection (default: `Ctrl+Alt+R`)
- **Auto-capture checkbox**: Enable/disable with sensitivity and interval controls
- **Save folder**: Where captured images are stored

### Session Panel
- Shows capture count and workflow progress
- **Step 1 Card**: Create PPT (shows status, action button)
- **Step 2 Card**: Extract with AI (shows status, action button)
- **Discard**: Clear current session

### Collapsible Sections
- **Process Existing Files**: PPT from Folder, AI Extract from Files
- **AI Settings**: Model selection, prompts, options

## Requirements

- **Windows** (uses Windows-specific capture APIs)
- **Python 3.10+**
- **Local option**: Ollama daemon at `http://localhost:11434`
- **Cloud option**: Azure AI or OpenAI API credentials

## Installation

```powershell
git clone https://github.com/bryanlan/vizextractor.git
cd vizextractor
pip install -r requirements.txt
```

Required packages: `ollama`, `openai`, `PySide6`, `pillow`, `imagehash`, `keyboard`, `pywin32`, `numpy`, `python-pptx`

## Quick Start

### Using Auto-Capture for Presentations

1. **Start the app**: `python ollama_extractor.py`
2. **Configure save folder** (optional)
3. **Switch to your presentation** (PowerPoint, Teams, browser, etc.)
4. **Press `Ctrl+Alt+R`** to start auto-capture
   - Red "RECORDING" banner appears at top of screen
   - App detects slide changes automatically
5. **Navigate through slides** - app captures each new slide
6. **Press `Ctrl+Alt+R` again** to stop
7. **Click "Create PPT"** to review, crop, and generate PowerPoint
8. **Click "Extract"** to run AI extraction on cropped images

### Using Manual Capture

1. **Press `Ctrl+Alt+P`** anytime to capture current window
2. Repeat for each slide/screen you want
3. Process with Step 1 (PPT) and/or Step 2 (AI)

### Processing Existing Files

1. Expand **"Process Existing Files"** section
2. Click **"PPT from Folder"** or **"AI Extract from Files"**
3. Select images → Review thumbnails → Crop → Process

## Configuration

Settings persist to `config.json`:

```json
{
  "manual_capture_hotkey": "ctrl+alt+p",
  "auto_capture_hotkey": "ctrl+alt+r",
  "auto_capture_interval_sec": 5.0,
  "auto_capture_threshold": 15,
  "extraction_model": "Ollama (local): llava:13b",
  "summary_model": "Ollama (local): llama3:8b",
  "include_summary": true,
  "defer_processing": true
}
```

### Auto-Capture Tuning

- **Sensitivity (1-50)**: Lower = more sensitive to changes. Default 15 works well for most presentations.
- **Interval (500-10000ms)**: How often to check for changes. Default 5000ms (5 seconds) balances responsiveness with CPU usage.

## Troubleshooting

**Hotkeys not working**
- Run app as Administrator
- Check for conflicts with other apps
- Verify hotkeys in UI match your expectations

**Auto-capture too sensitive / not sensitive enough**
- Adjust sensitivity slider (lower = more captures)
- Check console output for distance values to tune threshold

**Auto-capture not detecting slides**
- Increase polling interval if slides change slowly
- Decrease sensitivity if small changes should trigger capture

**Cloud models not appearing**
- Click Refresh Models
- Verify endpoint configuration in Manage Endpoints
- Use Test button to verify extraction works

## Development

Built with:
- **PySide6**: Qt-based GUI
- **Ollama Python SDK**: Local model inference
- **OpenAI SDK**: Cloud endpoint compatibility
- **PIL/Pillow**: Image processing
- **imagehash**: Perceptual hashing for slide detection
- **python-pptx**: PowerPoint generation

## License

MIT License - See LICENSE file for details

## Contributing

Issues and pull requests welcome at https://github.com/bryanlan/vizextractor
