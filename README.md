# Watermark Remover GUI

[![Build and Release](https://github.com/yhbstf/watermark-remover-gui/actions/workflows/release.yml/badge.svg)](https://github.com/yhbstf/watermark-remover-gui/actions/workflows/release.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A Tkinter desktop tool that removes watermarks from images, Word `.docx`, PowerPoint `.pptx`, Excel `.xlsx`, PDF, and video files using the [IOPaint](https://github.com/Sanster/IOPaint) LaMa inpainting model (and OpenCV inpainting for video).

UI auto-switches between English and 中文 based on your system locale (override with `WMR_LANG=en` or `WMR_LANG=zh`).

## Features

### Selection tools
- **Rectangle, brush, polygon, color-pick** selection modes (hotkeys `R` / `B` / `P` / `C`).
- **Live mask overlay** — selected pixels painted in translucent red.
- **Rectangle editing** — click to select, drag to move, `Delete` to remove.
- **Feathering slider** for softer mask edges.
- **Auto-detect** watermark candidates (MSER heuristic, corner-biased scoring).
- **OCR detection** of text watermarks via optional `pytesseract`.
- **Quick templates** for 豆包 AI / 抖音 / 快手 / 小红书 / B 站 / 微博 / Midjourney / Getty / Shutterstock, etc.
- **Save / load custom templates** as `.wmrtpl` JSON.

### Workflow
- **Multi-step undo / redo** (`Ctrl+Z` / `Ctrl+Y`, 20 steps). Covers selections too.
- **Hold Space** on the main canvas to compare with the previous step.
- **Zoom & pan** — wheel, `+` / `-` / `0`, Shift+drag, middle-drag.
- **Loupe** (`L`) — cursor-tracking 3× magnifier.
- **Drag-and-drop** files into the window (needs optional `tkinterdnd2`).
- **Recent files** menu, last directory remembered.
- **Config persisted** in `~/.watermark-remover.json` (language, model, window size, feather, brush, color tolerance, recents).

### Input formats
- Plain images (`.png`, `.jpg`, `.jpeg`, `.bmp`, `.webp`, `.tif`).
- Word `.docx`, PowerPoint `.pptx`, Excel `.xlsx` — thumbnail grid + large zoomable preview to pick which embedded image to edit.
- PDF `.pdf` — needs optional `pymupdf`.
- Video `.mp4` / `.avi` / `.mov` / `.mkv` / `.webm` — first frame as reference.

### Inpainting
- **Model switcher** — LaMa / MAT / FCF / LDM / Manga.
- **Batch: all doc images** — draw on one reference image, apply proportionally to every image in a `.docx` / `.pptx` / `.xlsx` / `.pdf`.
- **Batch: folder** — apply to every image in a directory.
- **Process video** — applies selection to every frame via OpenCV Telea inpainting (fast).
- **CLI mode** — `python watermark_remover_gui.py --cli --image in.png --mask mask.png --output out.png [--model lama]`.

## Download (click-to-run)

Pre-built binaries on the [Releases](https://github.com/yhbstf/watermark-remover-gui/releases) page:

- **Windows (x64)** — extract the zip, run `watermark-remover-gui.exe` inside the folder.
- **Linux (x64)** — extract the tar.gz, run `./watermark-remover-gui` inside the folder.
- **macOS (Apple Silicon)** — extract the zip, right-click `.app` → **Open** on first launch (macOS Gatekeeper warns about unsigned apps).

First click of **Remove Watermark** downloads LaMa weights (~200 MB) into the local cache. Subsequent runs are offline.

## Run from source

```bash
pip install -r requirements.txt
python watermark_remover_gui.py
```

Optional extras:

```bash
pip install pymupdf          # PDF support
pip install pytesseract      # OCR text watermark detection (also needs Tesseract binary)
pip install tkinterdnd2      # Drag-and-drop
```

Requires Python 3.9+.

## Shortcuts

| Key / Mouse | Action |
|---|---|
| Left-drag | Draw a region (rect mode) / paint (brush) / add vertex (polygon) / sample color (color) |
| `Enter` | Close polygon |
| Shift+drag | Pan |
| Middle-drag | Pan |
| Mouse wheel | Zoom at cursor |
| `R` / `B` / `P` / `C` | Switch to Rect / Brush / Polygon / Color mode |
| `L` | Toggle loupe |
| `Delete` | Remove selected rectangle |
| `+` / `=` / `-` | Zoom in / out |
| `0` | Fit to window |
| `Esc` | Clear all regions |
| `Ctrl+Z` / `Ctrl+Y` | Undo / redo |
| Hold `Space` | Show previous step (before/after compare) |

## Notes

- Source documents (`.docx` / `.pptx` / `.xlsx` / `.pdf`) are backed up as `<name>.backup` before write-back.
- Temp files live under `D:/temp/watermark_remover_*` on Windows, `~/.cache/watermark-remover/` on Linux/macOS. Edit the `TEMP_ROOT` constants in the script to relocate.
- For text watermarks drop the brush size and paint tightly; fat masks wipe out surrounding content.
- Auto-detect is a cheap heuristic — it may catch non-watermark text (subtitles, labels). Use it as a starting point, then trim with rect-edit or delete unwanted boxes.

## License

MIT — see [LICENSE](LICENSE).
