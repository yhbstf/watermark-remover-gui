# Watermark Remover GUI

[![Build and Release](https://github.com/yhbstf/watermark-remover-gui/actions/workflows/release.yml/badge.svg)](https://github.com/yhbstf/watermark-remover-gui/actions/workflows/release.yml)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)

A Tkinter desktop tool that removes watermarks from images — including images embedded in Word `.docx` files — using the [IOPaint](https://github.com/Sanster/IOPaint) LaMa inpainting model.

## Features

- Open a standalone image file, or pick an image directly from a Word `.docx` document (with thumbnail preview).
- **Multi-region selection**: drag as many boxes as you want before running; every box is filled into the mask in a single pass.
- **Zoom & pan**: mouse wheel zooms around the cursor, Shift+drag or middle-mouse drag pans. Frame small watermarks precisely on large images.
- **Batch across Word**: draw regions once, then apply the *same relative regions* to every image in the `.docx` in one click.
- Save the result back into the original Word document (a `.backup` is created automatically), or export PNG/JPG.

## Download (click-to-run)

Pre-built binaries are published on the [Releases](https://github.com/yhbstf/watermark-remover-gui/releases) page:

- **Windows (x64)** — `watermark-remover-gui-windows-x64.zip`: extract, run `watermark-remover-gui.exe` inside the folder.
- **Linux (x64)** — `watermark-remover-gui-linux-x64.tar.gz`: extract, run `./watermark-remover-gui` inside the folder.
- **macOS (Apple Silicon)** — `watermark-remover-gui-macos-arm64.zip`: extract, right-click `watermark-remover-gui.app` → **Open** (first launch only — macOS Gatekeeper warns about unsigned apps).

The first time you click **Remove Watermark**, `iopaint` downloads LaMa weights (~200 MB) to the user cache. Subsequent runs are offline.

## Run from source

### Requirements

- Python 3.9+
- Windows / macOS / Linux
- Dependencies in `requirements.txt`:
  - `opencv-python`, `numpy`, `pillow`, `python-docx`, `iopaint` (pulls in `torch`)

### Install

```bash
pip install -r requirements.txt
```

### Run

```bash
python watermark_remover_gui.py
```

## Usage

### Single image

1. **File → Open Image File** or **Open from Word** and pick your source.
2. When opening from Word, a selector window lists all embedded images with a preview; pick the one you want to fix.
3. Drag one or more rectangles over the watermark(s). Confirmed boxes are red; the one you're currently dragging is a dashed orange.
4. Click **Remove Watermark**. LaMa runs via `iopaint` and fills every box in a single pass.
5. Click **Save to Word** (only when the image came from a `.docx`) or **Export PNG**.

### Batch-remove every image in a Word doc

1. **File → Open from Word**, pick any image as a reference.
2. Draw your watermark region(s) on it.
3. Click **Batch: Apply to ALL Word Images**.
4. Your regions are converted to fractions of the reference image's size and re-applied proportionally to every image in the document. A `.backup` copy of the `.docx` is written before anything is modified.

The batch button is useful when the same template watermark sits at roughly the same relative position in every figure (e.g. a logo in the top-right of every screenshot).

## Shortcuts

| Key / Mouse | Action |
|---|---|
| Left-drag | Draw a new region |
| Shift + Left-drag | Pan |
| Middle-drag | Pan |
| Mouse wheel | Zoom at cursor |
| `+` / `=` | Zoom in |
| `-` | Zoom out |
| `0` | Fit to window |
| `Esc` | Clear all regions |
| `Ctrl+Z` | Reset image to original |

Buttons also exist for **Undo Last Region** and **Clear All Regions**.

## Notes

- The original `.docx` is always backed up as `<name>.docx.backup` before **Save to Word** or **Batch**.
- Temporary files live under `D:/temp/watermark_remover_temp/` and `D:/temp/watermark_remover_batch/`. Edit the `TEMP_ROOT` / `BATCH_TEMP_ROOT` constants at the top of the script to relocate.
- If `iopaint` fails, verify the environment with `python -m iopaint --help`.
- Frame the watermark tightly — an oversized box will wipe out surrounding content; a tiny box can leave residue. Splitting one big watermark into several smaller boxes usually beats one giant region.

## License

MIT — see [LICENSE](LICENSE).
