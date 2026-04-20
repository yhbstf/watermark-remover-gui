# Watermark Remover GUI

A Tkinter-based desktop tool that removes watermarks from images (including images embedded in Word `.docx` documents) using the [IOPaint](https://github.com/Sanster/IOPaint) LaMa inpainting model.

## Features

- Open a standalone image file, or pick an image directly from a Word `.docx` document (with preview).
- Drag a rectangle over the watermark area.
- One click to run LaMa inpainting and remove the watermark.
- Save the result back into the original Word document (a `.backup` is created automatically), or export a PNG/JPG.
- Reset (`Ctrl+Z`) and clear selection (`Esc`) shortcuts.

## Requirements

- Python 3.9+
- Windows/macOS/Linux (tested on Windows 11)
- Dependencies in `requirements.txt`:
  - `opencv-python`, `numpy`, `pillow`, `python-docx`, `iopaint`

The first time you run the inpainting step, `iopaint` will download the LaMa model weights (~200 MB).

## Install

```bash
pip install -r requirements.txt
```

## Run

```bash
python watermark_remover_gui.py
```

## Usage

1. **File → Open from Word** (or **Open Image File**) and pick your source.
2. When opening from Word, a selector window lists all embedded images with a preview; pick the one you want to fix.
3. In the main canvas, **click-drag** a red rectangle over the watermark.
4. Click **Remove Watermark**. The LaMa model runs via `iopaint` and replaces the region.
5. Click **Save to Word** (only works when the image came from a `.docx`) or **Export PNG** to save the result.

### Shortcuts

| Key | Action |
|-----|--------|
| `Esc` | Clear the current selection |
| `Ctrl+Z` | Reset the image back to the original |

## Notes

- When saving back to Word, the original `.docx` is backed up as `<name>.docx.backup` before being overwritten.
- Temporary files are written to `D:/temp/watermark_remover_temp/` (edit `remove_watermark` in the script if you want a different location).
- If `iopaint` fails, check that you can run `python -m iopaint --help` in the same environment.

## License

MIT — see [LICENSE](LICENSE).
