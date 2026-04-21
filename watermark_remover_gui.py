# -*- coding: utf-8 -*-
"""
Watermark Remover GUI v4 — multi-region selection, zoom/pan, Word batch.
"""

import os
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from docx import Document


TEMP_ROOT = Path("D:/temp/watermark_remover_temp")
BATCH_TEMP_ROOT = Path("D:/temp/watermark_remover_batch")
EXTRACT_ROOT = Path("D:/temp/docx_temp_extract")
BATCH_EXTRACT_ROOT = Path("D:/temp/docx_temp_extract_batch")


class WatermarkRemover:
    SHIFT_MASK = 0x0001

    def __init__(self, root):
        self.root = root
        self.root.title("Watermark Remover v4")
        self.root.geometry("1400x900")

        self.cv_image = None
        self.original_image = None
        self.photo = None

        self.rects = []
        self.in_progress = None
        self.drawing = False

        self.base_scale = 1.0
        self.zoom = 1.0
        self.pan_x = 0.0
        self.pan_y = 0.0

        self.panning = False
        self.pan_anchor = None

        self.docx_path = None
        self.image_rId = None
        self.doc = None

        self._build_menu()
        self._build_layout()
        self._bind_events()

        self.log(
            "Ready.\n"
            "1. File -> Open image / Open from Word\n"
            "2. Drag to draw boxes (multiple supported)\n"
            "3. Remove Watermark  OR  Batch across Word\n"
            "4. Save to Word / Export PNG\n"
            "Wheel = zoom at cursor\n"
            "Shift+drag or middle-drag = pan\n"
            "+/-/0 = zoom in/out/fit, Esc = clear, Ctrl+Z = reset"
        )

    def _build_menu(self):
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Image File", command=self.open_image)
        file_menu.add_command(label="Open from Word", command=self.open_from_docx)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

    def _build_layout(self):
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        left = tk.Frame(main_frame)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(
            left,
            text="Preview - drag to box, Shift+drag/middle-drag to pan, wheel to zoom",
            font=("Arial", 11, "bold"),
        ).pack()
        self.canvas = tk.Canvas(left, bg="gray", cursor="crosshair")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        right = tk.Frame(main_frame, width=320)
        right.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10)
        tk.Label(right, text="Control Panel", font=("Arial", 12, "bold")).pack()

        tk.Label(right, text="Current file:", font=("Arial", 10)).pack(anchor="w", pady=(10, 5))
        self.file_text = tk.Text(right, height=2, width=40)
        self.file_text.pack(anchor="w", padx=5)

        tk.Label(right, text="Regions:", font=("Arial", 10)).pack(anchor="w", pady=(15, 5))
        self.coord_text = tk.Text(right, height=6, width=40)
        self.coord_text.pack(anchor="w", padx=5)

        btn_frame = tk.Frame(right)
        btn_frame.pack(anchor="w", padx=5, pady=10, fill=tk.X)

        tk.Button(btn_frame, text="Undo Last Region", command=self.undo_last_region,
                  bg="lightblue").pack(fill=tk.X, pady=3)
        tk.Button(btn_frame, text="Clear All Regions (Esc)", command=self.clear_selection,
                  bg="lightblue").pack(fill=tk.X, pady=3)
        self.process_btn = tk.Button(btn_frame, text="Remove Watermark",
                                     command=self.remove_watermark, bg="lightgreen",
                                     font=("Arial", 11, "bold"))
        self.process_btn.pack(fill=tk.X, pady=5)
        self.batch_btn = tk.Button(btn_frame, text="Batch: Apply to ALL Word Images",
                                   command=self.batch_remove_word, bg="#c5e1a5",
                                   font=("Arial", 10, "bold"))
        self.batch_btn.pack(fill=tk.X, pady=3)
        tk.Button(btn_frame, text="Reset to Original (Ctrl+Z)", command=self.reset_image,
                  bg="lightgray").pack(fill=tk.X, pady=3)

        zoom_row = tk.Frame(btn_frame)
        zoom_row.pack(fill=tk.X, pady=3)
        tk.Button(zoom_row, text="Zoom -", command=lambda: self.zoom_step(1 / 1.25)).pack(
            side=tk.LEFT, expand=True, fill=tk.X)
        tk.Button(zoom_row, text="Fit (0)", command=self.reset_view).pack(
            side=tk.LEFT, expand=True, fill=tk.X)
        tk.Button(zoom_row, text="Zoom +", command=lambda: self.zoom_step(1.25)).pack(
            side=tk.LEFT, expand=True, fill=tk.X)

        tk.Button(btn_frame, text="Save to Word", command=self.save_to_docx,
                  bg="lightyellow").pack(fill=tk.X, pady=3)
        tk.Button(btn_frame, text="Export PNG", command=self.save_image,
                  bg="lightcoral").pack(fill=tk.X, pady=3)

        tk.Label(right, text="Status:", font=("Arial", 10)).pack(anchor="w", pady=(15, 5))
        self.status_text = tk.Text(right, height=10, width=40)
        self.status_text.pack(anchor="w", padx=5, fill=tk.BOTH, expand=True)

    def _bind_events(self):
        self.canvas.bind("<Configure>", lambda e: self.display_image())
        self.canvas.bind("<Button-1>", self.on_left_down)
        self.canvas.bind("<B1-Motion>", self.on_left_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_left_up)
        self.canvas.bind("<Button-2>", self.on_pan_start)
        self.canvas.bind("<B2-Motion>", self.on_pan_move)
        self.canvas.bind("<ButtonRelease-2>", self.on_pan_end)
        self.canvas.bind("<MouseWheel>", self.on_wheel)
        self.canvas.bind("<Button-4>", lambda e: self._zoom_around(e.x, e.y, 1.1))
        self.canvas.bind("<Button-5>", lambda e: self._zoom_around(e.x, e.y, 1 / 1.1))

        self.root.bind("<Escape>", lambda e: self.clear_selection())
        self.root.bind("<Control-z>", lambda e: self.reset_image())
        self.root.bind("<plus>", lambda e: self.zoom_step(1.25))
        self.root.bind("<equal>", lambda e: self.zoom_step(1.25))
        self.root.bind("<minus>", lambda e: self.zoom_step(1 / 1.25))
        self.root.bind("<Key-0>", lambda e: self.reset_view())

    def log(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update()

    def open_image(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp"), ("All files", "*.*")]
        )
        if not filepath:
            return
        try:
            img = cv2.imdecode(np.fromfile(filepath, dtype=np.uint8), cv2.IMREAD_COLOR)
        except Exception as e:
            self.log("Read error: " + str(e))
            img = None
        if img is None:
            messagebox.showerror("Error", "Cannot open image:\n" + filepath)
            return
        self.cv_image = img
        self.original_image = img.copy()
        self.docx_path = None
        self.image_rId = None
        self.doc = None
        self._reset_selection_state()
        self.file_text.delete(1.0, tk.END)
        self.file_text.insert(tk.END, "Type: Image file\n" + Path(filepath).name)
        self.log("Opened: " + Path(filepath).name)
        self.log("Size: {}x{}".format(img.shape[1], img.shape[0]))
        self.reset_view()

    def open_from_docx(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")]
        )
        if not filepath:
            return
        try:
            doc = Document(filepath)
        except Exception as e:
            messagebox.showerror("Error", "Cannot open Word file: " + str(e))
            return
        self.doc = doc
        self.docx_path = filepath
        images_info = []
        for rel_id, rel in doc.part.rels.items():
            if 'image' in rel.reltype:
                images_info.append({
                    'rId': rel_id,
                    'name': Path(rel.target_part.partname).name,
                    'size': len(rel.target_part.blob),
                })
        if not images_info:
            messagebox.showerror("Error", "No images in Word document")
            return
        self.log("Opened: " + Path(filepath).name)
        self.log("Found {} images".format(len(images_info)))
        self.show_image_selector(images_info)

    def show_image_selector(self, images_info):
        selector = tk.Toplevel(self.root)
        selector.title("Select Image")
        selector.geometry("720x520")
        tk.Label(selector, text="Select image to edit:", font=("Arial", 11, "bold")).pack(pady=10)

        main_frame = tk.Frame(selector)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        list_frame = tk.Frame(main_frame)
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tk.Label(list_frame, text="Images:", font=("Arial", 10, "bold")).pack(anchor="w")
        frame = tk.Frame(list_frame)
        frame.pack(fill=tk.BOTH, expand=True)
        scrollbar = tk.Scrollbar(frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox = tk.Listbox(frame, yscrollcommand=scrollbar.set, height=12)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        for info in images_info:
            size_mb = info['size'] / 1024 / 1024
            listbox.insert(tk.END, "{} ({:.2f} MB)".format(info['name'], size_mb))

        preview_frame = tk.Frame(main_frame, width=280, bg="gray")
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(10, 0))
        tk.Label(preview_frame, text="Preview:", font=("Arial", 10, "bold"),
                 bg="gray", fg="white").pack()
        self.preview_canvas = tk.Canvas(preview_frame, bg="gray", width=260, height=360)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        def show_preview(event=None):
            sel = listbox.curselection()
            if not sel:
                return
            info = images_info[sel[0]]
            rel = self.doc.part.rels[info['rId']]
            nparr = np.frombuffer(rel.target_part.blob, np.uint8)
            img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
            if img is None:
                return
            h, w = img.shape[:2]
            scale = min(260 / w, 360 / h, 1.0)
            nw, nh = max(1, int(w * scale)), max(1, int(h * scale))
            resized = cv2.resize(img, (nw, nh))
            rgb = cv2.cvtColor(resized, cv2.COLOR_BGR2RGB)
            photo = ImageTk.PhotoImage(Image.fromarray(rgb))
            self.preview_canvas.delete("all")
            self.preview_canvas.create_image(130, 180, image=photo)
            self.preview_canvas.image = photo

        listbox.bind("<<ListboxSelect>>", show_preview)
        if images_info:
            listbox.selection_set(0)
            show_preview()

        def do_select():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("Warning", "Please select an image")
                return
            info = images_info[sel[0]]
            self.image_rId = info['rId']
            rel = self.doc.part.rels[self.image_rId]
            nparr = np.frombuffer(rel.target_part.blob, np.uint8)
            img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
            if img is None:
                messagebox.showerror("Error", "Cannot decode image from Word")
                return
            self.cv_image = img
            self.original_image = img.copy()
            self._reset_selection_state()
            self.file_text.delete(1.0, tk.END)
            self.file_text.insert(
                tk.END,
                "Type: Word document\n{}\nrId: {}".format(info['name'], self.image_rId),
            )
            self.log("Loaded: " + info['name'])
            self.log("Size: {}x{}".format(img.shape[1], img.shape[0]))
            self.reset_view()
            selector.destroy()

        tk.Button(selector, text="Confirm", command=do_select, bg="lightgreen",
                  font=("Arial", 11)).pack(pady=10)

    def reset_view(self):
        if self.cv_image is None:
            self.zoom = 1.0
            self.pan_x = 0.0
            self.pan_y = 0.0
            self.canvas.delete("all")
            return
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 50 or canvas_h < 50:
            canvas_w, canvas_h = 800, 600
        h, w = self.cv_image.shape[:2]
        self.base_scale = min(canvas_w / w, canvas_h / h, 1.0)
        self.zoom = 1.0
        eff = self.base_scale * self.zoom
        self.pan_x = (canvas_w - w * eff) / 2
        self.pan_y = (canvas_h - h * eff) / 2
        self.display_image()

    def zoom_step(self, factor):
        if self.cv_image is None:
            return
        cx = self.canvas.winfo_width() / 2
        cy = self.canvas.winfo_height() / 2
        self._zoom_around(cx, cy, factor)

    def _zoom_around(self, cx, cy, factor):
        if self.cv_image is None:
            return
        new_zoom = max(0.1, min(20.0, self.zoom * factor))
        if abs(new_zoom - self.zoom) < 1e-6:
            return
        ix, iy = self._canvas_to_image(cx, cy)
        self.zoom = new_zoom
        eff = self.base_scale * self.zoom
        self.pan_x = cx - ix * eff
        self.pan_y = cy - iy * eff
        self.display_image()

    def on_wheel(self, event):
        if self.cv_image is None:
            return
        factor = 1.1 if event.delta > 0 else (1 / 1.1)
        self._zoom_around(event.x, event.y, factor)

    def display_image(self):
        if self.cv_image is None:
            return
        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 50 or canvas_h < 50:
            canvas_w, canvas_h = 800, 600
        h, w = self.cv_image.shape[:2]
        self.base_scale = min(canvas_w / w, canvas_h / h, 1.0)
        eff = self.base_scale * self.zoom
        nw = max(1, int(w * eff))
        nh = max(1, int(h * eff))
        # Cap extreme zooms to avoid huge PIL allocations.
        if nw * nh > 40_000_000:
            cap = (40_000_000 / (nw * nh)) ** 0.5
            nw = max(1, int(nw * cap))
            nh = max(1, int(nh * cap))
        rgb = cv2.cvtColor(self.cv_image, cv2.COLOR_BGR2RGB)
        resized = cv2.resize(rgb, (nw, nh))
        self.photo = ImageTk.PhotoImage(Image.fromarray(resized))
        self.canvas.delete("all")
        self.canvas.create_image(self.pan_x, self.pan_y, anchor=tk.NW,
                                 image=self.photo, tags="img")
        self._redraw_rects()

    def _image_to_canvas(self, ix, iy):
        eff = self.base_scale * self.zoom
        return self.pan_x + ix * eff, self.pan_y + iy * eff

    def _canvas_to_image(self, cx, cy):
        eff = self.base_scale * self.zoom
        if eff <= 0 or self.cv_image is None:
            return 0, 0
        x = (cx - self.pan_x) / eff
        y = (cy - self.pan_y) / eff
        h, w = self.cv_image.shape[:2]
        return max(0, min(w, x)), max(0, min(h, y))

    def _redraw_rects(self):
        self.canvas.delete("rect")
        for (x1, y1, x2, y2) in self.rects:
            cx1, cy1 = self._image_to_canvas(x1, y1)
            cx2, cy2 = self._image_to_canvas(x2, y2)
            self.canvas.create_rectangle(cx1, cy1, cx2, cy2,
                                         outline="red", width=2, tags="rect")
        if self.in_progress:
            x1, y1, x2, y2 = self.in_progress
            cx1, cy1 = self._image_to_canvas(x1, y1)
            cx2, cy2 = self._image_to_canvas(x2, y2)
            self.canvas.create_rectangle(cx1, cy1, cx2, cy2,
                                         outline="orange", width=2,
                                         dash=(4, 2), tags="rect")

    def on_left_down(self, event):
        if self.cv_image is None:
            return
        if event.state & self.SHIFT_MASK:
            self.on_pan_start(event)
            return
        ix, iy = self._canvas_to_image(event.x, event.y)
        self.in_progress = (ix, iy, ix, iy)
        self.drawing = True
        self._redraw_rects()

    def on_left_drag(self, event):
        if self.panning:
            self.on_pan_move(event)
            return
        if not self.drawing or self.cv_image is None:
            return
        ix, iy = self._canvas_to_image(event.x, event.y)
        x1, y1, _, _ = self.in_progress
        self.in_progress = (x1, y1, ix, iy)
        self._redraw_rects()
        self._update_coord_text()

    def on_left_up(self, event):
        if self.panning:
            self.on_pan_end(event)
            return
        if not self.drawing:
            return
        self.drawing = False
        if self.in_progress:
            x1, y1, x2, y2 = self.in_progress
            x1, x2 = min(x1, x2), max(x1, x2)
            y1, y2 = min(y1, y2), max(y1, y2)
            if (x2 - x1) >= 2 and (y2 - y1) >= 2:
                self.rects.append((x1, y1, x2, y2))
            self.in_progress = None
            self._redraw_rects()
            self._update_coord_text()

    def on_pan_start(self, event):
        if self.cv_image is None:
            return
        self.panning = True
        self.pan_anchor = (event.x, event.y, self.pan_x, self.pan_y)
        self.canvas.config(cursor="fleur")

    def on_pan_move(self, event):
        if not self.panning or not self.pan_anchor:
            return
        ex0, ey0, px0, py0 = self.pan_anchor
        self.pan_x = px0 + (event.x - ex0)
        self.pan_y = py0 + (event.y - ey0)
        self.display_image()

    def on_pan_end(self, event):
        self.panning = False
        self.pan_anchor = None
        self.canvas.config(cursor="crosshair")

    def undo_last_region(self):
        if self.rects:
            self.rects.pop()
            self._redraw_rects()
            self._update_coord_text()
            self.log("Removed last region")

    def clear_selection(self):
        self.rects = []
        self.in_progress = None
        self.drawing = False
        self._redraw_rects()
        self._update_coord_text()
        self.log("Cleared all regions")

    def _reset_selection_state(self):
        self.rects = []
        self.in_progress = None
        self.drawing = False
        self.coord_text.delete(1.0, tk.END)

    def _update_coord_text(self):
        self.coord_text.delete(1.0, tk.END)
        text = "Total regions: {}\n".format(len(self.rects))
        start = max(0, len(self.rects) - 4)
        for i in range(start, len(self.rects)):
            x1, y1, x2, y2 = self.rects[i]
            text += "  #{}: x=[{}:{}] y=[{}:{}] {}x{}\n".format(
                i + 1, int(x1), int(x2), int(y1), int(y2),
                int(x2 - x1), int(y2 - y1),
            )
        if self.in_progress:
            x1, y1, x2, y2 = self.in_progress
            x1, x2 = min(x1, x2), max(x1, x2)
            y1, y2 = min(y1, y2), max(y1, y2)
            text += "drawing: x=[{}:{}] y=[{}:{}]\n".format(
                int(x1), int(x2), int(y1), int(y2)
            )
        self.coord_text.insert(tk.END, text)

    def reset_image(self):
        if self.original_image is None:
            return
        self.cv_image = self.original_image.copy()
        self._reset_selection_state()
        self.reset_view()
        self.log("Reset to original")

    def _build_mask_for_rects(self, rects, h, w):
        mask = np.zeros((h, w), dtype=np.uint8)
        for (x1, y1, x2, y2) in rects:
            ix1 = max(0, int(min(x1, x2)))
            iy1 = max(0, int(min(y1, y2)))
            ix2 = min(w, int(max(x1, x2)))
            iy2 = min(h, int(max(y1, y2)))
            if ix2 > ix1 and iy2 > iy1:
                mask[iy1:iy2, ix1:ix2] = 255
        return mask

    def _run_iopaint(self, img_bgr, mask, temp_dir):
        temp_dir.mkdir(exist_ok=True, parents=True)
        img_path = temp_dir / "input.png"
        mask_path = temp_dir / "mask.png"
        out_dir = temp_dir / "output"
        if out_dir.exists():
            shutil.rmtree(out_dir)
        out_dir.mkdir(parents=True)
        cv2.imwrite(str(img_path), img_bgr)
        cv2.imwrite(str(mask_path), mask)
        subprocess.run(
            [sys.executable, '-m', 'iopaint', 'run',
             '--model', 'lama',
             '--image', str(img_path),
             '--mask', str(mask_path),
             '--output', str(out_dir)],
            check=True, capture_output=True, timeout=300, text=True,
            encoding='utf-8', errors='replace',
        )
        result = cv2.imread(str(out_dir / "input.png"))
        if result is None:
            raise RuntimeError("iopaint produced no output at " + str(out_dir / "input.png"))
        return result

    def remove_watermark(self):
        if self.cv_image is None:
            messagebox.showerror("Error", "Please open an image first")
            return
        if not self.rects:
            messagebox.showerror("Error", "Draw at least one region first")
            return
        h, w = self.cv_image.shape[:2]
        mask = self._build_mask_for_rects(self.rects, h, w)
        if mask.max() == 0:
            messagebox.showerror("Error", "Regions are empty after clipping")
            return
        self.log("Removing watermark from {} region(s)...".format(len(self.rects)))
        self.process_btn.config(state="disabled")
        self.batch_btn.config(state="disabled")
        self.root.config(cursor="watch")
        self.root.update()
        try:
            result = self._run_iopaint(self.cv_image, mask, TEMP_ROOT)
            self.cv_image = result
            self.rects = []
            self.in_progress = None
            self._update_coord_text()
            self.display_image()
            self.log("Done.")
            messagebox.showinfo("Success", "Watermark removed")
        except subprocess.CalledProcessError as e:
            err = (e.stderr or e.stdout or "")[-800:]
            self.log("iopaint failed:\n" + err)
            messagebox.showerror("Error", "iopaint failed:\n" + err)
        except subprocess.TimeoutExpired:
            self.log("iopaint timed out (>5min)")
            messagebox.showerror("Error", "iopaint timed out (>5min)")
        except Exception as e:
            self.log("Error: " + str(e))
            messagebox.showerror("Error", "Failed: " + str(e))
        finally:
            self.process_btn.config(state="normal")
            self.batch_btn.config(state="normal")
            self.root.config(cursor="")

    def batch_remove_word(self):
        if not self.docx_path or not self.doc:
            messagebox.showerror(
                "Error", "Open a Word document first (File -> Open from Word)")
            return
        if self.original_image is None or not self.rects:
            messagebox.showerror(
                "Error", "Draw at least one region on the current image first")
            return
        # Convert rects to fractions of the current image; reapply proportionally per image.
        h0, w0 = self.original_image.shape[:2]
        rel_rects = [(x1 / w0, y1 / h0, x2 / w0, y2 / h0)
                     for (x1, y1, x2, y2) in self.rects]

        img_rels = [(rid, rel) for rid, rel in self.doc.part.rels.items()
                    if 'image' in rel.reltype]
        if not messagebox.askyesno(
            "Batch Remove",
            "Apply the {} region(s) you drew to ALL {} images in:\n"
            "{}\n\n"
            "Regions scale proportionally per image.\n"
            "A .backup copy of the .docx will be created.\n\nContinue?".format(
                len(self.rects), len(img_rels), self.docx_path)
        ):
            return

        backup_path = str(self.docx_path) + ".backup"
        shutil.copy(self.docx_path, backup_path)
        self.log("Backup: " + backup_path)

        self.process_btn.config(state="disabled")
        self.batch_btn.config(state="disabled")
        self.root.config(cursor="watch")
        self.root.update()

        processed = {}
        failed = []
        try:
            for idx, (rid, rel) in enumerate(img_rels, 1):
                partname = rel.target_part.partname
                name = Path(partname).name
                self.log("[{}/{}] {}".format(idx, len(img_rels), name))
                self.root.update()
                try:
                    nparr = np.frombuffer(rel.target_part.blob, np.uint8)
                    img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
                    if img is None:
                        self.log("  skip (cannot decode)")
                        failed.append(name)
                        continue
                    h, w = img.shape[:2]
                    abs_rects = [(rx1 * w, ry1 * h, rx2 * w, ry2 * h)
                                 for (rx1, ry1, rx2, ry2) in rel_rects]
                    mask = self._build_mask_for_rects(abs_rects, h, w)
                    if mask.max() == 0:
                        self.log("  skip (empty mask)")
                        failed.append(name)
                        continue
                    result = self._run_iopaint(img, mask, BATCH_TEMP_ROOT)
                    ext = Path(partname).suffix.lower() or ".png"
                    ok, buf = cv2.imencode(ext, result)
                    if not ok:
                        self.log("  encode failed")
                        failed.append(name)
                        continue
                    processed[partname] = buf.tobytes()
                    self.log("  ok")
                except subprocess.CalledProcessError as e:
                    err = (e.stderr or e.stdout or "")[-300:]
                    self.log("  FAILED: iopaint: " + err)
                    failed.append(name)
                except Exception as e:
                    self.log("  FAILED: " + str(e))
                    failed.append(name)

            if not processed:
                messagebox.showerror("Batch Remove", "No images were processed.")
                return

            if BATCH_EXTRACT_ROOT.exists():
                shutil.rmtree(BATCH_EXTRACT_ROOT)
            BATCH_EXTRACT_ROOT.mkdir(parents=True)
            with zipfile.ZipFile(self.docx_path, 'r') as zf:
                zf.extractall(BATCH_EXTRACT_ROOT)
            for partname, bts in processed.items():
                p = BATCH_EXTRACT_ROOT / partname.lstrip('/')
                p.parent.mkdir(parents=True, exist_ok=True)
                with open(p, 'wb') as f:
                    f.write(bts)
            with zipfile.ZipFile(self.docx_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root_dir, _, files in os.walk(BATCH_EXTRACT_ROOT):
                    for file in files:
                        fp = Path(root_dir) / file
                        zf.write(fp, fp.relative_to(BATCH_EXTRACT_ROOT))
            shutil.rmtree(BATCH_EXTRACT_ROOT)

            self.doc = Document(self.docx_path)

            msg = "Batch complete.\nProcessed: {}\nFailed: {}\nBackup: {}".format(
                len(processed), len(failed), backup_path)
            if failed:
                msg += "\n\nFailed images:\n" + "\n".join("  " + f for f in failed[:10])
                if len(failed) > 10:
                    msg += "\n  ... and {} more".format(len(failed) - 10)
            self.log(msg)
            messagebox.showinfo("Batch Remove", msg)
        finally:
            self.process_btn.config(state="normal")
            self.batch_btn.config(state="normal")
            self.root.config(cursor="")

    def save_to_docx(self):
        if self.cv_image is None:
            messagebox.showerror("Error", "No image to save")
            return
        if not self.docx_path or not self.image_rId:
            messagebox.showerror(
                "Error", "This image was not opened from Word. Use Export PNG instead.")
            return
        try:
            rel = self.doc.part.rels[self.image_rId]
            image_part_name = rel.target_part.partname
            backup_path = str(self.docx_path) + ".backup"
            shutil.copy(self.docx_path, backup_path)
            ext = Path(image_part_name).suffix.lower() or ".png"
            ok, buffer = cv2.imencode(ext, self.cv_image)
            if not ok:
                raise RuntimeError("cv2.imencode failed")
            image_data = buffer.tobytes()

            if EXTRACT_ROOT.exists():
                shutil.rmtree(EXTRACT_ROOT)
            EXTRACT_ROOT.mkdir(parents=True)
            with zipfile.ZipFile(self.docx_path, 'r') as zf:
                zf.extractall(EXTRACT_ROOT)
            image_path = EXTRACT_ROOT / image_part_name.lstrip('/')
            image_path.parent.mkdir(parents=True, exist_ok=True)
            with open(image_path, 'wb') as f:
                f.write(image_data)
            with zipfile.ZipFile(self.docx_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root_dir, _, files in os.walk(EXTRACT_ROOT):
                    for file in files:
                        fp = Path(root_dir) / file
                        zf.write(fp, fp.relative_to(EXTRACT_ROOT))
            shutil.rmtree(EXTRACT_ROOT)
            self.log("Saved to Word (backup: {})".format(backup_path))
            messagebox.showinfo("Success", "Saved to:\n" + self.docx_path)
        except Exception as e:
            self.log("Error: " + str(e))
            messagebox.showerror("Error", "Failed to save: " + str(e))

    def save_image(self):
        if self.cv_image is None:
            messagebox.showerror("Error", "No image to save")
            return
        filepath = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg"), ("All files", "*.*")],
        )
        if not filepath:
            return
        ext = Path(filepath).suffix.lower() or ".png"
        ok, buf = cv2.imencode(ext, self.cv_image)
        if ok:
            buf.tofile(filepath)
            self.log("Exported: " + Path(filepath).name)
            messagebox.showinfo("Success", "Saved to:\n" + filepath)
        else:
            messagebox.showerror("Error", "Failed to encode image")


if __name__ == "__main__":
    root = tk.Tk()
    app = WatermarkRemover(root)
    root.mainloop()
