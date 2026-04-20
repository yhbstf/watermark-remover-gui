# -*- coding: utf-8 -*-
"""
Watermark Remover GUI v3 - Extract from Word and Remove Watermarks
1. Open Word doc and select image
2. Drag to select watermark area (red rectangle)
3. Click Remove to remove watermark
4. Save back to Word or export PNG
"""

import cv2
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import subprocess
import sys
import os
import shutil
import zipfile
from pathlib import Path
from docx import Document

class WatermarkRemover:
    def __init__(self, root):
        self.root = root
        self.root.title("Watermark Remover v3")
        self.root.geometry("1400x900")

        self.cv_image = None
        self.original_image = None
        self.photo = None
        self.start_point = None
        self.end_point = None
        self.drawing = False
        self.scale_factor = 1.0
        self.display_scale = 1.0
        self.img_offset_x = 0
        self.img_offset_y = 0

        self.docx_path = None
        self.image_rId = None
        self.doc = None

        # Menu
        menubar = tk.Menu(root)
        root.config(menu=menubar)

        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="Open Image File", command=self.open_image)
        file_menu.add_command(label="Open from Word", command=self.open_from_docx)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=root.quit)

        # Main frame
        main_frame = tk.Frame(root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left: Image display
        left_frame = tk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        tk.Label(left_frame, text="Preview (Drag to select watermark area)", font=("Arial", 12, "bold")).pack()

        self.canvas = tk.Canvas(left_frame, bg="gray", cursor="crosshair")
        self.canvas.pack(fill=tk.BOTH, expand=True)
        self.canvas.bind("<Button-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)
        self.canvas.bind("<Configure>", lambda e: self.display_image())

        root.bind("<Escape>", lambda e: self.clear_selection())
        root.bind("<Control-z>", lambda e: self.reset_image())

        # Right: Control panel
        right_frame = tk.Frame(main_frame, width=300)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=10)

        tk.Label(right_frame, text="Control Panel", font=("Arial", 12, "bold")).pack()

        # File info
        tk.Label(right_frame, text="Current file:", font=("Arial", 10)).pack(anchor="w", pady=(10, 5))
        self.file_text = tk.Text(right_frame, height=2, width=40)
        self.file_text.pack(anchor="w", padx=5)

        # Coordinates
        tk.Label(right_frame, text="Selected area:", font=("Arial", 10)).pack(anchor="w", pady=(15, 5))
        self.coord_text = tk.Text(right_frame, height=4, width=40)
        self.coord_text.pack(anchor="w", padx=5)

        # Buttons
        button_frame = tk.Frame(right_frame)
        button_frame.pack(anchor="w", padx=5, pady=15)

        tk.Button(button_frame, text="Clear Selection (Esc)", command=self.clear_selection, bg="lightblue").pack(fill=tk.X, pady=5)
        self.process_btn = tk.Button(button_frame, text="Remove Watermark", command=self.remove_watermark, bg="lightgreen", font=("Arial", 11, "bold"))
        self.process_btn.pack(fill=tk.X, pady=5)
        tk.Button(button_frame, text="Reset to Original (Ctrl+Z)", command=self.reset_image, bg="lightgray").pack(fill=tk.X, pady=5)
        tk.Button(button_frame, text="Save to Word", command=self.save_to_docx, bg="lightyellow").pack(fill=tk.X, pady=5)
        tk.Button(button_frame, text="Export PNG", command=self.save_image, bg="lightcoral").pack(fill=tk.X, pady=5)

        # Status
        tk.Label(right_frame, text="Status:", font=("Arial", 10)).pack(anchor="w", pady=(15, 5))
        self.status_text = tk.Text(right_frame, height=8, width=40)
        self.status_text.pack(anchor="w", padx=5)

        self.log("Ready. Please:\n1. Open from Word\n2. Drag to select watermark\n3. Click Remove Watermark\n4. Save result")

    def log(self, message):
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.root.update()

    def open_image(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp"), ("All files", "*.*")]
        )
        if filepath:
            try:
                self.cv_image = cv2.imdecode(np.fromfile(filepath, dtype=np.uint8), cv2.IMREAD_COLOR)
            except Exception as e:
                self.cv_image = None
                self.log("Read error: " + str(e))
            if self.cv_image is None:
                messagebox.showerror("Error", "Cannot open image:\n" + filepath)
                return

            self.original_image = self.cv_image.copy()
            self.docx_path = None
            self.image_rId = None
            self.start_point = None
            self.end_point = None
            self.coord_text.delete(1.0, tk.END)

            self.log("Opened: " + Path(filepath).name)
            self.log("Size: {}x{}".format(self.cv_image.shape[1], self.cv_image.shape[0]))
            self.file_text.delete(1.0, tk.END)
            self.file_text.insert(tk.END, "Type: Image file\n" + Path(filepath).name)
            self.display_image()

    def open_from_docx(self):
        filepath = filedialog.askopenfilename(
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")]
        )
        if not filepath:
            return

        try:
            self.doc = Document(filepath)
            self.docx_path = filepath

            images_info = []
            for rel_id, rel in self.doc.part.rels.items():
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

        except Exception as e:
            messagebox.showerror("Error", "Cannot open Word file: " + str(e))
            self.log("Error: " + str(e))

    def show_image_selector(self, images_info):
        selector = tk.Toplevel(self.root)
        selector.title("Select Image")
        selector.geometry("700x500")

        tk.Label(selector, text="Select image to process:", font=("Arial", 11, "bold")).pack(pady=10)

        # Main frame with listbox and preview
        main_frame = tk.Frame(selector)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left: List
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

        for i, info in enumerate(images_info):
            size_mb = info['size'] / 1024 / 1024
            listbox.insert(tk.END, "{} ({:.2f} MB)".format(info['name'], size_mb))

        # Right: Preview
        preview_frame = tk.Frame(main_frame, width=280, bg="gray")
        preview_frame.pack(side=tk.RIGHT, fill=tk.BOTH, padx=(10, 0))

        tk.Label(preview_frame, text="Preview:", font=("Arial", 10, "bold"), bg="gray", fg="white").pack()

        self.preview_canvas = tk.Canvas(preview_frame, bg="gray", width=260, height=350)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        def show_preview(event=None):
            selection = listbox.curselection()
            if not selection:
                return

            idx = selection[0]
            info = images_info[idx]

            # Load and display preview
            rel = self.doc.part.rels[info['rId']]
            image_data = rel.target_part.blob

            nparr = np.frombuffer(image_data, np.uint8)
            img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

            if img is not None:
                # Resize for preview
                h, w = img.shape[:2]
                scale = min(260 / w, 350 / h, 1.0)
                new_w, new_h = int(w * scale), int(h * scale)
                img_resized = cv2.resize(img, (new_w, new_h))

                # Convert to RGB and PIL
                img_rgb = cv2.cvtColor(img_resized, cv2.COLOR_BGR2RGB)
                pil_img = Image.fromarray(img_rgb)
                photo = ImageTk.PhotoImage(pil_img)

                self.preview_canvas.create_image(130, 175, image=photo)
                self.preview_canvas.image = photo  # Keep reference

        listbox.bind("<<ListboxSelect>>", show_preview)

        # Initial preview
        if images_info:
            listbox.selection_set(0)
            show_preview()

        def select_image():
            selection = listbox.curselection()
            if not selection:
                messagebox.showwarning("Warning", "Please select an image")
                return

            idx = selection[0]
            info = images_info[idx]
            self.image_rId = info['rId']

            rel = self.doc.part.rels[self.image_rId]
            image_data = rel.target_part.blob

            nparr = np.frombuffer(image_data, np.uint8)
            self.cv_image = cv2.imdecode(nparr, cv2.IMREAD_COLOR)

            if self.cv_image is None:
                messagebox.showerror("Error", "Cannot decode image from Word")
                return

            self.original_image = self.cv_image.copy()
            self.start_point = None
            self.end_point = None
            self.coord_text.delete(1.0, tk.END)

            self.log("Loaded: " + info['name'])
            self.log("Size: {}x{}".format(self.cv_image.shape[1], self.cv_image.shape[0]))

            self.file_text.delete(1.0, tk.END)
            self.file_text.insert(tk.END, "Type: Word document\n{}\nrId: {}".format(info['name'], self.image_rId))

            self.display_image()
            selector.destroy()

        tk.Button(selector, text="Confirm", command=select_image, bg="lightgreen", font=("Arial", 11)).pack(pady=10)

    def display_image(self):
        if self.cv_image is None:
            return

        canvas_w = self.canvas.winfo_width()
        canvas_h = self.canvas.winfo_height()
        if canvas_w < 100 or canvas_h < 100:
            canvas_w, canvas_h = 800, 600

        h, w = self.cv_image.shape[:2]
        scale = min(canvas_w / w, canvas_h / h, 1.0)
        new_w, new_h = max(1, int(w * scale)), max(1, int(h * scale))

        image_rgb = cv2.cvtColor(self.cv_image, cv2.COLOR_BGR2RGB)
        img_resized = cv2.resize(image_rgb, (new_w, new_h))

        self.display_scale = scale
        self.scale_factor = 1.0 / scale
        self.img_offset_x = (canvas_w - new_w) // 2
        self.img_offset_y = (canvas_h - new_h) // 2

        self.photo = ImageTk.PhotoImage(Image.fromarray(img_resized))

        self.canvas.delete("all")
        self.canvas.create_image(self.img_offset_x, self.img_offset_y,
                                 anchor=tk.NW, image=self.photo, tags="img")
        self._redraw_selection()

    def _redraw_selection(self):
        self.canvas.delete("rect")
        if self.start_point and self.end_point:
            x1 = self.img_offset_x + self.start_point[0] * self.display_scale
            y1 = self.img_offset_y + self.start_point[1] * self.display_scale
            x2 = self.img_offset_x + self.end_point[0] * self.display_scale
            y2 = self.img_offset_y + self.end_point[1] * self.display_scale
            self.canvas.create_rectangle(x1, y1, x2, y2, outline="red", width=2, tags="rect")

    def _canvas_to_image(self, event_x, event_y):
        x = (event_x - self.img_offset_x) * self.scale_factor
        y = (event_y - self.img_offset_y) * self.scale_factor
        h, w = self.cv_image.shape[:2]
        return max(0, min(w, x)), max(0, min(h, y))

    def on_mouse_down(self, event):
        if self.cv_image is None:
            return
        self.start_point = self._canvas_to_image(event.x, event.y)
        self.end_point = self.start_point
        self.drawing = True

    def on_mouse_drag(self, event):
        if self.drawing and self.cv_image is not None:
            self.end_point = self._canvas_to_image(event.x, event.y)
            self._redraw_selection()
            self.update_coords()

    def on_mouse_up(self, event):
        self.drawing = False
        if self.start_point and self.end_point:
            self._redraw_selection()
            self.update_coords()

    def update_coords(self):
        if self.start_point and self.end_point:
            x1, y1 = self.start_point
            x2, y2 = self.end_point

            x1, x2 = min(x1, x2), max(x1, x2)
            y1, y2 = min(y1, y2), max(y1, y2)

            text = "x=[{}:{}]\ny=[{}:{}]\nSize: {}x{}".format(
                int(x1), int(x2), int(y1), int(y2), int(x2-x1), int(y2-y1)
            )
            self.coord_text.delete(1.0, tk.END)
            self.coord_text.insert(tk.END, text)

    def clear_selection(self):
        self.start_point = None
        self.end_point = None
        self.coord_text.delete(1.0, tk.END)
        self._redraw_selection()
        self.log("Selection cleared")

    def reset_image(self):
        if self.original_image is None:
            return
        self.cv_image = self.original_image.copy()
        self.start_point = None
        self.end_point = None
        self.coord_text.delete(1.0, tk.END)
        self.display_image()
        self.log("Reset to original")

    def remove_watermark(self):
        if self.cv_image is None:
            messagebox.showerror("Error", "Please open an image first")
            return

        if not self.start_point or not self.end_point:
            messagebox.showerror("Error", "Please select watermark area")
            return

        x1, y1 = self.start_point
        x2, y2 = self.end_point

        x1, x2 = int(min(x1, x2)), int(max(x1, x2))
        y1, y2 = int(min(y1, y2)), int(max(y1, y2))

        h, w = self.cv_image.shape[:2]
        x1, x2 = max(0, x1), min(w, x2)
        y1, y2 = max(0, y1), min(h, y2)

        if x2 - x1 < 2 or y2 - y1 < 2:
            messagebox.showerror("Error", "Selection area too small")
            return

        self.log("Removing watermark x=[{}:{}], y=[{}:{}]...".format(x1, x2, y1, y2))

        mask = np.zeros((h, w), dtype=np.uint8)
        mask[y1:y2, x1:x2] = 255

        temp_dir = Path("D:/temp/watermark_remover_temp")
        temp_dir.mkdir(exist_ok=True, parents=True)

        img_path = temp_dir / "input.png"
        mask_path = temp_dir / "mask.png"
        out_dir = temp_dir / "output"
        out_dir.mkdir(exist_ok=True, parents=True)

        cv2.imwrite(str(img_path), self.cv_image)
        cv2.imwrite(str(mask_path), mask)

        self.process_btn.config(state="disabled")
        self.root.config(cursor="watch")
        self.root.update()

        try:
            self.log("Running LaMa model...")
            subprocess.run([
                sys.executable, '-m', 'iopaint', 'run',
                '--model', 'lama',
                '--image', str(img_path),
                '--mask', str(mask_path),
                '--output', str(out_dir),
            ], check=True, capture_output=True, timeout=300, text=True,
               encoding='utf-8', errors='replace')

            result_path = out_dir / "input.png"
            result = cv2.imread(str(result_path))
            if result is None:
                raise RuntimeError("iopaint produced no output at " + str(result_path))
            self.cv_image = result
            self.start_point = None
            self.end_point = None
            self.coord_text.delete(1.0, tk.END)
            self.display_image()

            self.log("Done! Watermark removed.")
            messagebox.showinfo("Success", "Watermark removed successfully")

        except subprocess.CalledProcessError as e:
            err = (e.stderr or e.stdout or "")[-800:]
            self.log("iopaint failed:\n" + err)
            messagebox.showerror("Error", "iopaint failed:\n" + err)
        except subprocess.TimeoutExpired:
            self.log("iopaint timed out (>5min)")
            messagebox.showerror("Error", "iopaint timed out (>5min)")
        except Exception as e:
            self.log("Error: " + str(e))
            messagebox.showerror("Error", "Failed to remove watermark: " + str(e))
        finally:
            self.process_btn.config(state="normal")
            self.root.config(cursor="")

    def save_to_docx(self):
        if self.cv_image is None:
            messagebox.showerror("Error", "No image to save")
            return

        if not self.docx_path or not self.image_rId:
            messagebox.showerror("Error", "This image was not opened from Word. Use Export PNG instead.")
            return

        try:
            # Get image part name from relationship
            rel = self.doc.part.rels[self.image_rId]
            image_part_name = rel.target_part.partname

            # Create backup
            backup_path = str(self.docx_path) + ".backup"
            shutil.copy(self.docx_path, backup_path)

            # Encode new image
            _, buffer = cv2.imencode('.png', self.cv_image)
            image_data = buffer.tobytes()

            # Modify docx (it's a zip file)
            # Extract all files
            temp_extract = Path("D:/temp/docx_temp_extract")
            if temp_extract.exists():
                shutil.rmtree(temp_extract)
            temp_extract.mkdir()

            with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
                zip_ref.extractall(temp_extract)

            # Replace the image file
            image_path = temp_extract / image_part_name.lstrip('/')
            image_path.parent.mkdir(parents=True, exist_ok=True)
            with open(image_path, 'wb') as f:
                f.write(image_data)

            # Re-zip the docx
            with zipfile.ZipFile(self.docx_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for root, dirs, files in os.walk(temp_extract):
                    for file in files:
                        file_path = Path(root) / file
                        arcname = file_path.relative_to(temp_extract)
                        zip_ref.write(file_path, arcname)

            # Cleanup
            shutil.rmtree(temp_extract)

            self.log("Saved to Word document")
            messagebox.showinfo("Success", "Saved to Word:\n" + self.docx_path)

        except Exception as e:
            self.log("Error: " + str(e))
            messagebox.showerror("Error", "Failed to save: " + str(e))

    def save_image(self):
        if self.cv_image is None:
            messagebox.showerror("Error", "No image to save")
            return

        filepath = filedialog.asksaveasfilename(
            defaultextension=".png",
            filetypes=[("PNG files", "*.png"), ("JPEG files", "*.jpg"), ("All files", "*.*")]
        )
        if filepath:
            ext = Path(filepath).suffix or ".png"
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
