import os
import sys
import glob
import math
import shutil
import tempfile
import threading
from tkinter import messagebox, filedialog
import tkinter as tk
import tkinter.ttk as ttk
import customtkinter as ctk
from PIL import Image, ImageTk, ImageDraw
import fitz  # PyMuPDF
import win32com.client  # For PPT conversion on Windows

RES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "res")
CACHE_DIR = os.path.join(RES_DIR, ".cache")

# A4 Size at 300 DPI
A4_W_PTS, A4_H_PTS = 595.0, 842.0
COLUMNS = 4

# Settings
PPT_AVAILABLE = False
try:
    import win32com.client

    PPT_AVAILABLE = True
except ImportError:
    pass


class PDFRenderer:
    def __init__(self, path):
        self.doc = fitz.open(path)
        self.path = path
        self.doc_path = path

    def get_page_count(self):
        return self.doc.page_count

    def get_page_image(self, index, width=None):
        page = self.doc.load_page(index)
        if width:
            pix = page.get_pixmap(width=width)
        else:
            pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img

    def close(self):
        if self.doc:
            self.doc.close()

        # Debounce timer for resize
        self.resize_timer = None
        self.center_resize_timer = None
        self.cached_preview_pages = (
            []
        )  # Store PIL images of pages (Full A4 or high res)

        # Initial Render
        self._refresh_preview()


# Try importing dependencies
MISSING_DEPS = []

try:
    import customtkinter as ctk

    ctk.set_appearance_mode("Dark")
    ctk.set_default_color_theme("blue")
except ImportError:
    MISSING_DEPS.append("customtkinter")
    import tkinter as ctk  # Fallback

try:
    import fitz  # PyMuPDF
except ImportError:
    MISSING_DEPS.append("pymupdf")

try:
    from PIL import Image, ImageTk, ImageDraw
except ImportError:
    MISSING_DEPS.append("Pillow")

try:
    import win32com.client

    PPT_AVAILABLE = True
except ImportError:
    # Try comtypes as fallback
    try:
        import comtypes.client

        PPT_AVAILABLE = True
        win32com = comtypes
    except ImportError:
        PPT_AVAILABLE = False

# Configuration
RES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "res")
CACHE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "res", ".cache")

if not os.path.exists(RES_DIR):
    os.makedirs(RES_DIR)
if not os.path.exists(CACHE_DIR):
    os.makedirs(CACHE_DIR)

# Constants
A4_WIDTH = 2480
A4_HEIGHT = 3508
COLUMNS = 4
MARGIN = 40
GUTTER = 20


def check_deps():
    if MISSING_DEPS:
        root = tk.Tk()
        root.withdraw()
        msg = (
            "Missing required libraries:\n"
            + "\n".join(MISSING_DEPS)
            + "\n\nPlease install them using:\npip install "
            + " ".join(MISSING_DEPS)
        )
        messagebox.showerror("Missing Dependencies", msg)
        sys.exit(1)


class PDFRenderer:
    def __init__(self, doc_path):
        self.doc_path = doc_path
        self.doc = fitz.open(doc_path)

    def get_page_count(self):
        return len(self.doc)

    def get_page_image(self, page_num, zoom=1.0):
        if 0 <= page_num < len(self.doc):
            page = self.doc.load_page(page_num)
            mat = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            return img
        return None

    def close(self):
        if self.doc:
            self.doc.close()


def convert_ppt_to_pdf(ppt_path):
    if not os.path.exists(CACHE_DIR):
        os.makedirs(CACHE_DIR)

    if not PPT_AVAILABLE:
        return None

    base_name = os.path.splitext(os.path.basename(ppt_path))[0]
    pdf_path = os.path.join(CACHE_DIR, f"{base_name}.pdf")

    # Check if cached pdf is valid
    if os.path.exists(pdf_path):
        if os.path.getmtime(pdf_path) > os.path.getmtime(ppt_path):
            return pdf_path

    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        # app.Visible = 1 # Errors sometimes if hidden?
        # For pywin32, usually invisible works if WithWindow=False
        try:
            pres = app.Presentations.Open(ppt_path, WithWindow=False)
        except:
            # Fallback to visible if error
            pres = app.Presentations.Open(ppt_path)

        pres.SaveAs(pdf_path, 32)  # 32 = ppSaveAsPDF
        pres.Close()
        # app.Quit() # Keep it open for speed if we convert multiple? safer to leave it.
        return pdf_path
    except Exception as e:
        print(f"PPT Conversion Error: {e}")
        return None


class CheatSheetMaker(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Cheatsheet Maker")
        self.geometry("1400x900")

        self.current_pdf_renderer = None
        self.current_page_index = 0
        self.user_zoom = 1.0
        self.cheatsheet_items = []

        self.grid_columnconfigure(1, weight=1)
        # Variables for hit testing
        self.page_hit_boxes = []
        self.view_images = {}
        self.page_layouts = []
        self.loading_thread = None

        self._init_ui()
        self._load_files()

    def _init_ui(self):
        # Main Layout: PanedWindow
        # We need a style for the Sash to make it visible in dark mode
        # Tkinter PanedWindow sash styling is limited.
        # We use a bg color that shows through.

        self.paned = tk.PanedWindow(
            self, orient="horizontal", bg="#1a1a1a", sashwidth=8, sashrelief="flat"
        )
        self.paned.pack(fill="both", expand=True)

        # --- LEFT SIDEBAR ---
        self.sidebar = ctk.CTkFrame(self.paned, corner_radius=0)
        # PanedWindow expects widgets to be added via .add

        self.lbl_title = ctk.CTkLabel(
            self.sidebar, text="Resources", font=ctk.CTkFont(size=20, weight="bold")
        )
        self.lbl_title.pack(pady=20)

        self.btn_refresh = ctk.CTkButton(
            self.sidebar, text="Refresh", command=self._load_files
        )
        self.btn_refresh.pack(pady=10)

        # Treeview for File Explorer
        self.tree_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        self.tree_frame.pack(fill="both", expand=True, padx=5, pady=5)

        # Style Treeview
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "Treeview",
            background="#2b2b2b",
            fieldbackground="#2b2b2b",
            foreground="white",
            borderwidth=0,
            rowheight=25,
        )
        style.map("Treeview", background=[("selected", "#1f538d")])

        self.file_tree = ttk.Treeview(self.tree_frame, show="tree", selectmode="browse")
        self.file_tree.pack(side="left", fill="both", expand=True)

        self.tree_scroll = ctk.CTkScrollbar(
            self.tree_frame, orientation="vertical", command=self.file_tree.yview
        )
        self.tree_scroll.pack(side="right", fill="y")
        self.file_tree.configure(yscrollcommand=self.tree_scroll.set)

        # Bind double click or selection
        self.file_tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        self.paned.add(self.sidebar, minsize=200)

        # --- CENTER VIEW ---
        self.center_frame = ctk.CTkFrame(self.paned)
        self.center_frame.grid_rowconfigure(1, weight=1)
        self.center_frame.grid_columnconfigure(0, weight=1)

        self.lbl_file_name = ctk.CTkLabel(
            self.center_frame, text="Select a file", font=ctk.CTkFont(size=16)
        )
        self.lbl_file_name.grid(row=0, column=0, pady=5)

        self.canvas_frame = ctk.CTkFrame(self.center_frame, fg_color="transparent")
        self.canvas_frame.grid(row=1, column=0, sticky="nsew")
        self.canvas_frame.grid_rowconfigure(0, weight=1)
        self.canvas_frame.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(self.canvas_frame, bg="#2b2b2b", highlightthickness=0)
        self.canvas.grid(row=0, column=0, sticky="nsew")

        self.v_scroll = ctk.CTkScrollbar(
            self.canvas_frame, orientation="vertical", command=self.canvas.yview
        )
        self.v_scroll.grid(row=0, column=1, sticky="ns")

        self.h_scroll = ctk.CTkScrollbar(
            self.canvas_frame, orientation="horizontal", command=self.canvas.xview
        )
        self.h_scroll.grid(row=1, column=0, sticky="ew")

        self.canvas.configure(
            yscrollcommand=self.v_scroll.set, xscrollcommand=self.h_scroll.set
        )

        # Bind events
        self.canvas.bind("<Control-MouseWheel>", self._on_zoom_wheel)
        self.canvas.bind("<MouseWheel>", self._on_mouse_scroll)
        self.canvas.bind("<Button-4>", self._on_mouse_scroll)
        self.canvas.bind("<Button-5>", self._on_mouse_scroll)
        self.canvas.bind("<Enter>", lambda e: self._bind_mouse_scroll(e))
        self.canvas.bind("<Leave>", lambda e: self._unbind_mouse_scroll(e))
        self.canvas.bind("<Motion>", self._on_mouse_move)
        self.canvas.bind("<Button-1>", self._on_canvas_click)
        self.canvas.bind("<Configure>", self._on_canvas_resize)

        # Controls
        self.ctrl_frame = ctk.CTkFrame(self.center_frame, fg_color="transparent")
        self.ctrl_frame.grid(row=2, column=0, pady=10)

        self.btn_prev = ctk.CTkButton(
            self.ctrl_frame, text="<", width=40, command=self._prev
        )
        self.btn_prev.pack(side="left", padx=5)

        self.ent_page = ctk.CTkEntry(self.ctrl_frame, width=50, justify="center")
        self.ent_page.pack(side="left", padx=5)
        self.ent_page.bind("<Return>", self._jump)

        self.lbl_pages = ctk.CTkLabel(self.ctrl_frame, text="/ 0")
        self.lbl_pages.pack(side="left", padx=5)

        self.btn_next = ctk.CTkButton(
            self.ctrl_frame, text=">", width=40, command=self._next
        )
        self.btn_next.pack(side="left", padx=5)

        self.paned.add(self.center_frame, minsize=1300)  # Add center

        # --- RIGHT SIDEBAR ---
        self.right_frame = ctk.CTkFrame(self.paned, corner_radius=0)

        self.lbl_cs = ctk.CTkLabel(
            self.right_frame,
            text="Cheatsheet Preview",
            font=ctk.CTkFont(size=20, weight="bold"),
        )
        self.lbl_cs.pack(pady=10)

        self.btn_frame = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        self.btn_frame.pack(fill="x", padx=10)

        self.btn_export = ctk.CTkButton(
            self.btn_frame, text="Export PDF", command=self._export, width=100
        )
        self.btn_export.pack(side="left", padx=5)

        self.btn_undo = ctk.CTkButton(
            self.btn_frame,
            text="Undo",
            fg_color="orange",
            command=self._undo_last,
            width=60,
        )
        self.btn_undo.pack(side="left", padx=5)

        self.btn_clear = ctk.CTkButton(
            self.btn_frame, text="Clear", fg_color="red", command=self._clear, width=60
        )
        self.btn_clear.pack(side="left", padx=5)

        # Real-time Preview Area (Canvas based)
        self.preview_frame = ctk.CTkFrame(self.right_frame, fg_color="transparent")
        self.preview_frame.pack(fill="both", expand=True, padx=5, pady=5)
        self.preview_frame.grid_rowconfigure(0, weight=1)
        self.preview_frame.grid_columnconfigure(0, weight=1)

        self.preview_canvas = tk.Canvas(
            self.preview_frame, bg="#2b2b2b", highlightthickness=0
        )
        self.preview_canvas.grid(row=0, column=0, sticky="nsew")

        self.pv_scroll = ctk.CTkScrollbar(
            self.preview_frame,
            orientation="vertical",
            command=self.preview_canvas.yview,
        )
        self.pv_scroll.grid(row=0, column=1, sticky="ns")

        self.ph_scroll = ctk.CTkScrollbar(
            self.preview_frame,
            orientation="horizontal",
            command=self.preview_canvas.xview,
        )
        self.ph_scroll.grid(row=1, column=0, sticky="ew")

        self.preview_canvas.configure(
            yscrollcommand=self.pv_scroll.set, xscrollcommand=self.ph_scroll.set
        )

        # Preview State
        self.preview_zoom = 0.33
        self.preview_tk_images = {}  # Keep references

        # Bind Preview Events
        self.preview_canvas.bind("<Control-MouseWheel>", self._on_preview_zoom_wheel)
        self.preview_canvas.bind("<MouseWheel>", self._on_preview_scroll)
        self.preview_canvas.bind("<Button-4>", self._on_preview_scroll)
        self.preview_canvas.bind("<Button-5>", self._on_preview_scroll)

        self.paned.add(self.right_frame, minsize=200)

        # Debounce timer for resize
        self.resize_timer = None
        self.center_resize_timer = None
        self.cached_preview_pages = (
            []
        )  # Store PIL images of pages (Full A4 or high res)

        # Bind Resize on right sidebar (to adjust preview images)
        # self.preview_scroll.bind("<Configure>", self._on_preview_resize) # No longer needed

        # Initial Render
        self._refresh_preview()

    def _load_files(self):
        # Clear tree
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        # Recursive load
        self._populate_tree(RES_DIR, "")

    def _populate_tree(self, path, parent_node):
        try:
            # Sort: directories first, then files
            items = os.listdir(path)
            items.sort(
                key=lambda x: (not os.path.isdir(os.path.join(path, x)), x.lower())
            )

            for item in items:
                # Ignore . files
                if item.startswith("."):
                    continue

                full_path = os.path.join(path, item)
                is_dir = os.path.isdir(full_path)

                if is_dir:
                    node = self.file_tree.insert(
                        parent_node, "end", text=item, open=False, image=""
                    )
                    self._populate_tree(full_path, node)
                else:
                    valid_exts = [".pdf", ".ppt", ".pptx"]
                    ext = os.path.splitext(item)[1].lower()
                    if ext in valid_exts:
                        self.file_tree.insert(
                            parent_node, "end", text=item, values=[full_path]
                        )
        except Exception as e:
            print(f"Tree error {path}: {e}")

    def _on_tree_select(self, event):
        selected = self.file_tree.selection()
        if not selected:
            return

        item = self.file_tree.item(selected[0])
        # Check if it has values (file path store)
        if item["values"]:
            file_path = item["values"][0]
            self._select_file(file_path)

    def _select_file(self, path):
        # Stop existing thread if any (simple flag check could be added for robustness)
        self.loading_thread = None

        target_path = path
        ext = os.path.splitext(path)[1].lower()

        if ext in [".ppt", ".pptx"]:
            self.lbl_file_name.configure(text="Converting PPT...")
            self.update()
            target_path = convert_ppt_to_pdf(path)
            if not target_path:
                messagebox.showerror("Error", "Could not convert PPT.")
                self.lbl_file_name.configure(text="Error converting file")
                return

        if self.current_pdf_renderer:
            self.current_pdf_renderer.close()

        try:
            self.current_pdf_renderer = PDFRenderer(target_path)
            self.current_page_index = 0
            self.lbl_file_name.configure(text=os.path.basename(path))

            # Reset View
            # self.canvas.delete("all") # This is now handled by _update_view
            # self.page_layouts = [] # This is now handled by _update_view
            # self.view_images = {} # This is now handled by _update_view
            # self.canvas.yview_moveto(0) # This is now handled by _update_view

            # Start loading
            # threading.Thread(target=self._load_pdf_layout, daemon=True).start() # This is now handled by _update_view
            self._update_view()

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def _load_pdf_layout(self):
        """Phase 1: Calculate layout (Fast)"""
        if not self.current_pdf_renderer:
            return

        # Dispatch
        self.after(0, self._calculate_layout_and_start_render)

    def _calculate_layout_and_start_render(self):
        if not self.current_pdf_renderer:
            return

        cv_w = self.canvas.winfo_width()
        if cv_w < 100:
            cv_w = 600

        target_width = (cv_w - 40) * self.user_zoom
        if target_width < 100:
            target_width = 100

        total_pages = self.current_pdf_renderer.get_page_count()
        spacing = 20
        current_y = 20

        self.page_layouts = []

        for i in range(total_pages):
            page = self.current_pdf_renderer.doc.load_page(i)
            rect = page.rect

            # Calculate height maintaining aspect ratio
            scale = target_width / rect.width
            h = int(rect.height * scale)

            # Draw placeholder
            self.canvas.create_rectangle(
                20,
                current_y,
                20 + target_width,
                current_y + h,
                fill="#333333",
                outline="#444444",
                tags=f"ph_{i}",
            )
            self.canvas.create_text(
                20 + target_width / 2,
                current_y + h / 2,
                text=f"Loading Page {i+1}...",
                fill="gray",
                tags=f"txt_{i}",
            )

            self.page_layouts.append((current_y, h, i))
            current_y += h + spacing

        self.canvas.configure(scrollregion=(0, 0, cv_w, current_y))

        # Start Phase 2: Render in background
        self.loading_stop_event = threading.Event()
        self.loading_thread = threading.Thread(
            target=self._render_pages_background,
            args=(target_width, cv_w, self.loading_stop_event),
            daemon=True,
        )
        self.loading_thread.start()

    def _render_pages_background(self, target_width, canvas_width, stop_event):
        if not self.current_pdf_renderer:
            return

        for y, h, i in self.page_layouts:
            if stop_event.is_set():
                break

            try:
                page = self.current_pdf_renderer.doc.load_page(i)
                rect = page.rect
                scale = target_width / rect.width
                mat = fitz.Matrix(scale, scale)
                pix = page.get_pixmap(matrix=mat)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                self.after(
                    0,
                    lambda idx=i, image=img, cy=y, cw=canvas_width: self._update_page_image(
                        idx, image, cy, cw
                    ),
                )
            except Exception as e:
                print(f"Render error p{i}: {e}")

    def _update_page_image(self, idx, img, current_y, canvas_width):
        # Remove placeholder
        self.canvas.delete(f"ph_{idx}")
        self.canvas.delete(f"txt_{idx}")

        # Create PhotoImage
        tk_img = ImageTk.PhotoImage(img)
        self.view_images[idx] = tk_img  # Cache

        # Center horizontally
        # Recalculate based on current canvas_width?
        # Actually canvas_width passed from background might be stale if resized rapidly?
        # But we are in main thread now.
        # Let's use current canvas width for positioning if possible, or the one used for generation.
        # If we use current, but image was generated for old width, centering might look weird if mismatch large.
        # But Phase 1 layout used 'canvas_width' to determine 'target_width'.
        # 'img' width is 'target_width'.

        x = (canvas_width - img.width) // 2
        # Ensure positive
        if x < 0:
            x = 0

        # Draw image
        self.canvas.create_image(
            x, current_y, image=tk_img, anchor="nw", tags=f"page_{idx}"
        )

        # Draw page number overlay (top right of page)
        self.canvas.create_text(
            x + img.width - 5,
            current_y + 5,
            text=f"{idx+1}",
            fill="white",
            anchor="ne",
            font=("Arial", 14, "bold"),
            tags=f"page_{idx}",
        )

    def _update_view(self):
        # Re-trigger layout calculation (e.g. after zoom)
        # Cancel old thread safely
        if hasattr(self, "loading_stop_event"):
            self.loading_stop_event.set()

        self.canvas.delete("all")
        self.view_images = {}

        # Delay slightly to allow layout to settle if called rapidly
        self.after(10, self._calculate_layout_and_start_render)

    def _bind_mouse_scroll(self, event):
        self.canvas.bind_all("<MouseWheel>", self._on_mouse_scroll)
        self.canvas.bind_all("<Button-4>", self._on_mouse_scroll)
        self.canvas.bind_all("<Button-5>", self._on_mouse_scroll)

    def _unbind_mouse_scroll(self, event):
        self.canvas.unbind_all("<MouseWheel>")
        self.canvas.unbind_all("<Button-4>")
        self.canvas.unbind_all("<Button-5>")

    def _on_mouse_scroll(self, event):
        if event.state & 0x0004:
            return

        if event.num == 5 or event.delta < 0:
            self.canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:
            self.canvas.yview_scroll(-1, "units")

    def _on_zoom_wheel(self, event):
        if event.delta > 0:
            self.user_zoom *= 1.1
        else:
            self.user_zoom /= 1.1

        if self.user_zoom < 0.2:
            self.user_zoom = 0.2
        if self.user_zoom > 5.0:
            self.user_zoom = 5.0

        self._update_view()
        return "break"

    def _on_mouse_move(self, event):
        wx = self.canvas.canvasx(event.x)
        wy = self.canvas.canvasy(event.y)
        self.canvas.delete("highlight")

        cv_w = self.canvas.winfo_width()

        for y, h, idx in self.page_layouts:
            # Re-estimate box for highlight
            # We assume centered
            # This estimate needs to match the render logic 'target_width'
            target_w = (cv_w - 40) * self.user_zoom
            if target_w < 100:
                target_w = 100

            x1 = (cv_w - target_w) // 2
            if x1 < 0:
                x1 = 0
            x2 = x1 + target_w
            y1 = y
            y2 = y + h

            if x1 <= wx <= x2 and y1 <= wy <= y2:
                self.canvas.create_rectangle(
                    x1, y1, x2, y2, outline="#00ff00", width=4, tags="highlight"
                )
                break

    def _on_canvas_click(self, event):
        wx = self.canvas.canvasx(event.x)
        wy = self.canvas.canvasy(event.y)

        cv_w = self.canvas.winfo_width()
        for y, h, idx in self.page_layouts:
            target_w = (cv_w - 40) * self.user_zoom
            if target_w < 100:
                target_w = 100

            x1 = (cv_w - target_w) // 2
            if x1 < 0:
                x1 = 0
            x2 = x1 + target_w
            y1 = y
            y2 = y + h

            if x1 <= wx <= x2 and y1 <= wy <= y2:
                self._add_page_by_index(idx)
                break

    def _add_page_by_index(self, page_idx):
        if not self.current_pdf_renderer:
            return
        item = {
            "file": self.current_pdf_renderer.doc_path,
            "page": page_idx,
            "orig_name": self.lbl_file_name.cget("text"),
        }
        self.cheatsheet_items.append(item)
        self._update_cheatsheet_cache_and_refresh()

    def _undo_last(self):
        if self.cheatsheet_items:
            self.cheatsheet_items.pop()
            self._update_cheatsheet_cache_and_refresh()

    def _clear(self):
        self.cheatsheet_items = []
        self._update_cheatsheet_cache_and_refresh()

    def _update_cheatsheet_cache_and_refresh(self):
        if not self.cheatsheet_items:
            self.cached_preview_pages = []
        else:
            self.cached_preview_pages = self._render_pages()

        self._refresh_preview()

    # --- PREVIEW CANVAS LOGIC ---
    def _refresh_preview(self):
        self.preview_canvas.delete("all")
        self.preview_tk_images = {}

        if not self.cached_preview_pages:
            self.preview_canvas.create_text(
                150, 100, text="Empty Cheatsheet", fill="white", font=("Arial", 16)
            )
            return

        current_y = 20
        spacing = 20
        max_w = 0

        for i, p in enumerate(self.cached_preview_pages):
            # Calculate display size based on zoom
            target_w = int(p.width * self.preview_zoom)
            target_h = int(p.height * self.preview_zoom)

            p_resized = p.resize((target_w, target_h), Image.Resampling.LANCZOS)
            tk_img = ImageTk.PhotoImage(p_resized)
            self.preview_tk_images[i] = tk_img  # Keep ref

            x = 20  # fixed left margin? or centered?
            # Let's simple left align for now, or center in scrollregion

            self.preview_canvas.create_image(x, current_y, image=tk_img, anchor="nw")
            self.preview_canvas.create_text(
                x, current_y - 10, text=f"Page {i+1}", fill="white", anchor="nw"
            )

            current_y += target_h + spacing
            if target_w + 40 > max_w:
                max_w = target_w + 40

        self.preview_canvas.configure(scrollregion=(0, 0, max_w, current_y))

    def _on_preview_scroll(self, event):
        if event.state & 0x0004:
            return  # Zoom handled separately
        if event.num == 5 or event.delta < 0:
            self.preview_canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta > 0:
            self.preview_canvas.yview_scroll(-1, "units")

    def _on_preview_zoom_wheel(self, event):
        if event.delta > 0:
            self.preview_zoom *= 1.1
        else:
            self.preview_zoom /= 1.1

        if self.preview_zoom < 0.1:
            self.preview_zoom = 0.1
        if self.preview_zoom > 3.0:
            self.preview_zoom = 3.0

        self._refresh_preview()
        return "break"

    def _on_preview_resize(self, event):
        pass  # No auto-resize needed now

    def _on_canvas_resize(self, event):
        if self.center_resize_timer:
            self.after_cancel(self.center_resize_timer)
        self.center_resize_timer = self.after(300, self._update_view)

    def _prev(self):  # unused but kept for compatibility
        if self.current_pdf_renderer and self.current_page_index > 0:
            self.current_page_index -= 1
            self._update_view()

    def _next(self):  # unused
        if (
            self.current_pdf_renderer
            and self.current_page_index < self.current_pdf_renderer.get_page_count() - 1
        ):
            self.current_page_index += 1
            self._update_view()

    def _jump(self, event=None):
        try:
            p = int(self.ent_page.get())
            if (
                self.current_pdf_renderer
                and 1 <= p <= self.current_pdf_renderer.get_page_count()
            ):
                target_idx = p - 1
                for y, h, i in self.page_layouts:
                    if i == target_idx:
                        _, _, _, scroll_h = self.canvas.bbox("all") or (0, 0, 0, 1000)
                        bounds = self.canvas.cget("scrollregion").split()
                        if bounds:
                            total_h = float(bounds[3])
                            if total_h > 0:
                                self.canvas.yview_moveto(y / total_h)
                        break
        except:
            pass

    def _render_pages(self):
        if not self.cheatsheet_items:
            return []

        full_a4_w, full_a4_h = A4_WIDTH, A4_HEIGHT

        pg_imgs = []
        pg_img = Image.new("RGB", (full_a4_w, full_a4_h), "white")
        draw = ImageDraw.Draw(pg_img)

        # User requested appropriate margins
        PAGE_MARGIN_V = 50
        PAGE_MARGIN_H = 40
        COL_GUTTER = 20
        ITEM_GUTTER = 0

        col_w = (full_a4_w - 2 * PAGE_MARGIN_H - (COLUMNS - 1) * COL_GUTTER) // COLUMNS

        cur_col = 0
        cur_x = PAGE_MARGIN_H
        cur_y = PAGE_MARGIN_V

        def draw_grid(drawing, img_h):
            line_y_top = 0
            line_y_btm = img_h
            for c in range(1, COLUMNS):
                lx = PAGE_MARGIN_H + c * col_w + (c - 1) * COL_GUTTER + COL_GUTTER // 2
                drawing.line([lx, line_y_top, lx, line_y_btm], fill="black", width=2)

        draw_grid(draw, full_a4_h)

        for i_item, item in enumerate(self.cheatsheet_items):
            try:
                doc = fitz.open(item["file"])
                p = doc.load_page(item["page"])
                zoom = col_w / p.rect.width
                pix = p.get_pixmap(matrix=fitz.Matrix(zoom, zoom))
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                # Check if item has a stored split preference
                # If not, and it overflows, we do not ask.
                if cur_y + img.height > (full_a4_h - PAGE_MARGIN_V):
                    cur_col += 1
                    cur_y = PAGE_MARGIN_V
                    if cur_col >= COLUMNS:
                        pg_imgs.append(pg_img)
                        pg_img = Image.new("RGB", (full_a4_w, full_a4_h), "white")
                        draw = ImageDraw.Draw(pg_img)
                        draw_grid(draw, full_a4_h)
                        cur_col = 0
                
                cur_x = PAGE_MARGIN_H + cur_col * (col_w + COL_GUTTER)
                pg_img.paste(img, (int(cur_x), int(cur_y)))
                cur_y += img.height + ITEM_GUTTER
                doc.close()
            except:
                pass

        pg_imgs.append(pg_img)
        return pg_imgs

    def _export(self):
        if not self.cheatsheet_items:
            messagebox.showwarning("Warning", "Empty!")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF文件", "*.pdf")])
        if not save_path:
            return

        out_doc = fitz.open()
        
        MARGIN_V = 20
        MARGIN_H = 15
        GUTTER = 10
        COL_WIDTH = (A4_W_PTS - 2 * MARGIN_H - (COLUMNS - 1) * GUTTER) / COLUMNS

        def add_target_page():
            p = out_doc.new_page(width=A4_W_PTS, height=A4_H_PTS)
            for c in range(1, COLUMNS):
                lx = MARGIN_H + c * COL_WIDTH + (c - 1) * GUTTER + GUTTER / 2
                p.draw_line((lx, 0), (lx, A4_H_PTS), color=(0.9, 0.9, 0.9), width=0.3)
            return p

        current_page = add_target_page()
        cur_col = 0
        cur_y = MARGIN_V

        for item in self.cheatsheet_items:
            try:
                src_doc = fitz.open(item["file"])
                src_page = src_doc.load_page(item["page"])
                
                scale = COL_WIDTH / src_page.rect.width
                target_h = src_page.rect.height * scale

                if cur_y + target_h > (A4_H_PTS - MARGIN_V):
                    cur_col += 1
                    cur_y = MARGIN_V
                    
                    if cur_col >= COLUMNS:
                        current_page = add_target_page()
                        cur_col = 0
                
                cur_x = MARGIN_H + cur_col * (COL_WIDTH + GUTTER)
                target_rect = fitz.Rect(cur_x, cur_y, cur_x + COL_WIDTH, cur_y + target_h)

                current_page.show_pdf_page(target_rect, src_doc, item["page"])
                
                cur_y += target_h + 5
                src_doc.close()
            except Exception as e:
                print(f"Export item error: {e}")

        try:
            out_doc.save(save_path, garbage=3, deflate=True)
            out_doc.close()
            messagebox.showinfo("Done", "Exported!")
        except Exception as e:
            messagebox.showerror("Error", f"failed: {e}")


if __name__ == "__main__":
    check_deps()
    app = CheatSheetMaker()
    app.mainloop()
