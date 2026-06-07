import math
import os
import platform
import sys
from dataclasses import dataclass

try:
    import tkinter as tk
    import tkinter.ttk as ttk
    from tkinter import filedialog, messagebox
except Exception as exc:
    print("tkinter is not available. On macOS Homebrew Python, install python-tk for your Python version.")
    print(f"Original error: {exc}")
    sys.exit(1)

MISSING_DEPS = []

try:
    import customtkinter as ctk
except ImportError:
    MISSING_DEPS.append("customtkinter")
    ctk = None

try:
    import fitz
except ImportError:
    MISSING_DEPS.append("pymupdf")
    fitz = None

try:
    from PIL import Image, ImageDraw, ImageTk
except ImportError:
    MISSING_DEPS.append("Pillow")
    Image = None
    ImageDraw = None
    ImageTk = None


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
RES_DIR = os.path.join(BASE_DIR, "res")
CACHE_DIR = os.path.join(RES_DIR, ".cache")

A4_W_PTS, A4_H_PTS = 595.0, 842.0
EXPORT_COLUMNS = 4
EXPORT_MARGIN_V = 20.0
EXPORT_MARGIN_H = 15.0
EXPORT_GUTTER = 10.0
EXPORT_ITEM_GAP = 5.0

PREVIEW_RENDER_SCALE = 3
PREVIEW_BASE_WIDTH = 760
PREVIEW_WIDTH = PREVIEW_BASE_WIDTH * PREVIEW_RENDER_SCALE
PREVIEW_HEIGHT = int(PREVIEW_WIDTH * A4_H_PTS / A4_W_PTS)

PPT_CLIENT = None
PPT_CLIENT_KIND = None

try:
    import win32com.client as _ppt_client

    PPT_CLIENT = _ppt_client
    PPT_CLIENT_KIND = "win32com"
except ImportError:
    try:
        import comtypes.client as _ppt_client

        PPT_CLIENT = _ppt_client
        PPT_CLIENT_KIND = "comtypes"
    except ImportError:
        PPT_CLIENT = None

PPT_AVAILABLE = platform.system() == "Windows" and PPT_CLIENT is not None

THEME = {
    "bg": "#0f1217",
    "panel": "#171b22",
    "panel_alt": "#1d222b",
    "surface": "#11151b",
    "surface_alt": "#202631",
    "canvas": "#0b0f14",
    "border": "#2c3440",
    "border_soft": "#222a34",
    "text": "#edf1f6",
    "muted": "#9aa4b2",
    "muted_soft": "#687383",
    "blue": "#3f7ec3",
    "blue_hover": "#4d8fd6",
    "green": "#2f9366",
    "green_hover": "#3ba978",
    "red": "#8f3a3a",
    "red_hover": "#a54848",
    "ghost": "#252c36",
    "ghost_hover": "#303846",
    "input": "#141920",
    "selection": "#315f93",
}

BUTTON_STYLES = {
    "primary": ("blue", "blue_hover", "text"),
    "success": ("green", "green_hover", "text"),
    "danger": ("red", "red_hover", "text"),
    "ghost": ("ghost", "ghost_hover", "text"),
}


@dataclass
class CheatItem:
    pdf_path: str
    page_index: int
    source_name: str
    scale: float = 1.0
    crop_left: float = 0.0
    crop_top: float = 0.0
    crop_right: float = 0.0
    crop_bottom: float = 0.0

    @property
    def label(self):
        crop_total = self.crop_left + self.crop_top + self.crop_right + self.crop_bottom
        suffix = ""
        if abs(self.scale - 1.0) > 0.005 or crop_total > 0.005:
            suffix = f"  {int(self.scale * 100)}%"
            if crop_total > 0.005:
                suffix += " crop"
        return f"{self.source_name}  p.{self.page_index + 1}{suffix}"


class PDFDocument:
    def __init__(self, path):
        self.path = path
        self.doc = fitz.open(path)

    @property
    def page_count(self):
        return len(self.doc)

    def page_rect(self, page_index):
        return self.doc.load_page(page_index).rect

    def render_page(self, page_index, target_width):
        page = self.doc.load_page(page_index)
        scale = max(0.05, min(float(target_width) / page.rect.width, 6.0))
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

    def close(self):
        if self.doc:
            self.doc.close()
            self.doc = None


def check_deps():
    if not MISSING_DEPS:
        return

    msg = (
        "Missing required libraries:\n"
        + "\n".join(MISSING_DEPS)
        + "\n\nInstall them with:\npython -m pip install -r requirements.txt"
    )

    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Missing Dependencies", msg)
        root.destroy()
    except Exception:
        print(msg)

    sys.exit(1)


def ensure_dirs():
    os.makedirs(RES_DIR, exist_ok=True)
    os.makedirs(CACHE_DIR, exist_ok=True)


def convert_ppt_to_pdf(ppt_path):
    if not PPT_AVAILABLE:
        return None

    ensure_dirs()
    base_name = os.path.splitext(os.path.basename(ppt_path))[0]
    pdf_path = os.path.join(CACHE_DIR, f"{base_name}.pdf")

    if os.path.exists(pdf_path) and os.path.getmtime(pdf_path) >= os.path.getmtime(ppt_path):
        return pdf_path

    app = None
    pres = None

    try:
        if PPT_CLIENT_KIND == "win32com":
            app = PPT_CLIENT.Dispatch("PowerPoint.Application")
        else:
            app = PPT_CLIENT.CreateObject("PowerPoint.Application")

        try:
            pres = app.Presentations.Open(ppt_path, WithWindow=False)
        except Exception:
            pres = app.Presentations.Open(ppt_path)

        pres.SaveAs(pdf_path, 32)
        return pdf_path
    except Exception as exc:
        print(f"PPT conversion failed: {exc}")
        return None
    finally:
        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass


AppBase = ctk.CTk if ctk is not None else tk.Tk


class CheatSheetMaker(AppBase):
    def __init__(self):
        super().__init__()

        self.title("Cheatsheet Maker")
        self.geometry("1420x900")
        self.minsize(1280, 720)

        self.current_doc = None
        self.current_source_name = ""
        self.current_page = 0
        self.reader_zoom = 1.0
        self.reader_image_refs = {}
        self.reader_page_boxes = []
        self.reader_render_after = None
        self.reader_restore_y_fraction = None

        self.cheat_items = []
        self.preview_pages = []
        self.preview_refs = []
        self.preview_item_layouts = []
        self.preview_hit_boxes = []
        self.preview_zoom = 0.48
        self.preview_restore_y_fraction = None
        self.preview_rebuild_after = None
        self.pending_preview_select_index = None
        self.active_wheel_target = None
        self.property_syncing = False

        self._init_style()
        self._init_ui()
        self._load_files()
        self._set_status("Ready")

    def _init_style(self):
        ctk.set_appearance_mode("Dark")
        ctk.set_default_color_theme("blue")
        self.configure(fg_color=THEME["bg"])

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.layout("Resource.Treeview", [("Treeview.treearea", {"sticky": "nswe"})])
        style.configure(
            "Resource.Treeview",
            background=THEME["surface"],
            fieldbackground=THEME["surface"],
            foreground=THEME["text"],
            borderwidth=0,
            rowheight=30,
            font=("Arial", 13),
            relief="flat",
            padding=0,
        )
        style.map(
            "Resource.Treeview",
            background=[("selected", THEME["selection"])],
            foreground=[("selected", THEME["text"])],
        )

    def _button(self, parent, text, command, width=72, style="primary", **kwargs):
        fg_key, hover_key, text_key = BUTTON_STYLES[style]
        return ctk.CTkButton(
            parent,
            text=text,
            width=width,
            height=34,
            corner_radius=7,
            border_width=0,
            command=command,
            fg_color=THEME[fg_key],
            hover_color=THEME[hover_key],
            text_color=THEME[text_key],
            font=ctk.CTkFont(size=13, weight="bold"),
            **kwargs,
        )

    def _panel_label(self, parent, text, size=20):
        return ctk.CTkLabel(
            parent,
            text=text,
            font=ctk.CTkFont(size=size, weight="bold"),
            text_color=THEME["text"],
            anchor="w",
        )

    def _init_ui(self):
        self.grid_rowconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self.main_paned = tk.PanedWindow(
            self,
            orient="horizontal",
            bg=THEME["border_soft"],
            bd=0,
            sashwidth=8,
            sashrelief="flat",
            showhandle=False,
            opaqueresize=False,
        )
        self.main_paned.grid(row=0, column=0, sticky="nsew")

        self.left_panel = ctk.CTkFrame(self.main_paned, width=284, corner_radius=0, fg_color=THEME["panel"])
        self.left_panel.grid_propagate(False)

        self.reader_panel = ctk.CTkFrame(self.main_paned, corner_radius=0, fg_color=THEME["bg"])

        self.preview_panel = ctk.CTkFrame(self.main_paned, width=360, corner_radius=0, fg_color=THEME["bg"])
        self.preview_panel.grid_propagate(False)

        self.right_panel = ctk.CTkFrame(self.main_paned, width=380, corner_radius=0, fg_color=THEME["panel"])
        self.right_panel.grid_propagate(False)

        self.main_paned.add(self.left_panel, minsize=210)
        self.main_paned.add(self.reader_panel, minsize=440)
        self.main_paned.add(self.preview_panel, minsize=260)
        self.main_paned.add(self.right_panel, minsize=300)

        self.status_bar = ctk.CTkLabel(
            self,
            text="",
            anchor="w",
            height=28,
            fg_color=THEME["surface"],
            text_color=THEME["muted"],
            font=ctk.CTkFont(size=12),
        )
        self.status_bar.grid(row=1, column=0, sticky="ew")

        self._build_left_panel()
        self._build_reader_panel()
        self._build_right_panel()
        self._build_preview_panel()
        self._bind_global_wheel()
        self.after(120, self._place_initial_sashes)

    def _place_initial_sashes(self):
        try:
            total_w = self.main_paned.winfo_width()
            if total_w <= 0:
                return
            self.main_paned.sash_place(0, 284, 0)
            control_w = 380
            preview_w = 360
            self.main_paned.sash_place(1, max(724, total_w - control_w - preview_w), 0)
            self.main_paned.sash_place(2, max(984, total_w - control_w), 0)
        except tk.TclError:
            pass

    def _build_left_panel(self):
        self.left_panel.grid_rowconfigure(2, weight=1)
        self.left_panel.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self.left_panel, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        header.grid_columnconfigure(0, weight=1)

        title = self._panel_label(header, "Resources", size=22)
        title.grid(row=0, column=0, sticky="w")

        self.refresh_button = self._button(header, "Refresh", self._load_files, width=86, style="ghost")
        self.refresh_button.grid(row=0, column=1, sticky="e")

        self.res_label = ctk.CTkLabel(
            self.left_panel,
            text=f"Folder: {os.path.basename(RES_DIR)}",
            text_color=THEME["muted"],
            font=ctk.CTkFont(size=12),
            anchor="w",
        )
        self.res_label.grid(row=1, column=0, sticky="ew", padx=20, pady=(0, 10))

        tree_shell = ctk.CTkFrame(
            self.left_panel,
            fg_color=THEME["surface"],
            border_color=THEME["border_soft"],
            border_width=1,
            corner_radius=10,
        )
        tree_shell.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 16))
        tree_shell.grid_rowconfigure(0, weight=1)
        tree_shell.grid_columnconfigure(0, weight=1)

        self.file_tree = ttk.Treeview(
            tree_shell,
            show="tree",
            selectmode="browse",
            style="Resource.Treeview",
            columns=("path",),
            displaycolumns=(),
            height=20,
        )
        self.file_tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)
        self.file_tree.bind("<<TreeviewSelect>>", self._on_tree_select)

        tree_scroll = ctk.CTkScrollbar(
            tree_shell,
            command=self.file_tree.yview,
            fg_color=THEME["surface"],
            button_color=THEME["ghost"],
            button_hover_color=THEME["ghost_hover"],
            width=12,
        )
        tree_scroll.grid(row=0, column=1, sticky="ns", pady=10, padx=(6, 8))
        self.file_tree.configure(yscrollcommand=tree_scroll.set)

    def _build_reader_panel(self):
        self.reader_panel.grid_rowconfigure(2, weight=1)
        self.reader_panel.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(self.reader_panel, fg_color=THEME["bg"], corner_radius=0)
        top.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 8))
        top.grid_columnconfigure(0, weight=1)

        self.file_title = ctk.CTkLabel(
            top,
            text="Select a PDF or PPT file",
            font=ctk.CTkFont(size=20, weight="bold"),
            text_color=THEME["text"],
            anchor="w",
        )
        self.file_title.grid(row=0, column=0, sticky="ew")

        self.file_meta = ctk.CTkLabel(
            top,
            text="",
            text_color=THEME["muted"],
            font=ctk.CTkFont(size=12),
            anchor="w",
        )
        self.file_meta.grid(row=1, column=0, sticky="ew", pady=(2, 0))

        toolbar = ctk.CTkFrame(
            self.reader_panel,
            fg_color=THEME["panel"],
            corner_radius=10,
            border_color=THEME["border_soft"],
            border_width=1,
        )
        toolbar.grid(row=1, column=0, sticky="ew", padx=18, pady=(0, 12))

        self.prev_button = self._button(toolbar, "Prev", self._prev_page, width=58, style="ghost")
        self.prev_button.pack(side="left", padx=(12, 5), pady=10)

        self.page_entry = ctk.CTkEntry(
            toolbar,
            width=58,
            height=34,
            justify="center",
            fg_color=THEME["input"],
            border_color=THEME["border"],
            text_color=THEME["text"],
            corner_radius=7,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.page_entry.pack(side="left", padx=5, pady=9)
        self.page_entry.bind("<Return>", self._jump_to_page)

        self.page_count_label = ctk.CTkLabel(
            toolbar,
            text="/ 0",
            width=44,
            text_color=THEME["muted"],
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.page_count_label.pack(side="left", padx=(0, 8), pady=9)

        self.next_button = self._button(toolbar, "Next", self._next_page, width=58, style="ghost")
        self.next_button.pack(side="left", padx=5, pady=9)

        self.add_button = self._button(toolbar, "Add Page", self._add_current_page, width=104, style="success")
        self.add_button.pack(side="left", padx=(12, 5), pady=9)

        self.zoom_out_button = self._button(
            toolbar,
            "-",
            lambda: self._change_reader_zoom(0.9),
            width=36,
            style="ghost",
        )
        self.zoom_out_button.pack(side="right", padx=(5, 12), pady=9)

        self.zoom_label = ctk.CTkLabel(
            toolbar,
            text="100%",
            width=54,
            text_color=THEME["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.zoom_label.pack(side="right", padx=4, pady=9)

        self.zoom_in_button = self._button(
            toolbar,
            "+",
            lambda: self._change_reader_zoom(1.1),
            width=36,
            style="ghost",
        )
        self.zoom_in_button.pack(side="right", padx=4, pady=9)

        self.fit_button = self._button(toolbar, "Fit", self._fit_reader_width, width=48, style="ghost")
        self.fit_button.pack(side="right", padx=4, pady=9)

        canvas_shell = ctk.CTkFrame(
            self.reader_panel,
            fg_color=THEME["canvas"],
            corner_radius=12,
            border_color=THEME["border_soft"],
            border_width=1,
        )
        self.reader_canvas_shell = canvas_shell
        canvas_shell.grid(row=2, column=0, sticky="nsew", padx=18, pady=(0, 16))
        canvas_shell.grid_rowconfigure(0, weight=1)
        canvas_shell.grid_columnconfigure(0, weight=1)

        self.reader_canvas = tk.Canvas(canvas_shell, bg=THEME["canvas"], highlightthickness=0, bd=0)
        self.reader_canvas.grid(row=0, column=0, sticky="nsew", padx=(1, 0), pady=(1, 0))

        reader_y = ctk.CTkScrollbar(
            canvas_shell,
            orientation="vertical",
            command=self._reader_yview,
            fg_color=THEME["canvas"],
            button_color=THEME["ghost"],
            button_hover_color=THEME["ghost_hover"],
            width=12,
        )
        reader_y.grid(row=0, column=1, sticky="ns", padx=(6, 8), pady=8)
        reader_x = ctk.CTkScrollbar(
            canvas_shell,
            orientation="horizontal",
            command=self.reader_canvas.xview,
            fg_color=THEME["canvas"],
            button_color=THEME["ghost"],
            button_hover_color=THEME["ghost_hover"],
            height=12,
        )
        reader_x.grid(row=1, column=0, sticky="ew", padx=8, pady=(6, 8))
        self.reader_canvas.configure(yscrollcommand=reader_y.set, xscrollcommand=reader_x.set)

        self.reader_canvas.bind("<Configure>", self._schedule_reader_render)
        self.reader_canvas.bind("<Button-1>", self._on_reader_click)
        self.reader_canvas.bind("<Enter>", lambda _event: self._activate_wheel("reader"))
        self.reader_canvas.bind("<Leave>", lambda _event: self._deactivate_wheel("reader"))

        self._sync_reader_buttons()
        self._draw_reader_empty()

    def _build_right_panel(self):
        self.right_panel.grid_rowconfigure(2, weight=1)
        self.right_panel.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        header.grid_columnconfigure(0, weight=1)

        title = self._panel_label(header, "Controls", size=22)
        title.grid(row=0, column=0, sticky="w")

        self.item_count_label = ctk.CTkLabel(
            header,
            text="0 pages",
            text_color=THEME["muted"],
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.item_count_label.grid(row=0, column=1, sticky="e")

        actions = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        actions.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))
        for col in range(3):
            actions.grid_columnconfigure(col, weight=1)

        self.export_button = self._button(actions, "Export PDF", self._export, style="primary")
        self.export_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.undo_button = self._button(actions, "Undo", self._undo_last, style="ghost")
        self.undo_button.grid(row=0, column=1, sticky="ew", padx=5)

        self.clear_button = self._button(actions, "Clear", self._clear_items, style="danger")
        self.clear_button.grid(row=0, column=2, sticky="ew", padx=(5, 0))

        list_shell = ctk.CTkFrame(
            self.right_panel,
            fg_color=THEME["surface"],
            border_color=THEME["border_soft"],
            border_width=1,
            corner_radius=10,
        )
        list_shell.grid(row=2, column=0, sticky="nsew", padx=16, pady=(0, 10))
        list_shell.grid_columnconfigure(0, weight=1)
        list_shell.grid_rowconfigure(0, weight=1)

        self.item_list = tk.Listbox(
            list_shell,
            height=7,
            activestyle="none",
            bg=THEME["surface"],
            fg=THEME["text"],
            selectbackground=THEME["selection"],
            selectforeground=THEME["text"],
            highlightthickness=0,
            borderwidth=0,
            relief="flat",
            font=("Arial", 13),
            exportselection=False,
        )
        self.item_list.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        self.item_list.bind("<<ListboxSelect>>", self._on_item_list_select)

        item_scroll = ctk.CTkScrollbar(
            list_shell,
            command=self.item_list.yview,
            fg_color=THEME["surface"],
            button_color=THEME["ghost"],
            button_hover_color=THEME["ghost_hover"],
            width=12,
        )
        item_scroll.grid(row=0, column=1, sticky="ns", pady=12, padx=(6, 10))
        self.item_list.configure(yscrollcommand=item_scroll.set)

        item_controls = ctk.CTkFrame(self.right_panel, fg_color="transparent")
        item_controls.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 10))
        for col in range(3):
            item_controls.grid_columnconfigure(col, weight=1)

        self.remove_button = self._button(item_controls, "Remove", self._remove_selected_item, style="ghost")
        self.remove_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.up_button = self._button(item_controls, "Up", lambda: self._move_selected_item(-1), style="ghost")
        self.up_button.grid(row=0, column=1, sticky="ew", padx=5)

        self.down_button = self._button(item_controls, "Down", lambda: self._move_selected_item(1), style="ghost")
        self.down_button.grid(row=0, column=2, sticky="ew", padx=(5, 0))

        self._build_property_panel(self.right_panel, row=4, padx=16, pady=(0, 16))

        self._sync_item_buttons()

    def _build_preview_panel(self):
        self.preview_panel.grid_rowconfigure(1, weight=1)
        self.preview_panel.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(self.preview_panel, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew", padx=18, pady=(18, 8))
        header.grid_columnconfigure(0, weight=1)

        title = self._panel_label(header, "Preview", size=22)
        title.grid(row=0, column=0, sticky="w")

        preview_shell = ctk.CTkFrame(
            self.preview_panel,
            fg_color=THEME["canvas"],
            corner_radius=12,
            border_color=THEME["border_soft"],
            border_width=1,
        )
        self.preview_shell = preview_shell
        preview_shell.grid(row=1, column=0, sticky="nsew", padx=16, pady=(0, 16))
        preview_shell.grid_rowconfigure(1, weight=1)
        preview_shell.grid_columnconfigure(0, weight=1)

        preview_toolbar = ctk.CTkFrame(preview_shell, fg_color=THEME["canvas"], corner_radius=0)
        preview_toolbar.grid(row=0, column=0, columnspan=2, sticky="ew", padx=10, pady=(10, 4))

        preview_title = ctk.CTkLabel(
            preview_toolbar,
            text="Preview",
            anchor="w",
            text_color=THEME["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        preview_title.pack(side="left", padx=12, pady=8)

        self.preview_zoom_out = self._button(
            preview_toolbar, "-", lambda: self._change_preview_zoom(0.9), width=34, style="ghost"
        )
        self.preview_zoom_out.pack(side="right", padx=(4, 12), pady=8)

        self.preview_zoom_label = ctk.CTkLabel(
            preview_toolbar,
            text="48%",
            width=48,
            text_color=THEME["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.preview_zoom_label.pack(side="right", padx=4, pady=8)

        self.preview_zoom_in = self._button(
            preview_toolbar, "+", lambda: self._change_preview_zoom(1.1), width=34, style="ghost"
        )
        self.preview_zoom_in.pack(side="right", padx=4, pady=8)

        self.preview_canvas = tk.Canvas(preview_shell, bg=THEME["canvas"], highlightthickness=0, bd=0)
        self.preview_canvas.grid(row=1, column=0, sticky="nsew", padx=(1, 0), pady=(0, 1))
        preview_y = ctk.CTkScrollbar(
            preview_shell,
            command=self.preview_canvas.yview,
            fg_color=THEME["canvas"],
            button_color=THEME["ghost"],
            button_hover_color=THEME["ghost_hover"],
            width=12,
        )
        preview_y.grid(row=1, column=1, sticky="ns", padx=(6, 8), pady=(4, 8))
        self.preview_canvas.configure(yscrollcommand=preview_y.set)
        self.preview_canvas.bind("<Button-1>", self._on_preview_click)
        self.preview_canvas.bind("<Enter>", lambda _event: self._activate_wheel("preview"))
        self.preview_canvas.bind("<Leave>", lambda _event: self._deactivate_wheel("preview"))

        self._refresh_preview()

    def _build_property_panel(self, parent, row=0, padx=0, pady=(0, 0)):
        panel = ctk.CTkFrame(
            parent,
            fg_color=THEME["surface"],
            border_color=THEME["border_soft"],
            border_width=1,
            corner_radius=10,
        )
        panel.grid(row=row, column=0, sticky="ew", padx=padx, pady=pady)
        panel.grid_columnconfigure(0, weight=1)

        self.prop_title = ctk.CTkLabel(
            panel,
            text="Item Properties",
            text_color=THEME["text"],
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        )
        self.prop_title.grid(row=0, column=0, sticky="ew", padx=12, pady=(10, 0))

        self.prop_selected_label = ctk.CTkLabel(
            panel,
            text="No item selected",
            text_color=THEME["muted"],
            font=ctk.CTkFont(size=12),
            anchor="w",
        )
        self.prop_selected_label.grid(row=1, column=0, sticky="ew", padx=12, pady=(2, 8))

        scale_row = ctk.CTkFrame(panel, fg_color="transparent")
        scale_row.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 4))
        scale_row.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(scale_row, text="Scale", text_color=THEME["muted"], width=46, anchor="w").grid(
            row=0, column=0, sticky="w"
        )
        self.scale_value_label = ctk.CTkLabel(
            scale_row,
            text="100%",
            text_color=THEME["text"],
            width=48,
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.scale_value_label.grid(row=0, column=2, sticky="e")

        self.scale_slider = ctk.CTkSlider(
            scale_row,
            from_=35,
            to=200,
            number_of_steps=165,
            command=self._on_scale_slider,
            fg_color=THEME["ghost"],
            progress_color=THEME["blue"],
            button_color=THEME["blue_hover"],
            button_hover_color=THEME["blue_hover"],
        )
        self.scale_slider.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(4, 0))

        crop_frame = ctk.CTkFrame(panel, fg_color="transparent")
        crop_frame.grid(row=3, column=0, sticky="ew", padx=12, pady=(6, 4))
        crop_frame.grid_columnconfigure(1, weight=1)

        self.crop_sliders = {}
        self.crop_value_labels = {}
        crop_rows = [
            ("crop_left", "Left"),
            ("crop_top", "Top"),
            ("crop_right", "Right"),
            ("crop_bottom", "Bottom"),
        ]
        for row, (key, label) in enumerate(crop_rows):
            ctk.CTkLabel(crop_frame, text=label, text_color=THEME["muted"], width=48, anchor="w").grid(
                row=row, column=0, sticky="w", pady=2
            )
            slider = ctk.CTkSlider(
                crop_frame,
                from_=0,
                to=70,
                number_of_steps=70,
                command=lambda value, crop_key=key: self._on_crop_slider(crop_key, value),
                fg_color=THEME["ghost"],
                progress_color=THEME["green"],
                button_color=THEME["green_hover"],
                button_hover_color=THEME["green_hover"],
            )
            slider.grid(row=row, column=1, sticky="ew", padx=8, pady=2)
            value_label = ctk.CTkLabel(
                crop_frame,
                text="0%",
                text_color=THEME["text"],
                width=38,
                font=ctk.CTkFont(size=12, weight="bold"),
            )
            value_label.grid(row=row, column=2, sticky="e", pady=2)
            self.crop_sliders[key] = slider
            self.crop_value_labels[key] = value_label

        prop_actions = ctk.CTkFrame(panel, fg_color="transparent")
        prop_actions.grid(row=4, column=0, sticky="ew", padx=12, pady=(4, 10))
        prop_actions.grid_columnconfigure(0, weight=1)
        prop_actions.grid_columnconfigure(1, weight=1)

        self.scale_down_button = self._button(
            prop_actions,
            "-10%",
            lambda: self._adjust_selected_scale(-0.1),
            style="ghost",
        )
        self.scale_down_button.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.scale_up_button = self._button(
            prop_actions,
            "+10%",
            lambda: self._adjust_selected_scale(0.1),
            style="ghost",
        )
        self.scale_up_button.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        self.reset_transform_button = self._button(
            prop_actions,
            "Reset",
            self._reset_selected_transform,
            style="ghost",
        )
        self.reset_transform_button.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(8, 0))

    def _set_status(self, text):
        self.status_bar.configure(text=f"  {text}")

    def _bind_global_wheel(self):
        scroll_sequences = (
            "<MouseWheel>",
            "<Button-4>",
            "<Button-5>",
            "<Shift-MouseWheel>",
            "<Shift-Button-4>",
            "<Shift-Button-5>",
        )
        zoom_sequences = (
            "<Control-MouseWheel>",
            "<Control-Button-4>",
            "<Control-Button-5>",
        )

        for sequence in scroll_sequences:
            self._safe_bind_all(sequence, self._dispatch_wheel)
        for sequence in zoom_sequences:
            self._safe_bind_all(sequence, self._dispatch_zoom_wheel)

    def _safe_bind_all(self, sequence, callback):
        try:
            self.bind_all(sequence, callback)
        except tk.TclError:
            pass

    def _activate_wheel(self, target):
        self.active_wheel_target = target

    def _deactivate_wheel(self, target):
        if self.active_wheel_target != target:
            return
        self.active_wheel_target = None

    def _dispatch_wheel(self, event):
        target = self._wheel_target_from_event(event) or self.active_wheel_target
        if target == "reader":
            return self._on_reader_wheel(event)
        if target == "preview":
            return self._on_preview_wheel(event)
        return None

    def _dispatch_zoom_wheel(self, event):
        target = self._wheel_target_from_event(event) or self.active_wheel_target
        if target == "reader":
            return self._on_reader_wheel(event, force_zoom=True)
        if target == "preview":
            return self._on_preview_wheel(event, force_zoom=True)
        return None

    def _wheel_target_from_event(self, event):
        widget = None
        if hasattr(event, "x_root") and hasattr(event, "y_root"):
            try:
                widget = self.winfo_containing(event.x_root, event.y_root)
            except tk.TclError:
                widget = None
        widget = widget or getattr(event, "widget", None)

        if self._is_widget_descendant(widget, self.reader_canvas_shell):
            return "reader"
        if self._is_widget_descendant(widget, self.preview_shell):
            return "preview"
        return None

    def _is_widget_descendant(self, widget, ancestor):
        while widget is not None:
            if widget == ancestor:
                return True
            try:
                widget = widget.master
            except AttributeError:
                return False
        return False

    def _load_files(self):
        ensure_dirs()
        self.file_tree.delete(*self.file_tree.get_children())
        self._populate_tree(RES_DIR, "")
        self._set_status(f"Loaded resources from {RES_DIR}")

    def _populate_tree(self, path, parent):
        try:
            entries = sorted(
                os.listdir(path),
                key=lambda name: (not os.path.isdir(os.path.join(path, name)), name.lower()),
            )
        except OSError as exc:
            self._set_status(f"Cannot read {path}: {exc}")
            return

        for name in entries:
            if name.startswith("."):
                continue

            full_path = os.path.join(path, name)
            if os.path.isdir(full_path):
                node = self.file_tree.insert(parent, "end", text=name, open=False)
                self._populate_tree(full_path, node)
                if not self.file_tree.get_children(node):
                    self.file_tree.delete(node)
                continue

            ext = os.path.splitext(name)[1].lower()
            if ext in {".pdf", ".ppt", ".pptx"}:
                self.file_tree.insert(parent, "end", text=name, values=(full_path,))

    def _on_tree_select(self, _event=None):
        selected = self.file_tree.selection()
        if not selected:
            return

        values = self.file_tree.item(selected[0], "values")
        if values:
            self._open_resource(values[0])

    def _open_resource(self, path):
        ext = os.path.splitext(path)[1].lower()
        pdf_path = path
        source_name = os.path.basename(path)

        if ext in {".ppt", ".pptx"}:
            if not PPT_AVAILABLE:
                messagebox.showwarning(
                    "PPT conversion unavailable",
                    "PPT conversion currently requires Windows, Microsoft PowerPoint, and pywin32/comtypes.",
                )
                return

            self._set_status("Converting PPT to PDF...")
            self.update_idletasks()
            pdf_path = convert_ppt_to_pdf(path)
            if not pdf_path:
                messagebox.showerror("Conversion Failed", "Could not convert this PPT file to PDF.")
                self._set_status("PPT conversion failed")
                return

        try:
            new_doc = PDFDocument(pdf_path)
        except Exception as exc:
            messagebox.showerror("Open Failed", str(exc))
            self._set_status(f"Open failed: {exc}")
            return

        if self.current_doc:
            self.current_doc.close()

        self.current_doc = new_doc
        self.current_source_name = source_name
        self.current_page = 0
        self.reader_zoom = 1.0
        self.file_title.configure(text=source_name)
        self.file_meta.configure(text=f"{self.current_doc.page_count} pages")
        self._sync_reader_buttons()
        self._schedule_reader_render()
        self._set_status(f"Opened {source_name}")

    def _sync_reader_buttons(self):
        has_doc = self.current_doc is not None and self.current_doc.page_count > 0
        total = self.current_doc.page_count if has_doc else 0

        self.page_entry.configure(state="normal")
        self.page_entry.delete(0, "end")
        self.page_entry.insert(0, str(self.current_page + 1) if has_doc else "")
        self.page_entry.configure(state="normal" if has_doc else "disabled")

        self.page_count_label.configure(text=f"/ {total}")

        self.prev_button.configure(state="normal" if has_doc and self.current_page > 0 else "disabled")
        self.next_button.configure(state="normal" if has_doc and self.current_page < total - 1 else "disabled")
        self.add_button.configure(state="normal" if has_doc else "disabled")
        self.fit_button.configure(state="normal" if has_doc else "disabled")
        self.zoom_in_button.configure(state="normal" if has_doc else "disabled")
        self.zoom_out_button.configure(state="normal" if has_doc else "disabled")
        self.zoom_label.configure(text=f"{int(self.reader_zoom * 100)}%")

    def _schedule_reader_render(self, _event=None):
        if self.reader_render_after:
            self.after_cancel(self.reader_render_after)
        self.reader_render_after = self.after(80, self._render_reader_pages)

    def _draw_reader_empty(self):
        self.reader_canvas.delete("all")
        width = max(self.reader_canvas.winfo_width(), 640)
        height = max(self.reader_canvas.winfo_height(), 420)
        card_w = min(420, width - 80)
        card_h = 108
        x0 = (width - card_w) / 2
        y0 = (height - card_h) / 2
        self.reader_canvas.create_rectangle(
            x0,
            y0,
            x0 + card_w,
            y0 + card_h,
            fill=THEME["surface"],
            outline=THEME["border_soft"],
            width=1,
        )
        self.reader_canvas.create_text(
            width / 2,
            height / 2 - 12,
            text="Select a file from Resources",
            fill=THEME["text"],
            font=("Arial", 17, "bold"),
        )
        self.reader_canvas.create_text(
            width / 2,
            height / 2 + 18,
            text="PDF pages will appear here",
            fill=THEME["muted_soft"],
            font=("Arial", 12),
        )
        self.reader_canvas.configure(scrollregion=(0, 0, width, height))

    def _render_reader_pages(self):
        self.reader_render_after = None
        self.reader_canvas.delete("all")
        self.reader_image_refs = {}
        self.reader_page_boxes = []

        if not self.current_doc:
            self._draw_reader_empty()
            return

        canvas_w = max(self.reader_canvas.winfo_width(), 640)
        canvas_h = max(self.reader_canvas.winfo_height(), 420)
        target_width = int(max(180, min((canvas_w - 72) * self.reader_zoom, 2600)))
        current_y = 28
        max_right = canvas_w
        page_gap = 28

        for page_index in range(self.current_doc.page_count):
            try:
                image = self.current_doc.render_page(page_index, target_width)
            except Exception as exc:
                self.reader_canvas.create_text(
                    canvas_w / 2,
                    current_y + 40,
                    text=f"Page {page_index + 1} render failed: {exc}",
                    fill="#ff9a9a",
                    font=("Arial", 15),
                )
                current_y += 96
                continue

            tk_image = ImageTk.PhotoImage(image)
            self.reader_image_refs[page_index] = tk_image

            x = max(30, (canvas_w - image.width) // 2)
            y = current_y

            self.reader_canvas.create_rectangle(
                x - 8,
                y - 8,
                x + image.width + 8,
                y + image.height + 10,
                fill="#05070a",
                outline=THEME["border_soft"],
                width=1,
            )
            self.reader_canvas.create_rectangle(
                x - 3,
                y - 3,
                x + image.width + 3,
                y + image.height + 3,
                fill="#f8fafc",
                outline="#3d4653",
                width=1,
            )
            self.reader_canvas.create_image(x, y, image=tk_image, anchor="nw")
            self.reader_canvas.create_rectangle(
                x,
                y,
                x + image.width,
                y + image.height,
                outline="#cfd6df",
                width=1,
            )
            self.reader_canvas.create_text(
                x + image.width - 10,
                y + 10,
                text=f"{page_index + 1}",
                fill="#27313d",
                anchor="ne",
                font=("Arial", 14, "bold"),
            )

            self.reader_page_boxes.append((page_index, x, y, x + image.width, y + image.height))
            current_y += image.height + page_gap
            max_right = max(max_right, x + image.width + 30)

        scroll_h = max(canvas_h, current_y)
        self.reader_canvas.configure(scrollregion=(0, 0, max_right, scroll_h))
        if self.reader_restore_y_fraction is not None:
            self.reader_canvas.yview_moveto(self.reader_restore_y_fraction)
            self.reader_restore_y_fraction = None
            self._update_current_page_from_view()
        else:
            self._scroll_to_reader_page(self.current_page, update_buttons=False)
        self._sync_reader_buttons()

    def _wheel_units(self, event, amount=3):
        if getattr(event, "num", None) == 4:
            return -amount
        if getattr(event, "num", None) == 5:
            return amount
        if event.delta > 0:
            return -amount
        if event.delta < 0:
            return amount
        return 0

    def _ctrl_pressed(self, event):
        return bool(getattr(event, "state", 0) & 0x0004)

    def _shift_pressed(self, event):
        return bool(getattr(event, "state", 0) & 0x0001)

    def _wheel_pixels(self, event):
        if getattr(event, "num", None) == 4:
            return -90
        if getattr(event, "num", None) == 5:
            return 90

        delta = getattr(event, "delta", 0)
        if not delta:
            return 0

        if platform.system() == "Darwin":
            return -delta * 90

        return -(delta / 120) * 90

    def _wheel_zoom_factor(self, event):
        if getattr(event, "num", None) == 4:
            steps = 1.0
        elif getattr(event, "num", None) == 5:
            steps = -1.0
        else:
            delta = getattr(event, "delta", 0)
            if not delta:
                return 1.0
            if platform.system() == "Darwin":
                steps = delta
            else:
                steps = delta / 120

        return max(0.82, min(1.22, 1.1**steps))

    def _canvas_view_fraction(self, canvas):
        top, _bottom = canvas.yview()
        return top

    def _scroll_canvas_pixels(self, canvas, dx=0, dy=0):
        scrollregion = canvas.cget("scrollregion").split()
        if len(scrollregion) != 4:
            return False

        x0, y0, x1, y1 = [float(value) for value in scrollregion]
        total_w = max(1.0, x1 - x0)
        total_h = max(1.0, y1 - y0)
        view_w = max(1, canvas.winfo_width())
        view_h = max(1, canvas.winfo_height())
        moved = False

        if dx and total_w > view_w:
            left_frac, _right_frac = canvas.xview()
            left_px = left_frac * total_w
            new_left = max(0.0, min(left_px + dx, total_w - view_w))
            canvas.xview_moveto(new_left / total_w)
            moved = True

        if dy and total_h > view_h:
            top_frac, _bottom_frac = canvas.yview()
            top_px = top_frac * total_h
            new_top = max(0.0, min(top_px + dy, total_h - view_h))
            canvas.yview_moveto(new_top / total_h)
            moved = True

        return moved

    def _reader_yview(self, *args):
        self.reader_canvas.yview(*args)
        self.after_idle(self._update_current_page_from_view)

    def _on_reader_wheel(self, event, force_zoom=False):
        if force_zoom or self._ctrl_pressed(event):
            self._change_reader_zoom(self._wheel_zoom_factor(event))
            return "break"

        pixels = self._wheel_pixels(event)
        if not pixels:
            return "break"

        if self._shift_pressed(event):
            self._scroll_canvas_pixels(self.reader_canvas, dx=pixels)
        else:
            self._scroll_canvas_pixels(self.reader_canvas, dy=pixels)
        self.after_idle(self._update_current_page_from_view)
        return "break"

    def _update_current_page_from_view(self):
        if not self.reader_page_boxes:
            return

        center_y = self.reader_canvas.canvasy(self.reader_canvas.winfo_height() / 2)
        best_page = self.current_page
        best_distance = None

        for page_index, _x1, y1, _x2, y2 in self.reader_page_boxes:
            if y1 <= center_y <= y2:
                best_page = page_index
                break

            page_center = (y1 + y2) / 2
            distance = abs(center_y - page_center)
            if best_distance is None or distance < best_distance:
                best_distance = distance
                best_page = page_index

        if best_page != self.current_page:
            self.current_page = best_page
            self._sync_reader_buttons()

    def _on_preview_wheel(self, event, force_zoom=False):
        if force_zoom or self._ctrl_pressed(event):
            self._change_preview_zoom(self._wheel_zoom_factor(event))
            return "break"

        pixels = self._wheel_pixels(event)
        if pixels:
            self._scroll_canvas_pixels(self.preview_canvas, dy=pixels)
        return "break"

    def _on_reader_click(self, event):
        if not self.reader_page_boxes:
            return

        x = self.reader_canvas.canvasx(event.x)
        y = self.reader_canvas.canvasy(event.y)
        for page_index, left, top, right, bottom in self.reader_page_boxes:
            if left <= x <= right and top <= y <= bottom:
                self.current_page = page_index
                self._sync_reader_buttons()
                self._add_current_page()
                break

    def _change_reader_zoom(self, factor):
        self.reader_restore_y_fraction = self._canvas_view_fraction(self.reader_canvas)
        self.reader_zoom = max(0.35, min(self.reader_zoom * factor, 3.0))
        self._sync_reader_buttons()
        self._schedule_reader_render()

    def _fit_reader_width(self):
        self.reader_restore_y_fraction = self._canvas_view_fraction(self.reader_canvas)
        self.reader_zoom = 1.0
        self._sync_reader_buttons()
        self._schedule_reader_render()

    def _prev_page(self):
        if self.current_doc and self.current_page > 0:
            self.current_page -= 1
            self._scroll_to_reader_page(self.current_page)

    def _next_page(self):
        if self.current_doc and self.current_page < self.current_doc.page_count - 1:
            self.current_page += 1
            self._scroll_to_reader_page(self.current_page)

    def _jump_to_page(self, _event=None):
        if not self.current_doc:
            return

        try:
            target = int(self.page_entry.get()) - 1
        except ValueError:
            self._sync_reader_buttons()
            return

        if 0 <= target < self.current_doc.page_count:
            self.current_page = target
            self._scroll_to_reader_page(self.current_page)
        else:
            self._sync_reader_buttons()

    def _scroll_to_reader_page(self, page_index, update_buttons=True):
        target_box = None
        for box in self.reader_page_boxes:
            if box[0] == page_index:
                target_box = box
                break

        if target_box is None:
            if update_buttons:
                self._sync_reader_buttons()
            return

        _idx, _x1, y1, _x2, _y2 = target_box
        scrollregion = self.reader_canvas.cget("scrollregion").split()
        total_h = float(scrollregion[3]) if len(scrollregion) == 4 else 0
        if total_h > 0:
            self.reader_canvas.yview_moveto(max(0, min(y1 / total_h, 1)))
        self.reader_canvas.xview_moveto(0)
        if update_buttons:
            self._sync_reader_buttons()

    def _add_current_page(self):
        if not self.current_doc:
            return

        self.cheat_items.append(
            CheatItem(
                pdf_path=self.current_doc.path,
                page_index=self.current_page,
                source_name=self.current_source_name,
            )
        )
        self._refresh_items(select_index=len(self.cheat_items) - 1)
        self._set_status(f"Added {self.current_source_name} page {self.current_page + 1}")

    def _on_item_list_select(self, _event=None):
        self._sync_item_buttons()
        self._sync_property_controls()
        self._refresh_preview()

    def _refresh_items(self, select_index=None):
        self.item_list.delete(0, "end")
        for idx, item in enumerate(self.cheat_items, start=1):
            self.item_list.insert("end", f"{idx:02d}. {item.label}")

        if select_index is not None and self.cheat_items:
            select_index = max(0, min(select_index, len(self.cheat_items) - 1))
            self.item_list.selection_clear(0, "end")
            self.item_list.selection_set(select_index)
            self.item_list.see(select_index)

        self.item_count_label.configure(text=f"{len(self.cheat_items)} pages")
        self._sync_item_buttons()
        self._sync_property_controls()
        self._rebuild_preview()

    def _selected_item_index(self):
        selected = self.item_list.curselection()
        return selected[0] if selected else None

    def _select_item(self, index):
        if not 0 <= index < len(self.cheat_items):
            return

        self.item_list.selection_clear(0, "end")
        self.item_list.selection_set(index)
        self.item_list.see(index)
        self._sync_item_buttons()
        self._sync_property_controls()
        self._refresh_preview()

    def _sync_item_buttons(self):
        has_items = bool(self.cheat_items)
        selected = self._selected_item_index()

        self.export_button.configure(state="normal" if has_items else "disabled")
        self.undo_button.configure(state="normal" if has_items else "disabled")
        self.clear_button.configure(state="normal" if has_items else "disabled")
        self.remove_button.configure(state="normal" if selected is not None else "disabled")
        self.up_button.configure(state="normal" if selected is not None and selected > 0 else "disabled")
        self.down_button.configure(
            state="normal" if selected is not None and selected < len(self.cheat_items) - 1 else "disabled"
        )
        prop_state = "normal" if selected is not None else "disabled"
        for widget in [
            self.scale_slider,
            self.scale_down_button,
            self.scale_up_button,
            self.reset_transform_button,
            *self.crop_sliders.values(),
        ]:
            widget.configure(state=prop_state)

    def _sync_property_controls(self):
        selected = self._selected_item_index()
        self.property_syncing = True
        try:
            if selected is None or not 0 <= selected < len(self.cheat_items):
                self.prop_selected_label.configure(text="No item selected")
                self.scale_value_label.configure(text="100%")
                self.scale_slider.set(100)
                for key, slider in self.crop_sliders.items():
                    slider.set(0)
                    self.crop_value_labels[key].configure(text="0%")
                return

            item = self.cheat_items[selected]
            self.prop_selected_label.configure(text=f"{selected + 1:02d}. {item.label}")
            self.scale_slider.set(item.scale * 100)
            self.scale_value_label.configure(text=f"{int(item.scale * 100)}%")

            for key, slider in self.crop_sliders.items():
                value = getattr(item, key) * 100
                slider.set(value)
                self.crop_value_labels[key].configure(text=f"{int(round(value))}%")
        finally:
            self.property_syncing = False

    def _on_scale_slider(self, value):
        if self.property_syncing:
            return
        selected = self._selected_item_index()
        if selected is None:
            return

        item = self.cheat_items[selected]
        item.scale = max(0.35, min(float(value) / 100.0, 2.0))
        self.scale_value_label.configure(text=f"{int(item.scale * 100)}%")
        self._schedule_preview_rebuild(select_index=selected)

    def _adjust_selected_scale(self, delta):
        selected = self._selected_item_index()
        if selected is None:
            return

        item = self.cheat_items[selected]
        item.scale = max(0.35, min(item.scale + delta, 2.0))
        self._refresh_items(select_index=selected)

    def _on_crop_slider(self, crop_key, value):
        if self.property_syncing:
            return
        selected = self._selected_item_index()
        if selected is None:
            return

        item = self.cheat_items[selected]
        requested = max(0.0, min(float(value) / 100.0, 0.7))
        self._set_item_crop(item, crop_key, requested)
        actual = getattr(item, crop_key) * 100
        self.crop_value_labels[crop_key].configure(text=f"{int(round(actual))}%")
        self._schedule_preview_rebuild(select_index=selected)

    def _set_item_crop(self, item, crop_key, value):
        pair_key = {
            "crop_left": "crop_right",
            "crop_right": "crop_left",
            "crop_top": "crop_bottom",
            "crop_bottom": "crop_top",
        }[crop_key]
        max_total = 0.9
        pair_value = getattr(item, pair_key)
        setattr(item, crop_key, max(0.0, min(value, max_total - pair_value)))

    def _reset_selected_transform(self):
        selected = self._selected_item_index()
        if selected is None:
            return

        item = self.cheat_items[selected]
        item.scale = 1.0
        item.crop_left = 0.0
        item.crop_top = 0.0
        item.crop_right = 0.0
        item.crop_bottom = 0.0
        self._refresh_items(select_index=selected)

    def _undo_last(self):
        if not self.cheat_items:
            return
        removed = self.cheat_items.pop()
        self._refresh_items(select_index=len(self.cheat_items) - 1 if self.cheat_items else None)
        self._set_status(f"Removed {removed.label}")

    def _clear_items(self):
        if not self.cheat_items:
            return
        self.cheat_items.clear()
        self._refresh_items()
        self._set_status("Cleared cheatsheet")

    def _remove_selected_item(self):
        selected = self._selected_item_index()
        if selected is None:
            return

        removed = self.cheat_items.pop(selected)
        next_index = min(selected, len(self.cheat_items) - 1) if self.cheat_items else None
        self._refresh_items(select_index=next_index)
        self._set_status(f"Removed {removed.label}")

    def _move_selected_item(self, offset):
        selected = self._selected_item_index()
        if selected is None:
            return

        target = selected + offset
        if not 0 <= target < len(self.cheat_items):
            return

        self.cheat_items[selected], self.cheat_items[target] = self.cheat_items[target], self.cheat_items[selected]
        self._refresh_items(select_index=target)

    def _rebuild_preview(self):
        self.preview_pages = self._compose_preview_pages()
        self._refresh_preview()

    def _schedule_preview_rebuild(self, select_index=None, delay=140):
        self.pending_preview_select_index = select_index
        if self.preview_rebuild_after:
            self.after_cancel(self.preview_rebuild_after)
        self.preview_rebuild_after = self.after(delay, self._flush_preview_rebuild)

    def _flush_preview_rebuild(self):
        self.preview_rebuild_after = None
        select_index = self.pending_preview_select_index
        self.pending_preview_select_index = None
        self._refresh_items(select_index=select_index)

    def _compose_preview_pages(self):
        self.preview_item_layouts = []
        if not self.cheat_items:
            return []

        pages = []
        page_img = self._new_preview_page()
        draw = ImageDraw.Draw(page_img)
        self._draw_preview_grid(draw, PREVIEW_WIDTH, PREVIEW_HEIGHT)

        margin_v = int(EXPORT_MARGIN_V / A4_H_PTS * PREVIEW_HEIGHT)
        margin_h = int(EXPORT_MARGIN_H / A4_W_PTS * PREVIEW_WIDTH)
        gutter = int(EXPORT_GUTTER / A4_W_PTS * PREVIEW_WIDTH)
        gap = max(1, int(EXPORT_ITEM_GAP / A4_H_PTS * PREVIEW_HEIGHT))
        col_w = (PREVIEW_WIDTH - 2 * margin_h - (EXPORT_COLUMNS - 1) * gutter) // EXPORT_COLUMNS
        usable_h = PREVIEW_HEIGHT - 2 * margin_v

        col = 0
        y = margin_v
        doc_cache = {}

        try:
            for item_index, item in enumerate(self.cheat_items):
                try:
                    doc = doc_cache.get(item.pdf_path)
                    if doc is None:
                        doc = fitz.open(item.pdf_path)
                        doc_cache[item.pdf_path] = doc

                    src_page = doc.load_page(item.page_index)
                    clip_rect = self._item_clip_rect(item, src_page.rect)
                    target_w = max(24, col_w * item.scale)
                    scale = target_w / clip_rect.width
                    if clip_rect.height * scale > usable_h:
                        scale = usable_h / clip_rect.height

                    pix = src_page.get_pixmap(matrix=fitz.Matrix(scale, scale), clip=clip_rect, alpha=False)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

                    if y + img.height > PREVIEW_HEIGHT - margin_v + 1:
                        col += 1
                        y = margin_v
                        if col >= EXPORT_COLUMNS:
                            pages.append(page_img)
                            page_img = self._new_preview_page()
                            draw = ImageDraw.Draw(page_img)
                            self._draw_preview_grid(draw, PREVIEW_WIDTH, PREVIEW_HEIGHT)
                            col = 0

                    x = margin_h + col * (col_w + gutter) + max(0, (col_w - img.width) // 2)
                    page_img.paste(img, (int(x), int(y)))
                    self.preview_item_layouts.append(
                        {
                            "item_index": item_index,
                            "preview_page": len(pages),
                            "rect": (int(x), int(y), int(x + img.width), int(y + img.height)),
                        }
                    )
                    y += img.height + gap
                except Exception as exc:
                    print(f"Preview item failed: {exc}")
        finally:
            for doc in doc_cache.values():
                doc.close()

        pages.append(page_img)
        return pages

    def _item_clip_rect(self, item, page_rect):
        left = page_rect.x0 + page_rect.width * item.crop_left
        top = page_rect.y0 + page_rect.height * item.crop_top
        right = page_rect.x1 - page_rect.width * item.crop_right
        bottom = page_rect.y1 - page_rect.height * item.crop_bottom

        if right <= left:
            right = min(page_rect.x1, left + 1)
        if bottom <= top:
            bottom = min(page_rect.y1, top + 1)

        return fitz.Rect(left, top, right, bottom)

    def _new_preview_page(self):
        return Image.new("RGB", (PREVIEW_WIDTH, PREVIEW_HEIGHT), "white")

    def _draw_preview_grid(self, draw, page_w, page_h):
        margin_h = int(EXPORT_MARGIN_H / A4_W_PTS * page_w)
        gutter = int(EXPORT_GUTTER / A4_W_PTS * page_w)
        col_w = (page_w - 2 * margin_h - (EXPORT_COLUMNS - 1) * gutter) // EXPORT_COLUMNS
        for col in range(1, EXPORT_COLUMNS):
            x = margin_h + col * col_w + (col - 1) * gutter + gutter // 2
            draw.line((x, 0, x, page_h), fill=(220, 224, 230), width=PREVIEW_RENDER_SCALE)

    def _refresh_preview(self):
        self.preview_canvas.delete("all")
        self.preview_refs = []
        self.preview_hit_boxes = []

        if not self.preview_pages:
            width = max(self.preview_canvas.winfo_width(), 340)
            height = max(self.preview_canvas.winfo_height(), 420)
            card_w = min(260, width - 48)
            card_h = 92
            x0 = (width - card_w) / 2
            y0 = 78
            self.preview_canvas.create_rectangle(
                x0,
                y0,
                x0 + card_w,
                y0 + card_h,
                fill=THEME["surface"],
                outline=THEME["border_soft"],
                width=1,
            )
            self.preview_canvas.create_text(
                width / 2,
                y0 + 36,
                text="No pages selected",
                fill=THEME["text"],
                font=("Arial", 14, "bold"),
            )
            self.preview_canvas.create_text(
                width / 2,
                y0 + 62,
                text="Preview is empty",
                fill=THEME["muted_soft"],
                font=("Arial", 11),
            )
            self.preview_canvas.configure(scrollregion=(0, 0, width, height))
            self.preview_zoom_label.configure(text=f"{int(self.preview_zoom * 100)}%")
            self.preview_restore_y_fraction = None
            return

        canvas_w = max(self.preview_canvas.winfo_width(), 340)
        y = 18
        max_w = canvas_w
        selected_index = self._selected_item_index()

        for idx, page in enumerate(self.preview_pages, start=1):
            display_w = max(80, int(page.width * self.preview_zoom / PREVIEW_RENDER_SCALE))
            display_h = max(80, int(page.height * self.preview_zoom / PREVIEW_RENDER_SCALE))
            resized = page.resize((display_w, display_h), Image.Resampling.LANCZOS)
            ref = ImageTk.PhotoImage(resized)
            self.preview_refs.append(ref)

            x = max(16, (canvas_w - display_w) // 2)
            page_y = y + 10
            page_scale = display_w / page.width
            self.preview_canvas.create_text(
                x,
                y,
                text=f"Page {idx}",
                fill=THEME["muted"],
                anchor="sw",
                font=("Arial", 11),
            )
            self.preview_canvas.create_rectangle(
                x - 7,
                y + 3,
                x + display_w + 8,
                y + display_h + 17,
                fill="#05070a",
                outline=THEME["border_soft"],
            )
            self.preview_canvas.create_rectangle(
                x - 1,
                y + 9,
                x + display_w + 1,
                y + display_h + 11,
                fill="#f8fafc",
                outline="#3d4653",
            )
            self.preview_canvas.create_image(x, page_y, image=ref, anchor="nw")

            for layout in self.preview_item_layouts:
                if layout["preview_page"] != idx - 1:
                    continue

                item_index = layout["item_index"]
                rx1, ry1, rx2, ry2 = layout["rect"]
                cx1 = x + rx1 * page_scale
                cy1 = page_y + ry1 * page_scale
                cx2 = x + rx2 * page_scale
                cy2 = page_y + ry2 * page_scale
                self.preview_hit_boxes.append((item_index, cx1, cy1, cx2, cy2))

                if item_index == selected_index:
                    self.preview_canvas.create_rectangle(
                        cx1 - 2,
                        cy1 - 2,
                        cx2 + 2,
                        cy2 + 2,
                        outline=THEME["blue_hover"],
                        width=3,
                    )
                    self.preview_canvas.create_text(
                        cx1 + 5,
                        cy1 + 5,
                        text=f"{item_index + 1}",
                        fill=THEME["blue_hover"],
                        anchor="nw",
                        font=("Arial", 12, "bold"),
                    )
            y += display_h + 44
            max_w = max(max_w, x + display_w + 18)

        self.preview_canvas.configure(scrollregion=(0, 0, max_w, y))
        if self.preview_restore_y_fraction is not None:
            self.preview_canvas.yview_moveto(self.preview_restore_y_fraction)
            self.preview_restore_y_fraction = None
        self.preview_zoom_label.configure(text=f"{int(self.preview_zoom * 100)}%")

    def _change_preview_zoom(self, factor):
        self.preview_restore_y_fraction = self._canvas_view_fraction(self.preview_canvas)
        self.preview_zoom = max(0.2, min(self.preview_zoom * factor, 1.5))
        self._refresh_preview()

    def _on_preview_click(self, event):
        x = self.preview_canvas.canvasx(event.x)
        y = self.preview_canvas.canvasy(event.y)

        for item_index, x1, y1, x2, y2 in reversed(self.preview_hit_boxes):
            if x1 <= x <= x2 and y1 <= y <= y2:
                self._select_item(item_index)
                self._set_status(f"Selected item {item_index + 1}")
                return "break"
        return None

    def _export(self):
        if not self.cheat_items:
            messagebox.showwarning("Empty Cheatsheet", "Add at least one page before exporting.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Export cheatsheet",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not save_path:
            return

        try:
            self._export_pdf(save_path)
            self._create_blank_for_pdf(save_path)
            self._load_files()
            messagebox.showinfo("Exported", "PDF exported. A blank companion PDF was created beside it.")
            self._set_status(f"Exported {save_path}")
        except Exception as exc:
            messagebox.showerror("Export Failed", str(exc))
            self._set_status(f"Export failed: {exc}")

    def _export_pdf(self, save_path):
        out_doc = fitz.open()
        doc_cache = {}

        col_w = (A4_W_PTS - 2 * EXPORT_MARGIN_H - (EXPORT_COLUMNS - 1) * EXPORT_GUTTER) / EXPORT_COLUMNS
        usable_h = A4_H_PTS - 2 * EXPORT_MARGIN_V

        def get_doc(path):
            if path not in doc_cache:
                doc_cache[path] = fitz.open(path)
            return doc_cache[path]

        def new_target_page():
            page = out_doc.new_page(width=A4_W_PTS, height=A4_H_PTS)
            for col in range(1, EXPORT_COLUMNS):
                x = EXPORT_MARGIN_H + col * col_w + (col - 1) * EXPORT_GUTTER + EXPORT_GUTTER / 2
                page.draw_line((x, 0), (x, A4_H_PTS), color=(0.88, 0.88, 0.88), width=0.3)
            return page

        current_page = new_target_page()
        current_col = 0
        current_y = EXPORT_MARGIN_V

        try:
            for item in self.cheat_items:
                src_doc = get_doc(item.pdf_path)
                src_page = src_doc.load_page(item.page_index)
                clip_rect = self._item_clip_rect(item, src_page.rect)

                target_w = max(1.0, col_w * item.scale)
                target_h = clip_rect.height * (target_w / clip_rect.width)

                if target_h > usable_h:
                    scale = usable_h / clip_rect.height
                    target_w = clip_rect.width * scale
                    target_h = usable_h

                if current_y + target_h > A4_H_PTS - EXPORT_MARGIN_V + 0.01:
                    current_col += 1
                    current_y = EXPORT_MARGIN_V

                    if current_col >= EXPORT_COLUMNS:
                        current_page = new_target_page()
                        current_col = 0

                col_x = EXPORT_MARGIN_H + current_col * (col_w + EXPORT_GUTTER)
                target_x = col_x + (col_w - target_w) / 2
                target_rect = fitz.Rect(target_x, current_y, target_x + target_w, current_y + target_h)

                current_page.show_pdf_page(target_rect, src_doc, item.page_index, clip=clip_rect)
                current_y += target_h + EXPORT_ITEM_GAP

            out_doc.save(save_path, garbage=3, deflate=True)
        finally:
            out_doc.close()
            for doc in doc_cache.values():
                doc.close()

    def _create_blank_for_pdf(self, src_pdf_path):
        if not os.path.exists(src_pdf_path):
            raise FileNotFoundError(src_pdf_path)

        src = fitz.open(src_pdf_path)
        out = fitz.open()

        try:
            first_page = src.load_page(0)
            media = first_page.mediabox if hasattr(first_page, "mediabox") else first_page.rect
            page_w = media.width
            page_h = media.height

            safe_margin = 28.35
            left = safe_margin
            top = safe_margin
            right = max(page_w - safe_margin, left + 1)
            bottom = max(page_h - safe_margin, top + 1)

            def draw_dashed_line(page, x0, y0, x1, y1):
                dash_len = 6
                gap_len = 6
                dx = x1 - x0
                dy = y1 - y0
                dist = math.hypot(dx, dy)
                if dist == 0:
                    return

                ux = dx / dist
                uy = dy / dist
                pos = 0.0
                while pos < dist:
                    end = min(pos + dash_len, dist)
                    page.draw_line(
                        (x0 + ux * pos, y0 + uy * pos),
                        (x0 + ux * end, y0 + uy * end),
                        color=(0, 0, 0),
                        width=0.5,
                    )
                    pos = end + gap_len

            for _ in range(2):
                page = out.new_page(width=page_w, height=page_h)
                draw_dashed_line(page, left, top, right, top)
                draw_dashed_line(page, right, top, right, bottom)
                draw_dashed_line(page, right, bottom, left, bottom)
                draw_dashed_line(page, left, bottom, left, top)

                printable_w = right - left
                gutter = 10
                col_w = (printable_w - (EXPORT_COLUMNS - 1) * gutter) / EXPORT_COLUMNS
                for col in range(1, EXPORT_COLUMNS):
                    x = left + col * col_w + (col - 1) * gutter + gutter / 2
                    page.draw_line((x, top), (x, bottom), color=(0.6, 0.6, 0.6), width=0.5)

            out_dir = os.path.dirname(src_pdf_path)
            base = os.path.splitext(os.path.basename(src_pdf_path))[0]
            out.save(os.path.join(out_dir, f"{base}_白板.pdf"))
        finally:
            src.close()
            out.close()


if __name__ == "__main__":
    check_deps()
    ensure_dirs()
    app = CheatSheetMaker()
    app.mainloop()
