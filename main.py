"""Word to PDF Converter — a small local Mac app that batch-converts
.doc/.docx files to PDF by driving Microsoft Word, exactly like using
File > Save As > PDF, but for many files at once.

User-facing text is loaded from JSON translation files in ./locales
(one per language, e.g. en.json, de.json). The active language is
auto-detected from the macOS UI language, with English as the fallback.
Add a new language by dropping another <code>.json file into locales/.
"""

import json
import os
import re
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

WORD_EXTENSIONS = (".doc", ".docx")
SUPPORTED_LANGUAGES = ("en", "de")
DEFAULT_LANGUAGE = "en"

BG = "#f4f2f7"          # soft lavender-tinted background (matches the logo)
CARD_BG = "#ffffff"
ACCENT = "#7a4fd0"       # purple, drawn from the logo gradient
ACCENT_HOVER = "#653cba"
SUBTLE = "#efedf4"       # secondary button / input fill
SUBTLE_HOVER = "#e5e2ee"
TEXT_MAIN = "#1d1d1f"
TEXT_MUTED = "#6e6e73"
BORDER = "#e6e3ec"
SUCCESS = "#1a8a4a"
ERROR = "#c0392b"


# ---------- Internationalization ----------

_STRINGS = {}   # active language
_FALLBACK = {}  # English, used when a key is missing


def _resource_base():
    return os.environ.get("RESOURCEPATH") or os.path.dirname(os.path.abspath(__file__))


def _locales_dir():
    """Folder holding the JSON translation files, in both dev and the built app."""
    return os.path.join(_resource_base(), "locales")


def _assets_dir():
    """Folder holding image assets (the header logo), in both dev and the built app."""
    return os.path.join(_resource_base(), "assets")


def _read_locale(lang):
    path = os.path.join(_locales_dir(), f"{lang}.json")
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}


def detect_language(supported=SUPPORTED_LANGUAGES, default=DEFAULT_LANGUAGE):
    """Pick a language: an explicit WORDTOPDF_LANG override, else the macOS UI
    language, else the POSIX locale env, else the default."""
    override = os.environ.get("WORDTOPDF_LANG")
    if override and override in supported:
        return override

    candidates = []
    try:
        out = subprocess.check_output(
            ["defaults", "read", "-g", "AppleLanguages"],
            text=True, stderr=subprocess.DEVNULL,
        )
        candidates.extend(re.findall(r'"([A-Za-z\-_]+)"', out))
    except Exception:
        pass
    for env in ("LANGUAGE", "LC_ALL", "LC_MESSAGES", "LANG"):
        val = os.environ.get(env)
        if val:
            candidates.append(val)

    for cand in candidates:
        code = cand.replace("_", "-").split("-")[0].lower()
        if code in supported:
            return code
    return default


def load_language(lang):
    global _STRINGS, _FALLBACK
    _FALLBACK = _read_locale(DEFAULT_LANGUAGE)
    _STRINGS = _FALLBACK if lang == DEFAULT_LANGUAGE else _read_locale(lang)


def t(key, **kwargs):
    """Look up a translated string by key and fill in any {placeholders}."""
    text = _STRINGS.get(key) or _FALLBACK.get(key) or key
    if kwargs:
        try:
            return text.format(**kwargs)
        except (KeyError, IndexError, ValueError):
            return text
    return text


class WordToPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title(t("app_title"))
        # Min width keeps the button rows from clipping; height can shrink freely
        # because the content scrolls.
        self.root.minsize(560, 320)
        self._size_to_screen()
        self.root.configure(bg=BG)

        self.logo_image = self._load_logo()

        self.dest_folder = tk.StringVar()
        self.status_text = tk.StringVar(value=t("status_start"))
        self.same_as_source = tk.BooleanVar(value=True)

        # Batch naming
        self.naming_mode = tk.StringVar(value="keep")  # keep | affix | number
        self.prefix_text = tk.StringVar()
        self.suffix_text = tk.StringVar()
        self.base_name_text = tk.StringVar(value=t("name_default_base"))
        self.naming_preview = tk.StringVar()

        self.files = []  # ordered list of absolute paths
        self.file_included = {}  # path -> tk.BooleanVar
        self.file_rows = {}  # path -> status Label
        self.is_converting = False

        self._build_ui()
        self._render_file_list()

        for var in (self.prefix_text, self.suffix_text, self.base_name_text):
            var.trace_add("write", self._update_naming_preview)
        self._update_naming_state()

    def _size_to_screen(self):
        """Open at a comfortable size that always fits the current screen, centered."""
        self.root.update_idletasks()
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        w = min(700, max(560, sw - 120))
        h = min(900, sh - 140)
        x = max((sw - w) // 2, 0)
        y = max((sh - h) // 3, 24)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _load_logo(self):
        """Load the header logo (a downscaled copy of the app icon); None if missing."""
        try:
            img = tk.PhotoImage(file=os.path.join(_assets_dir(), "logo.png"))
            # 128px source → ~43px, a tidy header size after subsampling.
            return img.subsample(3, 3)
        except Exception:
            return None

    # ---------- UI construction ----------

    def _build_ui(self):
        # Pinned bottom bar (Convert + progress) — always visible, outside the scroll area.
        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(side="bottom", fill="x", padx=24, pady=(0, 16))

        bottom = tk.Frame(self.root, bg=BG)
        bottom.pack(side="bottom", fill="x", padx=24, pady=(6, 0))

        self.status_label = tk.Label(
            bottom, textvariable=self.status_text, font=("SF Pro Text", 11),
            bg=BG, fg=TEXT_MUTED, anchor="w", justify="left", wraplength=380,
        )
        self.status_label.pack(side="left", fill="x", expand=True)

        self.convert_button = self._button(
            bottom, t("btn_convert"), self.start_conversion, primary=True
        )
        self.convert_button.pack(side="right")

        # Everything else lives in a scrollable, width-adaptive content area, so
        # resizing the window never clips content — it just scrolls.
        outer = self._build_scroll_area(self.root)
        self.root.bind("<Configure>", self._on_resize)

        header = tk.Frame(outer, bg=BG)
        header.pack(fill="x", pady=(0, 18))

        if self.logo_image is not None:
            tk.Label(header, image=self.logo_image, bg=BG).pack(side="left", padx=(0, 14))

        header_text = tk.Frame(header, bg=BG)
        header_text.pack(side="left", fill="x", expand=True)

        tk.Label(
            header_text, text=t("app_title"), font=("SF Pro Display", 26, "bold"),
            bg=BG, fg=TEXT_MAIN, anchor="w",
        ).pack(fill="x")

        self.subtitle_label = tk.Label(
            header_text, text=t("subtitle"),
            font=("SF Pro Text", 12), bg=BG, fg=TEXT_MUTED, anchor="w",
            justify="left", wraplength=460,
        )
        self.subtitle_label.pack(fill="x", pady=(3, 0))

        # --- File list card ---
        list_card = self._card(outer)

        self._section_label(list_card, t("section_files"))

        button_row = tk.Frame(list_card, bg=CARD_BG)
        button_row.pack(fill="x", pady=(8, 0))

        add_files_btn = self._button(button_row, t("btn_add_files"), self.add_files, primary=True)
        add_files_btn.pack(side="left")
        add_folder_btn = self._button(button_row, t("btn_add_folder"), self.add_folder)
        add_folder_btn.pack(side="left", padx=(8, 0))

        clear_btn = self._button(button_row, t("btn_clear_all"), self.clear_files)
        clear_btn.pack(side="right")

        self.list_frame = tk.Frame(list_card, bg=CARD_BG)
        self.list_frame.pack(fill="both", expand=True, pady=(10, 0))

        # --- Naming card ---
        naming_card = self._card(outer)
        self._section_label(naming_card, t("section_names"))

        keep_row = tk.Frame(naming_card, bg=CARD_BG)
        keep_row.pack(fill="x", pady=(6, 0))
        tk.Radiobutton(
            keep_row, text=t("name_keep"), variable=self.naming_mode,
            value="keep", command=self._update_naming_state, font=("SF Pro Text", 12),
            bg=CARD_BG, fg=TEXT_MAIN, activebackground=CARD_BG, selectcolor=CARD_BG, anchor="w",
        ).pack(side="left")

        affix_row = tk.Frame(naming_card, bg=CARD_BG)
        affix_row.pack(fill="x", pady=(4, 0))
        tk.Radiobutton(
            affix_row, text=t("name_add_text"), variable=self.naming_mode, value="affix",
            command=self._update_naming_state, font=("SF Pro Text", 12), bg=CARD_BG,
            fg=TEXT_MAIN, activebackground=CARD_BG, selectcolor=CARD_BG, anchor="w",
        ).pack(side="left")
        tk.Label(affix_row, text=t("name_before"), font=("SF Pro Text", 11), bg=CARD_BG,
                 fg=TEXT_MUTED).pack(side="left", padx=(6, 3))
        self.prefix_entry = tk.Entry(affix_row, textvariable=self.prefix_text, width=9,
                                     font=("SF Pro Text", 12), relief="flat", bg=SUBTLE, fg=TEXT_MAIN)
        self.prefix_entry.pack(side="left", ipady=3)
        tk.Label(affix_row, text=t("name_and_after"), font=("SF Pro Text", 11), bg=CARD_BG,
                 fg=TEXT_MUTED).pack(side="left", padx=(6, 3))
        self.suffix_entry = tk.Entry(affix_row, textvariable=self.suffix_text, width=9,
                                     font=("SF Pro Text", 12), relief="flat", bg=SUBTLE, fg=TEXT_MAIN)
        self.suffix_entry.pack(side="left", ipady=3)

        number_row = tk.Frame(naming_card, bg=CARD_BG)
        number_row.pack(fill="x", pady=(4, 0))
        tk.Radiobutton(
            number_row, text=t("name_rename_number"), variable=self.naming_mode, value="number",
            command=self._update_naming_state, font=("SF Pro Text", 12), bg=CARD_BG,
            fg=TEXT_MAIN, activebackground=CARD_BG, selectcolor=CARD_BG, anchor="w",
        ).pack(side="left")
        self.base_entry = tk.Entry(number_row, textvariable=self.base_name_text, width=16,
                                   font=("SF Pro Text", 12), relief="flat", bg=SUBTLE, fg=TEXT_MAIN)
        self.base_entry.pack(side="left", ipady=3, padx=(8, 4))
        tk.Label(number_row, text=t("name_sequence"), font=("SF Pro Text", 11), bg=CARD_BG,
                 fg=TEXT_MUTED).pack(side="left")

        # Fields stay clickable: focusing one selects its mode, and each field's
        # text is kept in its own variable so switching modes never loses it.
        for entry, mode in ((self.prefix_entry, "affix"), (self.suffix_entry, "affix"),
                            (self.base_entry, "number")):
            entry.configure(highlightthickness=1, highlightbackground=CARD_BG, highlightcolor=ACCENT)
            entry.bind("<FocusIn>", lambda e, m=mode: self._select_naming_mode(m))

        tk.Label(naming_card, textvariable=self.naming_preview, font=("SF Pro Text", 11, "italic"),
                 bg=CARD_BG, fg=TEXT_MUTED, anchor="w").pack(fill="x", pady=(10, 0))

        # --- Destination card ---
        dest_card = self._card(outer)

        self._section_label(dest_card, t("section_dest"))
        check_row = tk.Frame(dest_card, bg=CARD_BG)
        check_row.pack(fill="x", pady=(6, 0))

        same_check = tk.Checkbutton(
            check_row, text=t("dest_same"), variable=self.same_as_source,
            command=self.toggle_dest_mode, font=("SF Pro Text", 12), bg=CARD_BG, fg=TEXT_MAIN,
            activebackground=CARD_BG, selectcolor=CARD_BG, anchor="w",
        )
        same_check.pack(fill="x")

        dest_row = tk.Frame(dest_card, bg=CARD_BG)
        dest_row.pack(fill="x", pady=(8, 0))

        self.dest_entry = tk.Entry(
            dest_row, textvariable=self.dest_folder, font=("SF Pro Text", 12),
            state="readonly", relief="flat", bg=SUBTLE, fg=TEXT_MAIN,
        )
        self.dest_entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

        self.dest_button = self._button(dest_row, t("btn_choose_folder"), self.pick_dest_folder)
        self.dest_button.pack(side="left")
        self.toggle_dest_mode()

    def _card(self, parent, expand=False):
        card = tk.Frame(parent, bg=CARD_BG, highlightbackground=BORDER, highlightthickness=1)
        if expand:
            card.pack(fill="both", expand=True, pady=(0, 14))
        else:
            card.pack(fill="x", pady=(0, 14))
        inner = tk.Frame(card, bg=CARD_BG)
        inner.pack(fill="both", expand=True, padx=20, pady=18)
        return inner

    def _section_label(self, parent, text):
        row = tk.Frame(parent, bg=CARD_BG)
        row.pack(fill="x")
        # Small accent bar to the left of each section heading.
        tk.Frame(row, bg=ACCENT, width=3, height=16).pack(side="left", padx=(0, 9))
        label = tk.Label(
            row, text=text, font=("SF Pro Text", 14, "bold"), bg=CARD_BG, fg=TEXT_MAIN, anchor="w",
        )
        label.pack(side="left", fill="x", expand=True)
        return row

    def _button(self, parent, text, command, primary=False):
        base = ACCENT if primary else SUBTLE
        hover = ACCENT_HOVER if primary else SUBTLE_HOVER
        fg = "#ffffff" if primary else TEXT_MAIN
        btn = tk.Label(
            parent, text=text, font=("SF Pro Text", 12, "bold" if primary else "normal"),
            bg=base, fg=fg, padx=16, pady=8, cursor="pointinghand",
        )
        btn.bind("<Button-1>", lambda e: command())
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover))
        btn.bind("<Leave>", lambda e: btn.configure(bg=base))
        # Remember the resting colour so disabled/again states can restore it.
        btn._base_bg = base
        return btn

    def _build_scroll_area(self, parent):
        """Return a padded content frame that scrolls vertically and whose width
        tracks the window, so content adapts on resize and is never clipped."""
        canvas = tk.Canvas(parent, bg=BG, highlightthickness=0)
        vbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        self._page_canvas = canvas

        inner = tk.Frame(canvas, bg=BG)
        window = canvas.create_window((0, 0), window=inner, anchor="nw")
        # Keep the scroll region in sync with the content height…
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        # …and make the content width follow the window (horizontal responsiveness).
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window, width=e.width))

        # Mouse-wheel / trackpad scrolling anywhere in the window.
        canvas.bind_all("<MouseWheel>", self._on_mousewheel)
        canvas.bind_all("<Button-4>", self._on_mousewheel)
        canvas.bind_all("<Button-5>", self._on_mousewheel)

        content = tk.Frame(inner, bg=BG)
        content.pack(fill="both", expand=True, padx=24, pady=20)
        return content

    def _on_mousewheel(self, event):
        canvas = getattr(self, "_page_canvas", None)
        if canvas is None:
            return
        if getattr(event, "num", None) == 4:       # Linux scroll up
            direction = -1
        elif getattr(event, "num", None) == 5:     # Linux scroll down
            direction = 1
        else:                                        # macOS / Windows
            direction = -1 if event.delta > 0 else 1
        # A few lines per wheel notch / trackpad tick so scrolling feels responsive.
        canvas.yview_scroll(direction * 3, "units")

    def _on_resize(self, event):
        if event.widget is self.root:
            # Let text wrap to the available width as the window resizes.
            self.status_label.configure(wraplength=max(event.width - 220, 160))
            if hasattr(self, "subtitle_label"):
                self.subtitle_label.configure(wraplength=max(event.width - 160, 200))

    # ---------- File selection ----------

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title=t("fd_choose_files"),
            filetypes=[(t("fd_word_documents"), "*.docx *.doc")],
        )
        if paths:
            self._add_paths(paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title=t("fd_choose_source_folder"))
        if not folder:
            return
        found = []
        for name in sorted(os.listdir(folder)):
            if name.startswith("~$"):
                continue
            if name.lower().endswith(WORD_EXTENSIONS):
                found.append(os.path.join(folder, name))
        if not found:
            messagebox.showinfo(t("dlg_no_found_title"), t("dlg_no_found_msg"))
            return
        self._add_paths(found)

    def _add_paths(self, paths):
        added = 0
        for path in paths:
            if path not in self.file_included:
                self.files.append(path)
                self.file_included[path] = tk.BooleanVar(value=True)
                added += 1
        self._render_file_list()
        self._update_naming_preview()
        if added:
            self.status_text.set(t("status_ready", n=len(self.files)))

    def remove_file(self, path):
        if path in self.file_included:
            self.files.remove(path)
            del self.file_included[path]
            self.file_rows.pop(path, None)
            self._render_file_list()
            self._update_naming_preview()

    def clear_files(self):
        if self.is_converting:
            return
        self.files = []
        self.file_included = {}
        self.file_rows = {}
        self._render_file_list()
        self._update_naming_preview()
        self.status_text.set(t("status_start"))

    def _render_file_list(self):
        for widget in self.list_frame.winfo_children():
            widget.destroy()
        self.file_rows = {}

        if not self.files:
            tk.Label(
                self.list_frame, text=t("list_empty"),
                font=("SF Pro Text", 12), bg=CARD_BG, fg=TEXT_MUTED,
                wraplength=340, justify="center",
            ).pack(pady=20)
            return

        for path in self.files:
            row = tk.Frame(self.list_frame, bg=CARD_BG)
            row.pack(fill="x", pady=3)

            check = tk.Checkbutton(
                row, variable=self.file_included[path], bg=CARD_BG,
                activebackground=CARD_BG, selectcolor=CARD_BG,
                state="disabled" if self.is_converting else "normal",
            )
            check.pack(side="left")

            name_label = tk.Label(
                row, text=os.path.basename(path), font=("SF Pro Text", 12),
                bg=CARD_BG, fg=TEXT_MAIN, anchor="w",
            )
            name_label.pack(side="left", fill="x", expand=True, padx=(4, 0))

            status_label = tk.Label(
                row, text=t("row_waiting"), font=("SF Pro Text", 11), bg=CARD_BG, fg=TEXT_MUTED,
            )
            status_label.pack(side="right")
            self.file_rows[path] = status_label

            if not self.is_converting:
                remove_btn = tk.Label(
                    row, text=t("row_remove"), font=("SF Pro Text", 11), bg=CARD_BG, fg=ACCENT,
                    cursor="pointinghand", padx=8,
                )
                remove_btn.bind("<Button-1>", lambda e, p=path: self.remove_file(p))
                remove_btn.pack(side="right", padx=(0, 10))

    # ---------- Destination ----------

    def toggle_dest_mode(self):
        if self.same_as_source.get():
            self.dest_button.configure(state="disabled")
            self.dest_button.unbind("<Button-1>")
            self.dest_folder.set(t("dest_same_hint"))
        else:
            self.dest_folder.set("")
            self.dest_button.bind("<Button-1>", lambda e: self.pick_dest_folder())

    def pick_dest_folder(self):
        folder = filedialog.askdirectory(title=t("fd_choose_dest_folder"))
        if folder:
            self.dest_folder.set(folder)

    # ---------- Naming ----------

    def _current_naming(self):
        return {
            "mode": self.naming_mode.get(),
            "prefix": self.prefix_text.get(),
            "suffix": self.suffix_text.get(),
            "base": self.base_name_text.get().strip() or t("name_default_base"),
        }

    def _select_naming_mode(self, mode):
        """Called when a naming field is clicked/focused — activate that field's mode."""
        if self.naming_mode.get() != mode:
            self.naming_mode.set(mode)
        self._update_naming_state()

    def _update_naming_state(self):
        # All fields stay editable so you can click straight into one; the active
        # group gets an accent outline. Text is preserved in each field's variable.
        mode = self.naming_mode.get()
        for entry, entry_mode in ((self.prefix_entry, "affix"), (self.suffix_entry, "affix"),
                                  (self.base_entry, "number")):
            active = (mode == entry_mode)
            entry.configure(highlightbackground=ACCENT if active else CARD_BG)
        self._update_naming_preview()

    def _update_naming_preview(self, *args):
        selected = [p for p in self.files if self.file_included[p].get()]
        total = max(len(selected), 1)
        if selected:
            example = os.path.splitext(os.path.basename(selected[0]))[0]
        else:
            example = "Report"
        stem = compute_pdf_stem(self._current_naming(), example, 0, total)
        self.naming_preview.set(t("name_preview", name=stem + ".pdf"))

    # ---------- Conversion ----------

    def start_conversion(self):
        if self.is_converting:
            return
        selected = [p for p in self.files if self.file_included[p].get()]
        if not selected:
            messagebox.showinfo(t("dlg_no_files_title"), t("dlg_no_files_msg"))
            return

        same_as_source = self.same_as_source.get()
        dest_folder = None
        if not same_as_source:
            dest_folder = self.dest_folder.get()
            if not dest_folder:
                messagebox.showinfo(t("dlg_no_dest_title"), t("dlg_no_dest_msg"))
                return
            os.makedirs(dest_folder, exist_ok=True)

        naming = self._current_naming()

        self.is_converting = True
        self._render_file_list()
        self.convert_button.configure(bg="#c9c9cf", cursor="arrow")
        self.convert_button.unbind("<Button-1>")
        self.progress.configure(maximum=len(selected), value=0)
        self.status_text.set(t("status_converting"))

        thread = threading.Thread(
            target=self._convert_all, args=(selected, same_as_source, dest_folder, naming), daemon=True
        )
        thread.start()

    def _convert_all(self, paths, same_as_source, dest_folder, naming):
        succeeded = 0
        failed = 0
        needs_enable_editing = []
        total = len(paths)
        for index, path in enumerate(paths):
            self.root.after(0, self._set_row_status, path, t("row_converting"), TEXT_MUTED)
            target_folder = os.path.dirname(path) if same_as_source else dest_folder
            original_stem = os.path.splitext(os.path.basename(path))[0]
            pdf_name = compute_pdf_stem(naming, original_stem, index, total) + ".pdf"
            pdf_path = os.path.join(target_folder, pdf_name)
            ok, error = convert_with_word(path, pdf_path)
            if ok:
                succeeded += 1
                self.root.after(0, self._set_row_status, path, t("row_done"), SUCCESS)
            else:
                failed += 1
                if error and "Timed out" in error:
                    needs_enable_editing.append(path)
                self.root.after(0, self._set_row_status, path, t("row_failed"), ERROR)
                self.root.after(0, self._log_error, path, error)
            self.root.after(0, self._advance_progress)

        self.root.after(0, self._finish_conversion, succeeded, failed, needs_enable_editing)

    def _set_row_status(self, path, text, color):
        label = self.file_rows.get(path)
        if label:
            label.configure(text=text, fg=color)

    def _advance_progress(self):
        self.progress.step(1)

    def _log_error(self, path, error):
        print(f"Failed to convert {path}: {error}")

    def _finish_conversion(self, succeeded, failed, needs_enable_editing):
        self.is_converting = False
        self._render_file_list()
        self.convert_button.configure(bg=ACCENT, cursor="pointinghand")
        self.convert_button.bind("<Button-1>", lambda e: self.start_conversion())

        parts = [t("summary_converted", n=succeeded)]
        if failed:
            parts.append(t("summary_failed", n=failed))
        self.status_text.set(t("status_done", summary=", ".join(parts)))

        if needs_enable_editing:
            names = "\n".join(os.path.basename(p) for p in needs_enable_editing)
            messagebox.showwarning(t("dlg_enable_title"), t("dlg_enable_msg", names=names))
        elif failed:
            messagebox.showwarning(t("dlg_failed_title"), t("dlg_failed_msg", n=failed))


def compute_pdf_stem(naming, original_stem, index, total):
    """Return the PDF file name (without extension) for one document, based on the
    chosen batch-naming mode. `index` is 0-based; `total` is the count being converted."""
    mode = naming.get("mode", "keep")
    if mode == "affix":
        return f"{naming.get('prefix', '')}{original_stem}{naming.get('suffix', '')}"
    if mode == "number":
        base = naming.get("base") or "Document"
        width = len(str(total))
        return f"{base} {str(index + 1).zfill(width)}"
    return original_stem


def convert_with_word(docx_path, pdf_path):
    """Drive Microsoft Word via AppleScript to open a document and export it as PDF,
    the same result as File > Save As > PDF."""
    script = f'''
    on run
        set srcPath to POSIX file "{docx_path}"
        tell application "Microsoft Word"
            try
                open srcPath
                delay 0.5
                set theDoc to active document
            on error errMsg
                return "OPEN_ERROR: " & errMsg
            end try
            try
                save as theDoc file name "{pdf_path}" file format format PDF
            on error errMsg
                try
                    close theDoc saving no
                end try
                return "SAVE_ERROR: " & errMsg
            end try
            try
                close theDoc saving no
            end try
        end tell
        return "OK"
    end run
    '''
    try:
        result = subprocess.run(
            ["osascript", "-e", script],
            capture_output=True, text=True, timeout=120,
        )
    except subprocess.TimeoutExpired:
        return False, "Timed out waiting for Microsoft Word."

    output = (result.stdout or "").strip()
    if result.returncode != 0:
        return False, (result.stderr or "Unknown AppleScript error").strip()
    if output != "OK":
        return False, output or "Unknown error"
    return True, None


def main():
    load_language(detect_language())
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.4)
    except tk.TclError:
        pass
    app = WordToPdfApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
