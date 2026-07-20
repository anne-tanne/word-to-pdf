"""Word to PDF Converter — a small local Mac app that batch-converts
.doc/.docx files to PDF by driving Microsoft Word, exactly like using
File > Save As > PDF, but for many files at once.
"""

import os
import subprocess
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

WORD_EXTENSIONS = (".doc", ".docx")

BG = "#f5f5f7"
CARD_BG = "#ffffff"
ACCENT = "#2f6fed"
ACCENT_HOVER = "#255bc4"
TEXT_MAIN = "#1d1d1f"
TEXT_MUTED = "#6e6e73"
BORDER = "#e2e2e6"
SUCCESS = "#1a8a4a"
ERROR = "#c0392b"


class WordToPdfApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Word to PDF Converter")
        self.root.geometry("640x720")
        self.root.minsize(560, 560)
        self.root.configure(bg=BG)

        self.dest_folder = tk.StringVar()
        self.status_text = tk.StringVar(value="Add Word documents to get started.")
        self.same_as_source = tk.BooleanVar(value=True)

        self.files = []  # ordered list of absolute paths
        self.file_included = {}  # path -> tk.BooleanVar
        self.file_rows = {}  # path -> status Label
        self.is_converting = False

        self._build_ui()
        self._render_file_list()

    # ---------- UI construction ----------

    def _build_ui(self):
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True, padx=24, pady=20)

        title = tk.Label(
            outer, text="Word to PDF Converter", font=("SF Pro Display", 20, "bold"),
            bg=BG, fg=TEXT_MAIN, anchor="w",
        )
        title.pack(fill="x")

        subtitle = tk.Label(
            outer, text="Add Word documents and convert them all to PDF in one go.",
            font=("SF Pro Text", 12), bg=BG, fg=TEXT_MUTED, anchor="w",
        )
        subtitle.pack(fill="x", pady=(2, 16))

        # --- File list card ---
        list_card = self._card(outer, expand=True)

        header_row = tk.Frame(list_card, bg=CARD_BG)
        header_row.pack(fill="x")

        self._section_label(header_row, "1. Word documents").pack(side="left")

        clear_btn = self._button(header_row, "Clear All", self.clear_files)
        clear_btn.pack(side="right")
        add_folder_btn = self._button(header_row, "Add Folder…", self.add_folder)
        add_folder_btn.pack(side="right", padx=(0, 8))
        add_files_btn = self._button(header_row, "Add Files…", self.add_files, primary=True)
        add_files_btn.pack(side="right", padx=(0, 8))

        list_container = tk.Frame(list_card, bg=CARD_BG)
        list_container.pack(fill="both", expand=True, pady=(10, 0))

        canvas = tk.Canvas(list_container, bg=CARD_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=canvas.yview)
        self.list_frame = tk.Frame(canvas, bg=CARD_BG)

        self.list_frame.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        list_window = canvas.create_window((0, 0), window=self.list_frame, anchor="nw", width=1)
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.bind("<Configure>", lambda e: canvas.itemconfig(list_window, width=e.width))
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.list_canvas = canvas

        # --- Destination card ---
        dest_card = self._card(outer)

        self._section_label(dest_card, "2. Save PDFs to")
        check_row = tk.Frame(dest_card, bg=CARD_BG)
        check_row.pack(fill="x", pady=(6, 0))

        same_check = tk.Checkbutton(
            check_row, text="Next to each Word document", variable=self.same_as_source,
            command=self.toggle_dest_mode, font=("SF Pro Text", 12), bg=CARD_BG, fg=TEXT_MAIN,
            activebackground=CARD_BG, selectcolor=CARD_BG, anchor="w",
        )
        same_check.pack(fill="x")

        dest_row = tk.Frame(dest_card, bg=CARD_BG)
        dest_row.pack(fill="x", pady=(8, 0))

        self.dest_entry = tk.Entry(
            dest_row, textvariable=self.dest_folder, font=("SF Pro Text", 12),
            state="readonly", relief="flat", bg="#f0f0f2", fg=TEXT_MAIN,
        )
        self.dest_entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 8))

        self.dest_button = self._button(dest_row, "Choose Folder…", self.pick_dest_folder)
        self.dest_button.pack(side="left")
        self.toggle_dest_mode()

        # --- Bottom action bar ---
        bottom = tk.Frame(outer, bg=BG)
        bottom.pack(fill="x")

        self.status_label = tk.Label(
            bottom, textvariable=self.status_text, font=("SF Pro Text", 11),
            bg=BG, fg=TEXT_MUTED, anchor="w", justify="left", wraplength=380,
        )
        self.status_label.pack(side="left", fill="x", expand=True)

        self.convert_button = self._button(
            bottom, "Convert to PDF", self.start_conversion, primary=True
        )
        self.convert_button.pack(side="right")

        self.progress = ttk.Progressbar(outer, mode="determinate")
        self.progress.pack(fill="x", pady=(10, 0))

    def _card(self, parent, expand=False):
        card = tk.Frame(parent, bg=CARD_BG, highlightbackground=BORDER, highlightthickness=1)
        if expand:
            card.pack(fill="both", expand=True, pady=(0, 12))
        else:
            card.pack(fill="x", pady=(0, 12))
        inner = tk.Frame(card, bg=CARD_BG)
        inner.pack(fill="both", expand=True, padx=16, pady=14)
        return inner

    def _section_label(self, parent, text):
        label = tk.Label(
            parent, text=text, font=("SF Pro Text", 13, "bold"), bg=CARD_BG, fg=TEXT_MAIN, anchor="w",
        )
        label.pack(fill="x")
        return label

    def _button(self, parent, text, command, primary=False):
        bg = ACCENT if primary else "#e8e8ed"
        fg = "#ffffff" if primary else TEXT_MAIN
        btn = tk.Label(
            parent, text=text, font=("SF Pro Text", 12, "bold" if primary else "normal"),
            bg=bg, fg=fg, padx=14, pady=7, cursor="pointinghand",
        )
        btn.bind("<Button-1>", lambda e: command())
        if primary:
            btn.bind("<Enter>", lambda e: btn.configure(bg=ACCENT_HOVER))
            btn.bind("<Leave>", lambda e: btn.configure(bg=ACCENT))
        return btn

    # ---------- File selection ----------

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Choose Word documents",
            filetypes=[("Word Documents", "*.docx *.doc")],
        )
        if paths:
            self._add_paths(paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Choose a folder of Word documents")
        if not folder:
            return
        found = []
        for name in sorted(os.listdir(folder)):
            if name.startswith("~$"):
                continue
            if name.lower().endswith(WORD_EXTENSIONS):
                found.append(os.path.join(folder, name))
        if not found:
            messagebox.showinfo("No files found", "That folder doesn't contain any .doc or .docx files.")
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
        if added:
            self.status_text.set(f"{len(self.files)} file(s) ready. Uncheck any you don't want to convert.")

    def remove_file(self, path):
        if path in self.file_included:
            self.files.remove(path)
            del self.file_included[path]
            self.file_rows.pop(path, None)
            self._render_file_list()

    def clear_files(self):
        if self.is_converting:
            return
        self.files = []
        self.file_included = {}
        self.file_rows = {}
        self._render_file_list()
        self.status_text.set("Add Word documents to get started.")

    def _render_file_list(self):
        for widget in self.list_frame.winfo_children():
            widget.destroy()
        self.file_rows = {}

        if not self.files:
            tk.Label(
                self.list_frame, text="No files yet — use “Add Files…” or “Add Folder…” above.",
                font=("SF Pro Text", 12), bg=CARD_BG, fg=TEXT_MUTED,
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
                row, text="Waiting", font=("SF Pro Text", 11), bg=CARD_BG, fg=TEXT_MUTED,
            )
            status_label.pack(side="right")
            self.file_rows[path] = status_label

            if not self.is_converting:
                remove_btn = tk.Label(
                    row, text="Remove", font=("SF Pro Text", 11), bg=CARD_BG, fg=ACCENT,
                    cursor="pointinghand", padx=8,
                )
                remove_btn.bind("<Button-1>", lambda e, p=path: self.remove_file(p))
                remove_btn.pack(side="right", padx=(0, 10))

    # ---------- Destination ----------

    def toggle_dest_mode(self):
        if self.same_as_source.get():
            self.dest_button.configure(state="disabled")
            self.dest_button.unbind("<Button-1>")
            self.dest_folder.set("Each PDF is saved next to its Word document")
        else:
            self.dest_folder.set("")
            self.dest_button.bind("<Button-1>", lambda e: self.pick_dest_folder())

    def pick_dest_folder(self):
        folder = filedialog.askdirectory(title="Choose folder to save PDFs")
        if folder:
            self.dest_folder.set(folder)

    # ---------- Conversion ----------

    def start_conversion(self):
        if self.is_converting:
            return
        selected = [p for p in self.files if self.file_included[p].get()]
        if not selected:
            messagebox.showinfo("No files selected", "Add Word documents and make sure at least one is checked.")
            return

        same_as_source = self.same_as_source.get()
        dest_folder = None
        if not same_as_source:
            dest_folder = self.dest_folder.get()
            if not dest_folder:
                messagebox.showinfo("No destination", "Choose where to save the PDFs.")
                return
            os.makedirs(dest_folder, exist_ok=True)

        self.is_converting = True
        self._render_file_list()
        self.convert_button.configure(bg="#c9c9cf", cursor="arrow")
        self.convert_button.unbind("<Button-1>")
        self.progress.configure(maximum=len(selected), value=0)
        self.status_text.set("Converting… Microsoft Word will open in the background.")

        thread = threading.Thread(target=self._convert_all, args=(selected, same_as_source, dest_folder), daemon=True)
        thread.start()

    def _convert_all(self, paths, same_as_source, dest_folder):
        succeeded = 0
        failed = 0
        needs_enable_editing = []
        for path in paths:
            self.root.after(0, self._set_row_status, path, "Converting…", TEXT_MUTED)
            target_folder = os.path.dirname(path) if same_as_source else dest_folder
            pdf_name = os.path.splitext(os.path.basename(path))[0] + ".pdf"
            pdf_path = os.path.join(target_folder, pdf_name)
            ok, error = convert_with_word(path, pdf_path)
            if ok:
                succeeded += 1
                self.root.after(0, self._set_row_status, path, "Done", SUCCESS)
            else:
                failed += 1
                if error and "Timed out" in error:
                    needs_enable_editing.append(path)
                self.root.after(0, self._set_row_status, path, "Failed", ERROR)
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

        parts = [f"{succeeded} converted"]
        if failed:
            parts.append(f"{failed} failed")
        self.status_text.set("Done — " + ", ".join(parts) + ".")

        if needs_enable_editing:
            names = "\n".join(os.path.basename(p) for p in needs_enable_editing)
            messagebox.showwarning(
                "Some files need one manual step",
                "These files came from an email or download, and Word is showing its "
                "\"Enable Editing\" banner for them, which blocks automatic conversion:\n\n" + names +
                "\n\nOpen each one in Word once, click \"Enable Editing\", save, then run "
                "the conversion again.",
            )
        elif failed:
            messagebox.showwarning(
                "Some files failed",
                f"{failed} file(s) could not be converted. Make sure Microsoft Word is installed "
                "and the files aren't open or password-protected.",
            )


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
    root = tk.Tk()
    try:
        root.tk.call("tk", "scaling", 1.4)
    except tk.TclError:
        pass
    app = WordToPdfApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
