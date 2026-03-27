import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import win32com.client
import pythoncom
import os
import sys
import threading
from tkinterdnd2 import DND_FILES, TkinterDnD


# -------------------------
# Utilities
# -------------------------
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def normalize_path(p):
    return os.path.abspath(os.path.normpath(p))


# -------------------------
# PowerPoint Merger
# -------------------------
class PowerPointMerger:
    def merge(self, files, output_file, progress_cb, status_cb):
        pythoncom.CoInitialize()
        app = None
        presentation = None

        try:
            app = win32com.client.Dispatch("PowerPoint.Application")

            first = files[0]
            status_cb(first, "...טוען בסיס")
            presentation = app.Presentations.Open(first, WithWindow=False)
            status_cb(first, "סיים")
            progress_cb()

            for f in files[1:]:
                status_cb(f, "...ממזג")
                presentation.Slides.InsertFromFile(f, presentation.Slides.Count)
                status_cb(f, "סיים")
                progress_cb()

            presentation.SaveAs(output_file)

        finally:
            if presentation:
                presentation.Close()
            if app:
                app.Quit()
            pythoncom.CoUninitialize()


# -------------------------
# UI
# -------------------------
class PPTXMergerApp:
    STATUS_PENDING = "ממתין"

    def __init__(self, root):
        self.root = root
        self.root.title("מיזוג מצגות")

        self._drag_item = None
        self._drag_start_y = None

        self.last_output_dir = ""

        self._setup_window()
        self._build_ui()

    # -------------------------
    def _setup_window(self):
        w, h = 550, 580
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        x, y = int((sw - w) / 2), int((sh - h) / 2)
        self.root.geometry(f"{w}x{h}+{x}+{y}")

        try:
            self.root.iconbitmap(resource_path("fire.ico"))
        except:
            pass

    def _build_ui(self):
        tk.Label(self.root, text="מיזוג מצגות", font=("Helvetica", 14, "bold")).pack(pady=5)
        tk.Label(self.root, text="(גרור קבצים או הוסף ידנית)", font=("Helvetica", 9)).pack()

        tk.Button(self.root, text="PPTX הוסף קבצי", command=self.add_files).pack(pady=5)

        self.tree = ttk.Treeview(
            self.root,
            columns=("File", "Size", "Status", "FullPath"),
            displaycolumns=("File", "Size", "Status"),
            show="headings",
            height=10,
        )

        self.tree.heading("File", text="שם קובץ")
        self.tree.heading("Size", text="גודל קובץ")
        self.tree.heading("Status", text="סטטוס")

        self.tree.column("File", width=240, anchor=tk.E)
        self.tree.column("Size", width=80, anchor=tk.CENTER)
        self.tree.column("Status", width=120, anchor=tk.CENTER)

        self.tree.pack(fill=tk.X, pady=5)

        # Drag & Drop
        self.tree.drop_target_register(DND_FILES)
        self.tree.dnd_bind("<<Drop>>", self.handle_drop)

        # Drag reorder
        self.tree.bind("<ButtonPress-1>", self.on_drag_start, add="+")
        self.tree.bind("<B1-Motion>", self.on_drag_motion, add="+")
        self.tree.bind("<ButtonRelease-1>", self.on_drag_release, add="+")

        self._insert_line = tk.Frame(self.tree, height=2, bg="red")

        # Move buttons
        move_frame = tk.Frame(self.root)
        move_frame.pack(pady=5)

        tk.Button(move_frame, text="↑ הזז למעלה", command=self.move_up, width=12).pack(side=tk.LEFT, padx=5)
        tk.Button(move_frame, text="↓ הזז למטה", command=self.move_down, width=12).pack(side=tk.LEFT, padx=5)

        tk.Button(self.root, text="נקה רשימה", command=self.clear_files).pack(pady=5)

        self.progress = ttk.Progressbar(self.root, length=450)
        self.progress.pack(pady=10)

        action_frame = tk.Frame(self.root)
        action_frame.pack()

        self.btn_merge = tk.Button(action_frame, text="מזג ושמור", command=self.merge_files, bg="#4CAF50", fg="white")
        self.btn_merge.pack(side=tk.RIGHT, padx=5, ipadx=20)

        self.btn_open = tk.Button(action_frame, text="פתח תיקיית יעד", command=self.open_output_folder, state=tk.DISABLED)
        self.btn_open.pack(side=tk.LEFT, padx=5, ipadx=20)

    # -------------------------
    # Drag logic
    # -------------------------
    def on_drag_start(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        self._drag_item = item
        self._drag_start_y = event.y

    def on_drag_motion(self, event):
        if not self._drag_item:
            return

        if abs(event.y - self._drag_start_y) < 5:
            return

        target = self.tree.identify_row(event.y)
        if not target:
            return

        bbox = self.tree.bbox(target)
        if not bbox:
            return

        y, h = bbox[1], bbox[3]

        line_y = y if event.y < y + h / 2 else y + h
        self._insert_line.place(x=0, y=line_y, relwidth=1)

    def on_drag_release(self, event):
        if not self._drag_item:
            return

        target = self.tree.identify_row(event.y)

        if target:
            idx = self.tree.index(target)
            bbox = self.tree.bbox(target)

            if bbox and event.y > bbox[1] + bbox[3] / 2:
                idx += 1

            self.tree.move(self._drag_item, "", idx)

        self._drag_item = None
        self._insert_line.place_forget()

    # -------------------------
    # File handling
    # -------------------------
    def get_file_size(self, path):
        size = os.path.getsize(path)
        return f"{size / 1024:.1f} KB" if size < 1024 * 1024 else f"{size / (1024 * 1024):.1f} MB"
        
    def add_file(self, path):
        path = normalize_path(path)
        if not os.path.exists(path):
            return

        self.tree.insert("", tk.END, values=(os.path.basename(path), self.get_file_size(path), self.STATUS_PENDING, path))

    def add_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PPT", "*.pptx *.ppt")])
        for f in files:
            self.add_file(f)

    def handle_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        for f in files:
            self.add_file(f.strip("{}"))

    def clear_files(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        self.progress["value"] = 0
        self.btn_open.config(state=tk.DISABLED)

    # -------------------------
    # Move buttons
    # -------------------------
    def move_up(self):
        for item in self.tree.selection():
            idx = self.tree.index(item)
            if idx > 0:
                self.tree.move(item, "", idx - 1)

    def move_down(self):
        for item in reversed(self.tree.selection()):
            idx = self.tree.index(item)
            if idx < len(self.tree.get_children()) - 1:
                self.tree.move(item, "", idx + 1)

    # -------------------------
    # Merge
    # -------------------------
    def merge_files(self):
        items = self.tree.get_children()
        if len(items) < 2:
            messagebox.showwarning("שגיאה", ".אנא בחר לפחות 2 קבצים למיזוג")
            return

        output = filedialog.asksaveasfilename(
            title="שמור קובץ ממוזג בשם",
            defaultextension=".pptx",
            filetypes=[("PowerPoint Files", "*.pptx")]
        )
        if not output:
            return

        files = [normalize_path(self.tree.item(i, "values")[3]) for i in items]

        self.progress["value"] = 0
        self.progress["maximum"] = len(files)
        self.btn_merge.config(state=tk.DISABLED)

        threading.Thread(target=self._merge_worker, args=(files, output), daemon=True).start()

    def _merge_worker(self, files, output):
        merger = PowerPointMerger()

        def progress():
            def update_bar():
                self.progress["value"] += 1
            self.root.after(0, update_bar)

        def status(path, text):
            for item in self.tree.get_children():
                if self.tree.item(item, "values")[3] == path:
                    self.root.after(0, lambda i=item: self.tree.set(i, "Status", text))
                    break

        try:
            merger.merge(files, output, progress, status)

            self.last_output_dir = os.path.dirname(output)

            self.root.after(0, lambda: self.btn_open.config(state=tk.NORMAL))
            self.root.after(0, lambda: messagebox.showinfo("הצלחה", f":הקבצים מוזגו בהצלחה ונשמרו ב\n{output}"))

        finally:
            self.root.after(0, lambda: self.btn_merge.config(state=tk.NORMAL))

    def open_output_folder(self):
        if self.last_output_dir:
            os.startfile(self.last_output_dir)


# -------------------------
if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = PPTXMergerApp(root)
    root.mainloop()