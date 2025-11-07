# app.py - FINAL: Full-width layout + scroll restored + logos fixed
import os, sys, csv, threading, time, tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlsxwriter
from PIL import Image, ImageTk

APP_TITLE = "Eureka"
BOSCH_RED = "#ED0006"
GREEN = "#28A745"
STATUS_BLUE = "#0066CC"
HOVER_RED = "#CC0000"
HOVER_GREEN = "#1E7E34"
MAX_PREVIEW = 500

# -----------------------------
# Progress Bar Widget
# -----------------------------
class CanvasProgress(tk.Frame):
    def __init__(self, master, label="Progress", width=280, height=16):
        super().__init__(master, bg="white")
        self.w = width; self.h = height
        tk.Label(self, text=label, bg="white", fg="#333", font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=5)
        self.canvas = tk.Canvas(self, width=self.w, height=self.h, highlightthickness=0, bg="white")
        self.canvas.pack(padx=5, pady=1)
        self.canvas.create_rectangle(0, 0, self.w, self.h, fill="#E6E6E6", outline="")
        self.fill = self.canvas.create_rectangle(0, 0, 0, self.h, fill=BOSCH_RED, outline="")
        info = tk.Frame(self, bg="white"); info.pack(fill="x", padx=5)
        self.lbl_left   = tk.Label(info, text="Processed: 0", anchor="w", bg="white", font=("Segoe UI", 8)); self.lbl_left.pack(side="left")
        self.lbl_center = tk.Label(info, text="0%", bg="white", font=("Segoe UI", 8)); self.lbl_center.pack(side="left", expand=True)
        self.lbl_right  = tk.Label(info, text="Total: 0", anchor="e", bg="white", font=("Segoe UI", 8)); self.lbl_right.pack(side="right")
        self.total = self.value = 0; self.start_ts = None

    def start(self, total):
        self.total = max(0, int(total)); self.value = 0; self.start_ts = time.time(); self._update(0)
    def update(self, value):
        self.value = int(value); self._update(self.value)
    def complete(self):
        self.update(self.total)

    def _update(self, value):
        try:
            frac = value / self.total if self.total else 0
            fill_w = int(self.w * frac)
            self.canvas.coords(self.fill, 0, 0, fill_w, self.h)
            elapsed = int(time.time() - self.start_ts) if self.start_ts else 0
            rate = value / elapsed if elapsed else 0
            eta = int((self.total - value) / rate) if rate else 0
            self.lbl_left.config(text=f"Processed: {value:,}")
            self.lbl_center.config(text=f"{int(frac*100)}% | ETA {eta}s")
            self.lbl_right.config(text=f"Total: {self.total:,}")
        except Exception as e:
            pass

# -----------------------------
# Value Normalizer
# -----------------------------
def normalize_value(val, match_case, ignore_zeros, decimal_mode, dec_exact):
    s = str(val).strip()
    s = s.replace('₹', '').replace('$', '').replace('€', '').replace(',', '').strip()
    if not match_case:
        s = s.lower()
    if ignore_zeros:
        s = s.lstrip('0')
        if not s: s = '0'
    if decimal_mode == "ignore" and not dec_exact:
        if '.' in s:
            s = s.rstrip('0').rstrip('.') if '.' in s else s
    elif decimal_mode == "ignore_decimal":
        s = s.split('.')[0]
    return s

# -----------------------------
# Fast header + row count
# -----------------------------
def headers_and_rowcount(path):
    with open(path, 'r', encoding='utf-8-sig', newline='') as f:
        r = csv.reader(f)
        headers = next(r, [])
        count = sum(1 for row in r if any(str(c).strip() for c in row))
    return headers, count

# -----------------------------
# Streaming compare - ROW FIX
# -----------------------------
def compare_streamed(src_path, tgt_path, progress_cb=None, stop_event=None,
                     match_case=False, ignore_zeros=False, decimal_mode="ignore", dec_exact=False):
    src_h, src_n = headers_and_rowcount(src_path)
    tgt_h, tgt_n = headers_and_rowcount(tgt_path)
    all_h = sorted(set(src_h) | set(tgt_h))

    preview = []
    spill = tempfile.NamedTemporaryFile(prefix="dirt_", suffix=".csv", delete=False, mode="w", newline="", encoding="utf-8")
    writer = csv.writer(spill); writer.writerow(["Row","Column","Source","Target"])

    rows_diff = cells_diff = processed = 0
    total = min(src_n, tgt_n)

    with open(src_path, "r", encoding="utf-8-sig", newline="") as sf, \
         open(tgt_path, "r", encoding="utf-8-sig", newline="") as tf:
        sr = csv.DictReader(sf); tr = csv.DictReader(tf)
        for src_row, tgt_row in zip(sr, tr):
            if stop_event and stop_event.is_set(): break
            diff_in_row = False
            row_idx = processed + 2  # ← Excel row = processed + 2
            for col in all_h:
                sv = normalize_value(src_row.get(col, ""), match_case, ignore_zeros, decimal_mode, dec_exact)
                tv = normalize_value(tgt_row.get(col, ""), match_case, ignore_zeros, decimal_mode, dec_exact)
                if sv != tv:
                    rec = (row_idx, col, src_row.get(col, ""), tgt_row.get(col, ""))
                    if len(preview) < MAX_PREVIEW: preview.append(rec)
                    writer.writerow(rec); cells_diff += 1; diff_in_row = True
            if diff_in_row: rows_diff += 1
            processed += 1
            if progress_cb and processed % 1000 == 0:
                progress_cb(processed, total)

    spill.close()
    stats = {
        "src_rows": src_n, "tgt_rows": tgt_n, "rows_match": int(src_n==tgt_n),
        "src_cols": len(src_h), "tgt_cols": len(tgt_h), "cols_match": int(len(src_h)==len(tgt_h)),
        "rows_with_diff": rows_diff, "cells_with_diff": cells_diff,
        "extra_src_rows": src_n - total, "extra_tgt_rows": tgt_n - total,
        "spill_path": spill.name
    }
    return stats, preview

# -----------------------------
# Export
# -----------------------------
def export_report(path, stats, prog=None, stop=None):
    with xlsxwriter.Workbook(path) as wb:
        bold = wb.add_format({'bold':True, 'font_color':BOSCH_RED})
        ws = wb.add_worksheet("Summary")
        for i, k in enumerate(["src_rows","tgt_rows","rows_match","src_cols","tgt_cols","cols_match",
                               "rows_with_diff","cells_with_diff","extra_src_rows","extra_tgt_rows"]):
            ws.write(i, 0, k, bold); ws.write(i, 1, stats.get(k,""))
        if prog: prog.update(1)

        ws2 = wb.add_worksheet("Cell Differences")
        ws2.write_row(0,0,["Row","Column","Source","Target"], bold)
        spill = stats.get("spill_path")
        if spill and os.path.exists(spill):
            with open(spill, "r", encoding="utf-8", newline="") as f:
                r = csv.reader(f); next(r,None)
                total = stats.get("cells_with_diff",0)
                if prog: prog.start(max(1,total))
                for i, row in enumerate(r, 1):
                    if stop and stop.is_set(): break
                    ws2.write_row(i,0,row)
                    if prog and i % 1000 == 0: prog.update(i)
        if spill and os.path.exists(spill):
            try: os.unlink(spill)
            except: pass

# -----------------------------
# UI Proxy
# -----------------------------
class UIProxy:
    def __init__(self, app, bar): self.app=app; self.bar=bar
    def start(self, t): self.app.ui(self.bar.start, t)
    def update(self, v): self.app.ui(self.bar.update, v)
    def complete(self): self.app.ui(self.bar.complete)

# -----------------------------
# Main App - FINAL
# -----------------------------
class DIRTApp:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.root.configure(bg="white")
        self.root.state('zoomed')  # ← FULL SCREEN

        self.stop = threading.Event()
        self.stats = None

        self.var_case = tk.BooleanVar(value=False)
        self.var_zeros = tk.BooleanVar(value=False)
        self.var_dec_exact = tk.BooleanVar(value=False)
        self.var_decimal_mode = tk.StringVar(value="ignore")

        self.src_filename = tk.StringVar()
        self.tgt_filename = tk.StringVar()

        self.summary_labels = {}

        self._ui()

    def _ui(self):
        # Global font
        self.root.option_add("*Font", ("Segoe UI", 9))
        self.root.option_add("*Label.Font", ("Segoe UI", 9))
        self.root.option_add("*Button.Font", ("Segoe UI", 10, "bold"))

        # Scrollable canvas
        canvas = tk.Canvas(self.root, bg="white")
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="white")

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Header - FULL WIDTH
        header = tk.Frame(scrollable_frame, bg="white")
        header.pack(fill="x", pady=10, padx=20)
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=1)

        # EUREKA LOGO (LEFT)
        left_frame = tk.Frame(header, bg="white")
        left_frame.grid(row=0, column=0, sticky="w")
        try:
            e_path = "E.png"
            if getattr(sys, '_MEIPASS', None):
                e_path = os.path.join(sys._MEIPASS, "E.png")
            e_img = Image.open(e_path)
            e_img = e_img.resize((120, 120), Image.Resampling.LANCZOS)
            e_photo = ImageTk.PhotoImage(e_img)
            tk.Label(left_frame, image=e_photo, bg="white").pack()
            left_frame.image = e_photo
        except:
            tk.Label(left_frame, text="EUREKA", font=("Segoe UI", 32, "bold"), fg=BOSCH_RED, bg="white").pack()
        tk.Label(left_frame, text="Developed by ASTRA Project Team", font=("Segoe UI", 10), fg="gray", bg="white").pack()

        # BOSCH LOGO (RIGHT)
        try:
            base_path = os.path.dirname(sys.argv[0]) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
            bosch_path = os.path.join(base_path, "Bosch.png")
            if getattr(sys, '_MEIPASS', None):
                bosch_path = os.path.join(sys._MEIPASS, "Bosch.png")
            bosch_img = Image.open(bosch_path)
            bosch_img = bosch_img.resize((100, 100), Image.Resampling.LANCZOS)
            bosch_photo = ImageTk.PhotoImage(bosch_img)
            tk.Label(header, image=bosch_photo, bg="white").grid(row=0, column=1, sticky="e")
            header.image2 = bosch_photo
        except:
            tk.Label(header, text="BOSCH", font=("Segoe UI", 24, "bold"), fg=BOSCH_RED, bg="white").grid(row=0, column=1, sticky="e")

        # Load + Progress Row - FULL WIDTH
        top_row = tk.Frame(scrollable_frame, bg="white")
        top_row.pack(fill="x", padx=20, pady=10)
        top_row.grid_columnconfigure(0, weight=1)
        top_row.grid_columnconfigure(1, weight=1)

        # Load Data + Filters
        load_box = tk.LabelFrame(top_row, text="LOAD DATA", bg="white", fg=BOSCH_RED, font=("Segoe UI", 10, "bold"))
        load_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        f = tk.Frame(load_box, bg="white")
        f.pack(fill="x", padx=10, pady=5)
        for i, (txt, key, var) in enumerate([
            ("Source CSV:", "src", self.src_filename),
            ("Target CSV:", "tgt", self.tgt_filename)
        ]):
            tk.Label(f, text=txt, bg="white").grid(row=i, column=0, sticky="w", pady=3)
            e = tk.Entry(f, width=50)
            e.grid(row=i, column=1, sticky="ew", padx=5)
            tk.Button(f, text="Browse", command=lambda k=key: getattr(self,f"browse_{k}")()).grid(row=i, column=2, padx=5)
            tk.Label(f, textvariable=var, bg="white", fg="#333").grid(row=i, column=3, sticky="w", padx=10)
            setattr(self, f"{key}_entry", e)
        f.grid_columnconfigure(1, weight=1)

        filters_inner = tk.Frame(load_box, bg="white")
        filters_inner.pack(fill="x", padx=10, pady=5)
        tk.Checkbutton(filters_inner, text="Match case (case-sensitive)", variable=self.var_case, bg="white").grid(row=0, column=0, sticky="w", padx=5)
        tk.Checkbutton(filters_inner, text="Ignore leading zeros", variable=self.var_zeros, bg="white").grid(row=0, column=1, sticky="w", padx=5)

        tk.Label(filters_inner, text="Decimals:", bg="white").grid(row=1, column=0, sticky="w", padx=5, pady=2)
        ttk.Radiobutton(filters_inner, text="Ignore trailing zeros", variable=self.var_decimal_mode, value="ignore").grid(row=1, column=1, sticky="w", padx=5)
        ttk.Radiobutton(filters_inner, text="Strict text match", variable=self.var_decimal_mode, value="strict").grid(row=2, column=1, sticky="w", padx=5)
        ttk.Radiobutton(filters_inner, text="Ignore decimals entirely", variable=self.var_decimal_mode, value="ignore_decimal").grid(row=3, column=1, sticky="w", padx=5)

        tk.Checkbutton(filters_inner, text="Compare decimals exactly", variable=self.var_dec_exact, bg="white").grid(row=4, column=0, columnspan=2, sticky="w", padx=5, pady=2)

        # Progress Box - FULL WIDTH
        prog_box = tk.LabelFrame(top_row, text="PROGRESS & STATUS", bg="white", fg=BOSCH_RED, font=("Segoe UI", 10, "bold"))
        prog_box.grid(row=0, column=1, sticky="nsew", padx=(10, 0))
        self.status_label = tk.Label(prog_box, text="STATUS: Ready.", bg="white", fg=STATUS_BLUE, font=("Segoe UI", 10, "bold"), anchor="w")
        self.status_label.pack(fill="x", padx=10, pady=5)
        self.p_cmp = CanvasProgress(prog_box, "Compare Progress", width=400)
        self.p_cmp.pack(fill="x", padx=10, pady=2)
        self.p_exp = CanvasProgress(prog_box, "Export Progress", width=400)
        self.p_exp.pack(fill="x", padx=10, pady=2)

        # Buttons
        btn_frame = tk.Frame(scrollable_frame, bg="white")
        btn_frame.pack(pady=10)

        def hover(btn, normal, hover):
            btn.bind("<Enter>", lambda e: btn.config(bg=hover))
            btn.bind("<Leave>", lambda e: btn.config(bg=normal))

        self.btn_cmp = tk.Button(btn_frame, text="COMPARE", bg=BOSCH_RED, fg="white", width=15, command=self.start_compare)
        self.btn_cmp.pack(side="left", padx=8)
        hover(self.btn_cmp, BOSCH_RED, HOVER_RED)

        self.btn_can = tk.Button(btn_frame, text="CANCEL", bg=BOSCH_RED, fg="white", width=15, command=self.cancel)
        self.btn_can.pack(side="left", padx=8)
        hover(self.btn_can, BOSCH_RED, HOVER_RED)

        self.btn_reset = tk.Button(btn_frame, text="RESET", bg=BOSCH_RED, fg="white", width=15, command=self.reset)
        self.btn_reset.pack(side="left", padx=8)
        hover(self.btn_reset, BOSCH_RED, HOVER_RED)

        self.btn_exp = tk.Button(btn_frame, text="EXPORT EXCEL", bg=GREEN, fg="white", width=15, command=self.export)
        self.btn_exp.pack(side="left", padx=8)
        hover(self.btn_exp, GREEN, HOVER_GREEN)

        # Summary - FULL WIDTH
        self.summary_frame = tk.LabelFrame(scrollable_frame, text="SUMMARY", bg="white", fg=BOSCH_RED, font=("Segoe UI", 10, "bold"))
        self.summary_frame.pack(fill="x", padx=20, pady=10)
        self.summary_frame.grid_columnconfigure(0, weight=1)
        self.summary_frame.grid_columnconfigure(1, weight=1)

        labels = [
            ("Source Rows", "Target Rows"),
            ("Source Cols", "Target Cols"),
            ("Rows with diff", "Cols Match"),
            ("Extra src rows", "Cells with diff"),
            ("Extra tgt rows", "")
        ]

        for row_idx, (l1, l2) in enumerate(labels):
            tk.Label(self.summary_frame, text=l1 + ":", bg="white", anchor="w").grid(row=row_idx, column=0, sticky="w", padx=15, pady=3)
            self.summary_labels[l1] = tk.Label(self.summary_frame, text="", bg="white", fg="blue", anchor="w")
            self.summary_labels[l1].grid(row=row_idx, column=1, sticky="w", padx=5, pady=3)

            if l2:
                tk.Label(self.summary_frame, text=l2 + ":", bg="white", anchor="w").grid(row=row_idx, column=2, sticky="w", padx=15, pady=3)
                self.summary_labels[l2] = tk.Label(self.summary_frame, text="", bg="white", fg="blue", anchor="w")
                self.summary_labels[l2].grid(row=row_idx, column=3, sticky="w", padx=5, pady=3)

        # Preview - FULL WIDTH + ZEBRA
        preview_frame = tk.LabelFrame(scrollable_frame, text="DIFFERENCES PREVIEW (500 max)", bg="white", fg=BOSCH_RED, font=("Segoe UI", 10, "bold"))
        preview_frame.pack(fill="both", expand=True, padx=20, pady=10)
        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        tf = tk.Frame(preview_frame)
        tf.grid(row=0, column=0, sticky="nsew")
        tf.grid_rowconfigure(0, weight=1)
        tf.grid_columnconfigure(0, weight=1)

        self.tree = ttk.Treeview(tf, columns=("Row","Col","Src","Tgt"), show="headings", height=20)
        for c,w in zip(("Row","Col","Src","Tgt"), (80,180,300,300)):
            self.tree.heading(c, text=c)
            self.tree.column(c, width=w, anchor="center")
        self.tree.tag_configure("even", background="#F9F9F9")
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.grid(row=0, column=1, sticky="ns")
        self.tree.grid(row=0, column=0, sticky="nsew")

        exp_frame = tk.Frame(preview_frame, bg="white")
        exp_frame.grid(row=1, column=0, pady=5)
        self.btn_exp_table = tk.Button(exp_frame, text="EXPORT EXCEL", bg=GREEN, fg="white", command=self.export)
        self.btn_exp_table.pack()
        hover(self.btn_exp_table, GREEN, HOVER_GREEN)

        # Scroll for whole app
        def _on_mousewheel(event):
            canvas.yview_scroll(-1*(event.delta//120), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def ui(self, fn, *a, **k):
        self.root.after(0, lambda: fn(*a, **k))

    def browse_src(self): self._browse("src", self.src_filename)
    def browse_tgt(self): self._browse("tgt", self.tgt_filename)
    def _browse(self, which, var):
        f = filedialog.askopenfilename(filetypes=[("CSV","*.csv")])
        if f:
            entry = getattr(self, f"{which}_entry")
            entry.delete(0, "end"); entry.insert(0, f)
            var.set(os.path.basename(f))

    def cancel(self):
        if self.stop.is_set():
            messagebox.showwarning("In Progress", "Operation already stopping.")
        else:
            self.stop.set(); self.status("STATUS: Cancelling...")

    def reset(self):
        if self.stop.is_set() or self.btn_cmp['state'] == 'disabled':
            messagebox.showwarning("In Progress", "Please wait until current operation completes.")
            return
        self.src_entry.delete(0, "end"); self.tgt_entry.delete(0, "end")
        self.src_filename.set(""); self.tgt_filename.set("")
        self.var_case.set(False); self.var_zeros.set(False); self.var_dec_exact.set(False)
        self.var_decimal_mode.set("ignore")
        for i in self.tree.get_children(): self.tree.delete(i)
        for val in self.summary_labels.values(): val.config(text="")
        self.status("STATUS: Ready.")
        self.p_cmp.start(1); self.p_cmp.update(0); self.p_exp.start(1); self.p_exp.update(0)

    def status(self, txt):
        self.ui(self.status_label.config, text=txt)

    def start_compare(self):
        if self.btn_cmp['state'] == 'disabled':
            messagebox.showwarning("In Progress", "Comparison already running.")
            return
        src = self.src_entry.get().strip()
        tgt = self.tgt_entry.get().strip()
        if not (src and tgt and os.path.exists(src) and os.path.exists(tgt)):
            messagebox.showerror("Error", "Please select two valid CSV files.")
            return

        match_case = self.var_case.get()
        ignore_zeros = self.var_zeros.get()
        dec_exact = self.var_dec_exact.get()
        decimal_mode = self.var_decimal_mode.get()

        self.stop.clear(); self.btn_cmp.config(state="disabled"); self.btn_can.config(state="normal")
        threading.Thread(
            target=self._compare,
            args=(src, tgt, match_case, ignore_zeros, decimal_mode, dec_exact),
            daemon=True
        ).start()

    def _compare(self, src, tgt, match_case, ignore_zeros, decimal_mode, dec_exact):
        try:
            self.status("STATUS: Scanning files...")
            total = min(*[headers_and_rowcount(p)[1] for p in (src,tgt)])
            self.ui(self.p_cmp.start, total or 1)

            def prog(v,_): self.ui(self.p_cmp.update, v)
            stats, preview = compare_streamed(src, tgt, prog, self.stop,
                                              match_case, ignore_zeros, decimal_mode, dec_exact)
            if self.stop.is_set():
                self.status("STATUS: Cancelled."); return

            self.stats = stats
            self.ui(lambda: self._show(preview, stats))
        except Exception as e:
            self.ui(messagebox.showerror, "Error", str(e))
        finally:
            self.ui(self.btn_cmp.config, state="normal")
            self.ui(self.btn_can.config, state="disabled")

    def _show(self, preview, stats):
        for i in self.tree.get_children(): self.tree.delete(i)
        for idx, r in enumerate(preview):
            tag = "even" if idx % 2 == 0 else ""
            self.tree.insert("", "end", values=r, tags=(tag,))
        self.p_cmp.complete()

        values = {
            "Source Rows": stats["src_rows"],
            "Target Rows": stats["tgt_rows"],
            "Source Cols": stats["src_cols"],
            "Target Cols": stats["tgt_cols"],
            "Rows with diff": stats["rows_with_diff"],
            "Cols Match": "Yes" if stats["cols_match"] else "No",
            "Cells with diff": stats["cells_with_diff"],
            "Extra src rows": stats["extra_src_rows"],
            "Extra tgt rows": stats["extra_tgt_rows"]
        }

        for label, val in values.items():
            self.summary_labels[label].config(text=str(val))

        diff_text = f"STATUS: Done — {stats['rows_with_diff']:,} rows, {stats['cells_with_diff']:,} cells differ."
        self.status(diff_text)

    def export(self):
        if not self.stats:
            messagebox.showwarning("No Data", "Run comparison first.")
            return
        if self.btn_exp['state'] == 'disabled':
            messagebox.showwarning("In Progress", "Export already running.")
            return
        out = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel","*.xlsx")])
        if not out: return
        self.stop.clear(); self.btn_exp.config(state="disabled"); self.btn_can.config(state="normal")
        threading.Thread(target=self._export, args=(out,), daemon=True).start()

    def _export(self, path):
        try:
            p = UIProxy(self, self.p_exp)
            p.start(self.stats.get("cells_with_diff",1))
            export_report(path, self.stats, p, self.stop)
            if not self.stop.is_set():
                self.status(f"STATUS: Saved: {os.path.basename(path)}")
                self.ui(messagebox.showinfo, "Saved", f"Report saved:\n{path}")
            p.complete()
        except Exception as e:
            self.ui(messagebox.showerror, "Error", str(e))
        finally:
            self.ui(self.btn_exp.config, state="normal")
            self.ui(self.btn_can.config, state="disabled")

def main():
    root = tk.Tk()
    DIRTApp(root)
    root.mainloop()

if __name__ == "__main__":
    if getattr(sys, 'frozen', False):
        os.chdir(sys._MEIPASS)
    main()