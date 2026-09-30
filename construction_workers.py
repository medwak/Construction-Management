"""
Construction Workforce Manager
------------------------------
A simple Windows desktop app for managing construction workers:
  * Workers  - add, edit, delete, search and filter worker records
  * Attendance - mark each worker Present / Half day / Absent per day
  * Payroll  - calculate days worked and pay due for any date range

Uses only the Python standard library (tkinter + sqlite3).
Data is stored in "workforce.db" next to this script / the .exe.
"""

import csv
import os
import sqlite3
import sys
import tkinter as tk
import tkinter.font as tkfont
from datetime import date, datetime, timedelta
from tkinter import filedialog, messagebox, ttk

APP_TITLE = "Construction Workforce Manager"
CURRENCY = "KES"  # change to your currency symbol/code if needed

TRADES = [
    "General Labourer", "Mason", "Carpenter", "Steel Fixer", "Plumber",
    "Electrician", "Painter", "Welder", "Tiler", "Machine Operator",
    "Foreman", "Other",
]
STATUSES = ["Active", "Inactive"]
ATTENDANCE_TAGS = {"Present": "present", "Half day": "half", "Absent": "absent"}


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def app_dir():
    """Folder where the database lives (works for .py and PyInstaller .exe)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


DB_PATH = os.path.join(app_dir(), "workforce.db")


def parse_date(text):
    try:
        return datetime.strptime(text.strip(), "%Y-%m-%d").date()
    except (ValueError, AttributeError):
        return None


def money(value):
    return f"{CURRENCY} {value or 0:,.2f}"


def fmt_days(value):
    return f"{value:g}"


def center_on(win, parent):
    win.update_idletasks()
    x = parent.winfo_rootx() + (parent.winfo_width() - win.winfo_width()) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - win.winfo_height()) // 3
    win.geometry(f"+{max(x, 0)}+{max(y, 0)}")


def sort_tree(tree, col, reverse):
    """Click a column heading to sort (numbers and money sort numerically)."""
    def key(v):
        s = v.replace(CURRENCY, "").replace(",", "").strip()
        try:
            return (0, float(s), "")
        except ValueError:
            return (1, 0.0, v.lower())

    items = [(tree.set(k, col), k) for k in tree.get_children("")]
    items.sort(key=lambda t: key(t[0]), reverse=reverse)
    for i, (_, k) in enumerate(items):
        tree.move(k, "", i)
    tree.heading(col, command=lambda: sort_tree(tree, col, not reverse))


def make_tree(parent, columns, widths, numeric=()):
    frame = ttk.Frame(parent)
    tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="extended")
    for col, width in zip(columns, widths):
        tree.heading(col, text=col, command=lambda c=col: sort_tree(tree, c, False))
        tree.column(col, width=width, minwidth=60, anchor="e" if col in numeric else "w")
    vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)
    return frame, tree


# --------------------------------------------------------------------------- #
# Database
# --------------------------------------------------------------------------- #
class Database:
    WORKER_FIELDS = ("full_name", "id_number", "phone", "trade", "daily_rate",
                     "site", "start_date", "status", "notes")

    def __init__(self, path):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS workers (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name   TEXT NOT NULL,
                id_number   TEXT,
                phone       TEXT,
                trade       TEXT,
                daily_rate  REAL NOT NULL DEFAULT 0,
                site        TEXT,
                start_date  TEXT,
                status      TEXT NOT NULL DEFAULT 'Active',
                notes       TEXT
            );
            CREATE TABLE IF NOT EXISTS attendance (
                worker_id INTEGER NOT NULL REFERENCES workers(id) ON DELETE CASCADE,
                day       TEXT NOT NULL,
                status    TEXT NOT NULL,
                PRIMARY KEY (worker_id, day)
            );
        """)
        self.conn.commit()

    # ---- workers ----
    def list_workers(self, search="", status="All", site="All"):
        sql, params = "SELECT * FROM workers WHERE 1=1", []
        if search:
            like = f"%{search}%"
            sql += (" AND (full_name LIKE ? OR id_number LIKE ? OR phone LIKE ?"
                    " OR trade LIKE ? OR site LIKE ?)")
            params += [like] * 5
        if status != "All":
            sql += " AND status = ?"
            params.append(status)
        if site != "All":
            sql += " AND site = ?"
            params.append(site)
        sql += " ORDER BY full_name COLLATE NOCASE"
        return self.conn.execute(sql, params).fetchall()

    def get_worker(self, wid):
        return self.conn.execute("SELECT * FROM workers WHERE id = ?", (wid,)).fetchone()

    def add_worker(self, data):
        cols = ", ".join(self.WORKER_FIELDS)
        marks = ", ".join("?" for _ in self.WORKER_FIELDS)
        self.conn.execute(f"INSERT INTO workers ({cols}) VALUES ({marks})",
                          [data[f] for f in self.WORKER_FIELDS])
        self.conn.commit()

    def update_worker(self, wid, data):
        sets = ", ".join(f"{f} = ?" for f in self.WORKER_FIELDS)
        self.conn.execute(f"UPDATE workers SET {sets} WHERE id = ?",
                          [data[f] for f in self.WORKER_FIELDS] + [wid])
        self.conn.commit()

    def delete_workers(self, ids):
        self.conn.executemany("DELETE FROM workers WHERE id = ?", [(i,) for i in ids])
        self.conn.commit()

    def sites(self):
        rows = self.conn.execute(
            "SELECT DISTINCT site FROM workers WHERE site IS NOT NULL AND site <> '' "
            "ORDER BY site COLLATE NOCASE").fetchall()
        return [r["site"] for r in rows]

    def counts(self):
        r = self.conn.execute(
            "SELECT COUNT(*) AS total, "
            "SUM(CASE WHEN status = 'Active' THEN 1 ELSE 0 END) AS active FROM workers"
        ).fetchone()
        return r["total"] or 0, r["active"] or 0

    # ---- attendance ----
    def attendance_for_day(self, day, site="All"):
        sql = ("SELECT w.id, w.full_name, w.trade, w.site, a.status "
               "FROM workers w LEFT JOIN attendance a "
               "ON a.worker_id = w.id AND a.day = ? "
               "WHERE w.status = 'Active'")
        params = [day]
        if site != "All":
            sql += " AND w.site = ?"
            params.append(site)
        sql += " ORDER BY w.full_name COLLATE NOCASE"
        return self.conn.execute(sql, params).fetchall()

    def set_attendance(self, worker_ids, day, status):
        if status is None:
            self.conn.executemany("DELETE FROM attendance WHERE worker_id = ? AND day = ?",
                                  [(w, day) for w in worker_ids])
        else:
            self.conn.executemany(
                "INSERT OR REPLACE INTO attendance (worker_id, day, status) VALUES (?, ?, ?)",
                [(w, day, status) for w in worker_ids])
        self.conn.commit()

    # ---- payroll ----
    def payroll(self, start, end, site="All"):
        sql = ("SELECT w.id, w.full_name, w.trade, w.site, w.daily_rate, "
               "SUM(CASE a.status WHEN 'Present' THEN 1.0 "
               "                  WHEN 'Half day' THEN 0.5 ELSE 0 END) AS days, "
               "SUM(CASE WHEN a.status = 'Absent' THEN 1 ELSE 0 END) AS absences "
               "FROM workers w JOIN attendance a ON a.worker_id = w.id "
               "WHERE a.day BETWEEN ? AND ?")
        params = [start, end]
        if site != "All":
            sql += " AND w.site = ?"
            params.append(site)
        sql += " GROUP BY w.id ORDER BY w.full_name COLLATE NOCASE"
        return self.conn.execute(sql, params).fetchall()


# --------------------------------------------------------------------------- #
# Add / edit worker dialog
# --------------------------------------------------------------------------- #
class WorkerDialog(tk.Toplevel):
    def __init__(self, parent, title, sites, worker=None):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)
        self.resizable(False, False)
        self.result = None
        w = dict(worker) if worker else {}

        frm = ttk.Frame(self, padding=18)
        frm.grid(sticky="nsew")
        frm.columnconfigure(1, weight=1)

        self.vars = {k: tk.StringVar() for k in
                     ("full_name", "id_number", "phone", "trade", "daily_rate",
                      "site", "start_date", "status")}
        for k in ("full_name", "id_number", "phone", "trade", "site"):
            self.vars[k].set(w.get(k) or "")
        rate = w.get("daily_rate")
        self.vars["daily_rate"].set("" if rate is None else f"{rate:g}")
        self.vars["start_date"].set(w.get("start_date") or date.today().isoformat())
        self.vars["status"].set(w.get("status") or "Active")

        def row(r, label, widget):
            ttk.Label(frm, text=label).grid(row=r, column=0, sticky="w", pady=5, padx=(0, 14))
            widget.grid(row=r, column=1, sticky="ew", pady=5)
            return widget

        name = row(0, "Full name *", ttk.Entry(frm, textvariable=self.vars["full_name"], width=36))
        row(1, "ID number", ttk.Entry(frm, textvariable=self.vars["id_number"]))
        row(2, "Phone", ttk.Entry(frm, textvariable=self.vars["phone"]))
        row(3, "Trade / role", ttk.Combobox(frm, textvariable=self.vars["trade"], values=TRADES))
        row(4, f"Daily rate ({CURRENCY}) *", ttk.Entry(frm, textvariable=self.vars["daily_rate"]))
        row(5, "Site", ttk.Combobox(frm, textvariable=self.vars["site"], values=sites))
        row(6, "Start date (YYYY-MM-DD)", ttk.Entry(frm, textvariable=self.vars["start_date"]))
        row(7, "Status", ttk.Combobox(frm, textvariable=self.vars["status"],
                                      values=STATUSES, state="readonly"))

        ttk.Label(frm, text="Notes").grid(row=8, column=0, sticky="nw", pady=5)
        self.notes = tk.Text(frm, width=36, height=4, wrap="word",
                             font=tkfont.nametofont("TkDefaultFont"))
        self.notes.grid(row=8, column=1, sticky="ew", pady=5)
        self.notes.insert("1.0", w.get("notes") or "")

        btns = ttk.Frame(frm)
        btns.grid(row=9, column=0, columnspan=2, sticky="e", pady=(14, 0))
        ttk.Button(btns, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="Save", command=self.save).pack(side="right", padx=(0, 8))

        self.bind("<Return>", lambda e: None if e.widget is self.notes else self.save())
        self.bind("<Escape>", lambda e: self.destroy())
        center_on(self, parent)
        name.focus_set()
        self.grab_set()
        self.wait_window()

    def save(self):
        name = self.vars["full_name"].get().strip()
        if not name:
            messagebox.showwarning("Missing name", "Please enter the worker's full name.", parent=self)
            return
        try:
            rate = float(self.vars["daily_rate"].get().replace(",", "").strip())
            if rate < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Invalid rate",
                                   "Daily rate must be a number, e.g. 1500.", parent=self)
            return
        start = self.vars["start_date"].get().strip()
        if start and not parse_date(start):
            messagebox.showwarning("Invalid date",
                                   "Start date must look like 2026-09-28.", parent=self)
            return
        self.result = {
            "full_name": name,
            "id_number": self.vars["id_number"].get().strip(),
            "phone": self.vars["phone"].get().strip(),
            "trade": self.vars["trade"].get().strip(),
            "daily_rate": rate,
            "site": self.vars["site"].get().strip(),
            "start_date": start,
            "status": self.vars["status"].get() or "Active",
            "notes": self.notes.get("1.0", "end").strip(),
        }
        self.destroy()


# --------------------------------------------------------------------------- #
# Main window
# --------------------------------------------------------------------------- #
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("1100x680")
        self.minsize(860, 520)
        self.db = Database(DB_PATH)
        self.site_combos = []
        self.setup_style()

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(10, 0))
        self.build_workers_tab(nb)
        self.build_attendance_tab(nb)
        self.build_payroll_tab(nb)
        nb.bind("<<NotebookTabChanged>>", lambda e: self.on_tab_change(nb))

        self.status = ttk.Label(self, anchor="w", padding=(12, 6))
        self.status.pack(fill="x")

        self.refresh_all()

    # ---------- look & feel ----------
    def setup_style(self):
        style = ttk.Style(self)
        for theme in ("vista", "winnative", "clam"):
            if theme in style.theme_names():
                style.theme_use(theme)
                break
        base = tkfont.nametofont("TkDefaultFont")
        base.configure(size=10)
        tkfont.nametofont("TkTextFont").configure(size=10)
        bold = base.copy()
        bold.configure(weight="bold")
        self.bold_font = bold
        style.configure("Treeview", rowheight=base.metrics("linespace") + 10)
        style.configure("Treeview.Heading", font=bold)
        style.configure("Total.TLabel", font=bold)

    # ---------- shared ----------
    def site_filter(self, parent, var, callback):
        ttk.Label(parent, text="Site").pack(side="left")
        cb = ttk.Combobox(parent, textvariable=var, state="readonly", width=18)
        cb.pack(side="left", padx=(6, 16))
        cb.bind("<<ComboboxSelected>>", lambda e: callback())
        self.site_combos.append((cb, var))
        return cb

    def refresh_sites(self):
        values = ["All"] + self.db.sites()
        for cb, var in self.site_combos:
            cb["values"] = values
            if var.get() not in values:
                var.set("All")

    def refresh_all(self):
        self.refresh_sites()
        self.refresh_workers()
        self.refresh_attendance()
        self.refresh_payroll(silent=True)

    def update_status_bar(self):
        total, active = self.db.counts()
        self.status.config(text=f"{total} workers  •  {active} active  •  "
                                f"{len(self.db.sites())} sites  •  Data: {DB_PATH}")

    def on_tab_change(self, nb):
        idx = nb.index(nb.select())
        if idx == 1:
            self.refresh_attendance()
        elif idx == 2:
            self.refresh_payroll(silent=True)

    def export_csv(self, tree, default_name, extra_rows=()):
        if not tree.get_children():
            messagebox.showinfo("Nothing to export", "There are no rows to export.", parent=self)
            return
        path = filedialog.asksaveasfilename(
            parent=self, defaultextension=".csv", initialfile=default_name,
            filetypes=[("CSV (Excel)", "*.csv"), ("All files", "*.*")])
        if not path:
            return
        try:
            with open(path, "w", newline="", encoding="utf-8-sig") as f:
                writer = csv.writer(f)
                writer.writerow(tree["columns"])
                for item in tree.get_children():
                    writer.writerow(tree.item(item, "values"))
                for r in extra_rows:
                    writer.writerow(r)
        except OSError as e:
            messagebox.showerror("Export failed", str(e), parent=self)
            return
        messagebox.showinfo("Exported", f"Saved to:\n{path}", parent=self)

    # ================= Workers tab =================
    def build_workers_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)
        nb.add(tab, text="   Workers   ")

        bar = ttk.Frame(tab)
        bar.pack(fill="x")
        ttk.Label(bar, text="Search").pack(side="left")
        self.w_search = tk.StringVar()
        ttk.Entry(bar, textvariable=self.w_search, width=28).pack(side="left", padx=(6, 16))
        self.w_search.trace_add("write", lambda *a: self.refresh_workers())

        ttk.Label(bar, text="Status").pack(side="left")
        self.w_status = tk.StringVar(value="All")
        cb = ttk.Combobox(bar, textvariable=self.w_status, values=["All"] + STATUSES,
                          state="readonly", width=10)
        cb.pack(side="left", padx=(6, 16))
        cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_workers())

        self.w_site = tk.StringVar(value="All")
        self.site_filter(bar, self.w_site, self.refresh_workers)

        cols = ("Name", "ID No.", "Phone", "Trade", "Site", "Daily rate", "Start date", "Status")
        frame, self.w_tree = make_tree(tab, cols, (210, 110, 120, 140, 150, 120, 100, 80),
                                       numeric={"Daily rate"})
        frame.pack(fill="both", expand=True, pady=10)
        self.w_tree.tag_configure("inactive", foreground="#8a8a8a")
        self.w_tree.bind("<Double-1>", lambda e: self.edit_worker())
        self.w_tree.bind("<Return>", lambda e: self.edit_worker())
        self.w_tree.bind("<Delete>", lambda e: self.delete_workers())

        btns = ttk.Frame(tab)
        btns.pack(fill="x")
        ttk.Button(btns, text="＋ Add worker", command=self.add_worker).pack(side="left")
        ttk.Button(btns, text="Edit", command=self.edit_worker).pack(side="left", padx=6)
        ttk.Button(btns, text="Delete", command=self.delete_workers).pack(side="left")
        ttk.Button(btns, text="Export to CSV",
                   command=lambda: self.export_csv(self.w_tree, "workers.csv")).pack(side="right")

    def refresh_workers(self):
        rows = self.db.list_workers(self.w_search.get().strip(),
                                    self.w_status.get(), self.w_site.get())
        self.w_tree.delete(*self.w_tree.get_children())
        for r in rows:
            self.w_tree.insert(
                "", "end", iid=str(r["id"]),
                values=(r["full_name"], r["id_number"] or "", r["phone"] or "",
                        r["trade"] or "", r["site"] or "", money(r["daily_rate"]),
                        r["start_date"] or "", r["status"]),
                tags=() if r["status"] == "Active" else ("inactive",))
        self.update_status_bar()

    def add_worker(self):
        dlg = WorkerDialog(self, "Add worker", self.db.sites())
        if dlg.result:
            self.db.add_worker(dlg.result)
            self.refresh_all()

    def edit_worker(self):
        sel = self.w_tree.selection()
        if not sel:
            messagebox.showinfo("Edit worker", "Select a worker to edit.", parent=self)
            return
        wid = int(sel[0])
        dlg = WorkerDialog(self, "Edit worker", self.db.sites(), self.db.get_worker(wid))
        if dlg.result:
            self.db.update_worker(wid, dlg.result)
            self.refresh_all()
            if self.w_tree.exists(str(wid)):
                self.w_tree.selection_set(str(wid))
                self.w_tree.see(str(wid))

    def delete_workers(self):
        sel = self.w_tree.selection()
        if not sel:
            messagebox.showinfo("Delete", "Select one or more workers to delete.", parent=self)
            return
        names = [self.w_tree.set(i, "Name") for i in sel]
        who = names[0] if len(names) == 1 else f"{len(names)} workers"
        if messagebox.askyesno(
                "Delete worker",
                f"Delete {who} and all their attendance records?\n\n"
                "This cannot be undone. Tip: set the status to Inactive instead "
                "if you want to keep their history.",
                icon="warning", parent=self):
            self.db.delete_workers([int(i) for i in sel])
            self.refresh_all()

    # ================= Attendance tab =================
    def build_attendance_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)
        nb.add(tab, text="   Attendance   ")

        bar = ttk.Frame(tab)
        bar.pack(fill="x")
        ttk.Label(bar, text="Date").pack(side="left")
        self.a_date = tk.StringVar(value=date.today().isoformat())
        ttk.Button(bar, text="◀", width=3,
                   command=lambda: self.shift_day(-1)).pack(side="left", padx=(6, 2))
        ent = ttk.Entry(bar, textvariable=self.a_date, width=12, justify="center")
        ent.pack(side="left")
        ent.bind("<Return>", lambda e: self.refresh_attendance())
        ent.bind("<FocusOut>", lambda e: self.refresh_attendance())
        ttk.Button(bar, text="▶", width=3,
                   command=lambda: self.shift_day(1)).pack(side="left", padx=(2, 6))
        ttk.Button(bar, text="Today", command=self.go_today).pack(side="left")
        self.a_day_label = ttk.Label(bar, style="Total.TLabel")
        self.a_day_label.pack(side="left", padx=16)

        self.a_site = tk.StringVar(value="All")
        self.site_filter(bar, self.a_site, self.refresh_attendance)

        cols = ("Name", "Trade", "Site", "Attendance")
        frame, self.a_tree = make_tree(tab, cols, (260, 180, 200, 140))
        frame.pack(fill="both", expand=True, pady=10)
        self.a_tree.tag_configure("present", background="#e3f4e6")
        self.a_tree.tag_configure("half", background="#fff3dc")
        self.a_tree.tag_configure("absent", background="#fbe4e4")
        for key, status in (("p", "Present"), ("h", "Half day"), ("a", "Absent")):
            for k in (key, key.upper()):
                self.a_tree.bind(f"<KeyPress-{k}>", lambda e, s=status: self.mark(s))
        self.a_tree.bind("<BackSpace>", lambda e: self.mark(None))
        self.a_tree.bind("<Control-a>", lambda e: self.a_tree.selection_set(
            self.a_tree.get_children()))

        btns = ttk.Frame(tab)
        btns.pack(fill="x")
        ttk.Button(btns, text="✔ Present (P)",
                   command=lambda: self.mark("Present")).pack(side="left")
        ttk.Button(btns, text="½ Half day (H)",
                   command=lambda: self.mark("Half day")).pack(side="left", padx=6)
        ttk.Button(btns, text="✘ Absent (A)",
                   command=lambda: self.mark("Absent")).pack(side="left")
        ttk.Button(btns, text="Clear",
                   command=lambda: self.mark(None)).pack(side="left", padx=6)
        ttk.Button(btns, text="Mark all unmarked as Present",
                   command=self.mark_unmarked_present).pack(side="right")

        self.a_summary = ttk.Label(tab, padding=(0, 8, 0, 0))
        self.a_summary.pack(fill="x")
        ttk.Label(tab, foreground="#6b6b6b",
                  text="Tip: Ctrl/Shift+click to select several workers, then press "
                       "P, H or A. Only Active workers are listed.").pack(fill="x")

    def current_day(self):
        return parse_date(self.a_date.get())

    def shift_day(self, n):
        d = self.current_day() or date.today()
        self.a_date.set((d + timedelta(days=n)).isoformat())
        self.refresh_attendance()

    def go_today(self):
        self.a_date.set(date.today().isoformat())
        self.refresh_attendance()

    def refresh_attendance(self):
        d = self.current_day()
        selected = set(self.a_tree.selection())
        self.a_tree.delete(*self.a_tree.get_children())
        if not d:
            self.a_day_label.config(text="Invalid date — use YYYY-MM-DD")
            self.a_summary.config(text="")
            return
        self.a_day_label.config(text=d.strftime("%A, %d %B %Y"))
        counts = {"Present": 0, "Half day": 0, "Absent": 0, "": 0}
        for r in self.db.attendance_for_day(d.isoformat(), self.a_site.get()):
            status = r["status"] or ""
            counts[status] = counts.get(status, 0) + 1
            iid = str(r["id"])
            self.a_tree.insert("", "end", iid=iid,
                               values=(r["full_name"], r["trade"] or "", r["site"] or "",
                                       status or "—"),
                               tags=(ATTENDANCE_TAGS.get(status, ""),))
        keep = [i for i in selected if self.a_tree.exists(i)]
        if keep:
            self.a_tree.selection_set(keep)
        self.a_summary.config(
            text=f"Present: {counts['Present']}     Half day: {counts['Half day']}     "
                 f"Absent: {counts['Absent']}     Not marked: {counts['']}")

    def mark(self, status):
        d = self.current_day()
        if not d:
            messagebox.showwarning("Invalid date", "Enter a date like 2026-09-28.", parent=self)
            return
        sel = self.a_tree.selection()
        if not sel:
            messagebox.showinfo("Attendance", "Select one or more workers first.", parent=self)
            return
        self.db.set_attendance([int(i) for i in sel], d.isoformat(), status)
        self.refresh_attendance()
        self.a_tree.focus_set()

    def mark_unmarked_present(self):
        d = self.current_day()
        if not d:
            return
        ids = [int(i) for i in self.a_tree.get_children()
               if self.a_tree.set(i, "Attendance") == "—"]
        if not ids:
            messagebox.showinfo("Attendance", "Everyone is already marked.", parent=self)
            return
        self.db.set_attendance(ids, d.isoformat(), "Present")
        self.refresh_attendance()

    # ================= Payroll tab =================
    def build_payroll_tab(self, nb):
        tab = ttk.Frame(nb, padding=12)
        nb.add(tab, text="   Payroll   ")

        today = date.today()
        self.p_from = tk.StringVar(value=(today - timedelta(days=today.weekday())).isoformat())
        self.p_to = tk.StringVar(value=today.isoformat())

        bar = ttk.Frame(tab)
        bar.pack(fill="x")
        ttk.Label(bar, text="From").pack(side="left")
        e1 = ttk.Entry(bar, textvariable=self.p_from, width=12, justify="center")
        e1.pack(side="left", padx=(6, 12))
        ttk.Label(bar, text="To").pack(side="left")
        e2 = ttk.Entry(bar, textvariable=self.p_to, width=12, justify="center")
        e2.pack(side="left", padx=(6, 12))
        for e in (e1, e2):
            e.bind("<Return>", lambda ev: self.refresh_payroll())
        ttk.Button(bar, text="This week", command=self.this_week).pack(side="left")
        ttk.Button(bar, text="This month", command=self.this_month).pack(side="left", padx=(6, 16))

        self.p_site = tk.StringVar(value="All")
        self.site_filter(bar, self.p_site, lambda: self.refresh_payroll(silent=True))
        ttk.Button(bar, text="Calculate", command=self.refresh_payroll).pack(side="left")

        cols = ("Name", "Trade", "Site", "Days worked", "Absences", "Daily rate", "Amount due")
        frame, self.p_tree = make_tree(tab, cols, (220, 150, 160, 110, 90, 120, 140),
                                       numeric={"Days worked", "Absences",
                                                "Daily rate", "Amount due"})
        frame.pack(fill="both", expand=True, pady=10)

        bottom = ttk.Frame(tab)
        bottom.pack(fill="x")
        self.p_total = ttk.Label(bottom, style="Total.TLabel")
        self.p_total.pack(side="left")
        ttk.Button(bottom, text="Export to CSV", command=self.export_payroll).pack(side="right")
        ttk.Label(tab, foreground="#6b6b6b",
                  text="Pay = days worked × daily rate (a half day counts as 0.5). "
                       "Based on the attendance records in the chosen period."
                  ).pack(fill="x", pady=(6, 0))
        self.p_total_value = 0.0

    def this_week(self):
        t = date.today()
        self.p_from.set((t - timedelta(days=t.weekday())).isoformat())
        self.p_to.set(t.isoformat())
        self.refresh_payroll()

    def this_month(self):
        t = date.today()
        self.p_from.set(t.replace(day=1).isoformat())
        self.p_to.set(t.isoformat())
        self.refresh_payroll()

    def refresh_payroll(self, silent=False):
        start, end = parse_date(self.p_from.get()), parse_date(self.p_to.get())
        if not start or not end or start > end:
            if not silent:
                messagebox.showwarning("Invalid period",
                                       "Enter valid dates (YYYY-MM-DD) with From before To.",
                                       parent=self)
            return
        self.p_tree.delete(*self.p_tree.get_children())
        total_amount, total_days = 0.0, 0.0
        for r in self.db.payroll(start.isoformat(), end.isoformat(), self.p_site.get()):
            days = r["days"] or 0.0
            amount = days * (r["daily_rate"] or 0)
            total_amount += amount
            total_days += days
            self.p_tree.insert("", "end", iid=str(r["id"]),
                               values=(r["full_name"], r["trade"] or "", r["site"] or "",
                                       fmt_days(days), r["absences"] or 0,
                                       money(r["daily_rate"]), money(amount)))
        n = len(self.p_tree.get_children())
        self.p_total_value = total_amount
        self.p_total.config(
            text=f"{start:%d %b %Y} – {end:%d %b %Y}:   {n} workers   •   "
                 f"{fmt_days(total_days)} worker-days   •   Total due: {money(total_amount)}")

    def export_payroll(self):
        self.export_csv(
            self.p_tree, f"payroll_{self.p_from.get()}_to_{self.p_to.get()}.csv",
            extra_rows=[("TOTAL", "", "", "", "", "", money(self.p_total_value))])


# --------------------------------------------------------------------------- #
def main():
    if sys.platform == "win32":
        try:  # crisp text on high-DPI screens
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()


if __name__ == "__main__":
    main()
