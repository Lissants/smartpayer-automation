"""
smartpayer_gui.py
Single Tkinter GUI for Smartpayer Automation + letter generation.

Run with:  python smartpayer_gui.py
"""

import json
import os
import re
import subprocess
import sys
import threading
import uuid
from datetime import date, datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkcalendar import Calendar

os.chdir(os.path.dirname(os.path.abspath(__file__)))

import smartpayer_automation as sa


# =============================================================================
# PATHS
# =============================================================================

BASE_DIR = Path(__file__).resolve().parent
SCRIPT_DIR = BASE_DIR / "script"
GENERATOR = SCRIPT_DIR / "smartpayer_letter_generator.py"
EMAILER = SCRIPT_DIR / "smartpayer_emailer.py"
LETTER_DEFAULTS = SCRIPT_DIR / "smartpayer_letter_defaults.json"
EMAIL_CFG = SCRIPT_DIR / "smartpayer_email_config.json"
OUTPUT_DIR = BASE_DIR / "OUTPUT"
LETTER_OUTPUT_DIR = BASE_DIR / "Generated_Letters"


# =============================================================================
# COLOURS
# =============================================================================

BG          = "#F5F7FA"
PANEL_BG    = "#FFFFFF"
ACCENT      = "#1E3A5F"
ACCENT_LITE = "#2E5FA3"
BTN_FG      = "#FFFFFF"
SUCCESS     = "#217A3C"
WARN        = "#B45309"
BORDER      = "#D1D5DB"
TEXT        = "#111827"
SUBTEXT     = "#6B7280"
DANGER      = "#B91C1C"

PAD = 12


# =============================================================================
# LETTER DEFAULTS SCHEMA
# =============================================================================

DEFAULT_SCHEMA = [
    ("Company", [
        ("company_name", "Company Name", "PT. Amerta Indah Otsuka"),
        ("company_abbr", "Company Abbreviation", "AIO"),
    ]),
    ("Date & Period", [
        ("letter_date", "Letter Date", "(always today's date)"),
        ("month_name", "Month Name", "(auto-derived from filename)"),
        ("month_year", "Month & Year", "(auto-derived from filename)"),
        ("month_roman", "Roman Month", "(always today's date)"),
        ("letter_number", "Letter Number", "(always today's date + batch counter)"),
        ("year_num", "Last two digits of year", "(always today's date)"),
    ]),
    ("Financial", [
        ("deposit_rate", "Deposit Interest Rate (%/year)", "2"),
        ("deposit_days", "Deposit Days", "365"),
        ("daily_rate", "Daily Reward Rate (%/day)", "(auto-derived from XLSX)"),
        ("annual_rate", "Annualized Reward Rate (%/year)", "(auto-derived from XLSX)"),
        ("rate_multiplier_num", "Rate Multiplier Number", "(auto-derived from XLSX)"),
        ("rate_multiplier_word", "Rate Multiplier Word", "(auto-derived from XLSX)"),
    ]),
    ("Invoice Period", [
        ("end_of_month_date", "End of Next Month", "(auto-derived from filename)"),
        ("discount_from", "Discount Period From", "(auto-derived from XLSX)"),
        ("discount_until", "Discount Period Until", "(auto-derived from XLSX)"),
    ]),
    ("Finance Team", [
        ("team1_honorific_name", "Team 1 Honorific Name", "Bapak/Ibu ..."),
        ("team2_honorific_name", "Team 2 Honorific Name", "Bapak/Ibu ..."),
        ("team_email", "Finance Team Email", "finance@example.com"),
        ("team3_honorific_name", "Team 3 Honorific Name", "Bapak/Ibu ..."),
        ("team3_email", "Team 3 Email", "finance3@example.com"),
        ("team4_honorific_name", "Team 4 Honorific Name", "Bapak/Ibu ..."),
        ("team4_email", "Team 4 Email", "finance4@example.com"),
    ]),
]

AUTO_TODAY_KEYS = {"letter_date", "month_roman", "letter_number", "year_num"}
AUTO_DERIVED_KEYS = {
    "month_name", "month_year", "daily_rate", "annual_rate",
    "rate_multiplier_num", "rate_multiplier_word", "end_of_month_date",
    "discount_from", "discount_until",
}


# =============================================================================
# HELPERS
# =============================================================================

def load_json(path):
    if path.exists():
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def organize_files_by_prefix(folder_path: Path, progress_cb=None) -> tuple:
    """
    Organize files in folder_path into subfolders based on filename prefix.
    Pattern: "SmartPayer <Month> <Year> *" → folder "SmartPayer <Month> <Year>"

    Returns: (files_moved, folders_created)
    """
    import shutil

    pattern = re.compile(
        r'^(SmartPayer\s+[A-Za-z]+\s+\d{4})\s+',
        re.IGNORECASE
    )

    files = [f for f in folder_path.iterdir()
             if f.is_file() and not f.name.startswith("~$")]

    if not files:
        if progress_cb:
            progress_cb("No files found to organize.")
        return 0, 0

    groups = {}
    for f in files:
        match = pattern.match(f.name)
        if match:
            prefix = match.group(1).strip()
            groups.setdefault(prefix, []).append(f)

    if not groups:
        if progress_cb:
            progress_cb("No files match 'SmartPayer <Month> <Year>' pattern.")
        return 0, 0

    moved = 0
    created = 0

    for prefix, file_list in sorted(groups.items()):
        safe_name = re.sub(r'[\/:*?"<>|]', '_', prefix)
        target_folder = folder_path / safe_name
        target_folder.mkdir(exist_ok=True)
        created += 1

        if progress_cb:
            progress_cb(f"  Creating: {safe_name}/ ({len(file_list)} files)")

        for f in file_list:
            try:
                dest = target_folder / f.name
                if dest.exists():
                    stem, suffix = f.stem.rsplit('_', 1) if '_' in f.stem else (f.stem, '')
                    counter = 1
                    while dest.exists():
                        new_name = f"{stem}_{counter}{suffix}{f.suffix}"
                        dest = target_folder / new_name
                        counter += 1
                shutil.move(str(f), str(dest))
                moved += 1
                if progress_cb:
                    progress_cb(f"    → {f.name}")
            except Exception as e:
                if progress_cb:
                    progress_cb(f"    ✗ Error moving {f.name}: {e}")

    return moved, created


def make_button(parent, text, command,
                width=18, bg=ACCENT, fg=BTN_FG,
                font_size=9, bold=True):
    """Standard flat button used throughout the app."""
    weight = "bold" if bold else "normal"
    btn = tk.Button(
        parent,
        text=text,
        command=command,
        width=width,
        bg=bg,
        fg=fg,
        font=("Segoe UI", font_size, weight),
        relief="flat",
        bd=0,
        padx=10,
        pady=6,
        cursor="hand2",
        activebackground=ACCENT_LITE,
        activeforeground=BTN_FG,
    )
    btn.bind("<Enter>", lambda e: btn.config(bg=ACCENT_LITE))
    btn.bind("<Leave>", lambda e: btn.config(bg=bg))
    return btn


def make_entry(parent, var, width=None, readonly=False):
    entry = tk.Entry(parent, textvariable=var, width=width,
                     font=("Segoe UI", 9), relief="solid", bd=1)
    if readonly:
        entry.config(state="readonly", readonlybackground="#EEF2F7", fg=SUBTEXT)
    return entry


class SectionLabel(tk.Label):
    def __init__(self, parent, text, **kwargs):
        super().__init__(parent, text=text, bg=PANEL_BG,
                         fg=ACCENT, font=("Segoe UI", 10, "bold"), **kwargs)


class HRule(tk.Frame):
    def __init__(self, parent, **kwargs):
        super().__init__(parent, bg=BORDER, height=1, **kwargs)


def scroll_frame(parent):
    outer = tk.Frame(parent, bg=PANEL_BG)
    canvas = tk.Canvas(outer, bg=PANEL_BG, highlightthickness=0)
    scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    inner = tk.Frame(canvas, bg=PANEL_BG)
    inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.configure(yscrollcommand=scroll.set)
    canvas.pack(side="left", fill="both", expand=True)
    scroll.pack(side="right", fill="y")
    return outer, inner


# =============================================================================
# CALENDAR POPUP / DATE PICKER
# =============================================================================

class CalendarPopup(tk.Toplevel):
    def __init__(self, parent, title, initial_date=None, callback=None):
        super().__init__(parent)
        self.title(title)
        self.resizable(False, False)
        self.grab_set()
        self.configure(bg=PANEL_BG)
        self.result = None
        self.callback = callback

        today = initial_date or date.today()
        self.cal = Calendar(
            self,
            selectmode="day",
            year=today.year,
            month=today.month,
            day=today.day,
            date_pattern="dd-mm-yyyy",
            background=ACCENT,
            foreground=BTN_FG,
            selectbackground=ACCENT_LITE,
            headersbackground=ACCENT,
            headersforeground=BTN_FG,
            normalbackground=PANEL_BG,
            normalforeground=TEXT,
            weekendbackground=PANEL_BG,
            weekendforeground=TEXT,
            othermonthbackground="#E5E7EB",
            othermonthforeground=SUBTEXT,
            bordercolor=BORDER,
        )
        self.cal.pack(padx=PAD, pady=PAD)

        btn_frame = tk.Frame(self, bg=PANEL_BG)
        btn_frame.pack(fill="x", padx=PAD, pady=(0, PAD))
        make_button(btn_frame, "Select", self._select, width=10,
                    bg=SUCCESS).pack(side="right", padx=(6, 0))
        make_button(btn_frame, "Cancel", self.destroy, width=10,
                    bg="#6B7280").pack(side="right")
        self._center(parent)

    def _select(self):
        raw = self.cal.get_date()
        d, m, y = raw.split("-")
        self.result = date(int(y), int(m), int(d))
        if self.callback:
            self.callback(self.result)
        self.destroy()

    def _center(self, parent):
        self.update_idletasks()
        px = parent.winfo_rootx() + parent.winfo_width() // 2 - self.winfo_width() // 2
        py = parent.winfo_rooty() + parent.winfo_height() // 2 - self.winfo_height() // 2
        self.geometry(f"+{px}+{py}")


class DatePickerRow(tk.Frame):
    def __init__(self, parent, label, optional=False):
        super().__init__(parent, bg=PANEL_BG)
        self._optional = optional
        self._date = None

        lbl_text = label + (" (optional)" if optional else "")
        tk.Label(self, text=lbl_text, bg=PANEL_BG, fg=TEXT,
                 font=("Segoe UI", 9), width=15,
                 anchor="w").pack(side="left")

        self._display = tk.StringVar(value="Not selected")
        tk.Label(self, textvariable=self._display,
                 bg=PANEL_BG, fg=SUBTEXT,
                 font=("Segoe UI", 9), width=14,
                 anchor="w").pack(side="left", padx=6)

        make_button(self, "Pick date", self._open_cal,
                    width=9).pack(side="left", padx=4)

        if optional:
            make_button(self, "Clear", self._clear,
                        width=9, bg="#6B7280").pack(side="left", padx=2)

    def _open_cal(self):
        CalendarPopup(self.winfo_toplevel(),
                      title="Select date",
                      initial_date=self._date or date.today(),
                      callback=self._set_date)

    def _set_date(self, d):
        self._date = d
        self._display.set(d.strftime("%d-%b-%Y"))

    def _clear(self):
        self._date = None
        self._display.set("Not selected")

    def get(self):
        return self._date

    def is_valid(self):
        return self._optional or self._date is not None


# =============================================================================
# DIALOGS
# =============================================================================

class DefaultsDialog(tk.Toplevel):
    def __init__(self, parent, mode="edit"):
        super().__init__(parent)
        self.title("Default Variables")
        self.configure(bg=BG)
        self.geometry("760x700")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        self.mode = mode
        self.entries = {}
        self.defaults = load_json(LETTER_DEFAULTS)

        header = tk.Frame(self, bg=ACCENT, height=54)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="View Defaults" if mode == "show" else "Edit Default Variables",
                 bg=ACCENT, fg=BTN_FG, font=("Segoe UI", 13, "bold")).pack(
            side="left", padx=PAD * 2)

        note = tk.Label(
            self,
            text="Date-derived fields are mandatory and always use today's date during the pipeline. "
                 "Filename/XLSX-derived fields are shown for context and are not manually edited.",
            bg=BG, fg=SUBTEXT, font=("Segoe UI", 9),
            wraplength=700, justify="left")
        note.pack(anchor="w", padx=PAD * 2, pady=(PAD, 4))

        outer, inner = scroll_frame(self)
        outer.pack(fill="both", expand=True, padx=PAD * 2, pady=PAD)

        for group_name, items in DEFAULT_SCHEMA:
            SectionLabel(inner, group_name).pack(anchor="w", pady=(10, 4))
            HRule(inner).pack(fill="x", pady=(0, 6))
            for key, label, example in items:
                row = tk.Frame(inner, bg=PANEL_BG)
                row.pack(fill="x", pady=3)
                tk.Label(row, text=label, bg=PANEL_BG, fg=TEXT,
                         font=("Segoe UI", 9), width=28, anchor="w").pack(side="left")
                current = self._display_value(key, example)
                if mode == "show":
                    tk.Label(row, text=current, bg=PANEL_BG, fg=SUBTEXT,
                             font=("Segoe UI", 9), anchor="w").pack(
                        side="left", fill="x", expand=True)
                else:
                    readonly = key in AUTO_TODAY_KEYS or key in AUTO_DERIVED_KEYS
                    var = tk.StringVar(value=current)
                    make_entry(row, var, readonly=readonly).pack(
                        side="left", fill="x", expand=True, ipady=4)
                    self.entries[key] = var

        foot = tk.Frame(self, bg=BG)
        foot.pack(fill="x", padx=PAD * 2, pady=(0, PAD))
        if mode == "show":
            make_button(foot, "Close", self.destroy, width=12).pack(side="right")
        else:
            make_button(foot, "Save Defaults", self._save, width=16,
                        bg=SUCCESS).pack(side="right", padx=(6, 0))
            make_button(foot, "Cancel", self.destroy, width=12,
                        bg="#6B7280").pack(side="right")
        self._center(parent)

    def _display_value(self, key, example):
        if key in AUTO_TODAY_KEYS or key in AUTO_DERIVED_KEYS:
            return example
        return str(self.defaults.get(key) or example)

    def _save(self):
        data = load_json(LETTER_DEFAULTS)
        data["auto_today"] = True
        data["auto_rates"] = True
        for key, var in self.entries.items():
            if key in AUTO_TODAY_KEYS or key in AUTO_DERIVED_KEYS:
                data.pop(key, None)
                continue
            val = var.get().strip()
            if val:
                data[key] = val
        save_json(LETTER_DEFAULTS, data)
        messagebox.showinfo("Saved", "Defaults saved.", parent=self)
        self.destroy()

    def _center(self, parent):
        self.update_idletasks()
        x = parent.winfo_rootx() + parent.winfo_width() // 2 - self.winfo_width() // 2
        y = parent.winfo_rooty() + parent.winfo_height() // 2 - self.winfo_height() // 2
        self.geometry(f"+{x}+{y}")


class RecipientsDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Email Recipients")
        self.configure(bg=BG)
        self.geometry("800x560")
        self.resizable(True, True)
        self.transient(parent)
        self.grab_set()
        # Each entry: (frame, name_var, to_var, cc_var)
        self._rows = []
        self._count_var = tk.StringVar(value="0 penerima")

        header = tk.Frame(self, bg=ACCENT, height=54)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="Email Recipients", bg=ACCENT, fg=BTN_FG,
                 font=("Segoe UI", 13, "bold")).pack(side="left", padx=PAD * 2)
        tk.Label(header, textvariable=self._count_var, bg=ACCENT, fg="#93C5FD",
                 font=("Segoe UI", 10)).pack(side="right", padx=PAD * 2)

        tk.Label(self, text="Client name matches the XLSX filename. Separate addresses with semicolons.",
                 bg=BG, fg=SUBTEXT, font=("Segoe UI", 9)).pack(
            anchor="w", padx=PAD * 2, pady=(PAD, 4))

        # ── Search bar ────────────────────────────────────────────
        search_frame = tk.Frame(self, bg=BG)
        search_frame.pack(fill="x", padx=PAD * 2, pady=(0, 6))
        tk.Label(search_frame, text="Cari:", bg=BG, fg=TEXT,
                 font=("Segoe UI", 9)).pack(side="left", padx=(0, 6))
        self._search_var = tk.StringVar()
        self._search_var.trace_add("write", self._filter_rows)
        search_entry = make_entry(search_frame, self._search_var, width=30)
        search_entry.pack(side="left", ipady=4)
        make_button(search_frame, "Hapus Filter", self._clear_search,
                    width=12, bg="#6B7280", font_size=8).pack(side="left", padx=(6, 0))
        # ──────────────────────────────────────────────────────────

        labels = tk.Frame(self, bg=BG)
        labels.pack(fill="x", padx=PAD * 2)
        for text, width in [("Client Name", 24), ("TO", 34), ("CC", 28), ("", 4)]:
            tk.Label(labels, text=text, bg=BG, fg=ACCENT,
                     font=("Segoe UI", 9, "bold"), width=width,
                     anchor="w").pack(side="left", padx=2)

        panel = tk.Frame(self, bg=PANEL_BG, highlightbackground=BORDER,
                         highlightthickness=1)
        panel.pack(fill="both", expand=True, padx=PAD * 2, pady=PAD)
        outer, self._inner = scroll_frame(panel)
        outer.pack(fill="both", expand=True, padx=PAD, pady=PAD)

        cfg = load_json(EMAIL_CFG)
        for name, val in cfg.get("recipients", {}).items():
            self._add(name, "; ".join(val.get("to", [])), "; ".join(val.get("cc", [])))

        foot = tk.Frame(self, bg=BG)
        foot.pack(fill="x", padx=PAD * 2, pady=(0, PAD))
        make_button(foot, "+ Add Client", self._add, width=13).pack(side="left")
        make_button(foot, "Import XLSX", self._import_xlsx,
                    width=13, bg=SUCCESS).pack(side="left", padx=6)
        make_button(foot, "Delete All", self._delete_all,
                    width=12, bg=DANGER).pack(side="left")
        make_button(foot, "Save", self._save,
                    width=12, bg=SUCCESS).pack(side="right", padx=(6, 0))
        make_button(foot, "Cancel", self.destroy,
                    width=12, bg="#6B7280").pack(side="right")
        self._center(parent)

    # -- count -----------------------------------------------------------------

    def _update_count(self):
        n = len(self._rows)
        self._count_var.set(f"{n} penerima")

    # -- search ----------------------------------------------------------------

    def _filter_rows(self, *_):
        query = self._search_var.get().strip().upper()
        for frame, nv, _tv, _cv in self._rows:
            if not query or query in nv.get().upper():
                frame.pack(fill="x", pady=3)
            else:
                frame.pack_forget()

    def _clear_search(self):
        self._search_var.set("")

    # -- row management --------------------------------------------------------

    def _add(self, name="", to="", cc=""):
        frame = tk.Frame(self._inner, bg=PANEL_BG)
        frame.pack(fill="x", pady=3)
        nv, tv, cv = tk.StringVar(value=name), tk.StringVar(value=to), tk.StringVar(value=cc)
        make_entry(frame, nv, width=24).pack(side="left", padx=2, ipady=4)
        make_entry(frame, tv, width=34).pack(side="left", padx=2, ipady=4)
        make_entry(frame, cv, width=28).pack(side="left", padx=2, ipady=4)

        def delete():
            frame.destroy()
            self._rows[:] = [r for r in self._rows if r[0] is not frame]
            self._update_count()

        tk.Button(frame, text="X", command=delete, bg=PANEL_BG, fg=DANGER,
                  relief="flat", cursor="hand2").pack(side="left", padx=2)
        self._rows.append((frame, nv, tv, cv))
        self._update_count()

    def _delete_all(self):
        if not self._rows:
            return
        if not messagebox.askyesno("Confirm delete",
                                   f"Delete all {len(self._rows)} recipient entries?",
                                   parent=self):
            return
        for child in list(self._inner.children.values()):
            child.destroy()
        self._rows.clear()
        self._update_count()

    def _import_xlsx(self):
        if not EMAILER.exists():
            messagebox.showerror("Missing", f"smartpayer_emailer.py not found:\n{EMAILER}",
                                 parent=self)
            return
        path = filedialog.askopenfilename(
            title="Choose recipient list XLSX",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
            parent=self)
        if not path:
            return
        mode = "merge" if messagebox.askyesno(
            "Import mode",
            "Choose Yes to merge with existing recipients.\nChoose No to replace the list.",
            parent=self,
            default="yes") else "replace"
        try:
            result = subprocess.run(
                [sys.executable, str(EMAILER), "--import-recipients", path,
                 "--import-mode", mode],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
        except Exception as exc:
            messagebox.showerror("Import failed", str(exc), parent=self)
            return
        if result.returncode != 0:
            messagebox.showerror("Import failed",
                                 (result.stderr or result.stdout or "Unknown error")[:600],
                                 parent=self)
            return
        for child in list(self._inner.children.values()):
            child.destroy()
        self._rows.clear()
        cfg = load_json(EMAIL_CFG)
        for name, val in cfg.get("recipients", {}).items():
            self._add(name, "; ".join(val.get("to", [])), "; ".join(val.get("cc", [])))
        self._update_count()
        messagebox.showinfo("Import complete", "Recipients imported.", parent=self)

    def _save(self):
        cfg = load_json(EMAIL_CFG)
        cfg["recipients"] = {}
        for _frame, nv, tv, cv in self._rows:
            name = nv.get().strip().upper()
            if not name:
                continue
            cfg["recipients"][name] = {
                "to": [a.strip() for a in tv.get().split(";") if a.strip()],
                "cc": [a.strip() for a in cv.get().split(";") if a.strip()],
            }
        save_json(EMAIL_CFG, cfg)
        messagebox.showinfo("Saved", f"{len(cfg['recipients'])} recipient(s) saved.",
                            parent=self)
        self.destroy()

    def _center(self, parent):
        self.update_idletasks()
        x = parent.winfo_rootx() + parent.winfo_width() // 2 - self.winfo_width() // 2
        y = parent.winfo_rooty() + parent.winfo_height() // 2 - self.winfo_height() // 2
        self.geometry(f"+{x}+{y}")


# =============================================================================
# MAIN APPLICATION WINDOW
# =============================================================================

class SmartpayerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Smartpayer Automation")
        self.resizable(True, True)
        self.minsize(980, 600)
        self.configure(bg=BG)
        self._running = False
        self._run_btn = None
        self._retry_btn = None
        self._failed_email_pdfs = []
        self._run_started_at = None
        self._retry_started_at = None
        self._auto_send = tk.BooleanVar(value=True)
        self._input_file_var = tk.StringVar(value="")
        self._build_ui()
        self._center()

    # -- UI BUILD ----------------------------------------------------------------

    def _build_ui(self):
        header = tk.Frame(self, bg=ACCENT, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="Smartpayer Automation",
                 bg=ACCENT, fg=BTN_FG,
                 font=("Segoe UI", 14, "bold")).pack(side="left", padx=PAD * 2)

        content = tk.Frame(self, bg=BG)
        content.pack(fill="both", expand=True, padx=PAD * 2, pady=PAD * 2)

        left_shell = tk.Frame(content, bg=BG, width=465)
        left_shell.pack(side="left", fill="y", padx=(0, PAD))
        left_shell.pack_propagate(False)

        left_canvas = tk.Canvas(left_shell, bg=BG, highlightthickness=0,
                                width=445)
        left_scroll = ttk.Scrollbar(left_shell, orient="vertical",
                                    command=left_canvas.yview)
        left_canvas.configure(yscrollcommand=left_scroll.set)
        left_canvas.pack(side="left", fill="both", expand=True)
        left_scroll.pack(side="right", fill="y")

        left = tk.Frame(left_canvas, bg=BG)
        left_window = left_canvas.create_window((0, 0), window=left, anchor="nw")

        def _sync_scroll_region(_event=None):
            left_canvas.configure(scrollregion=left_canvas.bbox("all"))

        def _sync_inner_width(event):
            left_canvas.itemconfigure(left_window, width=event.width)

        def _on_left_mousewheel(event):
            left_canvas.yview_scroll(-1 * (event.delta // 120), "units")

        left.bind("<Configure>", _sync_scroll_region)
        left_canvas.bind("<Configure>", _sync_inner_width)
        left_canvas.bind("<Enter>", lambda _e: left_canvas.bind_all(
            "<MouseWheel>", _on_left_mousewheel))
        left_canvas.bind("<Leave>", lambda _e: left_canvas.unbind_all("<MouseWheel>"))

        self._build_rate_panel(left)
        tk.Frame(left, bg=BG, height=PAD).pack()
        self._build_input_panel(left)
        tk.Frame(left, bg=BG, height=PAD).pack()
        self._build_dates_panel(left)
        tk.Frame(left, bg=BG, height=PAD).pack()
        self._build_defaults_panel(left)
        tk.Frame(left, bg=BG, height=PAD).pack()
        self._build_email_panel(left)
        tk.Frame(left, bg=BG, height=PAD).pack()
        self._build_run_panel(left)

        right = tk.Frame(content, bg=BG)
        right.pack(side="left", fill="both", expand=True)
        self._build_log_panel(right)

    def _build_rate_panel(self, parent):
        panel = tk.Frame(parent, bg=PANEL_BG,
                         highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="x")

        SectionLabel(panel, "Daily Rate").pack(anchor="w", padx=PAD, pady=(PAD, 4))
        HRule(panel).pack(fill="x", padx=PAD)

        inner = tk.Frame(panel, bg=PANEL_BG)
        inner.pack(fill="x", padx=PAD, pady=PAD)

        tk.Label(inner, text="Current rate is 0.022% per day.",
                 bg=PANEL_BG, fg=SUBTEXT, font=("Segoe UI", 9)).pack(anchor="w")
        tk.Label(inner, text="Keep the current rate?",
                 bg=PANEL_BG, fg=TEXT, font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 4))

        btn_row = tk.Frame(inner, bg=PANEL_BG)
        btn_row.pack(anchor="w")
        self._keep_rate = tk.BooleanVar(value=True)
        self._btn_yes = make_button(btn_row, "Yes - keep 0.022%",
                                    lambda: self._set_keep_rate(True),
                                    width=18, bg=SUCCESS)
        self._btn_yes.pack(side="left", padx=(0, 8))
        self._btn_no = make_button(btn_row, "No - change rate",
                                   lambda: self._set_keep_rate(False),
                                   width=17, bg=WARN)
        self._btn_no.pack(side="left")

        self._rate_frame = tk.Frame(inner, bg=PANEL_BG)
        tk.Label(self._rate_frame,
                 text="Enter new rate (e.g. 0.022 means 0.022% per day):",
                 bg=PANEL_BG, fg=TEXT, font=("Segoe UI", 9)).pack(anchor="w", pady=(8, 2))
        entry_row = tk.Frame(self._rate_frame, bg=PANEL_BG)
        entry_row.pack(anchor="w")
        self._rate_var = tk.StringVar(value="0.022")
        make_entry(entry_row, self._rate_var, width=10).pack(side="left")
        tk.Label(entry_row, text="%", bg=PANEL_BG, fg=TEXT,
                 font=("Segoe UI", 10)).pack(side="left", padx=4)

        self._rate_status = tk.Label(inner, text="Using rate: 0.022% per day",
                                     bg=PANEL_BG, fg=SUCCESS,
                                     font=("Segoe UI", 9, "italic"))
        self._rate_status.pack(anchor="w", pady=(6, 0))

    def _build_input_panel(self, parent):
        panel = tk.Frame(parent, bg=PANEL_BG,
                         highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="x")

        SectionLabel(panel, "Input Workbook").pack(
            anchor="w", padx=PAD, pady=(PAD, 4))
        HRule(panel).pack(fill="x", padx=PAD)

        inner = tk.Frame(panel, bg=PANEL_BG)
        inner.pack(fill="x", padx=PAD, pady=PAD)

        tk.Label(inner,
                 text="Choose a master data XLSX, or leave blank to use the INPUT folder.",
                 bg=PANEL_BG, fg=SUBTEXT, font=("Segoe UI", 9),
                 wraplength=390, justify="left").pack(anchor="w", pady=(0, 6))

        row = tk.Frame(inner, bg=PANEL_BG)
        row.pack(fill="x")
        make_entry(row, self._input_file_var, readonly=True).pack(
            side="left", fill="x", expand=True, ipady=4)
        make_button(row, "Browse", self._browse_input_file,
                    width=9).pack(side="left", padx=(6, 0))
        make_button(row, "Clear", self._clear_input_file,
                    width=8, bg="#6B7280").pack(side="left", padx=(6, 0))

    def _build_dates_panel(self, parent):
        panel = tk.Frame(parent, bg=PANEL_BG,
                         highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="x")

        SectionLabel(panel, "Payment Dates").pack(
            anchor="w", padx=PAD, pady=(PAD, 4))
        HRule(panel).pack(fill="x", padx=PAD)

        inner = tk.Frame(panel, bg=PANEL_BG)
        inner.pack(fill="x", padx=PAD, pady=PAD)

        self._date_pickers = []
        for i in range(1, 5):
            row = DatePickerRow(inner, f"Date {i}", optional=(i == 4))
            row.pack(anchor="w", pady=3)
            self._date_pickers.append(row)

    def _build_defaults_panel(self, parent):
        panel = tk.Frame(parent, bg=PANEL_BG,
                         highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="x")
        SectionLabel(panel, "Default Variables").pack(anchor="w", padx=PAD, pady=(PAD, 4))
        HRule(panel).pack(fill="x", padx=PAD)
        inner = tk.Frame(panel, bg=PANEL_BG)
        inner.pack(fill="x", padx=PAD, pady=PAD)
        make_button(inner, "Show Defaults", lambda: DefaultsDialog(self, "show"),
                    width=16).pack(side="left", padx=(0, 8))
        make_button(inner, "Edit Defaults", lambda: DefaultsDialog(self, "edit"),
                    width=16, bg=SUCCESS).pack(side="left")

    def _build_email_panel(self, parent):
        panel = tk.Frame(parent, bg=PANEL_BG,
                         highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="x")
        SectionLabel(panel, "Email Automation").pack(anchor="w", padx=PAD, pady=(PAD, 4))
        HRule(panel).pack(fill="x", padx=PAD)
        inner = tk.Frame(panel, bg=PANEL_BG)
        inner.pack(fill="x", padx=PAD, pady=PAD)

        btn_row = tk.Frame(inner, bg=PANEL_BG)
        btn_row.pack(fill="x")
        make_button(btn_row, "Recipients List", lambda: RecipientsDialog(self),
                    width=16).pack(side="left", padx=(0, 8))
        make_button(btn_row, "Import Recipients", self._import_recipients,
                    width=17, bg=SUCCESS).pack(side="left")

        self._email_status_var = tk.StringVar(value="")
        tk.Label(inner, textvariable=self._email_status_var,
                 bg=PANEL_BG, fg=SUBTEXT, font=("Segoe UI", 8, "italic"),
                 wraplength=390, justify="left").pack(fill="x", anchor="w", pady=(8, 0))
        self._refresh_email_status()

    def _build_run_panel(self, parent):
        panel = tk.Frame(parent, bg=PANEL_BG,
                         highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="x")

        inner = tk.Frame(panel, bg=PANEL_BG)
        inner.pack(fill="x", padx=PAD, pady=PAD)

        tk.Checkbutton(inner, text="Auto-send email", variable=self._auto_send,
                       bg=PANEL_BG, fg=TEXT, selectcolor=PANEL_BG,
                       activebackground=PANEL_BG, font=("Segoe UI", 9),
                       cursor="hand2").pack(anchor="w", pady=(0, 8))

        self._run_btn = make_button(
            inner, "Run Automation", self._on_run,
            width=22, font_size=10)
        self._run_btn.pack(anchor="w")

        self._retry_btn = make_button(
            inner, "Retry Failed Emails", self._retry_failed_emails,
            width=22, bg=WARN, font_size=9)
        self._retry_btn.pack(anchor="w", pady=(6, 0))
        self._retry_btn.config(state="disabled")

        self._organize_btn = make_button(
            inner, "Organise PDF Output", self._organize_files,
            width=22, bg="#8B5CF6", font_size=9)
        self._organize_btn.pack(anchor="w", pady=(6, 0))
        self._organize_btn.config(state="normal")

        pb_frame = tk.Frame(inner, bg=PANEL_BG)
        pb_frame.pack(fill="x", pady=(10, 0))
        self._progress = ttk.Progressbar(pb_frame, mode="indeterminate", length=340)
        self._progress.pack(fill="x")

        self._status_var = tk.StringVar(value="Ready.")
        tk.Label(inner, textvariable=self._status_var,
                 bg=PANEL_BG, fg=SUBTEXT, font=("Segoe UI", 8, "italic"),
                 wraplength=340, justify="left").pack(anchor="w", pady=(4, 0))

    def _build_log_panel(self, parent):
        panel = tk.Frame(parent, bg=PANEL_BG,
                         highlightbackground=BORDER, highlightthickness=1)
        panel.pack(fill="both", expand=True)

        hdr = tk.Frame(panel, bg=PANEL_BG)
        hdr.pack(fill="x", padx=PAD, pady=(PAD, 4))
        SectionLabel(hdr, "Activity Log").pack(side="left")
        make_button(hdr, "Export Log", self._export_log,
                    width=12, bg="#6B7280", font_size=8).pack(side="right")

        HRule(panel).pack(fill="x", padx=PAD)

        log_frame = tk.Frame(panel, bg=PANEL_BG)
        log_frame.pack(fill="both", expand=True, padx=PAD, pady=PAD)
        self._log = tk.Text(
            log_frame, width=42, height=26, font=("Consolas", 8),
            bg="#0F172A", fg="#94A3B8", insertbackground="#94A3B8",
            relief="flat", state="disabled", wrap="word")
        scroll = ttk.Scrollbar(log_frame, command=self._log.yview)
        self._log.configure(yscrollcommand=scroll.set)
        self._log.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self._log.tag_configure("ok", foreground="#4ADE80")
        self._log.tag_configure("warn", foreground="#FCD34D")
        self._log.tag_configure("error", foreground="#F87171")
        self._log.tag_configure("info", foreground="#94A3B8")
        self._log.tag_configure("done", foreground="#38BDF8",
                                font=("Consolas", 8, "bold"))

    # -- HELPERS -----------------------------------------------------------------

    def _center(self):
        self.update_idletasks()
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        w = self.winfo_reqwidth()
        h = self.winfo_reqheight()
        self.geometry(f"+{(sw - w) // 2}+{(sh - h) // 2}")

    def _log_write(self, message, tag="info"):
        self._log.configure(state="normal")
        ts = datetime.now().strftime("%H:%M:%S")
        self._log.insert("end", f"[{ts}]  {message}\n", tag)
        self._log.see("end")
        self._log.configure(state="disabled")

    def _set_status(self, msg):
        self._status_var.set(msg)
        self.update_idletasks()

    def _set_running(self, running):
        self._running = running
        state = "disabled" if running else "normal"
        self._run_btn.config(state=state)

        if self._retry_btn:
            retry_state = "normal" if (not running and self._failed_email_pdfs) else "disabled"
            self._retry_btn.config(state=retry_state)

        if self._organize_btn:
            self._organize_btn.config(state="disabled" if running else "normal")

    def _format_elapsed(self, started_at):
        if not started_at:
            return "00:00:00"
        total_seconds = max(0, int((datetime.now() - started_at).total_seconds()))
        hours, rem = divmod(total_seconds, 3600)
        minutes, seconds = divmod(rem, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"

    def _set_keep_rate(self, keep):
        self._keep_rate.set(keep)
        if keep:
            self._rate_frame.pack_forget()
            self._rate_status.config(text="Using rate: 0.022% per day", fg=SUCCESS)
            self._btn_yes.config(relief="sunken")
            self._btn_no.config(relief="flat")
        else:
            self._rate_frame.pack(fill="x", pady=(4, 0))
            self._rate_status.config(text="Enter the new rate above.", fg=WARN)
            self._btn_yes.config(relief="flat")
            self._btn_no.config(relief="sunken")

    def _get_daily_rate(self):
        if self._keep_rate.get():
            return 0.00022
        raw = self._rate_var.get().strip().replace(",", ".")
        return float(raw) / 100.0

    def _browse_input_file(self):
        path = filedialog.askopenfilename(
            title="Choose Smartpayer input workbook",
            initialdir=str(BASE_DIR / "INPUT"),
            filetypes=[("Excel workbooks", "*.xlsx"), ("All files", "*.*")])
        if path:
            self._input_file_var.set(path)

    def _clear_input_file(self):
        self._input_file_var.set("")

    def _refresh_email_status(self):
        cfg = load_json(EMAIL_CFG)
        count = len(cfg.get("recipients", {}))
        smtp = cfg.get("smtp", {})
        method = "SMTP" if smtp.get("username") else "Outlook"
        self._email_status_var.set(f"{count} recipient(s) configured. Send via {method}.")

    def _export_log(self):
        """Export the activity log to a .txt file."""
        content = self._log.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showinfo("Log Kosong", "Tidak ada log untuk diekspor.")
            return
        today = datetime.now()
        filename = f"log-smartpayer-{today.strftime('%d-%m-%Y')}.txt"
        out_path = BASE_DIR / filename
        try:
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(content)
            messagebox.showinfo(
                "Log Diekspor",
                f"Log berhasil diekspor ke:\n{filename}")
        except Exception as e:
            messagebox.showerror("Gagal Ekspor", f"Gagal mengekspor log:\n{e}")

    def _run_subprocess(self, args, progress_cb):
        result = subprocess.run(args, capture_output=True, text=True,
                                encoding="utf-8", errors="replace")
        for line in (result.stdout or "").splitlines():
            progress_cb(line)
        for line in (result.stderr or "").splitlines():
            progress_cb(line)
        if result.returncode != 0:
            raise RuntimeError((result.stderr or result.stdout or "Subprocess failed").strip())

    def _send_email_categorized(self, pdf, progress_cb):
        """
        Send email for the given PDF. Returns (success, failure_reason).
        failure_reason: None | 'no_recipient' | 'send_failed'
        """
        if not EMAILER.exists():
            return False, 'send_failed'
        if not pdf.exists():
            return False, 'send_failed'
        progress_cb(f"Sending email for {pdf.name}...")
        result = subprocess.run(
            [sys.executable, str(EMAILER), "--pdf", str(pdf)],
            capture_output=True, text=True, encoding="utf-8", errors="replace")
        output = (result.stdout or "") + (result.stderr or "")
        for line in output.splitlines():
            progress_cb(line)
        if result.returncode == 0:
            return True, None
        if "no recipient found" in output.lower():
            return False, 'no_recipient'
        return False, 'send_failed'

    def _send_email_for_pdf(self, pdf, progress_cb):
        """Legacy wrapper kept for compatibility."""
        if not EMAILER.exists():
            raise FileNotFoundError("smartpayer_emailer.py not found; email skipped.")
        if not pdf.exists():
            raise FileNotFoundError(f"PDF not found: {pdf.name}")
        progress_cb(f"Sending email for {pdf.name}...")
        self._run_subprocess(
            [sys.executable, str(EMAILER), "--pdf", str(pdf)],
            progress_cb)

    def _convert_docx_to_pdf_word(self, docx_path, pdf_path, progress_cb):
        """Convert a DOCX file to PDF using Microsoft Word COM. Returns True on success."""
        if sys.platform != "win32":
            progress_cb(f"  [!] PDF conversion only supported on Windows.")
            return False
        try:
            import comtypes.client
        except ImportError:
            progress_cb("  [!] comtypes not installed. Run: pip install comtypes")
            return False
        word = None
        try:
            progress_cb(f"  Mengkonversi {docx_path.name} ke PDF...")
            word = comtypes.client.CreateObject("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(str(docx_path.resolve()))
            doc.ExportAsFixedFormat(
                OutputFileName=str(pdf_path.resolve()),
                ExportFormat=17,
                OpenAfterExport=False,
                OptimizeFor=0,
                Range=0,
                Item=0,
                IncludeDocProps=True,
                KeepIRM=True,
                CreateBookmarks=0,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False,
            )
            doc.Close(False)
            word.Quit()
            word = None
            if pdf_path.exists():
                progress_cb(f"  [OK] PDF berhasil dibuat: {pdf_path.name}")
                return True
            progress_cb(f"  [!] PDF tidak ditemukan setelah konversi: {pdf_path.name}")
            return False
        except Exception as e:
            progress_cb(f"  [!] Gagal konversi {docx_path.name}: {e}")
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass
            return False

    def _find_docx_without_pdf(self):
        """
        Scan Generated_Letters for *_with_tables.docx files that have no
        corresponding *_letter.pdf. Returns list of (docx_path, pdf_path).
        """
        if not LETTER_OUTPUT_DIR.exists():
            return []
        result = []
        for docx in sorted(LETTER_OUTPUT_DIR.glob("*_with_tables.docx")):
            pdf_stem = docx.stem.replace("_with_tables", "_letter")
            pdf_path = docx.parent / f"{pdf_stem}.pdf"
            if not pdf_path.exists():
                result.append((docx, pdf_path))
        return result

    def _payer_name_from_letter_pdf(self, pdf):
        stem = Path(pdf).stem.replace("_", " ")
        match = re.match(
            r"^[Ss]mart\s*[Pp]ayer\s+[A-Za-z]+\s+\d{4}\s+(.+?)\s+letter$",
            stem,
            flags=re.IGNORECASE)
        if not match:
            return None
        return match.group(1).strip().upper()

    def _add_blank_recipients_for_failed_emails(self, failed_no_recipient):
        cfg = load_json(EMAIL_CFG)
        recips = cfg.setdefault("recipients", {})
        added = []
        for pdf, _reason in failed_no_recipient:
            payer_name = self._payer_name_from_letter_pdf(pdf)
            if not payer_name:
                continue
            exists = any(k.upper().strip() == payer_name for k in recips)
            if exists:
                continue
            recips[payer_name] = {"to": [], "cc": []}
            added.append(payer_name)
        if added:
            save_json(EMAIL_CFG, cfg)
            self._refresh_email_status()
            self._log_write(
                f"Menambahkan {len(added)} nama payer ke daftar penerima dengan TO/CC kosong: "
                + "; ".join(added),
                "warn")
        return added

    def _import_recipients(self):
        if not EMAILER.exists():
            messagebox.showerror("Missing", f"smartpayer_emailer.py not found:\n{EMAILER}")
            return
        path = filedialog.askopenfilename(
            title="Choose recipient list XLSX",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")])
        if not path:
            return
        mode = "merge" if messagebox.askyesno(
            "Import mode",
            "Choose Yes to merge with existing recipients.\nChoose No to replace the list.",
            default="yes") else "replace"
        self._log_write(f"Importing recipients from {Path(path).name} ({mode})...", "info")
        started_at = datetime.now()
        try:
            self._run_subprocess(
                [sys.executable, str(EMAILER), "--import-recipients", path,
                 "--import-mode", mode],
                lambda msg: self._log_write(msg, "info"))
        except Exception as exc:
            messagebox.showerror("Import failed", str(exc))
            self._log_write(f"Recipient import failed: {exc}", "error")
            return
        self._refresh_email_status()
        elapsed = self._format_elapsed(started_at)
        self._log_write(f"Recipient import complete. Elapsed time: {elapsed}", "done")
        messagebox.showinfo(
            "Import complete",
            f"Recipients imported successfully.\nElapsed time: {elapsed}")

    def _organize_files(self):
        if not LETTER_OUTPUT_DIR.exists():
            messagebox.showwarning(
                "No Output Folder",
                f"The Generated_Letters folder does not exist:\n{LETTER_OUTPUT_DIR}")
            return

        pattern = re.compile(r'^(SmartPayer\s+[A-Za-z]+\s+\d{4})\s+', re.IGNORECASE)
        matching = [f for f in LETTER_OUTPUT_DIR.iterdir()
                    if f.is_file() and pattern.match(f.name)]

        if not matching:
            messagebox.showinfo(
                "Nothing to Organize",
                f"No files matching 'SmartPayer <Month> <Year>' pattern found in:\n{LETTER_OUTPUT_DIR}")
            return

        prefixes = set()
        for f in matching:
            match = pattern.match(f.name)
            if match:
                prefixes.add(match.group(1).strip())

        preview = "\n".join(
            f"  • {p}/ ({sum(1 for f in matching if pattern.match(f.name) and pattern.match(f.name).group(1).strip() == p)} files)"
            for p in sorted(prefixes))

        if not messagebox.askyesno(
            "Confirm Organization",
            f"Organize {len(matching)} file(s) into {len(prefixes)} folder(s):\n\n{preview}\n\nProceed?"):
            return

        self._organize_btn.config(state="disabled")
        self._log_write(f"Organizing {len(matching)} file(s) by prefix...", "info")

        def worker():
            try:
                moved, created = organize_files_by_prefix(
                    LETTER_OUTPUT_DIR,
                    progress_cb=lambda msg: self.after(0, self._log_write, msg, "info"))
                if moved > 0:
                    self.after(0, lambda: messagebox.showinfo(
                        "Organization Complete",
                        f"Moved {moved} file(s) into {created} folder(s).\n"
                        f"Location: {LETTER_OUTPUT_DIR}"))
                    self.after(0, lambda: self._log_write(
                        f"Files organized: {moved} moved, {created} folders created", "done"))
                else:
                    self.after(0, lambda: messagebox.showwarning(
                        "Nothing Moved", "No files were moved. Check log for details."))
            except Exception as e:
                self.after(0, lambda: self._log_write(
                    f"Organisation error: {e}", "error"))
            finally:
                self.after(0, lambda: self._organize_btn.config(state="normal"))

        threading.Thread(target=worker, daemon=True).start()

    # -- RUN ---------------------------------------------------------------------

    def _on_run(self):
        if self._running:
            return

        for i, picker in enumerate(self._date_pickers):
            if not picker.is_valid():
                messagebox.showerror(
                    "Missing Date",
                    f"Date {i + 1} is required. Please select it.")
                return

        if not self._keep_rate.get():
            try:
                raw_rate = float(self._rate_var.get().replace(",", "."))
                if raw_rate <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("Invalid Rate",
                                     "Please enter a positive number, e.g. 0.022")
                return

        try:
            daily_rate = self._get_daily_rate()
        except Exception:
            messagebox.showerror("Invalid Rate", "Could not parse the daily rate.")
            return

        dates = [picker.get() for picker in self._date_pickers]
        input_path = self._input_file_var.get().strip()
        if input_path:
            selected = Path(input_path)
            if selected.suffix.lower() != ".xlsx" or selected.name.startswith("~$"):
                messagebox.showerror(
                    "Invalid Input",
                    "Please choose a valid .xlsx workbook.")
                return
            if not selected.exists():
                messagebox.showerror(
                    "Missing Input",
                    f"The selected input workbook does not exist:\n\n{input_path}")
                return

        params = {
            "daily_rate": daily_rate,
            "dates": dates,
            "input_path": input_path or None,
        }
        auto_send = self._auto_send.get()
        self._run_started_at = datetime.now()

        self._set_running(True)
        self._progress.start(10)

        self._log_write("=" * 54, "info")
        self._log_write("Memulai pipeline Smartpayer Automation...", "ok")
        self._log_write(
            f"Input file : {Path(input_path).name if input_path else 'INPUT folder'}",
            "info")
        self._log_write(f"Daily rate : {daily_rate * 100:.5f}%", "info")
        for i, d in enumerate(dates):
            self._log_write(
                f"Date {i + 1}     : "
                f"{d.strftime('%d-%b-%Y') if d else 'skipped'}", "info")
        self._log_write(f"Auto-send  : {'ON' if auto_send else 'OFF'}", "info")
        self._log_write("=" * 54, "info")

        thread = threading.Thread(
            target=self._run_thread, args=(params, auto_send), daemon=True)
        thread.start()

    def _run_thread(self, params, auto_send):
        try:
            progress = lambda msg: self.after(0, self._on_progress, msg)
            created = sa.run_automation(params, progress)
            result = self._run_letter_generation(auto_send, progress)
            self.after(0, self._on_success, created, result)
        except Exception as exc:
            self.after(0, self._on_error, str(exc))

    def _run_letter_generation(self, auto_send, progress_cb):
        if not GENERATOR.exists():
            raise FileNotFoundError(f"Letter generator not found: {GENERATOR}")

        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        xlsx_files = sorted(
            p for p in OUTPUT_DIR.glob("*.xlsx")
            if not p.name.startswith("~$")
        )
        if not xlsx_files:
            raise FileNotFoundError("No split .xlsx files found in the OUTPUT folder.")

        LETTER_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        token = uuid.uuid4().hex
        progress_cb(f"Starting letter generation for {len(xlsx_files)} OUTPUT file(s)...")
        self._run_subprocess(
            [sys.executable, str(GENERATOR), "--reset-batch", "--batch-token", token],
            progress_cb)

        generated = []
        failed_letters = []
        failed_pdf_conversion = []   # PDF path where DOCX exists but PDF wasn't created
        failed_no_recipient = []     # (pdf_path, reason) - no email recipient found
        failed_email_send = []       # (pdf_path, reason) - other send failures
        email_sent = 0

        for idx, xlsx in enumerate(xlsx_files, start=1):
            progress_cb(f"Generating letter {idx}/{len(xlsx_files)}: {xlsx.name}")
            try:
                self._run_subprocess(
                    [sys.executable, str(GENERATOR),
                     "--xlsx", str(xlsx),
                     "--output-dir", str(LETTER_OUTPUT_DIR),
                     "--use-defaults",
                     "--auto-today",
                     "--batch-token", token],
                    progress_cb)
            except Exception as exc:
                failed_letters.append((xlsx, str(exc)))
                progress_cb(f"Error generating letter for {xlsx.name}: {exc}")
                continue

            pdf = LETTER_OUTPUT_DIR / f"{xlsx.stem}_letter.pdf"
            docx = LETTER_OUTPUT_DIR / f"{xlsx.stem}_with_tables.docx"

            if not pdf.exists() and docx.exists():
                # Generator ran but PDF conversion failed — DOCX is available
                failed_pdf_conversion.append(pdf)
                generated.append(docx)
                progress_cb(f"  [!] Konversi PDF gagal untuk {xlsx.name}. DOCX tersedia.")
                continue

            output_file = pdf if pdf.exists() else docx
            generated.append(output_file)

            if auto_send and pdf.exists():
                success, reason = self._send_email_categorized(pdf, progress_cb)
                if success:
                    email_sent += 1
                elif reason == 'no_recipient':
                    failed_no_recipient.append((pdf, reason))
                    progress_cb(f"  [!] Tidak ada penerima email untuk {pdf.name}")
                else:
                    failed_email_send.append((pdf, str(reason)))
                    progress_cb(f"  [!] Gagal mengirim email untuk {pdf.name}")

        progress_cb("Letter generation complete.")
        return {
            "generated": generated,
            "email_sent": email_sent,
            "auto_send": auto_send,
            "failed_letters": failed_letters,
            "failed_pdf_conversion": failed_pdf_conversion,
            "failed_no_recipient": failed_no_recipient,
            "failed_email_send": failed_email_send,
        }

    def _on_progress(self, msg):
        lower = msg.lower()
        if "error" in lower or "failed" in lower or "[!]" in msg:
            tag = "error"
        elif "warn" in lower or "note:" in lower or "skipping" in lower:
            tag = "warn"
        elif msg == "DONE" or "complete" in lower or "done" in lower:
            tag = "done"
        else:
            tag = "ok"
        self._log_write(msg, tag)
        self._set_status(msg)

    def _retry_failed_emails(self):
        if self._running:
            return

        # Open recipients dialog for editing, then after user closes it
        # also scan for DOCX-only files to retry conversion
        dlg = RecipientsDialog(self)
        self.wait_window(dlg)
        self._refresh_email_status()

        # Collect email retries (PDFs with known failures)
        pdfs_to_retry = [pdf for pdf in self._failed_email_pdfs if pdf.exists()]

        # Scan for DOCX files without a corresponding PDF (conversion failures)
        docx_conversions = self._find_docx_without_pdf()

        if not pdfs_to_retry and not docx_conversions:
            messagebox.showinfo(
                "Tidak Ada yang Diulang",
                "Tidak ada email gagal atau konversi PDF yang tertunda.")
            self._failed_email_pdfs = []
            self._set_running(False)
            return

        summary_lines = []
        if pdfs_to_retry:
            summary_lines.append(f"• {len(pdfs_to_retry)} email gagal akan dicoba ulang")
        if docx_conversions:
            summary_lines.append(
                f"• {len(docx_conversions)} file DOCX akan dikonversi ke PDF:")
            for docx, pdf in docx_conversions[:5]:
                summary_lines.append(f"    - {docx.name}")
            if len(docx_conversions) > 5:
                summary_lines.append(f"    ... dan {len(docx_conversions) - 5} lainnya")

        if not messagebox.askyesno(
            "Konfirmasi Retry",
            "\n".join(summary_lines) + "\n\nLanjutkan?",
            default="yes"):
            return

        self._set_running(True)
        self._progress.start(10)
        self._retry_started_at = datetime.now()
        self._log_write("=" * 54, "info")
        self._log_write(
            f"Retry: {len(pdfs_to_retry)} email + {len(docx_conversions)} konversi PDF...",
            "warn")

        thread = threading.Thread(
            target=self._retry_failed_emails_thread,
            args=(pdfs_to_retry, docx_conversions),
            daemon=True)
        thread.start()

    def _retry_failed_emails_thread(self, pdfs_to_retry, docx_conversions):
        progress = lambda msg: self.after(0, self._on_progress, msg)

        # Step 1: Convert DOCX → PDF
        conv_success = []
        conv_failed = []
        for docx, pdf in docx_conversions:
            ok = self._convert_docx_to_pdf_word(docx, pdf, progress)
            if ok:
                conv_success.append(pdf)
            else:
                conv_failed.append(pdf)

        # Step 2: Retry emails — original failed list + newly converted PDFs
        all_pdfs = list(pdfs_to_retry) + conv_success
        email_ok = []
        still_failed = []
        for pdf in all_pdfs:
            if not pdf.exists():
                still_failed.append(pdf)
                continue
            success, reason = self._send_email_categorized(pdf, progress)
            if success:
                email_ok.append(pdf)
            else:
                still_failed.append(pdf)
                progress(f"  [!] Email masih gagal untuk {pdf.name}")

        self.after(0, self._on_retry_complete,
                   email_ok, still_failed, conv_success, conv_failed)

    def _on_retry_complete(self, email_ok, still_failed, conv_success, conv_failed):
        self._progress.stop()
        elapsed = self._format_elapsed(self._retry_started_at)
        self._retry_started_at = None
        self._failed_email_pdfs = [p for p in still_failed if p and p.exists()]
        self._set_running(False)

        # Write summary to activity log
        self._log_write("=" * 54, "done")
        self._log_write(f"Total Surat Berhasil Terkirim (Retry): {len(email_ok)}", "done")
        self._log_write(f"Total Surat Masih Gagal Terkirim: {len(still_failed)}",
                        "warn" if still_failed else "done")
        if still_failed:
            for i, pdf in enumerate(still_failed, 1):
                self._log_write(f"  {i}. {pdf.name}", "warn")
        self._log_write(f"Konversi PDF Berhasil: {len(conv_success)}", "done")
        self._log_write(f"Konversi PDF Gagal: {len(conv_failed)}",
                        "error" if conv_failed else "done")
        if conv_failed:
            for i, pdf in enumerate(conv_failed, 1):
                self._log_write(f"  {i}. {pdf.name}", "error")
        self._log_write(f"Waktu Proses Retry: {elapsed}", "done")
        self._log_write("=" * 54, "done")

        self._set_status(f"Retry selesai. Waktu: {elapsed}")

        # Build popup summary
        summary = (
            f"Total Surat Berhasil Terkirim (Retry): {len(email_ok)}\n"
            f"Total Surat Masih Gagal Terkirim: {len(still_failed)}\n"
            f"Konversi PDF Berhasil: {len(conv_success)}\n"
            f"Konversi PDF Gagal: {len(conv_failed)}\n"
            f"Waktu Proses Retry: {elapsed}"
        )
        if still_failed:
            messagebox.showwarning("Retry Selesai", summary)
        else:
            messagebox.showinfo("Retry Selesai", summary)

    def _on_success(self, created, result):
        self._progress.stop()
        elapsed = self._format_elapsed(self._run_started_at)
        self._run_started_at = None
        self._refresh_email_status()

        generated = result.get("generated", [])
        email_sent = result.get("email_sent", 0)
        auto_send = result.get("auto_send", False)
        failed_letters = result.get("failed_letters", [])
        failed_pdf_conversion = result.get("failed_pdf_conversion", [])
        failed_no_recipient = result.get("failed_no_recipient", [])
        failed_email_send = result.get("failed_email_send", [])

        total_email_failed = (len(failed_no_recipient) + len(failed_email_send)
                              + len(failed_pdf_conversion))

        # Collect failed PDFs for retry button
        self._failed_email_pdfs = []
        for pdf, _r in failed_no_recipient:
            if pdf and pdf.exists():
                self._failed_email_pdfs.append(pdf)
        for pdf, _r in failed_email_send:
            if pdf and pdf.exists():
                self._failed_email_pdfs.append(pdf)

        self._set_running(False)

        # Add blank recipients for no-recipient failures
        added_recipients = self._add_blank_recipients_for_failed_emails(failed_no_recipient)

        # ── Activity log summary ──────────────────────────────────
        self._log_write("=" * 54, "done")
        self._log_write(f"{len(generated)} Item telah Berhasil Diproses", "done")
        if auto_send:
            self._log_write(
                f"Total Surat Berhasil Terkirim: {email_sent}",
                "done")
            self._log_write(
                f"Total Surat Gagal Terkirim: {total_email_failed}",
                "warn" if total_email_failed else "done")

            if failed_no_recipient:
                self._log_write(
                    f"Total Surat Gagal Terkirim karena Email Penerima Tidak Ditemukan: "
                    f"{len(failed_no_recipient)}",
                    "warn")
                for i, (pdf, _) in enumerate(failed_no_recipient, 1):
                    self._log_write(f"  {i}. {pdf.name}", "warn")

            if failed_email_send:
                self._log_write(
                    f"Total Surat Gagal Terkirim (Error lainnya): {len(failed_email_send)}",
                    "error")
                for i, (pdf, _) in enumerate(failed_email_send, 1):
                    self._log_write(f"  {i}. {pdf.name}", "error")

            if failed_pdf_conversion:
                self._log_write(
                    f"Total Surat Gagal Terkirim karena Konversi DOCX menjadi PDF gagal: "
                    f"{len(failed_pdf_conversion)}",
                    "error")
                for i, pdf in enumerate(failed_pdf_conversion, 1):
                    self._log_write(f"  {i}. {pdf.name}", "error")

        if failed_letters:
            self._log_write(
                f"Letter generation gagal: {len(failed_letters)}", "error")

        self._log_write(f"Waktu Proses: {elapsed}", "done")
        self._log_write("=" * 54, "done")
        self._set_status(f"Pipeline selesai. Waktu: {elapsed}")

        # ── End-report popup ──────────────────────────────────────
        summary = f"{len(generated)} Item telah Berhasil Diproses\nWaktu Proses: {elapsed}"
        if auto_send:
            if total_email_failed > 0:
                summary += f"\nTotal {total_email_failed} Email Gagal Terkirim"
            if failed_pdf_conversion:
                summary += f"\n{len(failed_pdf_conversion)} Item gagal di-Convert menjadi PDF"
            if failed_no_recipient:
                summary += (
                    f"\n{len(failed_no_recipient)} Item Gagal Terkirim karena "
                    "Email Penerima Tidak Ditemukan")

        has_retryable = auto_send and (
            self._failed_email_pdfs or self._find_docx_without_pdf())

        if has_retryable:
            retry_now = messagebox.askyesno(
                "Hasil Pipeline",
                summary
                + "\n\nBeberapa email gagal dikirim. Lakukan Retry sekarang?",
                default="yes")
            if retry_now:
                self._retry_failed_emails()
            else:
                messagebox.showinfo(
                    "Selesai",
                    summary + "\n\nGunakan tombol 'Retry Failed Emails' untuk mencoba lagi.")
        else:
            messagebox.showinfo("Selesai", summary)

    def _on_error(self, error_msg):
        self._progress.stop()
        elapsed = self._format_elapsed(self._run_started_at)
        self._run_started_at = None
        self._set_running(False)
        self._log_write("=" * 54, "error")
        self._log_write(f"ERROR: {error_msg}", "error")
        self._log_write(f"Elapsed time before error: {elapsed}", "error")
        self._log_write("=" * 54, "error")
        self._set_status(f"Error occurred. Elapsed: {elapsed}")
        messagebox.showerror(
            "Error",
            f"An error occurred:\n\n{error_msg}\n\nElapsed time: {elapsed}")


# =============================================================================
# ENTRY POINT
# =============================================================================

if __name__ == "__main__":
    app = SmartpayerApp()
    app.mainloop()
