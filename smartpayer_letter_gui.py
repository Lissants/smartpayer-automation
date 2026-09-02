#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SmartPayer Letter Generator — Desktop GUI

Summary of changes vs. the previous version
-------------------------------------------
1. Added an "Always use today's date" checkbox at the top of the
   "Process Spreadsheet" card. Default ON. When checked, the five
   date-derived fields (letter_date, month_name, month_year, month_roman,
   year_num) are computed from datetime.now() and shown read-only.
   `letter_number` (just the mmdd prefix) and `end_of_month_date` are
   also auto-computed and displayed read-only. Toggling the box
   instantly recomputes and refreshes the previewed values, and the
   state is persisted in smartpayer_defaults.json so it survives GUI
   restarts.
2. Renamed `letter_date_num` to `year_num` (last two digits of the
   year). The Defaults dialog now shows it under the Date & Period
   group with label "Last two digits of the year (yy)".
3. Added a per-batch counter for the `nnn` suffix of `letter_number`.
   The GUI mints a fresh uuid4 batch token for every manual "Generate
   Letter" click, and a fresh token for every watcher scan-cycle that
   picks up new files, then passes it to the generator subprocess via
   --batch-token. A pre-launch `--reset-batch` call zeroes the counter
   so the first file in the batch gets …001.
4. `discount_from` and `discount_until` are no longer editable in the
   GUI — they are auto-derived from each XLSX's header date columns
   inside the generator. The Defaults dialog still shows them but as
   greyed-out informational entries with the placeholder text
   "(auto-derived from XLSX)".
5. The "Always use today's date" preference is stored in
   smartpayer_defaults.json under the key `auto_today` and is passed
   to the subprocess via --auto-today / --no-auto-today so the
   generator's behaviour stays in sync with the GUI's toggle even
   when defaults haven't been re-saved.

Existing behaviour preserved
----------------------------
- Watch mode and the email automation (recipients + SMTP dialogs) are
  unchanged. The watcher still spawns one subprocess per detected file
  and still passes the email auto-send flag through.
"""
import os,sys,json,threading,subprocess,time,uuid
from pathlib import Path
import tkinter as tk
from tkinter import ttk,filedialog,messagebox,scrolledtext
from datetime import datetime
import calendar

if sys.stdout and hasattr(sys.stdout,"reconfigure"):
    try: sys.stdout.reconfigure(encoding="utf-8",errors="replace")
    except: pass

SCRIPT_DIR = Path(__file__).parent.resolve()

GENERATOR  = SCRIPT_DIR / "smartpayer_letter_generator.py"
EMAILER    = SCRIPT_DIR / "smartpayer_emailer.py"
DEFAULTS_F = SCRIPT_DIR / "smartpayer_letter_defaults.json"
EMAIL_CFG  = SCRIPT_DIR / "smartpayer_email_config.json"
TEMPLATE   = SCRIPT_DIR / "Smart_Payer_Program_Letter_Template.docx"
OUTPUT_DIR = SCRIPT_DIR.parent / "Generated_Letters"

C={"bg":"#0F1923","panel":"#16222E","border":"#1E3448","accent":"#0E9EE8",
   "accent2":"#00D4AA","adim":"#0B7AB8","text":"#E8F1F8","tdim":"#7A9BB8",
   "ok":"#22C55E","warn":"#F59E0B","err":"#EF4444","white":"#FFFFFF",
   "ebg":"#1A2B3C","ebdr":"#2A4A66","hover":"#1C3048","purple":"#7C3AED",
   "plite":"#A78BFA","ebg_dim":"#152230"}  # NEW: ebg_dim for greyed read-only fields

FT=("Segoe UI",22,"bold"); FH=("Segoe UI",11,"bold")
FB=("Segoe UI",9);         FS=("Segoe UI",8)
FM=("Consolas",9)

# Variable schema. `letter_date_num` has been REMOVED and replaced by
# `year_num` (the last two digits of the year, e.g. "26" for 2026).
# `discount_from` / `discount_until` are kept in the schema only so the
# Defaults dialog can show them as "(auto-derived from XLSX)" — they are
# not user-editable anymore.
SCHEMA=[
 ("company_name","Company Full Name","Mencari Cinta Sejati"),
 ("company_abbr","Company Abbreviation","MCS"),
 ("letter_date","Letter Date (dd Month yyyy)","26 September 2026"),
 ("month_name","Month Name (auto from filename)","April"),
 ("month_year","Month & Year (auto from filename)","April 2026"),
 ("month_roman","Month (Roman Numerals)","IX"),
 ("letter_number","Letter Number (mmddnnn)","0926001"),
 ("year_num","Last two digits of the year (yy)","26"),
 ("deposit_rate","Deposit Interest Rate (%/year)","2"),
 ("deposit_days","Deposit Duration (days)","3"),
 ("daily_rate","Daily Reward Rate (%/day)","0.022"),
 ("annual_rate","Annualized Reward Rate (%/year)","8"),
 ("rate_multiplier_num","Rate Multiplier (number)","4"),
 ("rate_multiplier_word","Rate Multiplier (words/Indonesian)","empat"),
 ("end_of_month_date","End of Next Month (full date, auto from filename)","31 May 2026"),
 ("discount_from","Discount Period From (dd)","(auto-derived from XLSX)"),
 ("discount_until","Discount Period Until (dd Month yyyy)","(auto-derived from XLSX)"),
 ("team1_honorific_name","Finance Team 1 (Honorific + Name)","Bapak John"),
 ("team2_honorific_name","Finance Team 2 (Honorific + Name)","Bapak Arthur"),
 ("team_email","Finance Team Email","dagwon@amq.com"),
 ("team3_honorific_name","Finance Team 3 (Honorific + Name)","Bapak Edward"),
 ("team3_email","Finance Team 3 Email","edward@amq.com"),
 ("team4_honorific_name","Finance Team 4 (Honorific + Name)","Ibu Amber"),
 ("team4_email","Finance Team 4 Email","amber@amq.com"),
]

# Keys whose values come from datetime.now() when "Always use today's date"
# is on. These get rendered read-only / greyed in the Defaults dialog when
# the checkbox at the top of the dialog is ticked.
AUTO_TODAY_KEYS = {
    "letter_date", "month_roman", "year_num", "letter_number",
}
# Keys that come from the XLSX FILENAME per-file (the encoded billing
# month, e.g. "Smartpayer April 2026 X.xlsx" → April 2026 + next-month
# rollover for end_of_month_date). Always read-only in GUI.
FILENAME_DERIVED_KEYS = {"month_name", "month_year", "end_of_month_date"}
# Keys that come from the XLSX header per-file. Always read-only in GUI.
XLSX_DERIVED_KEYS = {"discount_from", "discount_until"}

_MONTH_ROMAN = ("I","II","III","IV","V","VI","VII","VIII","IX","X","XI","XII")

def compute_today_values():
    """
    Return a dict of date-derived fields computed from datetime.now().
    `letter_number` is just the mmdd prefix + '001' for display purposes;
    the real per-batch counter lives in the generator subprocess.
    """
    today = datetime.now()
    if today.month == 12:
        next_year, next_month = today.year + 1, 1
    else:
        next_year, next_month = today.year, today.month + 1
    _, eom_next = calendar.monthrange(next_year, next_month)
    return {
        "letter_date":       today.strftime("%d %B %Y"),
        "month_name":        today.strftime("%B"),
        "month_year":        today.strftime("%B %Y"),
        "month_roman":       _MONTH_ROMAN[today.month - 1],
        "year_num":          today.strftime("%y"),
        # Show "0512001" as the preview — the real counter is per-batch.
        "letter_number":     today.strftime("%m%d") + "001",
        "end_of_month_date": str(eom_next),
    }

def load_def():
    if DEFAULTS_F.exists():
        with open(DEFAULTS_F,encoding="utf-8") as f: return json.load(f)
    return {}
def save_def(d):
    with open(DEFAULTS_F,"w",encoding="utf-8") as f: json.dump(d,f,indent=2,ensure_ascii=False)
def load_ecfg():
    if EMAIL_CFG.exists():
        try:
            with open(EMAIL_CFG,encoding="utf-8") as f: return json.load(f)
        except: pass
    return {}
def save_ecfg(d):
    with open(EMAIL_CFG,"w",encoding="utf-8") as f: json.dump(d,f,indent=2,ensure_ascii=False)

def mkbtn(p,text,cmd,style="primary",**kw):
    pal={"primary":(C["accent"],C["adim"],C["white"]),
         "success":(C["accent2"],"#00AA88",C["bg"]),
         "ghost":(C["border"],C["hover"],C["tdim"]),
         "flat":(C["panel"],C["hover"],C["text"]),
         "danger":(C["err"],"#CC3333",C["white"]),
         "purple":(C["purple"],"#6D28D9",C["white"])}
    bg,hbg,fg=pal.get(style,pal["primary"])
    b=tk.Button(p,text=text,command=cmd,bg=bg,fg=fg,activebackground=hbg,
                activeforeground=fg,font=FH,relief="flat",cursor="hand2",
                padx=16,pady=8,bd=0,**kw)
    b.bind("<Enter>",lambda e:b.config(bg=hbg))
    b.bind("<Leave>",lambda e:b.config(bg=bg))
    return b

def card(p,**kw):
    return tk.Frame(p,bg=C["panel"],bd=0,highlightthickness=1,
                    highlightbackground=C["border"],**kw)

def mkentry(p,var,show="",width=None,readonly=False):
    """
    Create a styled Entry. When readonly=True the field is rendered with a
    dimmer background and is not editable, but its value still flows
    through to the saved defaults (we want this for the auto-computed
    today's-date fields).
    """
    kw=dict(textvariable=var,bg=C["ebg_dim"] if readonly else C["ebg"],
            fg=C["tdim"] if readonly else C["text"],font=FB,relief="flat",
            bd=0,insertbackground=C["accent"],highlightthickness=1,
            highlightbackground=C["ebdr"],highlightcolor=C["accent"])
    if show: kw["show"]=show
    if width: kw["width"]=width
    e=tk.Entry(p,**kw)
    if readonly:
        e.config(state="readonly", readonlybackground=C["ebg_dim"])
    return e

def scrollframe(p,bg=None):
    bg=bg or C["bg"]
    outer=tk.Frame(p,bg=bg)
    canvas=tk.Canvas(outer,bg=bg,highlightthickness=0)
    sb=ttk.Scrollbar(outer,orient="vertical",command=canvas.yview)
    inner=tk.Frame(canvas,bg=bg)
    win=canvas.create_window((0,0),window=inner,anchor="nw")
    inner.bind("<Configure>",lambda e:canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>",lambda e:canvas.itemconfig(win,width=e.width))
    canvas.configure(yscrollcommand=sb.set)
    sb.pack(side="right",fill="y"); canvas.pack(side="left",fill="both",expand=True)
    return outer,inner,canvas

class LogConsole(tk.Frame):
    def __init__(self,p,**kw):
        super().__init__(p,bg=C["bg"],**kw)
        h=tk.Frame(self,bg=C["panel"],pady=6); h.pack(fill="x")
        tk.Label(h,text="  \u258c Output Log",bg=C["panel"],fg=C["accent"],font=FH).pack(side="left")
        tk.Button(h,text="Clear",command=self.clear,bg=C["border"],fg=C["tdim"],
                  font=FS,relief="flat",cursor="hand2",padx=8,pady=2).pack(side="right",padx=8)
        self.txt=scrolledtext.ScrolledText(self,bg="#0A1520",fg=C["text"],font=FM,
                  relief="flat",bd=0,wrap="word",state="disabled")
        self.txt.pack(fill="both",expand=True)
        for tag,col in [("success",C["ok"]),("error",C["err"]),
                        ("warning",C["warn"]),("info",C["accent"]),("dim",C["tdim"])]:
            self.txt.tag_config(tag,foreground=col)
    def log(self,msg,tag=None):
        self.txt.config(state="normal")
        self.txt.insert("end",f"[{datetime.now().strftime('%H:%M:%S')}]  ","dim")
        self.txt.insert("end",msg+"\n",tag or "")
        self.txt.see("end"); self.txt.config(state="disabled")
    def clear(self):
        self.txt.config(state="normal"); self.txt.delete("1.0","end"); self.txt.config(state="disabled")


# ──────────────────────────────────────────────────────────────────────────────
# Subprocess launcher with per-batch counter support
# ──────────────────────────────────────────────────────────────────────────────
def _reset_batch(batch_token: str, log_fn):
    """
    Pre-launch step that calls the generator with --reset-batch so the next
    consume_batch_counter() returns 1. Runs synchronously and quickly.
    """
    try:
        subprocess.run(
            [sys.executable, str(GENERATOR),
             "--reset-batch", "--batch-token", batch_token],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=10,
        )
    except Exception as e:
        log_fn(f"[!] Could not reset batch counter: {e}", "warning")


def run_pipeline_thread(xlsx, use_def, log_fn, done_fn=None,
                        batch_token: str | None = None,
                        auto_today: bool | None = None):
    """
    Launch one generator subprocess for `xlsx`.

    `batch_token` / `auto_today` are forwarded as CLI flags so the
    generator can mint the per-batch nnn counter and honour the
    today-date override. The token is NOT reset here — callers reset
    it once at the start of a batch before calling this for the first
    file. The watcher calls _reset_batch() per scan-cycle; the manual
    "Generate Letter" handler calls it once per click.
    """
    def worker():
        args=[sys.executable,str(GENERATOR),"--xlsx",str(xlsx),"--output-dir",str(OUTPUT_DIR)]
        args.append("--use-defaults" if use_def else "--no-defaults")
        if batch_token:
            args += ["--batch-token", batch_token]
        if auto_today is True:
            args.append("--auto-today")
        elif auto_today is False:
            args.append("--no-auto-today")
        log_fn(f"\u25b6  Running pipeline for: {xlsx.name}","info")
        try:
            proc=subprocess.Popen(args,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,
                                  text=True,bufsize=1,encoding="utf-8",errors="replace")
            for line in proc.stdout:
                line=line.rstrip()
                if not line: continue
                tag=("success" if any(x in line for x in["\u2705","Done!","[OK]"])
                     else "error" if any(x in line for x in["\u274c","[!]","failed","Error"])
                     else "warning" if "\u26a0" in line else None)
                log_fn(line,tag)
            proc.wait()
            log_fn("\u2705  Pipeline completed successfully." if proc.returncode==0
                   else f"\u274c  Pipeline exited with code {proc.returncode}",
                   "success" if proc.returncode==0 else "error")
        except Exception as e: log_fn(f"\u274c  {e}","error")
        finally:
            if done_fn: done_fn()
    threading.Thread(target=worker,daemon=True).start()

def send_email_bg(pdf,log_fn):
    def worker():
        log_fn(f"\u2709  Sending email for {pdf.name}\u2026","info")
        r=subprocess.run([sys.executable,str(EMAILER),"--pdf",str(pdf)],
                         capture_output=True,text=True,encoding="utf-8",errors="replace")
        for line in (r.stdout+r.stderr).splitlines():
            if not line.strip(): continue
            tag=("success" if "[OK]" in line else "error" if "[X]" in line
                 else "warning" if "[!]" in line else None)
            log_fn(line,tag)
    threading.Thread(target=worker,daemon=True).start()

# ── Defaults dialog ──────────────────────────────────────────────────────────
class DefaultsDialog(tk.Toplevel):
    def __init__(self,p,mode="edit"):
        super().__init__(p)
        self.title("Default Variables")
        self.configure(bg=C["bg"]); self.geometry("760x760")
        self.resizable(True,True); self.transient(p); self.grab_set()
        self.mode=mode; self.entries={}
        # Track which entry widgets correspond to auto-today fields, so we
        # can grey/ungrey them live when the checkbox at the top of the
        # dialog is toggled.
        self._auto_entries: dict[str, tk.Entry] = {}
        self._auto_vars: dict[str, tk.StringVar] = {}

        defs = load_def()
        # The checkbox state is persisted under defaults["auto_today"].
        self._auto_today = tk.BooleanVar(value=bool(defs.get("auto_today", True)))

        hdr=tk.Frame(self,bg=C["accent"],pady=14,padx=20); hdr.pack(fill="x")
        icon="\U0001f441  " if mode=="show" else "\u2699  "
        tk.Label(hdr,text=icon+("View Defaults" if mode=="show" else "Edit Default Variables"),
                 bg=C["accent"],fg=C["white"],font=FT).pack(side="left")

        # ── "Always use today's date" toggle at the TOP of the dialog ───────
        # When ON, the date-derived rows render read-only and pull their
        # values from datetime.now(); when OFF, they become editable.
        # NOTE: month_name / month_year / end_of_month_date no longer fall
        # under this toggle — they're now always derived from the XLSX
        # FILENAME's billing month (📄 badge), regardless of this checkbox.
        top = tk.Frame(self, bg=C["bg"], padx=16, pady=8); top.pack(fill="x")
        cb_text = "\U0001f4c5  Always use today's date  (auto-fills letter date, Roman month, year-num, and letter number)"
        cb = tk.Checkbutton(top, text=cb_text, variable=self._auto_today,
                            bg=C["bg"], fg=C["accent2"], selectcolor=C["ebg"],
                            activebackground=C["bg"], font=FH, cursor="hand2",
                            command=self._apply_auto_today_to_fields)
        cb.pack(anchor="w")
        if mode=="show":
            cb.config(state="disabled")
        tk.Label(top,
                 text="\u2728 fields auto-fill from today's date (toggle above).  "
                      "\U0001f4c4 fields auto-fill from the XLSX filename's billing month.  "
                      "\U0001f4ca fields auto-fill from the XLSX header data.",
                 bg=C["bg"], fg=C["tdim"], font=FS,
                 wraplength=720, justify="left").pack(anchor="w", pady=(2,0))

        outer,inner,canvas=scrollframe(self); outer.pack(fill="both",expand=True,padx=16,pady=12)
        canvas.bind_all("<MouseWheel>",lambda e:canvas.yview_scroll(-1*(e.delta//120),"units"))

        # Use the new field list. Note `year_num` replaces `letter_date_num`
        # under "Date & Period", and discount fields stay under Invoice
        # Period but show as "(auto from XLSX)".
        groups=[("\U0001f3e2  Company",["company_name","company_abbr"]),
                ("\U0001f4c5  Date & Period",["letter_date","month_name","month_year","month_roman","letter_number","year_num"]),
                ("\U0001f4b0  Financial",["deposit_rate","deposit_days","daily_rate","annual_rate","rate_multiplier_num","rate_multiplier_word"]),
                ("\U0001f4c6  Invoice Period",["end_of_month_date","discount_from","discount_until"]),
                ("\U0001f465  Finance Team",["team1_honorific_name","team2_honorific_name","team_email","team3_honorific_name","team3_email","team4_honorific_name","team4_email"]),]
        sm={k:(l,e) for k,l,e in SCHEMA}
        today_vals = compute_today_values()

        for gn,keys in groups:
            gh=tk.Frame(inner,bg=C["bg"]); gh.pack(fill="x",pady=(14,4))
            tk.Label(gh,text=gn,bg=C["bg"],fg=C["accent"],font=FH).pack(side="left")
            tk.Frame(gh,bg=C["border"],height=1).pack(side="left",fill="x",expand=True,padx=(8,0))
            for key in keys:
                if key not in sm: continue
                lbl,ex=sm[key]
                # Mark auto rows with a leading icon so the user sees at a
                # glance which fields are computed:
                #   ✨  sparkle    = auto from today
                #   📄  page      = auto from XLSX filename's month/year
                #   📊  chart     = auto from XLSX header data
                badge = ""
                if key in AUTO_TODAY_KEYS:
                    badge = "\u2728 "
                elif key in FILENAME_DERIVED_KEYS:
                    badge = "\U0001f4c4 "    # page = auto from filename
                elif key in XLSX_DERIVED_KEYS:
                    badge = "\U0001f4ca "    # bar chart = auto from XLSX
                row=tk.Frame(inner,bg=C["bg"]); row.pack(fill="x",pady=2)
                tk.Label(row,text=badge+lbl,bg=C["bg"],fg=C["tdim"],font=FS,width=34,anchor="w").pack(side="left")

                # Compute the value to show:
                # - auto-today field AND auto_today is ON → today's value (read-only)
                # - filename-derived field → always show stub (read-only,
                #   since the value depends on the specific XLSX being run)
                # - XLSX-header-derived field → always show stub (read-only)
                # - otherwise → stored default or example placeholder
                stored = defs.get(key, "")
                if key in AUTO_TODAY_KEYS and self._auto_today.get():
                    initial = today_vals.get(key, stored or ex)
                elif key in FILENAME_DERIVED_KEYS:
                    initial = "(auto-derived from filename)"
                elif key in XLSX_DERIVED_KEYS:
                    initial = "(auto-derived from XLSX)"
                else:
                    initial = stored or ex

                if mode=="show":
                    tk.Label(row,text=initial if initial else f"(not set — e.g. {ex})",
                             bg=C["bg"],fg=C["text"] if initial else C["tdim"],font=FB,anchor="w").pack(side="left",fill="x",expand=True)
                else:
                    var=tk.StringVar(value=initial)
                    is_xlsx_derived     = (key in XLSX_DERIVED_KEYS)
                    is_filename_derived = (key in FILENAME_DERIVED_KEYS)
                    is_auto_now         = (key in AUTO_TODAY_KEYS and self._auto_today.get())
                    e=mkentry(row,var,readonly=(is_xlsx_derived or is_filename_derived or is_auto_now))
                    e.pack(side="left",fill="x",expand=True,ipady=4)

                    # Apply placeholder-style behaviour only for fields that
                    # are currently editable AND have no stored value.
                    if not stored and ex and not is_xlsx_derived and not is_filename_derived and not is_auto_now:
                        e.config(fg=C["tdim"])
                        e.bind("<FocusIn>",lambda ev,en=e,x=ex:(en.get()==x and (en.delete(0,"end"),en.config(fg=C["text"]))))
                        e.bind("<FocusOut>",lambda ev,en=e,x=ex,v=var:(not en.get() and (v.set(x),en.config(fg=C["tdim"]))))
                    self.entries[key]=var
                    if key in AUTO_TODAY_KEYS:
                        self._auto_entries[key] = e
                        self._auto_vars[key]    = var

        if mode!="show":
            foot=tk.Frame(self,bg=C["bg"],pady=12); foot.pack(fill="x",padx=16)
            mkbtn(foot,"\U0001f4be  Save Defaults",self._save,"success").pack(side="right",padx=4)
            mkbtn(foot,"\u2715  Cancel",self.destroy,"ghost").pack(side="right",padx=4)
        else:
            tk.Button(self,text="Close",command=self.destroy,bg=C["accent"],fg=C["white"],
                      font=FH,relief="flat",pady=8,cursor="hand2").pack(fill="x",padx=16,pady=12)
        self.update_idletasks()
        x=p.winfo_rootx()+p.winfo_width()//2-self.winfo_width()//2
        y=p.winfo_rooty()+p.winfo_height()//2-self.winfo_height()//2
        self.geometry(f"+{x}+{y}")

    def _apply_auto_today_to_fields(self):
        """
        Called whenever the user toggles the "Always use today's date"
        checkbox INSIDE the Defaults dialog. Recomputes today's values
        and flips the AUTO_TODAY_KEYS rows between read-only-greyed and
        editable, refreshing displayed values either way.
        """
        if self.mode == "show":
            return
        on = self._auto_today.get()
        today_vals = compute_today_values()
        defs = load_def()
        sm={k:(l,e) for k,l,e in SCHEMA}
        for key, e in self._auto_entries.items():
            var = self._auto_vars[key]
            if on:
                # Force the value to today's computed value and lock the entry.
                var.set(today_vals.get(key, ""))
                e.config(state="readonly", readonlybackground=C["ebg_dim"],
                         fg=C["tdim"])
            else:
                # Unlock and show whatever was stored, falling back to example.
                e.config(state="normal", bg=C["ebg"], fg=C["text"])
                stored = defs.get(key, "")
                if stored:
                    var.set(stored)
                else:
                    ex = sm.get(key, ("",""))[1]
                    var.set(ex)

    def _save(self):
        data=load_def(); sm={k:(l,e) for k,l,e in SCHEMA}
        # Persist the toggle state itself first.
        data["auto_today"] = self._auto_today.get()
        for key,var in self.entries.items():
            # Never save the XLSX-derived placeholders.
            if key in XLSX_DERIVED_KEYS:
                data.pop(key, None)
                continue
            v=var.get().strip(); ex=sm.get(key,("",""))[1]
            if v and v!=ex: data[key]=v
        save_def(data); messagebox.showinfo("Saved","Defaults saved.",parent=self); self.destroy()

# ── Recipients dialog ────────────────────────────────────────────────────────
class RecipientDialog(tk.Toplevel):
    def __init__(self,p):
        super().__init__(p)
        self.title("Email Recipients"); self.configure(bg=C["bg"])
        self.geometry("700x520"); self.resizable(True,True)
        self.transient(p); self.grab_set(); self._rows=[]
        hdr=tk.Frame(self,bg=C["purple"],pady=14,padx=20); hdr.pack(fill="x")
        tk.Label(hdr,text="\u2709  Email Recipients",bg=C["purple"],fg=C["white"],font=FT).pack(side="left")
        tk.Label(self,text="Client name matches the XLSX filename (case-insensitive). Separate addresses with  ;",
                 bg=C["bg"],fg=C["tdim"],font=FS,justify="left").pack(anchor="w",padx=16,pady=(8,4))
        hr=tk.Frame(self,bg=C["bg"]); hr.pack(fill="x",padx=16)
        for t,w in[("Client Name (UPPERCASE)",22),("TO  (semicolons)",30),("CC  (semicolons)",24),("",4)]:
            tk.Label(hr,text=t,bg=C["bg"],fg=C["tdim"],font=FS,width=w,anchor="w").pack(side="left",padx=2)
        outer,self._inner,canvas=scrollframe(self); outer.pack(fill="both",expand=True,padx=16,pady=4)
        canvas.bind_all("<MouseWheel>",lambda e:canvas.yview_scroll(-1*(e.delta//120),"units"))
        cfg=load_ecfg()
        for name,val in cfg.get("recipients",{}).items():
            self._add(name,"; ".join(val.get("to",[])),"; ".join(val.get("cc",[])))
        foot=tk.Frame(self,bg=C["bg"],pady=10); foot.pack(fill="x",padx=16)
        mkbtn(foot,"+ Add Client",self._add,"flat").pack(side="left",padx=(0,6))
        # NEW: bulk-import from an XLSX (e.g. List_RD_AR_Konfirmasi.xlsx). Same
        # underlying logic as the sidebar Import button — kept here too because
        # the Recipients dialog is the natural place users look for it.
        mkbtn(foot,"\U0001f4e5  Import XLSX\u2026",self._import_xlsx,"flat").pack(side="left",padx=6)
        mkbtn(foot,"\u2715  Cancel",self.destroy,"ghost").pack(side="right")
        mkbtn(foot,"\U0001f4be  Save",self._save,"success").pack(side="right",padx=6)
        # NEW: Delete All button (danger style)
        mkbtn(foot,"\U0001f5d1  Delete All",self._delete_all,"danger").pack(side="left",padx=6)
        mkbtn(foot,"\u2715  Cancel",self.destroy,"ghost").pack(side="right")
        mkbtn(foot,"\U0001f4be  Save",self._save,"success").pack(side="right",padx=6)        
        self.update_idletasks()
        x=p.winfo_rootx()+p.winfo_width()//2-self.winfo_width()//2
        y=p.winfo_rooty()+p.winfo_height()//2-self.winfo_height()//2
        self.geometry(f"+{x}+{y}")
    def _add(self,name="",to="",cc=""):
        row=tk.Frame(self._inner,bg=C["bg"]); row.pack(fill="x",pady=2)
        nv,tv,cv=tk.StringVar(value=name),tk.StringVar(value=to),tk.StringVar(value=cc)
        mkentry(row,nv,width=22).pack(side="left",padx=2,ipady=4)
        mkentry(row,tv,width=30).pack(side="left",padx=2,ipady=4)
        mkentry(row,cv,width=24).pack(side="left",padx=2,ipady=4)
        def delete(r=row,t=(nv,tv,cv)):
            r.destroy(); self._rows=[x for x in self._rows if x[0] is not t[0]]
        tk.Button(row,text="\u2715",command=delete,bg=C["border"],fg=C["err"],
                  font=FS,relief="flat",cursor="hand2",padx=4).pack(side="left",padx=2)
        self._rows.append((nv,tv,cv))
    def _import_xlsx(self):
        """
        In-dialog import: read an XLSX of recipients, merge into the config,
        then reload the rows shown in this dialog. Unlike the sidebar's
        equivalent, this one also asks the user whether to keep their
        currently-edited (but not yet saved) rows. To keep the logic simple
        and avoid losing unsaved edits silently, we offer two paths:
          - Cancel: do nothing
          - OK: discard unsaved rows, run merge import, reload from disk
        """
        if not EMAILER.exists():
            messagebox.showerror("Missing",
                f"smartpayer_emailer.py not found:\n{EMAILER}",
                parent=self); return
        if self._rows:
            ok=messagebox.askokcancel(
                "Discard unsaved edits?",
                "Importing will reload the recipients list from disk and "
                "discard any unsaved changes shown in this dialog.\n\n"
                "Continue?", parent=self)
            if not ok: return
        path=filedialog.askopenfilename(
            title="Choose recipient list XLSX",
            filetypes=[("Excel","*.xlsx"),("All files","*.*")],
            parent=self)
        if not path: return
        ans=messagebox.askyesnocancel(
            "Import mode",
            "Choose import mode:\n\n"
            "  YES   \u2192  MERGE  (keep existing; overwrite per row)\n"
            "  NO    \u2192  REPLACE (wipe existing first)\n"
            "  CANCEL \u2192  abort",
            parent=self)
        if ans is None: return
        mode="merge" if ans else "replace"
        try:
            r=subprocess.run([sys.executable,str(EMAILER),
                              "--import-recipients",path,
                              "--import-mode",mode],
                             capture_output=True,text=True,
                             encoding="utf-8",errors="replace")
        except Exception as e:
            messagebox.showerror("Import failed",str(e),parent=self); return
        if r.returncode!=0:
            messagebox.showerror("Import failed",
                (r.stderr or r.stdout or "Unknown error").strip()[:500],
                parent=self); return
        # Reload the dialog's rows from disk
        for row_widget in list(self._inner.children.values()):
            row_widget.destroy()
        self._rows.clear()
        cfg=load_ecfg()
        for name,val in cfg.get("recipients",{}).items():
            self._add(name,"; ".join(val.get("to",[])),"; ".join(val.get("cc",[])))
        summary=(r.stdout or "").strip().splitlines()[-1] if r.stdout else "Done."
        messagebox.showinfo("Import complete",summary,parent=self)
    def _delete_all(self):
        """
        Delete all recipient rows after user confirmation.
        Clears both the UI widgets and the internal _rows list.
        """
        if not self._rows:
            messagebox.showinfo("Nothing to delete", 
                "No recipients to clear.", parent=self)
            return
        
        # Confirm before destructive action
        if messagebox.askyesno("Confirm delete",
                f"Delete all {len(self._rows)} recipient entry/entries?\n\n"
                "This will clear the list but won't save until you click Save.",
                parent=self, icon="warning"):
            
            # Destroy all row widgets from the scrollable frame
            for row_widget in list(self._inner.children.values()):
                row_widget.destroy()
            
            # Clear the internal list that tracks StringVars
            self._rows.clear()
            
            # Optional: visual feedback
            messagebox.showinfo("Cleared", 
                "All recipients removed. Click Save to persist changes.", 
                parent=self)            
    def _save(self):
        cfg=load_ecfg(); cfg["recipients"]={}
        for nv,tv,cv in self._rows:
            n=nv.get().strip().upper()
            if n: cfg["recipients"][n]={"to":[a.strip() for a in tv.get().split(";") if a.strip()],
                                         "cc":[a.strip() for a in cv.get().split(";") if a.strip()]}
        save_ecfg(cfg); messagebox.showinfo("Saved",f"{len(cfg['recipients'])} recipient(s) saved.",parent=self); self.destroy()

# ── SMTP dialog ──────────────────────────────────────────────────────────────
class SmtpDialog(tk.Toplevel):
    def __init__(self,p):
        super().__init__(p)
        self.title("SMTP / Email Setup"); self.configure(bg=C["bg"])
        self.geometry("520x550"); self.resizable(False,False)
        self.transient(p); self.grab_set(); self._vars={}
        hdr=tk.Frame(self,bg=C["purple"],pady=14,padx=20); hdr.pack(fill="x")
        tk.Label(hdr,text="\u2699  SMTP / Email Setup",bg=C["purple"],fg=C["white"],font=FT).pack(side="left")
        body=tk.Frame(self,bg=C["bg"],padx=20,pady=16); body.pack(fill="both",expand=True)
        cfg=load_ecfg(); smtp=cfg.get("smtp",{})
        for lbl,key,defval,pw in[
            ("SMTP Host","host",smtp.get("host","smtp.gmail.com"),False),
            ("SMTP Port","port",str(smtp.get("port",587)),False),
            ("Username / Email","username",smtp.get("username",""),False),
            ("Password / App Key","password",smtp.get("password",""),True),
            ("Sender Display Name","sender_name",smtp.get("sender_name","ARRD Team"),False),
            ("Sender Email","sender_email",smtp.get("sender_email",""),False)]:
            row=tk.Frame(body,bg=C["bg"]); row.pack(fill="x",pady=4)
            tk.Label(row,text=lbl,bg=C["bg"],fg=C["tdim"],font=FS,width=22,anchor="w").pack(side="left")
            var=tk.StringVar(value=defval)
            mkentry(row,var,show="\u25cf" if pw else "").pack(side="left",fill="x",expand=True,ipady=5)
            self._vars[key]=var
        self._tls=tk.BooleanVar(value=smtp.get("use_tls",True))
        tk.Checkbutton(body,text="Use STARTTLS (port 587)",variable=self._tls,
                       bg=C["bg"],fg=C["text"],selectcolor=C["ebg"],
                       activebackground=C["bg"],font=FB,cursor="hand2").pack(anchor="w",pady=6)
        tk.Frame(body,bg=C["border"],height=1).pack(fill="x",pady=8)
        tk.Label(body,text="CC — always added to every email (separate with  ;):",
                 bg=C["bg"],fg=C["tdim"],font=FS).pack(anchor="w")
        self._cc=tk.StringVar(value="; ".join(cfg.get("cc_always",[])))
        mkentry(body,self._cc).pack(fill="x",ipady=4,pady=(2,0))
        tk.Frame(body,bg=C["border"],height=1).pack(fill="x",pady=8)
        self._outlook=tk.BooleanVar(value=cfg.get("outlook_fallback",{}).get("enabled",True))
        tk.Checkbutton(body,text="Use Outlook desktop app if SMTP not configured (Windows)",
                       variable=self._outlook,bg=C["bg"],fg=C["text"],selectcolor=C["ebg"],
                       activebackground=C["bg"],font=FB,cursor="hand2").pack(anchor="w")
        foot=tk.Frame(self,bg=C["bg"],pady=12); foot.pack(fill="x",padx=20)
        mkbtn(foot,"\U0001f4be  Save",self._save,"success").pack(side="right",padx=(6,0))
        mkbtn(foot,"\u2715  Cancel",self.destroy,"ghost").pack(side="right")
        mkbtn(foot,"\U0001f9ea  Test Send",self._test,"flat").pack(side="left")
        self.update_idletasks()
        x=p.winfo_rootx()+p.winfo_width()//2-self.winfo_width()//2
        y=p.winfo_rooty()+p.winfo_height()//2-self.winfo_height()//2
        self.geometry(f"+{x}+{y}")
    def _save(self):
        cfg=load_ecfg(); s=cfg.setdefault("smtp",{})
        s["host"]=self._vars["host"].get().strip()
        s["port"]=int(self._vars["port"].get().strip() or 587)
        s["use_tls"]=self._tls.get()
        s["username"]=self._vars["username"].get().strip()
        s["password"]=self._vars["password"].get().strip()
        s["sender_name"]=self._vars["sender_name"].get().strip()
        s["sender_email"]=self._vars["sender_email"].get().strip() or s["username"]
        cfg["cc_always"]=[a.strip() for a in self._cc.get().split(";") if a.strip()]
        cfg["outlook_fallback"]={"enabled":self._outlook.get()}
        if "email_template" not in cfg:
            cfg["email_template"]={
                "subject":"Smart Payer {client_name} Periode {month} {year}",
                "body":("Dear Bapak / Ibu Pimpinan RD,\r\n\r\n"
                        "Terlampir surat dari manajeman kami terkait program "
                        "\u201cSmart Payer\u201d periode {month} {year} "
                        "Invoice Jatuh tempo {next_month} {next_year}. "
                        "Mohon dapat diterima dan dipelajari.\r\n\r\n"
                        "Kami tunggu feedback nya.\r\n\r\n"
                        "Atas perhatian dan kerjasamanya kami ucapkan terima kasih.\r\n\r\n"
                        "Best Regards,\r\nARRD Team")}
        cfg.setdefault("recipients",{})
        save_ecfg(cfg); messagebox.showinfo("Saved","Email settings saved.",parent=self); self.destroy()
    def _test(self):
        self._save()
        def worker():
            if not EMAILER.exists():
                messagebox.showerror("Missing","smartpayer_emailer.py not found.",parent=self); return
            cfg=load_ecfg(); smtp=cfg.get("smtp",{})
            to=smtp.get("sender_email") or smtp.get("username","")
            if not to:
                messagebox.showwarning("No sender","Set a sender email first.",parent=self); return
            import copy,tempfile
            tc=copy.deepcopy(cfg); tc["recipients"]["__TEST__"]={"to":[to],"cc":[]}
            tmp=Path(tempfile.mktemp(suffix=".json"))
            with open(tmp,"w",encoding="utf-8") as f: json.dump(tc,f)
            fake=Path(tempfile.mktemp()); fake.write_bytes(b"%PDF-1.4 test")
            named=fake.parent/f"Smartpayer_April_2026___TEST___letter.pdf"
            fake.rename(named)
            r=subprocess.run([sys.executable,str(EMAILER),"--pdf",str(named)],
                             capture_output=True,text=True,
                             env={**os.environ,"SMARTPAYER_EMAIL_CONFIG":str(tmp)})
            for p2 in[named,tmp]:
                try: p2.unlink()
                except: pass
            out=(r.stdout+r.stderr).strip()
            if "[OK]" in out: messagebox.showinfo("Test OK",f"Test email sent to {to}.",parent=self)
            else: messagebox.showerror("Test Failed",out[-500:] or "Unknown error.",parent=self)
        threading.Thread(target=worker,daemon=True).start()

# ── Watch manager ────────────────────────────────────────────────────────────
class WatchMgr:
    """
    Polling-based watcher (kept as-is from the original). Now wraps each
    scan-cycle that picks up at least one new file in a fresh batch token
    so the per-batch counter restarts at 001 each cycle. Files within the
    same cycle are processed serially (the inner run_pipeline_thread
    function uses subprocess.Popen + proc.wait, so the next iteration
    blocks until the previous finishes), so counters increment cleanly
    001, 002, 003, … within one cycle.
    """
    def __init__(self,log,auto_today_provider=None):
        self._log=log; self._stop=threading.Event(); self._t=None
        # auto_today_provider is a callable returning the current bool so
        # the watcher always reads the live checkbox value, not a snapshot.
        self._auto_today_provider = auto_today_provider or (lambda: True)
        # Counters for bulk processing summary 
        self._processed_count = 0      # Files successfully generated
        self._email_sent_count = 0     # Emails successfully sent
        self._error_count = 0          # Files that failed
        self._counter_lock = threading.Lock()  # Thread-safe updates
    def start(self,folder,use_def,auto_send):
        self._stop.clear()
        self._t=threading.Thread(target=self._loop,args=(folder,use_def,auto_send),daemon=True)
        self._t.start()
    def stop(self): self._stop.set()
    def running(self): return self._t is not None and self._t.is_alive()
    def _increment_processed(self):
        """Thread-safe increment of processed counter."""
        with self._counter_lock:
            self._processed_count += 1
            count = self._processed_count
        self._log(f"  [✓] File processed ({count} total)", "success")

    def _increment_email_sent(self):
        """Thread-safe increment of email sent counter."""
        with self._counter_lock:
            self._email_sent_count += 1
            count = self._email_sent_count
        self._log(f"  [✉] Email sent ({count} total)", "success")

    def _increment_error(self):
        """Thread-safe increment of error counter."""
        with self._counter_lock:
            self._error_count += 1
            count = self._error_count
        self._log(f"  [✗] Error processing file ({count} errors)", "error")

    def _log_batch_summary(self, batch_token, file_count):
        """Log summary at end of batch/scan-cycle."""
        with self._counter_lock:
            processed = self._processed_count
            sent = self._email_sent_count
            errors = self._error_count
        
        self._log(f"\n  📊 Batch Summary ({batch_token}):", "info")
        self._log(f"     Files detected:  {file_count}", "info")
        self._log(f"     ✓ Processed:     {processed}", "success")
        self._log(f"     ✉ Emails sent:   {sent}", "success")
        if errors > 0:
            self._log(f"     ✗ Errors:        {errors}", "error")
        self._log(f"     {'─' * 40}", "info")    
    def _loop(self, folder, use_def, auto_send):
        self._log(f"\U0001f441  Watching: {folder}", "info")
        seen = set()
        
        while not self._stop.is_set():
            try:
                # Reset counters at start of each scan-cycle for per-batch reporting
                with self._counter_lock:
                    self._processed_count = 0
                    self._email_sent_count = 0
                    self._error_count = 0
                
                # Snapshot what's new this scan-cycle
                new_files = [f for f in Path(folder).glob("*.xlsx") if f not in seen]
                
                if new_files:
                    # Daily-scoped token: "daily-20260513" → resets automatically at midnight
                    batch_token = f"daily-{datetime.now().strftime('%Y%m%d')}"
                    auto_today = self._auto_today_provider()
                    
                    self._log(f"📦 New batch detected ({len(new_files)} file(s)) — token {batch_token}", "info")
                    
                    for f in new_files:
                        seen.add(f)
                        self._log(f"\U0001f4c2  New file: {f.name}", "info")
                        time.sleep(1)  # Let file finish writing
                        
                        done_ev = threading.Event()
                        
                        def done(f=f, ev=done_ev, auto_send=auto_send):
                            try:
                                pdf = OUTPUT_DIR / (f.stem + "_letter.pdf")
                                if pdf.exists():
                                    self._increment_processed()
                                    if auto_send and EMAILER.exists():
                                        try:
                                            send_email_bg(pdf, self._log)
                                            self._increment_email_sent()
                                        except Exception as e:
                                            self._log(f"  [!] Email send failed: {e}", "error")
                                            self._increment_error()
                                else:
                                    self._log(f"  [!] PDF not found for {f.name}", "error")
                                    self._increment_error()
                            finally:
                                ev.set()  # ✅ NEW: Signal completion ONLY after counter updates
                        
                        run_pipeline_thread(f, use_def, self._log, done_fn=done,
                                            batch_token=batch_token,
                                            auto_today=auto_today)
                        # Wait for this file to finish before starting the next
                        done_ev.wait(timeout=600)
                    
                    # Log summary after processing all files in this batch
                    self._log_batch_summary(batch_token, len(new_files))
                        
            except Exception as e:
                self._log(f"Watch error: {e}", "error")
            time.sleep(2)
        
        # Final summary when watcher stops
        self._log("\n\u23f9  Watch stopped. Final summary:", "info")
        with self._counter_lock:
            self._log(f"   Total processed:  {self._processed_count}", "success")
            self._log(f"   Total emails:     {self._email_sent_count}", "success")
            if self._error_count > 0:
                self._log(f"   Total errors:     {self._error_count}", "error")
# ── Main app ─────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SmartPayer Letter Generator")
        self.configure(bg=C["bg"]); self.geometry("1060x820"); self.minsize(860,640)

        # Load persisted state for the "Always use today's date" toggle.
        # Default to True so the first-time user gets the new behaviour.
        d = load_def()
        self._auto_today = tk.BooleanVar(value=bool(d.get("auto_today", True)))
        # Whenever the checkbox flips, save & refresh the live preview.
        self._auto_today.trace_add("write", lambda *_: self._on_auto_today_changed())

        # WatchMgr reads the live checkbox via a closure so toggling at
        # runtime affects future scan-cycles immediately.
        self._wm=WatchMgr(self._log, auto_today_provider=self._auto_today.get)
        self._busy=False; self._xlsx=None
        self._wfv=tk.StringVar(); self._wdef=tk.BooleanVar(value=True)
        self._pdef=tk.BooleanVar(value=True); self._asend=tk.BooleanVar(value=False)
        self._wasend=tk.BooleanVar(value=False)

        # StringVars for the live "today's values" preview row in the
        # sidebar — updated by _on_auto_today_changed().
        self._preview_date  = tk.StringVar()
        self._preview_lno   = tk.StringVar()
        self._preview_eom   = tk.StringVar()

        self._build(); self._refresh_preview(); self._check()

    def _on_auto_today_changed(self):
        # Persist immediately so a GUI crash doesn't lose the toggle.
        d = load_def()
        d["auto_today"] = self._auto_today.get()
        save_def(d)
        self._refresh_preview()

    def _refresh_preview(self):
        """Refresh the read-only preview labels shown under the checkbox."""
        v = compute_today_values()
        if self._auto_today.get():
            self._preview_date.set(
                f"\u2728  Today: {v['letter_date']}  "
                f"\u00b7  Letter No. prefix: {v['letter_number'][:4]}\u2026  "
                f"\u00b7  \U0001f4c4 month + end-of-month: from XLSX filename"
            )
        else:
            d = load_def()
            self._preview_date.set(f"\u270f  Manual mode \u2014 using stored defaults  "
                                   f"(letter_date={d.get('letter_date','?')}, "
                                   f"letter_number={d.get('letter_number','?')})")

    def _build(self):
        # Banner
        ban=tk.Frame(self,bg=C["accent"]); ban.pack(fill="x")
        tk.Frame(ban,bg=C["adim"],width=6).pack(side="left",fill="y")
        a=tk.Frame(ban,bg=C["accent"],padx=20,pady=14); a.pack(side="left",fill="both",expand=True)
        tk.Label(a,text="SmartPayer Letter Generator",bg=C["accent"],fg=C["white"],font=FT).pack(side="left")
        tk.Label(a,text="  Automate \u00b7 Fill \u00b7 Export",bg=C["accent"],
                 fg="#BDE8FF",font=("Segoe UI",10,"italic")).pack(side="left",pady=4)
        self._stv=tk.StringVar(value="Ready")
        self._stl=tk.Label(ban,textvariable=self._stv,bg=C["adim"],fg=C["white"],
                           font=("Segoe UI",9,"bold"),padx=14,pady=6)
        self._stl.pack(side="right",padx=12)
        # Main
        main=tk.Frame(self,bg=C["bg"]); main.pack(fill="both",expand=True)
        # Scrollable sidebar
        sbo=tk.Frame(main,bg=C["bg"],width=416); sbo.pack(side="left",fill="y",padx=(12,0),pady=12)
        sbo.pack_propagate(False)
        sbc=tk.Canvas(sbo,bg=C["bg"],highlightthickness=0,width=396)
        sbs=ttk.Scrollbar(sbo,orient="vertical",command=sbc.yview)
        sbc.configure(yscrollcommand=sbs.set); sbs.pack(side="right",fill="y"); sbc.pack(side="left",fill="both",expand=True)
        sb=tk.Frame(sbc,bg=C["bg"]); wid=sbc.create_window((0,0),window=sb,anchor="nw")
        sb.bind("<Configure>",lambda e:sbc.configure(scrollregion=sbc.bbox("all")))
        sbc.bind("<Configure>",lambda e:sbc.itemconfig(wid,width=e.width))
        def _scr(e): sbc.yview_scroll(-1*(e.delta//120),"units")
        sbo.bind("<Enter>",lambda e:sbc.bind_all("<MouseWheel>",_scr))
        sbo.bind("<Leave>",lambda e:sbc.unbind_all("<MouseWheel>"))
        self._sidebar(sb)
        # Log
        lf=tk.Frame(main,bg=C["bg"]); lf.pack(side="left",fill="both",expand=True,padx=(8,12),pady=12)
        self._con=LogConsole(lf); self._con.pack(fill="both",expand=True)

    def _reset_watch_counters(self):
        """Reset watch mode counters via the WatchMgr."""
        if hasattr(self, '_wm') and self._wm.running():
            with self._wm._counter_lock:
                self._wm._processed_count = 0
                self._wm._email_sent_count = 0
                self._wm._error_count = 0
            self._log("  [🔄] Watch counters reset", "info")
        else:
            self._log("  [!] Watcher not running", "warning")        

    def _sidebar(self,p):
        # S1 Process
        s=card(p); s.pack(fill="x",pady=(0,10))
        tk.Frame(s,bg=C["accent"],height=3).pack(fill="x")
        tk.Label(s,text="\U0001f4c4  Process Spreadsheet",bg=C["panel"],fg=C["accent"],
                 font=FH,pady=10,padx=14,anchor="w").pack(fill="x")
        b=tk.Frame(s,bg=C["panel"],padx=14,pady=4); b.pack(fill="x")

        # ── "Always use today's date" CHECKBOX AT THE TOP (change #1) ───────
        # Default ON, persisted in defaults JSON. Toggling it instantly
        # refreshes the preview line directly below.
        top_cb_row = tk.Frame(b, bg=C["panel"]); top_cb_row.pack(fill="x", pady=(4,2))
        tk.Checkbutton(top_cb_row,
                       text="\U0001f4c5  Always use today's date",
                       variable=self._auto_today,
                       bg=C["panel"], fg=C["accent2"], selectcolor=C["ebg"],
                       activebackground=C["panel"], font=FH, cursor="hand2"
                       ).pack(side="left")
        # Live preview line — re-rendered by _refresh_preview()
        tk.Label(b, textvariable=self._preview_date, bg=C["panel"], fg=C["tdim"],
                 font=FS, anchor="w", justify="left", wraplength=360
                 ).pack(fill="x", padx=2, pady=(0,4))

        # File picker row
        fr=tk.Frame(b,bg=C["panel"]); fr.pack(fill="x",pady=4)
        self._xlbl=tk.StringVar(value="No file selected")
        tk.Label(fr,textvariable=self._xlbl,bg=C["ebg"],fg=C["tdim"],font=FS,
                 anchor="w",padx=8,pady=6).pack(side="left",fill="x",expand=True)
        tk.Button(fr,text="Browse\u2026",command=self._browse_x,bg=C["border"],fg=C["text"],
                  font=FS,relief="flat",cursor="hand2",padx=8,pady=6).pack(side="right",padx=(4,0))
        for var,txt in[(self._pdef,"Use default variables"),(self._asend,"Auto-send email after generating")]:
            r=tk.Frame(b,bg=C["panel"]); r.pack(fill="x",pady=2)
            tk.Checkbutton(r,text=txt,variable=var,bg=C["panel"],fg=C["text"],
                           selectcolor=C["ebg"],activebackground=C["panel"],
                           font=FB,cursor="hand2").pack(side="left")
        br=tk.Frame(b,bg=C["panel"],pady=8); br.pack(fill="x")
        self._pbtn=mkbtn(br,"\u25b6  Generate Letter",self._process,"primary"); self._pbtn.pack(fill="x")

        # S2 Watch
        s=card(p); s.pack(fill="x",pady=(0,10))
        tk.Frame(s,bg=C["accent2"],height=3).pack(fill="x")
        tk.Label(s,text="\U0001f501  Auto Trigger \u2014 Watch Folder",bg=C["panel"],fg=C["accent2"],
                 font=FH,pady=10,padx=14,anchor="w").pack(fill="x")
        b=tk.Frame(s,bg=C["panel"],padx=14,pady=4); b.pack(fill="x")
        wr=tk.Frame(b,bg=C["panel"]); wr.pack(fill="x",pady=4)
        mkentry(wr,self._wfv).pack(side="left",fill="x",expand=True,ipady=5,padx=(0,4))
        tk.Button(wr,text="\U0001f4c2",command=self._browse_w,bg=C["border"],fg=C["text"],
                  font=FS,relief="flat",cursor="hand2",padx=8,pady=5).pack(side="right")
        for var,txt in[(self._wdef,"Auto-use defaults when triggered"),(self._wasend,"Auto-send email for each generated letter")]:
            r=tk.Frame(b,bg=C["panel"]); r.pack(fill="x",pady=2)
            tk.Checkbutton(r,text=txt,variable=var,bg=C["panel"],fg=C["text"],
                           selectcolor=C["ebg"],activebackground=C["panel"],
                           font=FB,cursor="hand2").pack(side="left")
        br=tk.Frame(b,bg=C["panel"],pady=8); br.pack(fill="x")
        self._wbtn=mkbtn(br,"\u25b6  Start Watching",self._toggle_watch,"success"); self._wbtn.pack(fill="x")
        self._wsv=tk.StringVar(value="\u23f9  Not watching")
        tk.Label(b,textvariable=self._wsv,bg=C["panel"],fg=C["tdim"],font=FS,pady=4).pack(anchor="w")

        # Optional: Clear counters button
        br=tk.Frame(b,bg=C["panel"],pady=4); br.pack(fill="x")
        mkbtn(br,"\U0001f504  Reset Counters", self._reset_watch_counters, "ghost").pack(side="left")        

        # S3 Defaults
        s=card(p); s.pack(fill="x",pady=(0,10))
        tk.Frame(s,bg=C["warn"],height=3).pack(fill="x")
        tk.Label(s,text="\u2699  Default Variables",bg=C["panel"],fg=C["warn"],
                 font=FH,pady=10,padx=14,anchor="w").pack(fill="x")
        b=tk.Frame(s,bg=C["panel"],padx=14,pady=8); b.pack(fill="x")
        br=tk.Frame(b,bg=C["panel"]); br.pack(fill="x",pady=4)
        mkbtn(br,"\U0001f4cb  Show",lambda:DefaultsDialog(self,"show"),"flat").pack(side="left",fill="x",expand=True,padx=(0,4))
        mkbtn(br,"\u270f  Edit",self._edit_def,"flat").pack(side="left",fill="x",expand=True,padx=(4,0))
        self._dsv=tk.StringVar(); tk.Label(b,textvariable=self._dsv,bg=C["panel"],fg=C["tdim"],font=FS,pady=4).pack(anchor="w")
        self._ref_def()

        # S4 Email
        s=card(p); s.pack(fill="x",pady=(0,10))
        tk.Frame(s,bg=C["purple"],height=3).pack(fill="x")
        tk.Label(s,text="\u2709  Email Automation",bg=C["panel"],fg=C["plite"],
                 font=FH,pady=10,padx=14,anchor="w").pack(fill="x")
        b=tk.Frame(s,bg=C["panel"],padx=14,pady=8); b.pack(fill="x")
        br=tk.Frame(b,bg=C["panel"]); br.pack(fill="x",pady=4)
        mkbtn(br,"\U0001f4cb  Recipients",self._edit_recip,"flat").pack(side="left",fill="x",expand=True,padx=(0,4))
        mkbtn(br,"\u2699  SMTP",self._edit_smtp,"flat").pack(side="left",fill="x",expand=True,padx=(2,2))
        # NEW: bulk-import recipients from XLSX (e.g. List_RD_AR_Konfirmasi.xlsx)
        mkbtn(br,"\U0001f4e5  Import",self._import_recip,"flat").pack(side="left",fill="x",expand=True,padx=(4,0))
        self._esv=tk.StringVar(); tk.Label(b,textvariable=self._esv,bg=C["panel"],fg=C["tdim"],font=FS,pady=4).pack(anchor="w")
        self._ref_email()

        # S4b Test Auto-Mailer (NEW — separate testing pipeline for existing PDFs)
        # Unchecked by default. When the checkbox is unchecked the "Send" button
        # is disabled, so accidentally clicking through cannot trigger a real
        # send. Distinct from the Process and Watch card auto-send toggles.
        s=card(p); s.pack(fill="x",pady=(0,10))
        tk.Frame(s,bg=C["warn"],height=3).pack(fill="x")
        tk.Label(s,text="\U0001f9ea  Test Auto-Mailer (Existing PDFs)",
                 bg=C["panel"],fg=C["warn"],
                 font=FH,pady=10,padx=14,anchor="w").pack(fill="x")
        b=tk.Frame(s,bg=C["panel"],padx=14,pady=4); b.pack(fill="x")
        tk.Label(b,
                 text="Re-sends every *_letter.pdf already in a chosen folder\n"
                      "(default: ./output). Useful for verifying SMTP / Outlook\n"
                      "setup against PDFs you generated earlier.",
                 bg=C["panel"],fg=C["tdim"],font=FS,
                 justify="left",anchor="w",wraplength=360).pack(fill="x",pady=(2,4))
        # Folder picker — defaults to ./output
        tr=tk.Frame(b,bg=C["panel"]); tr.pack(fill="x",pady=4)
        self._tmfv=tk.StringVar(value=str(OUTPUT_DIR))
        mkentry(tr,self._tmfv).pack(side="left",fill="x",expand=True,ipady=5,padx=(0,4))
        tk.Button(tr,text="\U0001f4c2",command=self._browse_tm,bg=C["border"],fg=C["text"],
                  font=FS,relief="flat",cursor="hand2",padx=8,pady=5).pack(side="right")
        # The arming checkbox — OFF by default per spec.
        self._tm_armed=tk.BooleanVar(value=False)
        self._tm_armed.trace_add("write", lambda *_: self._refresh_tm_button())
        cb=tk.Checkbutton(b,text="Enable test auto-mailer (required to send)",
                          variable=self._tm_armed,bg=C["panel"],fg=C["warn"],
                          selectcolor=C["ebg"],activebackground=C["panel"],
                          font=FB,cursor="hand2")
        cb.pack(anchor="w",pady=(2,2))
        br=tk.Frame(b,bg=C["panel"],pady=8); br.pack(fill="x")
        self._tmbtn=mkbtn(br,"\u25b6  Send All PDFs in Folder",self._test_mailer,"flat")
        self._tmbtn.pack(fill="x")
        self._tmsv=tk.StringVar(value="\u26a0  Disabled — tick the checkbox to arm")
        tk.Label(b,textvariable=self._tmsv,bg=C["panel"],fg=C["tdim"],font=FS,pady=4).pack(anchor="w")
        self._refresh_tm_button()

        # S5 Output
        s=card(p); s.pack(fill="x")
        b=tk.Frame(s,bg=C["panel"],padx=14,pady=10); b.pack(fill="x")
        mkbtn(b,"\U0001f4c1  Open Output Folder",self._open_out,"ghost").pack(fill="x")

    def _log(self,msg,tag=None): self.after(0,lambda:self._con.log(msg,tag))
    def _setstatus(self,t,c=None): self.after(0,lambda:(self._stv.set(t),self._stl.config(bg=c or C["adim"])))

    def _browse_x(self):
        p=filedialog.askopenfilename(title="Select XLSX",filetypes=[("Excel","*.xlsx"),("All","*.*")])
        if p:
            self._xlsx=Path(p); lbl=self._xlsx.name
            if len(lbl)>38: lbl="\u2026"+lbl[-36:]
            self._xlbl.set(lbl)

    def _browse_w(self):
        p=filedialog.askdirectory(title="Select watch folder")
        if p: self._wfv.set(p)

    def _process(self):
        """
        Handle a manual "Generate Letter" click. Each click is its own
        batch — mint a fresh token, reset the counter, then launch the
        subprocess. Since one click only ever processes one file, the
        counter will resolve to 001.
        """
        if self._busy: return
        if not self._xlsx or not self._xlsx.exists():
            messagebox.showwarning("No File","Please select an XLSX file first."); return
        if not GENERATOR.exists():
            messagebox.showerror("Missing",f"Generator script not found:\n{GENERATOR}"); return
        self._busy=True; self._pbtn.config(state="disabled",text="\u23f3  Processing\u2026")
        self._setstatus("Processing\u2026",C["warn"])
        snap=self._xlsx; do_send=self._asend.get()

        # NEW: mint a per-batch token + reset before launching.
        batch_token = uuid.uuid4().hex
        _reset_batch(batch_token, self._log)
        auto_today = self._auto_today.get()

        def done():
            self._busy=False
            self.after(0,lambda:(self._pbtn.config(state="normal",text="\u25b6  Generate Letter"),self._setstatus("Ready")))
            if do_send and EMAILER.exists():
                pdf=OUTPUT_DIR/(snap.stem+"_letter.pdf")
                if pdf.exists(): send_email_bg(pdf,self._log)
                else: self._log(f"[!] PDF not found for email: {snap.stem}_letter.pdf","warning")
        run_pipeline_thread(snap,self._pdef.get(),self._log,done_fn=done,
                            batch_token=batch_token,auto_today=auto_today)

    def _toggle_watch(self):
        if self._wm.running():
            self._wm.stop()
            self._wbtn.config(text="\u25b6  Start Watching",bg=C["accent2"])
            self._wsv.set("\u23f9  Not watching"); self._setstatus("Ready")
        else:
            f=self._wfv.get().strip()
            if not f or not Path(f).is_dir():
                messagebox.showwarning("No Folder","Please enter or browse to a valid folder."); return
            self._wm.start(Path(f),self._wdef.get(),self._wasend.get())
            self._wbtn.config(text="\u23f9  Stop Watching",bg=C["err"])
            self._wsv.set(f"\U0001f441  Watching: {Path(f).name}"); self._setstatus("Watching\u2026",C["accent2"])

    def _edit_def(self):
        dlg=DefaultsDialog(self,"edit"); self.wait_window(dlg)
        # Re-sync the sidebar checkbox + preview with whatever was saved
        # inside the dialog (the user may have toggled it there).
        d = load_def()
        self._auto_today.set(bool(d.get("auto_today", True)))
        self._ref_def()
        self._refresh_preview()
    def _ref_def(self):
        d=load_def()
        # Don't count the meta keys when reporting how many variables are set.
        meta_keys = ("auto_today", "auto_rates", "_comment_schema")
        meaningful = {k:v for k,v in d.items() if k not in meta_keys and v}
        self._dsv.set(f"\u2705  {len(meaningful)} variables configured" if meaningful else "\u26a0  No defaults set \u2014 click Edit")

    def _edit_recip(self):
        dlg=RecipientDialog(self); self.wait_window(dlg); self._ref_email()
    def _edit_smtp(self):
        dlg=SmtpDialog(self); self.wait_window(dlg); self._ref_email()

    def _import_recip(self):
        """
        Bulk-import recipients from an XLSX. Delegates to
        `smartpayer_emailer.py --import-recipients <xlsx>` so we stay
        consistent with the rest of the GUI's "shell out to a script"
        architecture and the same logic is reachable from the CLI.

        On a fresh config the user gets the choice between MERGE (default
        — keep existing keys, overwrite ones with the same NAME) and
        REPLACE (wipe first). MERGE is safer and matches the existing
        recipient-dialog behaviour, so it's the default.
        """
        if not EMAILER.exists():
            messagebox.showerror("Missing",
                f"smartpayer_emailer.py not found:\n{EMAILER}",
                parent=self); return
        path=filedialog.askopenfilename(
            title="Choose recipient list XLSX",
            filetypes=[("Excel","*.xlsx"),("All files","*.*")])
        if not path: return
        # Mode prompt: yes=merge (default), no=replace, cancel=abort
        ans=messagebox.askyesnocancel(
            "Import mode",
            "Choose import mode:\n\n"
            "  YES  \u2192  MERGE   (keep existing; overwrite per row)\n"
            "  NO   \u2192  REPLACE (wipe existing first)\n"
            "  CANCEL \u2192  abort\n\n"
            "Tip: MERGE is the safer choice if you've previously added "
            "recipients manually.",
            parent=self)
        if ans is None: return
        mode="merge" if ans else "replace"
        self._log(f"\U0001f4e5  Importing recipients from {Path(path).name} (mode={mode})\u2026","info")

        def worker():
            try:
                r=subprocess.run([sys.executable,str(EMAILER),
                                  "--import-recipients",path,
                                  "--import-mode",mode],
                                 capture_output=True,text=True,
                                 encoding="utf-8",errors="replace")
                for line in (r.stdout+r.stderr).splitlines():
                    if not line.strip(): continue
                    tag=("success" if "[OK]" in line else
                         "error"   if "[X]"  in line else
                         "warning" if "[!]"  in line else None)
                    self._log(line,tag)
                if r.returncode==0:
                    self.after(0,lambda:(self._ref_email(),
                        messagebox.showinfo("Import complete",
                            "Recipients imported successfully.\n"
                            "Open the Recipients dialog to review.",
                            parent=self)))
                else:
                    self.after(0,lambda:messagebox.showerror(
                        "Import failed",
                        "Import returned a non-zero exit code. "
                        "Check the log panel for details.",
                        parent=self))
            except Exception as e:
                self._log(f"\u274c  Import error: {e}","error")
        threading.Thread(target=worker,daemon=True).start()

    def _browse_tm(self):
        p=filedialog.askdirectory(initialdir=self._tmfv.get() or str(OUTPUT_DIR),
                                  title="Choose folder with generated PDFs")
        if p: self._tmfv.set(p)

    def _refresh_tm_button(self):
        """
        Keep the test-auto-mailer's send button disabled until the user
        ticks the arming checkbox. Also refresh the status label.
        """
        armed=self._tm_armed.get()
        if armed:
            self._tmbtn.config(state="normal")
            self._tmsv.set("\u2705  Armed \u2014 click Send to dispatch every PDF")
        else:
            self._tmbtn.config(state="disabled")
            self._tmsv.set("\u26a0  Disabled \u2014 tick the checkbox to arm")

    def _test_mailer(self):
        """
        Test pipeline: re-send every *_letter.pdf in a chosen folder.
        Completely separate from the Process card's auto-send and from
        the Watch card's per-file emailer.
        """
        if not self._tm_armed.get():
            messagebox.showwarning("Disabled",
                "Tick the 'Enable test auto-mailer' checkbox first.",
                parent=self); return
        if not EMAILER.exists():
            messagebox.showerror("Missing",
                f"smartpayer_emailer.py not found:\n{EMAILER}",
                parent=self); return
        folder=self._tmfv.get().strip()
        if not folder or not Path(folder).is_dir():
            messagebox.showwarning("No folder",
                "Pick a valid folder containing *_letter.pdf files.",
                parent=self); return
        pdfs=sorted(Path(folder).glob("*_letter.pdf"))
        if not pdfs:
            messagebox.showinfo("Nothing to send",
                f"No *_letter.pdf files found in:\n{folder}",
                parent=self); return
        if not messagebox.askyesno("Confirm send",
                f"This will dispatch {len(pdfs)} email(s) with the PDFs in:\n"
                f"  {folder}\n\nProceed?", parent=self):
            return

        self._tmbtn.config(state="disabled",text="\u23f3  Sending\u2026")
        self._setstatus("Test mailer running\u2026",C["warn"])

        def worker():
            try:
                r=subprocess.run([sys.executable,str(EMAILER),
                                  "--send-all",folder],
                                 capture_output=True,text=True,
                                 encoding="utf-8",errors="replace")
                for line in (r.stdout+r.stderr).splitlines():
                    if not line.strip(): continue
                    tag=("success" if "[OK]" in line or "[DONE]" in line else
                         "error"   if "[X]"  in line else
                         "warning" if "[!]"  in line else None)
                    self._log(line,tag)
            except Exception as e:
                self._log(f"\u274c  Test mailer error: {e}","error")
            finally:
                self.after(0,lambda:(
                    self._tmbtn.config(state="normal" if self._tm_armed.get() else "disabled",
                                       text="\u25b6  Send All PDFs in Folder"),
                    self._setstatus("Ready"),
                    self._refresh_tm_button()))
        threading.Thread(target=worker,daemon=True).start()

    def _ref_email(self):
        cfg=load_ecfg()
        if not cfg: self._esv.set("\u26a0  Not configured \u2014 click SMTP Setup"); return
        n=len(cfg.get("recipients",{}))
        method="SMTP" if cfg.get("smtp",{}).get("username") else "Outlook"
        self._esv.set(f"\u2705  {n} recipients \u00b7 Send via {method}")

    def _open_out(self):
        OUTPUT_DIR.mkdir(parents=True,exist_ok=True); p=str(OUTPUT_DIR)
        if sys.platform=="win32": os.startfile(p)
        elif sys.platform=="darwin": subprocess.Popen(["open",p])
        else: subprocess.Popen(["xdg-open",p])

    def _check(self):
        import shutil
        issues,warnings=[],[]
        if not TEMPLATE.exists(): issues.append(f"Template not found: {TEMPLATE.name}")
        if not GENERATOR.exists(): issues.append(f"Generator not found: {GENERATOR.name}")
        for pkg,mod in[("python-docx","docx"),("openpyxl","openpyxl"),("lxml","lxml")]:
            try: __import__(mod)
            except ImportError: issues.append(f"Missing package: {pkg}  (run install_prerequisites.bat)")
        lo=bool(shutil.which("soffice")) or any(Path(c).exists() for c in[
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice"])
        if not lo: warnings.append("LibreOffice not found \u2014 PDF export unavailable (libreoffice.org/download)")
        if not EMAILER.exists(): warnings.append("smartpayer_emailer.py not found \u2014 email features disabled")
        if issues:
            self._log("\u26a0  Setup issues:","warning")
            for i in issues: self._log(f"   \u2022 {i}","warning")
        if warnings:
            self._log("\u26a0  Optional:","warning")
            for w in warnings: self._log(f"   \u2022 {w}","warning")
        if not issues:
            d=load_def()
            self._log(f"\u2705  Ready. {'Defaults loaded.' if d else 'No defaults set \u2014 use Edit Defaults.'}",
                      "success" if not warnings else "info")

if __name__=="__main__":
    App().mainloop()
