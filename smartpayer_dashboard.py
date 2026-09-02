"""
smartpayer_dashboard.py
Success-rate dashboard for the Smartpayer Automation pipeline.

X-axis = RUN / EXECUTION MONTH (taken from each log file's date, e.g.
log-smartpayer-12-06-2026.txt -> June 2026).

Data sources (combined):
  1. Run logs   : log-smartpayer-*.txt  -> the monthly time-series (all rates)
  2. Generated_Letters/<cohort>/        -> current on-disk artefact cross-check
                                           (keyed by BILLING COHORT, not run month)

Run with:
    python -m streamlit run smartpayer_dashboard.py
(or double-click run_dashboard.bat)
"""

import calendar
import re
from datetime import date, datetime
from pathlib import Path

import altair as alt
import pandas as pd
import streamlit as st

BASE_DIR = Path(__file__).resolve().parent
LETTER_OUTPUT_DIR = BASE_DIR / "Generated_Letters"

BLACK = "#141414"
GREY = "#6B7280"
AMBER = "#B45309"
RED = "#B91C1C"
GREEN = "#217A3C"

# Log filename: log-smartpayer-DD-MM-YYYY.txt
LOG_NAME_RE = re.compile(r"log-smartpayer-(\d{2})-(\d{2})-(\d{4})\.txt$", re.IGNORECASE)

# Summary-line patterns (wording matches what the GUI writes to the log).
PATTERNS = {
    "processed":       re.compile(r"(\d+)\s+Item telah Berhasil Diproses"),
    "gen_failed":      re.compile(r"Letter generation gagal:\s*(\d+)"),
    "email_sent":      re.compile(r"Total Surat Berhasil Terkirim:\s*(\d+)"),
    "retry_sent":      re.compile(r"Total Surat Berhasil Terkirim \(Retry\):\s*(\d+)"),
    "email_failed":    re.compile(r"Total Surat Gagal Terkirim:\s*(\d+)"),
    "fail_no_recip":   re.compile(r"Email Penerima Tidak Ditemukan:\s*(\d+)"),
    "fail_other":      re.compile(r"Gagal Terkirim \(Error lainnya\):\s*(\d+)"),
    "fail_pdf_conv":   re.compile(r"Konversi DOCX menjadi PDF gagal:\s*(\d+)"),
    "retry_conv_ok":   re.compile(r"Konversi PDF Berhasil:\s*(\d+)"),
    "retry_conv_fail": re.compile(r"Konversi PDF Gagal:\s*(\d+)"),
}

METRIC_KEYS = list(PATTERNS.keys())


# =============================================================================
# DATA
# =============================================================================

def parse_logs(folder: Path) -> pd.DataFrame:
    """Aggregate every log-smartpayer-*.txt into per-run-month totals."""
    months = {}   # (year, month) -> dict of summed counts + file count
    for path in sorted(folder.glob("log-smartpayer-*.txt")):
        m = LOG_NAME_RE.search(path.name)
        if not m:
            continue
        _dd, mm, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
        key = (yyyy, mm)
        bucket = months.setdefault(
            key, {k: 0 for k in METRIC_KEYS} | {"log_files": 0, "runs": 0})
        bucket["log_files"] += 1
        try:
            text = path.read_text(encoding="utf-8", errors="replace")
        except Exception:
            continue
        bucket["runs"] += len(PATTERNS["processed"].findall(text))
        for k, pat in PATTERNS.items():
            bucket[k] += sum(int(x) for x in pat.findall(text))

    rows = []
    for (yyyy, mm), b in sorted(months.items()):
        processed = b["processed"]
        gen_failed = b["gen_failed"]
        sent_total = b["email_sent"] + b["retry_sent"]
        # net failed after retry recovery, floored at 0
        net_failed = max(0, b["email_failed"] - b["retry_sent"])
        # Genuine generation-time PDF failures only; retry conversion attempts
        # on ~$ Word temp files are noise and are excluded from the headline rate
        # (the Generated_Letters cohort panel reflects the true on-disk state).
        pdf_fail = b["fail_pdf_conv"]

        def rate(num, den):
            return round(100.0 * num / den, 1) if den else None

        rows.append({
            "year": yyyy, "month": mm,
            "month_key": f"{yyyy}-{mm:02d}",
            "label": f"{calendar.month_abbr[mm]} {yyyy}",
            "sort": date(yyyy, mm, 1),
            "log_files": b["log_files"], "runs": b["runs"],
            "processed": processed,
            "gen_failed": gen_failed,
            "email_sent": sent_total,
            "email_failed": net_failed,
            "fail_no_recip": b["fail_no_recip"],
            "fail_other": b["fail_other"],
            "fail_pdf_conv": b["fail_pdf_conv"],
            "retry_recovered": b["retry_sent"],
            # rates
            "email_success_rate": rate(sent_total, processed),
            "gen_success_rate": rate(processed, processed + gen_failed),
            "pdf_success_rate": rate(processed - pdf_fail, processed),
        })
    return pd.DataFrame(rows)


def scan_generated_letters(folder: Path) -> pd.DataFrame:
    """Count *_letter.pdf vs *_with_tables.docx per billing-cohort folder."""
    if not folder.exists():
        return pd.DataFrame()
    cohort_re = re.compile(r"(smartpayer\s+[a-z]+\s+\d{4})", re.IGNORECASE)
    cohorts = {}

    def add(cohort, pdfs, docx):
        c = cohorts.setdefault(cohort, {"pdf": 0, "docx": 0})
        c["pdf"] += pdfs
        c["docx"] += docx

    # Cohort sub-folders
    for sub in folder.iterdir():
        if sub.is_dir():
            pdfs = len([f for f in sub.glob("*_letter.pdf")
                        if not f.name.startswith("~$")])
            docx = len([f for f in sub.glob("*_with_tables.docx")
                        if not f.name.startswith("~$")])
            add(sub.name, pdfs, docx)
    # Loose files at top level
    for f in folder.glob("*"):
        if f.is_file() and not f.name.startswith("~$"):
            m = cohort_re.search(f.name)
            cohort = m.group(1).title() if m else "Uncategorised"
            if f.name.endswith("_letter.pdf"):
                add(cohort, 1, 0)
            elif f.name.endswith("_with_tables.docx"):
                add(cohort, 0, 1)

    rows = []
    for cohort, c in sorted(cohorts.items()):
        docx = c["docx"]
        rows.append({
            "cohort": cohort, "letters_docx": docx, "letters_pdf": c["pdf"],
            "pdf_conversion_rate": round(100.0 * c["pdf"] / docx, 1) if docx else None,
        })
    return pd.DataFrame(rows)


# =============================================================================
# STYLING
# =============================================================================

def inject_css():
    st.markdown(f"""
    <style>
      html, body, [class*="css"] {{ font-family:'GI Sans Text',Arial,sans-serif;
        color:{BLACK}; }}
      .stApp {{ background:#FFFFFF; }}
      .block-container {{ padding-top:3.2rem; max-width:1300px; }}
      .gi-header {{ display:flex; align-items:flex-end; justify-content:space-between;
        border-bottom:3px solid {BLACK}; padding:6px 0 12px 0; margin-bottom:14px; }}
      .gi-brand {{ font-size:26px; font-weight:800; letter-spacing:-0.5px;
        color:{BLACK}; line-height:1.15; padding-top:2px; }}
      .gi-brand small {{ display:block; font-size:12px; font-weight:600; color:{GREY};
        letter-spacing:2px; margin-top:4px; text-transform:uppercase; }}
      .gi-app {{ font-size:15px; font-weight:700; color:{BLACK}; text-align:right; }}
      .gi-app small {{ display:block; font-size:11px; font-weight:500; color:{GREY}; }}
      .gi-section {{ font-size:13px; font-weight:800; letter-spacing:1px;
        text-transform:uppercase; color:{BLACK}; margin:6px 0 2px 0; }}
      div[data-testid="stMetric"] {{ background:#F9FAFB; border:1px solid #E5E7EB;
        border-radius:10px; padding:12px 16px; }}
      [data-testid="stMetricValue"] {{ font-weight:800; }}
      .stButton > button {{ border-radius:8px; font-weight:700;
        border:1.5px solid {BLACK}; background:{BLACK}; color:#FFF; }}
      .stButton > button:hover {{ background:#FFF; color:{BLACK}; }}
      h1,h2,h3,h4 {{ color:{BLACK}; text-align:left; }}
    </style>
    """, unsafe_allow_html=True)


def header():
    st.markdown(
        '<div class="gi-header">'
        '<div class="gi-brand">Godrej<small>Consumer Products</small></div>'
        '<div class="gi-app">Smartpayer Success Dashboard'
        '<small>Monthly pipeline performance</small></div>'
        '</div>', unsafe_allow_html=True)


# =============================================================================
# CHARTS
# =============================================================================

def rate_chart(df, y, title, color):
    d = df.dropna(subset=[y])
    base = alt.Chart(d).encode(
        x=alt.X("label:N", sort=list(df["label"]), title=None,
                axis=alt.Axis(labelAngle=0)))
    bars = base.mark_bar(color=color, size=42).encode(
        y=alt.Y(f"{y}:Q", title="Success %", scale=alt.Scale(domain=[0, 100])),
        tooltip=[alt.Tooltip("label:N", title="Month"),
                 alt.Tooltip(f"{y}:Q", title=title, format=".1f")])
    labels = base.mark_text(dy=-8, fontWeight="bold", color=BLACK).encode(
        y=f"{y}:Q", text=alt.Text(f"{y}:Q", format=".1f"))
    return (bars + labels).properties(height=240, title=title)


def reasons_chart(df):
    reason_map = {
        "fail_no_recip": "No recipient",
        "fail_other": "Other send error",
        "fail_pdf_conv": "PDF conversion fail",
    }
    long = df.melt(
        id_vars=["label", "sort"], value_vars=list(reason_map),
        var_name="reason", value_name="count")
    long["reason"] = long["reason"].map(reason_map)
    long = long[long["count"] > 0]
    if long.empty:
        return None
    chart = alt.Chart(long).mark_bar(size=42).encode(
        x=alt.X("label:N", sort=list(df["label"]), title=None,
                axis=alt.Axis(labelAngle=0)),
        y=alt.Y("count:Q", title="Failed letters", stack="zero"),
        color=alt.Color("reason:N", title="Failure reason",
                        scale=alt.Scale(
                            domain=list(reason_map.values()),
                            range=[GREY, AMBER, RED])),
        tooltip=["label", "reason", "count"])
    return chart.properties(height=260)


def volume_chart(df):
    long = df.melt(
        id_vars=["label"], value_vars=["email_sent", "email_failed"],
        var_name="kind", value_name="count")
    long["kind"] = long["kind"].map(
        {"email_sent": "Sent", "email_failed": "Failed (net)"})
    return alt.Chart(long).mark_bar(size=42).encode(
        x=alt.X("label:N", sort=list(df["label"]), title=None,
                axis=alt.Axis(labelAngle=0)),
        y=alt.Y("count:Q", title="Letters", stack="zero"),
        color=alt.Color("kind:N", title=None,
                        scale=alt.Scale(domain=["Sent", "Failed (net)"],
                                        range=[BLACK, RED])),
        tooltip=["label", "kind", "count"]).properties(height=260)


# =============================================================================
# MAIN
# =============================================================================

def main():
    st.set_page_config(page_title="Smartpayer Dashboard",
                       page_icon="📊", layout="wide")
    inject_css()
    header()

    top = st.columns([4, 1])
    top[0].caption(
        "X-axis is the month each batch was **run** (from the log filename). "
        "Rates are parsed live from `log-smartpayer-*.txt`; the cohort panel "
        "cross-checks against files currently in `Generated_Letters/`.")
    if top[1].button("↻ Reload data", use_container_width=True):
        st.rerun()

    df = parse_logs(BASE_DIR)
    cohorts = scan_generated_letters(LETTER_OUTPUT_DIR)

    if df.empty:
        st.warning(
            "No `log-smartpayer-*.txt` files found in the project folder yet. "
            "Run the automation at least once, then reload this dashboard.")
    else:
        df = df.sort_values("sort").reset_index(drop=True)

        # -- KPI row (all-time totals) -----------------------------------
        processed = int(df["processed"].sum())
        sent = int(df["email_sent"].sum())
        failed = int(df["email_failed"].sum())
        gen_failed = int(df["gen_failed"].sum())
        pdf_fail = int(df["fail_pdf_conv"].sum())
        overall_email = round(100 * sent / processed, 1) if processed else 0
        overall_gen = round(100 * processed / (processed + gen_failed), 1) if (processed + gen_failed) else 0
        overall_pdf = round(100 * (processed - pdf_fail) / processed, 1) if processed else 0

        st.markdown('<div class="gi-section">All-time totals</div>',
                    unsafe_allow_html=True)
        k = st.columns(5)
        k[0].metric("Letters processed", f"{processed:,}")
        k[1].metric("Email send rate", f"{overall_email}%",
                    help="Emails sent (incl. retry) ÷ letters processed")
        k[2].metric("Generation rate", f"{overall_gen}%",
                    help="Letters generated OK ÷ (OK + failed)")
        k[3].metric("PDF conversion rate", f"{overall_pdf}%",
                    help="PDFs produced ÷ letters processed")
        k[4].metric("Failed emails (net)", f"{failed:,}")

        st.divider()

        # -- Success-rate charts -----------------------------------------
        st.markdown('<div class="gi-section">Success rate by run month</div>',
                    unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        c1.altair_chart(rate_chart(df, "email_success_rate", "Email send", BLACK),
                        use_container_width=True)
        c2.altair_chart(rate_chart(df, "gen_success_rate", "Letter generation", GREY),
                        use_container_width=True)
        c3.altair_chart(rate_chart(df, "pdf_success_rate", "PDF conversion", AMBER),
                        use_container_width=True)

        # -- Failure reasons + volume ------------------------------------
        st.divider()
        r1, r2 = st.columns(2)
        r1.markdown('<div class="gi-section">Failure reason breakdown</div>',
                    unsafe_allow_html=True)
        rc = reasons_chart(df)
        if rc is not None:
            r1.altair_chart(rc, use_container_width=True)
        else:
            r1.success("No email failures recorded across all run months.")
        r2.markdown('<div class="gi-section">Email volume (sent vs net failed)</div>',
                    unsafe_allow_html=True)
        r2.altair_chart(volume_chart(df), use_container_width=True)

        # -- Detail table -------------------------------------------------
        st.divider()
        st.markdown('<div class="gi-section">Per-month detail</div>',
                    unsafe_allow_html=True)
        show = df[[
            "label", "runs", "processed", "email_sent", "retry_recovered",
            "email_failed", "fail_no_recip", "fail_other", "fail_pdf_conv",
            "email_success_rate", "gen_success_rate", "pdf_success_rate"]].rename(columns={
                "label": "Month", "runs": "Runs", "processed": "Processed",
                "email_sent": "Sent", "retry_recovered": "Recovered (retry)",
                "email_failed": "Failed (net)", "fail_no_recip": "No recipient",
                "fail_other": "Other error", "fail_pdf_conv": "PDF fail",
                "email_success_rate": "Email %", "gen_success_rate": "Gen %",
                "pdf_success_rate": "PDF %"})
        st.dataframe(show, use_container_width=True, hide_index=True)
        st.download_button(
            "Export CSV", data=show.to_csv(index=False),
            file_name=f"smartpayer-success-{datetime.now():%Y%m%d}.csv")

    # -- Cohort cross-check (always shown) --------------------------------
    st.divider()
    st.markdown('<div class="gi-section">Cross-check · current artefacts by '
                'billing cohort (Generated_Letters/)</div>', unsafe_allow_html=True)
    st.caption("Keyed by billing cohort, not run month. PDF conversion rate here "
               "= letters with a PDF ÷ letters with a DOCX, from files on disk now.")
    if cohorts.empty:
        st.info("No cohort folders / letters found in Generated_Letters/ yet.")
    else:
        cc = cohorts.rename(columns={
            "cohort": "Cohort", "letters_docx": "DOCX", "letters_pdf": "PDF",
            "pdf_conversion_rate": "PDF conversion %"})
        st.dataframe(cc, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
