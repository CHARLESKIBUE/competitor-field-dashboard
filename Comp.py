import html
import math
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from io import BytesIO

import openpyxl
import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from openpyxl.drawing.image import Image as OpenPyxlImage
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt

# --- DATA SOURCES ---
JOTFORM_API_KEY = "6666042ab61ea98f9b378a4a7d"
NEW_SHEET_CSV = "https://docs.google.com/spreadsheets/d/1_k43t4QFDvDpisgFLs1lRykzOdIA4F_AC045SX-u78A/export?format=csv&gid=433581189"
OLD_SHEET_CSV = "https://docs.google.com/spreadsheets/d/1XeWdaYVyLdqfgF00nKBOjZ0OlhtcqRHb6HO0b1lND0Y/export?format=csv&gid=284416453"

PROJECT_COLORS = {"CORE": "#38BDF8", "IB": "#A78BFA", "Unassigned": "#64748B"}
ACTIVITY_COLORS = {"Paid Assets": "#34D399", "Unpaid Assets": "#FBBF24", "Competitor Activities": "#F87171"}

PAID_COL = "Does the store have any paid assets"
UNPAID_COL = "Does the store have any unpaid assets"
COMP_COL = "Are there any Competitor activities in the outlet"
DESC_COL = "Describe the competitor activities"
PAID_TYPES_COL = "Paid Asset Types"
UNPAID_TYPES_COL = "Unpaid Asset Types"
PHOTO_PAID = "Take Photo - Paid Assets"
PHOTO_UNPAID = "Take Photo - Unpaid Assets"
PHOTO_COMP = "Take Photo - Competitor Activities"
PHOTO_COMP_OTHERS = "Take Photo - Competitor Activities (Others)"
PHOTO_COLS = [PHOTO_PAID, PHOTO_UNPAID, PHOTO_COMP, PHOTO_COMP_OTHERS]

# 1. Page Configuration
st.set_page_config(
    page_title="Competitor Activities and Paid Assets",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

:root {
    --bg: #0B1220;
    --surface: #111A2E;
    --surface-2: #16213A;
    --border: #1F2B45;
    --text: #E2E8F0;
    --text-strong: #F8FAFC;
    --muted: #94A3B8;
    --accent: #38BDF8;
    --red: #F87171;
    --green: #34D399;
    --amber: #FBBF24;
}

html, body, [class*="css"], .stApp, .stMarkdown, button, input, textarea, select {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif !important;
}
.stApp { background-color: var(--bg); }
.block-container { padding-top: 2rem !important; padding-bottom: 3rem !important; max-width: 1400px; }
header[data-testid="stHeader"] { background: transparent; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background-color: var(--surface) !important;
    border-right: 1px solid var(--border) !important;
}
section[data-testid="stSidebar"] .sb-title {
    font-size: 0.72rem; font-weight: 700; letter-spacing: 0.08em; text-transform: uppercase;
    color: var(--muted); margin: 18px 0 6px 0;
}
section[data-testid="stSidebar"] [data-testid="stDateInput"] label p { font-size: 0.75rem; color: var(--muted); }
section[data-testid="stSidebar"] .sb-brand {
    font-size: 1.05rem; font-weight: 800; color: var(--text-strong); letter-spacing: -0.01em;
}
section[data-testid="stSidebar"] .sb-brand-sub { font-size: 0.8rem; color: var(--muted); margin-bottom: 8px; }

/* Header */
.dash-header { display: flex; justify-content: space-between; align-items: flex-end; gap: 16px; flex-wrap: wrap;
    padding-bottom: 18px; border-bottom: 1px solid var(--border); margin-bottom: 22px; }
.dash-eyebrow { font-size: 0.72rem; font-weight: 700; letter-spacing: 0.12em; text-transform: uppercase; color: var(--accent); }
.dash-title { font-size: 1.9rem; font-weight: 800; color: var(--text-strong); letter-spacing: -0.02em; line-height: 1.15; margin-top: 4px; }
.dash-meta { font-size: 0.82rem; color: var(--muted); text-align: right; line-height: 1.6; }
.dash-meta b { color: var(--text); font-weight: 600; }

/* KPI tiles */
.kpi-row { display: grid; grid-template-columns: repeat(5, minmax(0, 1fr)); gap: 12px; margin-bottom: 26px; }
@media (max-width: 1200px) { .kpi-row { grid-template-columns: repeat(3, minmax(0, 1fr)); } }
.kpi { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 16px 18px;
    border-top: 3px solid var(--kpi-accent, var(--accent)); }
.kpi-label { white-space: nowrap; overflow: hidden; text-overflow: ellipsis; font-size: 0.72rem; font-weight: 600; letter-spacing: 0.06em; text-transform: uppercase; color: var(--muted); }
.kpi-value { font-size: 1.75rem; font-weight: 800; color: var(--text-strong); margin-top: 6px; line-height: 1.1; font-variant-numeric: tabular-nums; }
.kpi-sub { font-size: 0.78rem; color: var(--muted); margin-top: 4px; }

/* Section headings */
.section-title { font-size: 1.02rem; font-weight: 700; color: var(--text-strong); margin: 6px 0 2px 0; }
.section-sub { font-size: 0.8rem; color: var(--muted); margin-bottom: 10px; }

/* Cards */
.card { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 16px 18px; margin-bottom: 14px; }
.card-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; flex-wrap: wrap; margin-bottom: 10px; }
.card-outlet { font-size: 1.02rem; font-weight: 700; color: var(--text-strong); }
.card-rep { font-size: 0.8rem; color: var(--muted); margin-top: 2px; }
.chip { display: inline-block; font-size: 0.7rem; font-weight: 600; padding: 3px 9px; border-radius: 999px;
    border: 1px solid var(--border); color: var(--muted); background: var(--surface-2); margin-left: 6px; white-space: nowrap; }
.chip-core { color: #7DD3FC; border-color: rgba(56,189,248,0.35); background: rgba(56,189,248,0.08); }
.chip-ib { color: #C4B5FD; border-color: rgba(167,139,250,0.35); background: rgba(167,139,250,0.08); }
.promo-list { list-style: none; padding: 0; margin: 0; }
.promo-list li { padding: 7px 0; border-top: 1px solid var(--border); font-size: 0.88rem; color: var(--text);
    display: flex; flex-wrap: wrap; align-items: center; gap: 8px; }
.promo-list li:first-child { border-top: none; }
.p-name { font-weight: 600; color: var(--text-strong); }
.p-now { font-weight: 700; color: var(--text-strong); font-variant-numeric: tabular-nums; }
.p-was { color: #64748B; text-decoration: line-through; font-variant-numeric: tabular-nums; }
.badge { font-size: 0.7rem; font-weight: 700; padding: 2px 8px; border-radius: 6px; white-space: nowrap; }
.badge-red { color: #FCA5A5; background: rgba(248,113,113,0.12); border: 1px solid rgba(248,113,113,0.35); }
.badge-muted { color: var(--muted); background: var(--surface-2); border: 1px solid var(--border); }
.badge-green { color: #6EE7B7; background: rgba(52,211,153,0.10); border: 1px solid rgba(52,211,153,0.35); }
.badge-amber { color: #FCD34D; background: rgba(251,191,36,0.10); border: 1px solid rgba(251,191,36,0.35); }
.empty-note { font-size: 0.85rem; color: var(--muted); }

.photo-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(140px, 1fr)); gap: 8px; margin-top: 12px; }
.photo-grid img { width: 100%; height: 140px; object-fit: cover; border-radius: 8px; border: 1px solid var(--border); background: var(--surface-2); }

.gallery-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(220px, 1fr)); gap: 14px; }
.g-item { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; overflow: hidden; }
.g-item img { width: 100%; height: 200px; object-fit: cover; display: block; background: var(--surface-2); }
.g-cap { padding: 10px 12px; }
.g-outlet { font-size: 0.86rem; font-weight: 700; color: var(--text-strong); }
.g-meta { font-size: 0.74rem; color: var(--muted); margin-top: 2px; }

/* Tabs */
button[data-baseweb="tab"] { font-size: 0.9rem !important; font-weight: 600 !important; color: var(--muted) !important; padding: 10px 4px !important; margin-right: 18px; }
button[data-baseweb="tab"][aria-selected="true"] { color: var(--text-strong) !important; }
div[data-baseweb="tab-highlight"] { background-color: var(--accent) !important; }
div[data-baseweb="tab-border"] { background-color: var(--border) !important; }

/* Dataframes */
div[data-testid="stDataFrame"] { border: 1px solid var(--border); border-radius: 10px; overflow: hidden; }

@media (max-width: 640px) {
    .dash-title { font-size: 1.45rem; }
    .dash-meta { text-align: left; }
    .kpi-row { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    .kpi-value { font-size: 1.5rem; }
}
</style>
""", unsafe_allow_html=True)


# --- TEXT & PRICE PARSING ---
MONTHS = {m: i for i, m in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}
MAX_VALID_DEPTH = 70.0
MIN_VALID_PRICE = 10.0

_NUM_DATE_RE = re.compile(r"(?<![\d.])(\d{1,2})\s*[/\-.]\s*(\d{1,2})(?:\s*[/\-.]\s*(\d{1,4}))?(?![\d])")
_TXT_DATE_RE = re.compile(r"(?<![\d])(\d{1,2})\s*(?:st|nd|rd|th)?\s*(?:of\s+)?(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\.?(?:\s*,?\s*(\d{2,4}))?", re.I)
_PACK_RE = re.compile(
    r"\d+(?:\.\d+)?\s*\+\s*\d+(?:\s*(?:pk|pack|packs|pcs|s)\b)?"
    r"|\d+\s*[*x×]\s*\d+(?:\.\d+)?\s*(?:m|cm|mm|g|kg|ml)?\b"
    r"|\b[x×]\s*\d+\b"
    r"|\d+\s*/\s*\d+\s*(?:g|gm|ml|kg|l)\b"
    r"|\d+(?:\.\d+)?\s*['’]?\s*(?:pk|pks|ppk|pck|pack|packs|ack|pkt|pcs|pc|pieces|rolls?|sheets?|ply|s|g|gm|gms|grams?|kg|kgs|ml|mls|l|ltr|ltrs|litres?|liters?|cm|mm|m|in1|in\s*1|ct|count|sachets?|bags?|wipes|tabs|tablets|caps|k)\b"
    r"|\d+\s*\*",
    re.I,
)
_WAS_RE = r"(?:was|from|frm|fr|fro|rrp|before|normal(?:ly)?|initial(?:ly)?|regular|usual(?:ly)?|old\s*price)"
_NOW_RE = r"(?:@|at|now|mow|nw|selling(?:\s*at)?|offer(?:\s*price)?|promo(?:\s*price)?|price|retailing(?:\s*at)?|going\s*for|to)"
_TOKEN_RE = re.compile(
    rf"(?P<was>\b{_WAS_RE}\b\s*[:\-]?)|(?P<now>(?:@|\b{_NOW_RE}\b)\s*[:\-]?)|(?P<num>\d[\d,]*(?:\.\d+)?)|(?P<word>[A-Za-z§][\w§'’&.]*)",
    re.I,
)
_FILLER = {
    "ksh", "kshs", "kes", "sh", "shs", "bob", "only", "each", "per", "offer", "offers", "ends", "end", "ending", "date", "dates",
    "on", "till", "until", "valid", "promo", "discount", "off", "for", "are",
    "while", "stocks", "stock", "last", "lasts", "wsl", "ksh.", "sh.", "selling", "it", "a", "price", "prices", "down",
    "reduced", "reduction", "start", "starts", "started", "starting", "e.g", "eg", "was", "now", "save", "saving",
    "savings", "new", "price:", "rsp", "untill", "untl", "unti", "stock's", "stocks'", "data", "white", "endsaug",
    "jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "sept", "oct", "nov", "dec", "january", "february",
    "march", "april", "june", "july", "august", "september", "october", "november", "december",
}


def _mask_packs(text):
    keep = []

    def sub(m):
        keep.append(m.group(0))
        return f" §{len(keep) - 1}§ "

    return _PACK_RE.sub(sub, text), keep


def _unmask(text, keep):
    return re.sub(r"§(\d+)§", lambda m: keep[int(m.group(1))], text)


def _to_date(day, month, year, ref):
    """Build a plausible offer date; fixes missing or mistyped years using the submission date."""
    try:
        day, month = int(day), int(month)
    except (TypeError, ValueError):
        return None
    if not (1 <= month <= 12 and 1 <= day <= 31):
        return None
    ref = ref if isinstance(ref, (pd.Timestamp, datetime)) and pd.notna(ref) else pd.Timestamp(datetime.now())
    ref = pd.Timestamp(ref).normalize()
    cands = []
    y = str(year).strip() if year not in (None, "") else ""
    if len(y) == 4 and y.isdigit():
        cands.append(int(y))
    elif len(y) == 2 and y.isdigit():
        cands.append(2000 + int(y))
    cands += [ref.year, ref.year + 1]
    for yr in cands:
        last_day = pd.Timestamp(year=yr, month=month, day=1).days_in_month
        d = pd.Timestamp(year=yr, month=month, day=min(day, last_day))
        if ref - pd.Timedelta(days=1) <= d <= ref + pd.Timedelta(days=183):
            return d
    return None


def _extract_dates(text, ref):
    """Remove dates from text; return (clean_text, [dates], while_stocks_last)."""
    found = []
    wsl = bool(re.search(r"\bwsl\b|while\s*stocks?\s*last", text, re.I))

    def num_sub(m):
        d = _to_date(m.group(1), m.group(2), m.group(3), ref)
        if d is not None:
            found.append(d)
        return " "

    def txt_sub(m):
        d = _to_date(m.group(1), MONTHS[m.group(2).lower()[:3]], m.group(3), ref)
        if d is not None:
            found.append(d)
        return " "

    text = _TXT_DATE_RE.sub(txt_sub, text)
    text = _NUM_DATE_RE.sub(num_sub, text)
    text = re.sub(r"(?i)\b(?:save|saving|savings)\s*(?:ksh\.?|kes|sh)?\s*\d[\d,]*(?:\.\d+)?\s*/?-?", " ", text)
    text = re.sub(r"(?i)(?<![\d])\d{1,2}\s*(?:st|nd|rd|th)\b", " ", text)
    text = re.sub(r"\(\s*\)|\(\s*$|^\s*\)", " ", text)
    return text, found, wsl


def _num(s):
    try:
        return float(s.replace(",", ""))
    except ValueError:
        return None


def _resolve(item):
    """Turn collected was/now/bare numbers into prices and a validity flag."""
    was, now, bare = item["was"], item["now"], item["bare"]
    if len(now) >= 2 and not was:
        was, now = [max(now)], [min(now)]
    if not was and now and bare:
        was = [b for b in bare if b > now[0]][:1]
    if was and not now and bare:
        now = [b for b in bare if b < was[0]][:1]
    if not was and not now and len(bare) >= 2:
        a, b = bare[0], bare[1]
        was, now = [max(a, b)], [min(a, b)]
    if not was and not now and len(bare) == 1:
        now = bare[:1]
    old_p = was[0] if was else None
    new_p = now[0] if now else None
    status, reason = "no_price", "No price captured"
    if new_p is not None and old_p is None:
        status, reason = "single", "Promo price only (no original price)"
    elif old_p is not None and new_p is None:
        status, reason = "single", "Original price only (no promo price)"
    elif old_p is not None and new_p is not None:
        if new_p >= old_p:
            status, reason = "review", "Promo price is not lower than original price"
        elif new_p < MIN_VALID_PRICE:
            status, reason = "review", f"Promo price below KSh {MIN_VALID_PRICE:.0f}"
        elif (old_p - new_p) / old_p * 100 > MAX_VALID_DEPTH:
            status, reason = "review", f"Depth above {MAX_VALID_DEPTH:.0f}% — likely a typing error"
        else:
            status, reason = "ok", ""
    return old_p, new_p, status, reason


def _parse_line(line):
    """Split one cleaned line into product items with raw price numbers."""
    masked, keep = _mask_packs(line)
    masked = re.sub(r"(\d)([A-Za-z@])", r"\1 \2", masked)
    masked = re.sub(r"([A-Za-z@:])(\d)", r"\1 \2", masked)
    items, cur, pending = [], None, None

    def new_item():
        return {"words": [], "was": [], "now": [], "bare": []}

    cur = new_item()
    for m in _TOKEN_RE.finditer(masked):
        kind = m.lastgroup
        tok = m.group(0)
        if kind == "word":
            w = tok.lower().strip(".")
            if w in _FILLER:
                continue
            has_price = bool(cur["was"] or cur["now"] or cur["bare"])
            if has_price:
                items.append(cur)
                cur = new_item()
            cur["words"].append(tok)
            pending = None
        elif kind == "was":
            pending = "was"
        elif kind == "now":
            if tok.strip().lower().startswith("to") and not (cur["was"] or cur["now"]):
                pending = None
                continue
            pending = "now"
        elif kind == "num":
            v = _num(tok)
            if v is None:
                continue
            if pending == "was":
                cur["was"].append(v)
            elif pending == "now":
                cur["now"].append(v)
            else:
                cur["bare"].append(v)
            pending = None
    items.append(cur)
    out, seen_price = [], False
    for it in items:
        has_num = bool(it["was"] or it["now"] or it["bare"])
        if seen_price and not has_num:
            continue
        seen_price = seen_price or has_num
        words = list(it["words"])
        if it["was"] and it["now"] and it["bare"]:
            words += [f"{v:g}" for v in it["bare"]]
            it["bare"] = []
        name = _unmask(" ".join(words), keep)
        name = re.sub(r"\s+", " ", name).strip(" -_*•:;,.@/()")
        it["product"] = name
        if name or it["was"] or it["now"] or it["bare"]:
            out.append(it)
    return out


def parse_price_discounts(text, submission_date=None):
    """Parse a rep's free-text competitor note into promo lines with validated prices and offer end dates."""
    if not isinstance(text, str) or not text.strip():
        return []
    ref = submission_date if isinstance(submission_date, (pd.Timestamp, datetime)) and pd.notna(submission_date) else None
    results = []
    for raw in re.split(r"[\r\n]+|(?<=\))\s*[,;]\s*|\s*;\s*", text):
        raw_clean = raw.strip().strip("-*•_ ").strip()
        if not raw_clean:
            continue
        body, dates, wsl = _extract_dates(raw_clean, ref)
        items = _parse_line(body)
        end_date = max(dates) if dates else None
        if not items:
            if end_date is not None:
                for prev in reversed(results):
                    if prev["end_date"] is None:
                        prev["end_date"] = end_date
                    else:
                        break
            continue
        for it in items:
            only_prices = not it["product"]
            target = None
            if only_prices and results:
                prev = results[-1]
                needs = (it["was"] and prev["_was"] == []) or (it["now"] and prev["_now"] == []) or (it["bare"] and not (prev["_was"] and prev["_now"]))
                if needs:
                    target = prev
            if target is not None:
                target["_was"] += it["was"]
                target["_now"] += it["now"]
                target["_bare"] += it["bare"]
                if end_date is not None and target["end_date"] is None:
                    target["end_date"] = end_date
                continue
            if only_prices:
                continue
            results.append({
                "product": it["product"], "_was": list(it["was"]), "_now": list(it["now"]),
                "_bare": list(it["bare"]), "end_date": end_date, "wsl": wsl, "raw_line": raw_clean,
            })
    out = []
    for r in results:
        old_p, new_p, status, reason = _resolve({"was": r["_was"], "now": r["_now"], "bare": r["_bare"]})
        valid = status == "ok"
        disc = (old_p - new_p) if valid else None
        end = r["end_date"]
        duration = None
        if end is not None and ref is not None:
            duration = int((end - pd.Timestamp(ref).normalize()).days) + 1
            if duration < 1:
                duration = None
        out.append({
            "product": r["product"], "old_price": old_p, "new_price": new_p,
            "discount_amount": disc, "discount_pct": (disc / old_p * 100) if valid else None,
            "start_date": pd.Timestamp(ref).strftime("%d/%m/%Y") if ref is not None else None,
            "end_date": end.strftime("%d/%m/%Y") if end is not None else ("While stocks last" if r["wsl"] else None),
            "end_ts": end, "duration_days": duration, "status": status, "reason": reason, "raw_line": r["raw_line"],
        })
    return out


# --- PHOTO LINKS ---
def drive_display_url(url):
    """Convert a Google Drive file link into a thumbnail link that renders in <img> tags and downloads reliably."""
    url = str(url).strip()
    m = re.search(r"(?:[?&]id=|/file/d/|/open\?id=)([A-Za-z0-9_-]{10,})", url)
    if "drive.google.com" in url and m:
        return f"https://drive.google.com/thumbnail?id={m.group(1)}&sz=w1200"
    return url


def img_src(url):
    url = str(url).strip()
    if "jotform.com" in url:
        return f"{url}{'&' if '?' in url else '?'}apiKey={JOTFORM_API_KEY}"
    return drive_display_url(url)


def split_urls(*vals):
    """Collect photo links from one or more cells (newline or comma separated)."""
    urls = []
    for val in vals:
        if val is None or (isinstance(val, float) and pd.isna(val)):
            continue
        for u in re.split(r"[\n\r,]+", str(val)):
            u = u.strip()
            if u.lower().startswith("http"):
                urls.append(drive_display_url(u))
    return urls


def row_photos(row, cols):
    urls = []
    for c in cols:
        urls.extend(split_urls(row.get(c)))
    return urls


# --- IMAGE DOWNLOADS FOR EXPORTS ---
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_and_resize_image(url):
    if not url:
        return None
    try:
        res = requests.get(img_src(url), headers={"User-Agent": "Mozilla/5.0"}, timeout=8)
        if res.status_code == 200:
            img = Image.open(BytesIO(res.content))
            img.thumbnail((600, 600))
            if img.mode not in ("RGB", "RGBA"):
                img = img.convert("RGB")
            out = BytesIO()
            img.save(out, format="PNG")
            return out.getvalue()
    except Exception:
        pass
    return None


def fetch_all_images_parallel(urls):
    unique_urls = list({u for u in urls if u})
    url_to_bytes = {}
    with ThreadPoolExecutor(max_workers=10) as executor:
        futures = {executor.submit(fetch_and_resize_image, u): u for u in unique_urls}
        for future in as_completed(futures):
            try:
                data = future.result()
                if data:
                    url_to_bytes[futures[future]] = data
            except Exception:
                pass
    return url_to_bytes


# --- REPORT EXPORTS ---
BRAND_NAVY = "00274E"
BRAND_HEAD = "005083"
BRAND_BLUE = "0079C2"
BRAND_TEAL = "10B895"
BRAND_CYAN = "00BEF2"
BRAND_RULE = "CFD8E4"
BRAND_ZEBRA = "F3F6FA"

APP_DIR = os.path.dirname(os.path.abspath(__file__))


def find_template():
    """Locate the Smollan PowerPoint template in the repo (templates/ folder or the main folder)."""
    for path in [os.path.join(APP_DIR, "templates", "Smollan_Core_Template.pptx"),
                 os.path.join(APP_DIR, "Smollan_Core_Template.pptx")]:
        if os.path.exists(path):
            return path
    for folder in [os.path.join(APP_DIR, "templates"), APP_DIR]:
        if os.path.isdir(folder):
            for name in sorted(os.listdir(folder)):
                if name.lower().endswith(".pptx") and "template" in name.lower():
                    return os.path.join(folder, name)
    return None


def _xl_title_block(ws, title, meta_line, width_cols):
    last = get_column_letter(width_cols)
    ws.merge_cells(f"A1:{last}1")
    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=BRAND_NAVY)
    ws["A1"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[1].height = 34
    ws.merge_cells(f"A2:{last}2")
    ws["A2"] = meta_line
    ws["A2"].font = Font(name="Calibri", size=10, color="FFFFFF")
    ws["A2"].fill = PatternFill("solid", fgColor=BRAND_HEAD)
    ws["A2"].alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[2].height = 20
    ws.sheet_view.showGridLines = False


def _xl_section(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text.upper())
    c.font = Font(name="Calibri", size=11, bold=True, color=BRAND_HEAD)
    return row + 1


def _xl_small_table(ws, row, col, headers, rows, formats=None, text_cols=(0,)):
    """Write a compact styled table; returns the next free row."""
    formats = formats or {}
    thin = Side(style="thin", color=BRAND_RULE)
    for j, h in enumerate(headers):
        c = ws.cell(row=row, column=col + j, value=h)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BRAND_NAVY)
        c.alignment = Alignment(horizontal="left" if j in text_cols else "right", vertical="center", indent=1 if (j in text_cols and j > 0) else 0)
    for i, r in enumerate(rows, start=1):
        for j, v in enumerate(r):
            c = ws.cell(row=row + i, column=col + j, value=v)
            c.font = Font(name="Calibri", size=10)
            c.border = Border(bottom=thin)
            c.alignment = Alignment(horizontal="left" if j in text_cols else "right", indent=1 if (j in text_cols and j > 0) else 0)
            if i % 2 == 0:
                c.fill = PatternFill("solid", fgColor=BRAND_ZEBRA)
            if j in formats:
                c.number_format = formats[j]
    return row + len(rows) + 2


def _xl_data_sheet(wb, title, df, col_formats, col_widths, link_cols=(), meta_line=""):
    ws = wb.create_sheet(title=title[:31])
    ncols = max(len(df.columns), 1)
    _xl_title_block(ws, title, meta_line, ncols)
    header_row = 4
    for j, h in enumerate(df.columns, start=1):
        c = ws.cell(row=header_row, column=j, value=h)
        c.font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BRAND_NAVY)
        c.alignment = Alignment(vertical="center", wrap_text=True)
    for i, r in enumerate(df.itertuples(index=False), start=header_row + 1):
        for j, v in enumerate(r, start=1):
            if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NaT:
                continue
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            c = ws.cell(row=i, column=j)
            h = df.columns[j - 1]
            if h in link_cols and isinstance(v, str) and v.startswith("http"):
                c.value = "View photo"
                c.hyperlink = v
                c.font = Font(name="Calibri", size=10, color=BRAND_BLUE, underline="single")
            else:
                c.value = v
                c.font = Font(name="Calibri", size=10)
            if h in col_formats:
                c.number_format = col_formats[h]
    for j, h in enumerate(df.columns, start=1):
        ws.column_dimensions[get_column_letter(j)].width = col_widths.get(h, 14)
    last_row = header_row + max(len(df), 1)
    if len(df):
        ref = f"A{header_row}:{get_column_letter(ncols)}{last_row}"
        tbl = Table(displayName=re.sub(r"[^A-Za-z0-9]", "", title)[:30] or "Data", ref=ref)
        tbl.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
        ws.add_table(tbl)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.row_dimensions[header_row].height = 30
    return ws, header_row


def generate_excel_report(ctx):
    meta_line = ctx["meta_line"]
    priced, promo, paid_u, unpaid_u = ctx["priced"], ctx["promo"], ctx["paid_unique"], ctx["unpaid_unique"]
    paid_src, unpaid_src = ctx["paid_df"], ctx["unpaid_df"]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    _xl_title_block(ws, "Competitor Activities & Paid Assets Report", meta_line, 10)
    for col, w in zip("ABCDEFGHIJ", [32, 13, 13, 11, 14, 24, 12, 12, 12, 12]):
        ws.column_dimensions[col].width = w

    # KPI band
    kpi_items = [
        ("Unique paid assets", len(paid_u), "#,##0"),
        ("Unique unpaid assets", len(unpaid_u), "#,##0"),
        ("Competitor outlets", ctx["comp_outlets"], "#,##0"),
        ("Price cuts captured", len(priced), "#,##0"),
        ("Avg promo depth", (priced["Depth %"].mean() / 100) if len(priced) else None, "0.0%"),
        ("Deepest price cut", (priced["Depth %"].max() / 100) if len(priced) else None, "0.0%"),
    ]
    kpi_cols = [1, 2, 3, 6, 7, 8]
    for (label, val, fmt), col in zip(kpi_items, kpi_cols):
        lc = ws.cell(row=4, column=col, value=label.upper())
        lc.font = Font(name="Calibri", size=8, bold=True, color="64748B")
        lc.alignment = Alignment(wrap_text=True, vertical="bottom")
        vc = ws.cell(row=5, column=col, value=val if val is not None else "—")
        vc.font = Font(name="Calibri", size=18, bold=True, color=BRAND_NAVY)
        vc.number_format = fmt
        vc.alignment = Alignment(horizontal="left")
        vc.border = Border(bottom=Side(style="medium", color=BRAND_TEAL))
    ws.row_dimensions[4].height = 26
    ws.row_dimensions[5].height = 30

    # Left column: by project + paid asset types
    r = _xl_section(ws, 7, 1, "By project")
    projects = sorted(set(paid_u["Project"].dropna()) | set(unpaid_u["Project"].dropna()) | set(ctx["comp_df"]["Select Your Project"].dropna()))
    proj_rows = []
    for p in projects:
        proj_rows.append([
            p,
            int((paid_u["Project"] == p).sum()),
            int((unpaid_u["Project"] == p).sum()),
        ])
    r = _xl_small_table(ws, r, 1, ["Project", "Unique paid", "Unique unpaid"], proj_rows, {1: "#,##0", 2: "#,##0"})

    r = _xl_section(ws, r, 1, "Unique paid assets by type")
    if len(paid_u):
        tc = paid_u.groupby("Asset").agg(n=("key", "size"), o=("key", "nunique")).sort_values("n", ascending=False)
        rows = [[a, int(x.n), int(x.o)] for a, x in tc.iterrows()]
    else:
        rows = []
    r_left = _xl_small_table(ws, r, 1, ["Asset type", "Unique assets", "Outlets"], rows, {1: "#,##0", 2: "#,##0"})

    # Right column: competitor brands
    r = _xl_section(ws, 7, 6, "Most active competitor brands")
    if len(priced):
        b = priced.groupby("Brand").agg(n=("Product", "size"), avg=("Depth %", "mean"), mx=("Depth %", "max"), o=("Outlet", "nunique")) \
            .sort_values("n", ascending=False).head(15)
        rows = [[br, int(x.n), int(x.o), x.avg / 100, x.mx / 100] for br, x in b.iterrows()]
    else:
        rows = []
    r_right = _xl_small_table(ws, r, 6, ["Brand", "Price cuts", "Outlets", "Avg depth", "Deepest"], rows,
                              {1: "#,##0", 2: "#,##0", 3: "0.0%", 4: "0.0%"})

    # Full width: deepest price cuts
    r = _xl_section(ws, max(r_left, r_right), 1, "Top 15 deepest competitor price cuts")
    if len(priced):
        top = priced.sort_values("Depth %", ascending=False).head(15)
        rows = [[x["Product"], x["Was (KSh)"], x["Now (KSh)"], x["Depth %"] / 100, x["Brand"], x["Outlet"],
                 x["Date"].to_pydatetime() if pd.notna(x["Date"]) else None, x["Ends"] or ""]
                for _, x in top.iterrows()]
    else:
        rows = []
    _xl_small_table(ws, r, 1, ["Product", "Was (KSh)", "Now (KSh)", "Depth", "Brand", "Outlet", "Date", "Ends"], rows,
                    {1: "#,##0", 2: "#,##0", 3: "0.0%", 6: "dd mmm yyyy"}, text_cols=(0, 4, 5, 6, 7))
    ws.freeze_panes = "A4"

    # Competitor price cuts (all parsed lines)
    comp_cols = ["Date", "Project", "Outlet", "Field Rep", "Brand", "Product", "Was (KSh)", "Now (KSh)", "Saving (KSh)", "Depth", "Ends", "Days Left", "Photo"]
    if len(priced):
        comp_tbl = priced.assign(**{"Depth": priced["Depth %"] / 100}).sort_values("Date", ascending=False)[comp_cols]
    else:
        comp_tbl = pd.DataFrame(columns=comp_cols)
    ws_c, hr = _xl_data_sheet(
        wb, "Competitor Price Cuts", comp_tbl,
        {"Date": "dd mmm yyyy", "Was (KSh)": "#,##0", "Now (KSh)": "#,##0", "Saving (KSh)": "#,##0", "Depth": "0.0%", "Days Left": "0"},
        {"Date": 13, "Project": 9, "Outlet": 24, "Field Rep": 26, "Brand": 13, "Product": 38, "Was (KSh)": 11, "Now (KSh)": 11,
         "Saving (KSh)": 12, "Depth": 10, "Ends": 13, "Days Left": 10, "Photo": 12},
        link_cols=("Photo",), meta_line=meta_line,
    )
    if len(comp_tbl):
        dcol = get_column_letter(list(comp_tbl.columns).index("Depth") + 1)
        ws_c.conditional_formatting.add(f"{dcol}{hr + 1}:{dcol}{hr + len(comp_tbl)}",
                                        DataBarRule(start_type="num", start_value=0, end_type="max", color="F87171"))

    # Unique asset sheets
    image_cache = fetch_all_images_parallel([u for u in ctx["paid_photos"].values() if u])
    for title, uniq, src, photos, embed in [
        ("Paid Assets", paid_u, paid_src, ctx["paid_photos"], True),
        ("Unpaid Assets", unpaid_u, unpaid_src, ctx["unpaid_photos"], False),
    ]:
        rows = []
        for i, a in uniq.iterrows():
            row = src.loc[a["idx"]]
            rows.append({
                "Outlet": a["Outlet"], "Project": a["Project"], "Asset": a["Asset"], "Visits Seen": int(a["Visits"]),
                "First Seen": a["First Seen"], "Last Seen": a["Last Seen"], "Last Reported By": row["Select Your Name"],
                "Photo": img_src(photos[i]) if photos.get(i) else None, **({"Image": None} if embed else {}),
            })
        cols = ["Outlet", "Project", "Asset", "Visits Seen", "First Seen", "Last Seen", "Last Reported By", "Photo"] + (["Image"] if embed else [])
        tdf = pd.DataFrame(rows, columns=cols)
        ws_a, hr = _xl_data_sheet(
            wb, title, tdf,
            {"First Seen": "dd mmm yyyy", "Last Seen": "dd mmm yyyy", "Visits Seen": "0"},
            {"Outlet": 26, "Project": 9, "Asset": 30, "Visits Seen": 11, "First Seen": 13, "Last Seen": 13, "Last Reported By": 26, "Photo": 12, "Image": 16},
            link_cols=("Photo",), meta_line=meta_line,
        )
        if embed and len(tdf):
            img_col = len(cols)
            for n, (i, _) in enumerate(uniq.iterrows(), start=hr + 1):
                url = photos.get(i)
                ws_a.row_dimensions[n].height = 66
                if url in image_cache:
                    try:
                        xl_img = OpenPyxlImage(BytesIO(image_cache[url]))
                        ratio = xl_img.width / max(xl_img.height, 1)
                        xl_img.height = 80
                        xl_img.width = int(80 * ratio)
                        if xl_img.width > 110:
                            xl_img.width, xl_img.height = 110, int(110 / ratio)
                        ws_a.add_image(xl_img, f"{get_column_letter(img_col)}{n}")
                    except Exception:
                        pass
            for n in range(hr + 1, hr + len(tdf) + 1):
                for j in range(1, img_col):
                    ws_a.cell(row=n, column=j).alignment = Alignment(vertical="center")

    # Raw data
    raw = ctx["raw_df"].drop(columns=[c for c in ("Parsed_Date", "Formatted_Timestamp") if c in ctx["raw_df"].columns])
    raw = raw.assign(**{"Submission Date": ctx["raw_df"]["Parsed_Date"]})
    _xl_data_sheet(wb, "Raw Data", raw, {"Submission Date": "dd mmm yyyy hh:mm"},
                   {c: (40 if "Photo" in c or c == DESC_COL else 18) for c in raw.columns}, meta_line=meta_line)

    for sh in wb.worksheets:
        sh.sheet_properties.tabColor = BRAND_NAVY if sh.title == "Summary" else BRAND_BLUE
        sh.page_setup.orientation = "landscape"
        sh.page_setup.fitToWidth = 1
        sh.page_setup.fitToHeight = 0
        sh.sheet_properties.pageSetUpPr.fitToPage = True

    out = BytesIO()
    wb.save(out)
    return out.getvalue()


# --- POWERPOINT (SMOLLAN CORE TEMPLATE) ---
def _rgb(hexstr):
    return RGBColor.from_string(hexstr)


def _template_prs():
    template_path = find_template()
    if template_path:
        prs = Presentation(template_path)
        sld_ids = prs.slides._sldIdLst
        for sld_id in list(sld_ids):
            prs.part.drop_rel(sld_id.rId)
            sld_ids.remove(sld_id)
        return prs, True
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    return prs, False


def _layout(prs, name, fallback=6):
    for l in prs.slide_layouts:
        if l.name == name:
            return l
    return prs.slide_layouts[min(fallback, len(prs.slide_layouts) - 1)]


def _ph(slide, idx):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            return ph
    return None


def _set_ph(slide, idx, text):
    ph = _ph(slide, idx)
    if ph is not None:
        ph.text_frame.text = text
    return ph


def _drop_ph(slide, idx):
    ph = _ph(slide, idx)
    if ph is not None:
        ph._element.getparent().remove(ph._element)


def _textbox(slide, left, top, width, height, text, size=12, bold=False, color="1F2937", font="Calibri", align=None):
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(0.05)
    tf.margin_top = tf.margin_bottom = Inches(0.02)
    p = tf.paragraphs[0]
    p.text = text
    p.font.size, p.font.bold, p.font.name = Pt(size), bold, font
    p.font.color.rgb = _rgb(color)
    if align:
        p.alignment = align
    return tb


def _content_slide(prs, heading, subheading):
    slide = prs.slides.add_slide(_layout(prs, "Title Slide - ONE LINE + SUB", 6))
    if _ph(slide, 0) is not None and _ph(slide, 2) is not None:
        _set_ph(slide, 0, heading.upper())
        _set_ph(slide, 2, subheading)
        _drop_ph(slide, 1)
    else:
        for ph in list(slide.placeholders):
            ph._element.getparent().remove(ph._element)
        _textbox(slide, Inches(0.64), Inches(0.5), Inches(12), Inches(0.5), heading.upper(), 24, True, BRAND_HEAD, "Montserrat")
        _textbox(slide, Inches(0.64), Inches(1.45), Inches(12), Inches(0.35), subheading, 13, False, BRAND_BLUE, "Montserrat")
    return slide


def _add_cover_picture(slide, img_bytes, left, top, width, height):
    """Place an image filling the box (crop to fit, like CSS object-fit: cover)."""
    with Image.open(BytesIO(img_bytes)) as im:
        iw, ih = im.size
    pic = slide.shapes.add_picture(BytesIO(img_bytes), left, top, width, height)
    box_ratio, img_ratio = width / height, iw / ih
    if img_ratio > box_ratio:
        excess = 1 - box_ratio / img_ratio
        pic.crop_left = pic.crop_right = excess / 2
    elif img_ratio < box_ratio:
        excess = 1 - img_ratio / box_ratio
        pic.crop_top = pic.crop_bottom = excess / 2
    return pic


def _pptx_table(slide, left, top, width, headers, rows, col_widths=None, font_size=11):
    shape = slide.shapes.add_table(len(rows) + 1, len(headers), left, top, width, Inches(0.32) * (len(rows) + 1))
    tbl = shape.table
    if col_widths:
        total = sum(col_widths)
        for j, w in enumerate(col_widths):
            tbl.columns[j].width = int(width * w / total)
    for j, h in enumerate(headers):
        cell = tbl.cell(0, j)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = _rgb(BRAND_NAVY)
        p = cell.text_frame.paragraphs[0]
        p.font.size, p.font.bold, p.font.name = Pt(font_size), True, "Calibri"
        p.font.color.rgb = _rgb("FFFFFF")
    for i, r in enumerate(rows, start=1):
        for j, v in enumerate(r):
            cell = tbl.cell(i, j)
            cell.text = str(v)
            cell.fill.solid()
            cell.fill.fore_color.rgb = _rgb(BRAND_ZEBRA if i % 2 == 0 else "FFFFFF")
            p = cell.text_frame.paragraphs[0]
            p.font.size, p.font.name = Pt(font_size), "Calibri"
            p.font.color.rgb = _rgb("1F2937")
    return shape


def generate_pptx_report(ctx):
    prs, has_template = _template_prs()
    paid_u, paid_src, photos = ctx["paid_unique"], ctx["paid_df"], ctx["paid_photos"]
    image_cache = fetch_all_images_parallel([u for u in photos.values() if u])
    today = datetime.now().strftime("%d %B %Y")

    # Cover
    cover = prs.slides.add_slide(_layout(prs, "COVER - Shape Dark Blue", 6))
    if has_template:
        _set_ph(cover, 0, "Paid Asset Executions Review")
        _set_ph(cover, 1, f"CHANDARIA {ctx['project_label'].upper()} · FIELD INTELLIGENCE")
        _set_ph(cover, 2, f"{ctx['period_txt']}  ·  {today}")
    else:
        _textbox(cover, Inches(0.8), Inches(2.6), Inches(11.7), Inches(1.0), "Paid Asset Executions Review", 36, True, BRAND_NAVY, "Arial")
        _textbox(cover, Inches(0.8), Inches(3.7), Inches(11.7), Inches(0.5), f"{ctx['period_txt']} · {today}", 16, False, "64748B")

    # Summary
    s = _content_slide(prs, "Paid assets summary", f"Unique paid assets · repeat visits to the same asset counted once · {ctx['period_txt']}")
    kpis = [
        ("UNIQUE PAID ASSETS", f"{len(paid_u):,}", BRAND_TEAL),
        ("OUTLETS WITH PAID ASSETS", f"{paid_u['key'].nunique():,}" if len(paid_u) else "0", BRAND_BLUE),
        ("UNIQUE UNPAID ASSETS", f"{len(ctx['unpaid_unique']):,}", BRAND_CYAN),
        ("ASSET TYPES", f"{paid_u['Asset'].nunique():,}" if len(paid_u) else "0", BRAND_HEAD),
    ]
    kx, ky, kw, kh, gap = Inches(0.64), Inches(2.05), Inches(2.85), Inches(1.0), Inches(0.23)
    for n, (label, value, color) in enumerate(kpis):
        left = kx + (kw + gap) * n
        box = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, ky, kw, kh)
        box.fill.solid()
        box.fill.fore_color.rgb = _rgb("F3F6FA")
        box.line.fill.background()
        bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, ky, Inches(0.07), kh)
        bar.fill.solid()
        bar.fill.fore_color.rgb = _rgb(color)
        bar.line.fill.background()
        _textbox(s, left + Inches(0.2), ky + Inches(0.12), kw - Inches(0.3), Inches(0.3), label, 9, True, "64748B", "Montserrat")
        _textbox(s, left + Inches(0.2), ky + Inches(0.38), kw - Inches(0.3), Inches(0.55), value, 26, True, BRAND_NAVY, "Montserrat")

    if len(paid_u):
        tc = paid_u["Asset"].value_counts().head(10).sort_values()
        cd = CategoryChartData()
        cd.categories = list(tc.index)
        cd.add_series("Unique assets", [int(v) for v in tc.values])
        gf = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.64), Inches(3.35), Inches(6.6), Inches(3.6), cd)
        ch = gf.chart
        ch.has_legend = False
        ch.has_title = True
        ch.chart_title.text_frame.text = "Unique paid assets by type"
        tp = ch.chart_title.text_frame.paragraphs[0]
        tp.font.size, tp.font.bold, tp.font.name = Pt(12), True, "Montserrat"
        tp.font.color.rgb = _rgb(BRAND_HEAD)
        plot = ch.plots[0]
        plot.gap_width = 60
        plot.has_data_labels = True
        plot.data_labels.font.size, plot.data_labels.font.name = Pt(10), "Calibri"
        plot.data_labels.number_format, plot.data_labels.number_format_is_linked = "0", False
        ser = plot.series[0]
        ser.format.fill.solid()
        ser.format.fill.fore_color.rgb = _rgb(BRAND_TEAL)
        ch.value_axis.visible = False
        ch.value_axis.has_major_gridlines = False
        ch.category_axis.tick_labels.font.size, ch.category_axis.tick_labels.font.name = Pt(11), "Calibri"
        ch.category_axis.tick_labels.font.color.rgb = _rgb("1F2937")
        plot.data_labels.font.color.rgb = _rgb("1F2937")
        plot.data_labels.font.bold = True
        ch.category_axis.format.line.color.rgb = _rgb(BRAND_RULE)

        by_proj = paid_u.groupby("Project").agg(a=("key", "size"), o=("key", "nunique")).reset_index()
        unpaid_proj = ctx["unpaid_unique"].groupby("Project").size() if len(ctx["unpaid_unique"]) else pd.Series(dtype=int)
        rows = [[r.Project, f"{r.a:,}", f"{r.o:,}", f"{int(unpaid_proj.get(r.Project, 0)):,}"] for r in by_proj.itertuples()]
        _textbox(s, Inches(7.6), Inches(3.42), Inches(5.1), Inches(0.3), "By project", 12, True, BRAND_HEAD, "Montserrat")
        _pptx_table(s, Inches(7.6), Inches(3.85), Inches(5.1), ["Project", "Unique paid", "Outlets", "Unique unpaid"], rows, [1.2, 1, 1, 1.1])

    # Photo slides: four unique assets per slide
    items = [(i, a) for i, a in paid_u.iterrows()]
    with_img = [(i, a) for i, a in items if photos.get(i) in image_cache]
    items = with_img or items
    pages = max(1, math.ceil(len(items) / 4))
    left0, top0 = Inches(0.64), Inches(1.98)
    tw, th, gx, gy = Inches(5.92), Inches(2.42), Inches(0.25), Inches(0.2)
    cap_h = Inches(0.42)
    for page in range(pages):
        chunk = items[page * 4:(page + 1) * 4]
        if not chunk:
            break
        slide = _content_slide(prs, "Paid asset executions", f"Outlet photos · page {page + 1} of {pages}")
        for n, (i, a) in enumerate(chunk):
            left = left0 + (tw + gx) * (n % 2)
            top = top0 + (th + gy) * (n // 2)
            url = photos.get(i)
            img_h = th - cap_h
            if url in image_cache:
                try:
                    _add_cover_picture(slide, image_cache[url], left, top, tw, img_h)
                except Exception:
                    pass
            else:
                ph = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, tw, img_h)
                ph.fill.solid()
                ph.fill.fore_color.rgb = _rgb("F3F6FA")
                ph.line.fill.background()
                _textbox(slide, left, top + img_h / 2 - Inches(0.15), tw, Inches(0.3), "Photo unavailable", 10, False, "94A3B8", align=PP_ALIGN.CENTER)
            cap = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top + img_h, tw, cap_h)
            cap.fill.solid()
            cap.fill.fore_color.rgb = _rgb(BRAND_NAVY)
            cap.line.fill.background()
            accent = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top + img_h, Inches(0.06), cap_h)
            accent.fill.solid()
            accent.fill.fore_color.rgb = _rgb(BRAND_TEAL)
            accent.line.fill.background()
            row = paid_src.loc[a["idx"]]
            last = a["Last Seen"].strftime("%d %b %Y") if pd.notna(a["Last Seen"]) else ""
            visits = int(a["Visits"])
            _textbox(slide, left + Inches(0.15), top + img_h + Inches(0.03), tw - Inches(0.25), Inches(0.2),
                     f"{a['Outlet']}  ·  {a['Asset']}", 11, True, "FFFFFF")
            _textbox(slide, left + Inches(0.15), top + img_h + Inches(0.21), tw - Inches(0.25), Inches(0.18),
                     f"{row['Select Your Name'].title()}  ·  last seen {last}  ·  {visits} visit{'s' if visits != 1 else ''}", 9, False, "CFD8E4")

    # Closing
    if has_template:
        close = prs.slides.add_slide(_layout(prs, "CLOSING Slide"))
        _set_ph(close, 1, "SMOLLAN KENYA   |   CHANDARIA FIELD INTELLIGENCE")
        _set_ph(close, 2, f"Generated {today}")

    out = BytesIO()
    prs.save(out)
    return out.getvalue()


# --- SHEET NORMALIZATION ---
def _find_col(df, *candidates):
    lower_map = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None


def _cols_starting(df, prefix):
    return [c for c in df.columns if c.lower().strip().startswith(prefix.lower())]


def _clean_text(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    s = str(val).strip()
    return "" if s.lower() in ["", "nan", "none", "no", "nil", "n/a", "na"] else s


def _clean_series(s):
    return s.fillna("").astype(str).str.strip().replace({"nan": "", "None": ""})


def _photo_col(df, cols):
    if not cols:
        return pd.Series([None] * len(df), index=df.index, dtype=object)
    return df[cols].apply(lambda r: ", ".join(split_urls(*r.tolist())) or None, axis=1)


def normalize_new_sheet(df):
    """'COMPETITORACTIVITES AND PAID ASSETS — Responses' sheet -> dashboard columns."""
    out = pd.DataFrame(index=df.index)
    ts_col = next((c for c in df.columns if c.lower().startswith("submission timestamp")), None) or _find_col(df, "Submission Date", "Timestamp")
    out["Submission Date"] = df[ts_col] if ts_col else None

    proj_col = _find_col(df, "SELECT YOUR PROJECT")
    out["Select Your Project"] = _clean_series(df[proj_col]).str.upper() if proj_col else ""

    names = pd.Series([""] * len(df), index=df.index, dtype=object)
    for c in _cols_starting(df, "select your name"):
        vals = _clean_series(df[c])
        names = names.where(names != "", vals)
    out["Select Your Name"] = names.str.upper()

    outlet_col = _find_col(df, "Type outlet name")
    out["Type Outlet Name"] = _clean_series(df[outlet_col]) if outlet_col else ""

    for canon in [PAID_COL, UNPAID_COL, COMP_COL]:
        c = _find_col(df, canon)
        out[canon] = _clean_series(df[c]) if c else ""

    desc_cols = _cols_starting(df, "describe the competitor activities")
    out[DESC_COL] = df.apply(lambda r: "\n".join(p for p in (_clean_text(r[c]) for c in desc_cols) if p) or None, axis=1) if desc_cols else None

    out[PAID_TYPES_COL] = None
    out[UNPAID_TYPES_COL] = None
    out[PHOTO_PAID] = _photo_col(df, [c for c in [_find_col(df, "Take a photo of Paid Assets")] if c])
    out[PHOTO_UNPAID] = _photo_col(df, [c for c in [_find_col(df, "Take a photo of unpaid Assets")] if c])
    out[PHOTO_COMP] = _photo_col(df, [c for c in [_find_col(df, "Take a photo of competitor activities")] if c])
    out[PHOTO_COMP_OTHERS] = _photo_col(df, [c for c in [_find_col(df, "Take a photo of competitor activities (others)")] if c])

    for c in ["Latitude", "Longitude"]:
        src = _find_col(df, c)
        out[c] = df[src] if src else None
    out["Source"] = "New form"
    return out


def normalize_old_sheet(df):
    """'PAID ASSETS/COMPETITOR ACTIVITIES' (Jotform) sheet -> dashboard columns."""
    out = pd.DataFrame(index=df.index)
    out["Submission Date"] = df[_find_col(df, "Submission Date")] if _find_col(df, "Submission Date") else None

    proj_col = _find_col(df, "Select Your Project")
    out["Select Your Project"] = _clean_series(df[proj_col]).str.upper() if proj_col else ""
    name_col = _find_col(df, "Select Your Name")
    out["Select Your Name"] = _clean_series(df[name_col]).str.upper() if name_col else ""
    outlet_col = _find_col(df, "Type Outlet Name")
    out["Type Outlet Name"] = _clean_series(df[outlet_col]) if outlet_col else ""

    paid = _find_col(df, "Does the store have any paid assets")
    out[PAID_COL] = _clean_series(df[paid]) if paid else ""
    unpaid = _find_col(df, "Does the store have any unpaid assets (Negotiated assets)", "Does the store have any unpaid assets")
    out[UNPAID_COL] = _clean_series(df[unpaid]) if unpaid else ""
    comp = _find_col(df, "Are there any Competitor activities in the outlet")
    out[COMP_COL] = _clean_series(df[comp]) if comp else ""

    desc = _find_col(df, "Describe the competitor activities")
    out[DESC_COL] = df[desc].apply(lambda v: _clean_text(v) or None) if desc else None

    ptypes = _find_col(df, "Select Paid assets in the outlet")
    out[PAID_TYPES_COL] = df[ptypes].apply(lambda v: _clean_text(v) or None) if ptypes else None
    utypes = _find_col(df, "Type asset negotiated")
    out[UNPAID_TYPES_COL] = df[utypes].apply(lambda v: _clean_text(v) or None) if utypes else None

    out[PHOTO_PAID] = _photo_col(df, _cols_starting(df, "take photo of paid assets"))
    out[PHOTO_UNPAID] = _photo_col(df, _cols_starting(df, "take photo of unpaid assets"))
    out[PHOTO_COMP] = _photo_col(df, _cols_starting(df, "take photo of competitor"))
    out[PHOTO_COMP_OTHERS] = None
    out["Latitude"] = None
    out["Longitude"] = None
    out["Source"] = "Previous form"
    return out


def _read_sheet(url):
    df = pd.read_csv(url)
    df.columns = df.columns.str.strip()
    return df.dropna(how="all")


@st.cache_data(ttl=300, show_spinner="Loading field submissions…")
def load_data():
    frames, warnings = [], []
    frames.append(normalize_new_sheet(_read_sheet(NEW_SHEET_CSV)))
    try:
        frames.append(normalize_old_sheet(_read_sheet(OLD_SHEET_CSV)))
    except Exception as e:
        warnings.append(f"Historical sheet could not be loaded: {e}")

    df = pd.concat(frames, ignore_index=True)
    df["Select Your Project"] = df["Select Your Project"].replace("", "Unassigned")
    df["Select Your Name"] = df["Select Your Name"].replace("", "UNKNOWN REP")
    df["Type Outlet Name"] = df["Type Outlet Name"].replace("", "Unknown Outlet")

    try:
        parsed = pd.to_datetime(df["Submission Date"], errors="coerce", format="mixed")
    except (TypeError, ValueError):
        parsed = pd.to_datetime(df["Submission Date"], errors="coerce")
    parsed = parsed.where((parsed.dt.year >= 2020) & (parsed.dt.year <= 2035))
    df["Parsed_Date"] = parsed
    df["Formatted_Timestamp"] = df["Parsed_Date"].dt.strftime("%d %b %Y, %H:%M").fillna("Recent Entry")
    df = df.sort_values("Parsed_Date", ascending=False, na_position="last").reset_index(drop=True)
    return df, warnings, datetime.now()


def is_yes(series):
    return series.astype(str).str.strip().str.lower() == "yes"


# --- UI HELPERS ---
def esc(v):
    return html.escape(str(v)) if v is not None else ""


def project_chip(project):
    cls = {"CORE": "chip chip-core", "IB": "chip chip-ib"}.get(str(project).upper(), "chip")
    return f"<span class='{cls}'>{esc(project)}</span>"


def style_fig(fig, height=330):
    fig.update_layout(
        template="plotly_dark", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        height=height, margin=dict(l=8, r=8, t=40, b=8),
        font=dict(family="Inter, sans-serif", color="#CBD5E1", size=12),
        title_font=dict(size=14, color="#F1F5F9"), title_x=0,
        legend=dict(orientation="h", yanchor="top", y=-0.12, xanchor="left", x=0, title=None, font=dict(size=11)),
        hoverlabel=dict(bgcolor="#111A2E", bordercolor="#1F2B45", font=dict(color="#E2E8F0")),
        bargap=0.28,
    )
    fig.update_xaxes(showgrid=False, linecolor="#1F2B45", title=None)
    fig.update_yaxes(gridcolor="#1A2540", zeroline=False, title=None)
    return fig


PLOT_CONFIG = {"displayModeBar": False}


def paginate(total, key, per_page):
    if total <= per_page:
        return 0, total
    pages = math.ceil(total / per_page)
    c1, c2 = st.columns([1, 5])
    page = c1.number_input("Page", min_value=1, max_value=pages, value=1, step=1, key=key)
    start = (page - 1) * per_page
    end = min(start + per_page, total)
    c2.markdown(f"<div class='section-sub' style='padding-top:34px'>Showing {start + 1}–{end} of {total} · page {page} of {pages}</div>", unsafe_allow_html=True)
    return start, end


def photos_html(urls, limit=6):
    if not urls:
        return ""
    tiles = "".join(
        f"<a href='{esc(img_src(u))}' target='_blank'><img src='{esc(img_src(u))}' loading='lazy'></a>"
        for u in urls[:limit]
    )
    more = f"<div class='section-sub' style='margin-top:6px'>+{len(urls) - limit} more photo(s)</div>" if len(urls) > limit else ""
    return f"<div class='photo-grid'>{tiles}</div>{more}"


def card_head(row):
    return (
        "<div class='card-head'>"
        f"<div><div class='card-outlet'>{esc(row['Type Outlet Name'])}</div>"
        f"<div class='card-rep'>{esc(row['Select Your Name'])}</div></div>"
        f"<div>{project_chip(row['Select Your Project'])}<span class='chip'>{esc(row['Formatted_Timestamp'])}</span></div>"
        "</div>"
    )


def promo_list_html(items):
    lis = []
    for it in items:
        extras = ""
        if it.get("end_ts") is not None:
            extras += f"<span class='badge badge-muted'>Ends {it['end_ts'].strftime('%d %b %Y')}</span>"
        elif it.get("end_date"):
            extras += f"<span class='badge badge-muted'>{esc(it['end_date'])}</span>"
        if it.get("duration_days"):
            extras += f"<span class='badge badge-muted'>{it['duration_days']} days left</span>"
        if it.get("status") == "ok":
            lis.append(
                f"<li><span class='p-name'>{esc(it['product'])}</span>"
                f"<span class='p-now'>KSh {it['new_price']:,.0f}</span>"
                f"<span class='p-was'>KSh {it['old_price']:,.0f}</span>"
                f"<span class='badge badge-red'>−{it['discount_pct']:.1f}% · KSh {it['discount_amount']:,.0f}</span>{extras}</li>"
            )
        elif it.get("status") == "single" and (it.get("new_price") or it.get("old_price")):
            price = it.get("new_price") or it.get("old_price")
            lis.append(f"<li><span class='p-name'>{esc(it['product'])}</span><span class='p-now'>KSh {price:,.0f}</span>{extras}</li>")
        elif it.get("status") == "review":
            lis.append(f"<li>{esc(it['raw_line'])}<span class='badge badge-amber'>check: {esc(it['reason'])}</span>{extras}</li>")
        else:
            lis.append(f"<li>{esc(it['raw_line'])}{extras}</li>")
    return f"<ul class='promo-list'>{''.join(lis)}</ul>"


def outlet_key(name):
    return re.sub(r"[^a-z0-9]+", " ", str(name).lower()).strip()


def unique_assets(src_df, type_col, default_label):
    """One row per outlet + asset type (repeat visits to the same asset count once); keeps the latest report."""
    cols = ["key", "Asset", "idx", "Visits", "First Seen", "Last Seen", "Outlet", "Project"]
    recs = []
    for idx, row in src_df.iterrows():
        raw = row[type_col]
        types = [t.strip().title() for t in raw.split(",") if t.strip()] if isinstance(raw, str) and raw.strip() else []
        for t in types or [default_label]:
            recs.append({"key": outlet_key(row["Type Outlet Name"]), "Asset": t, "idx": idx, "date": row["Parsed_Date"], "typed": bool(types)})
    if not recs:
        return pd.DataFrame(columns=cols)
    r = pd.DataFrame(recs)
    typed_keys = set(r.loc[r["typed"], "key"])
    r = r[r["typed"] | ~r["key"].isin(typed_keys)]
    agg = (r.sort_values("date", ascending=False, na_position="last")
             .groupby(["key", "Asset"], sort=False)
             .agg(idx=("idx", "first"), Visits=("idx", "nunique"), **{"First Seen": ("date", "min"), "Last Seen": ("date", "max")})
             .reset_index())
    agg["Outlet"] = agg["idx"].map(src_df["Type Outlet Name"])
    agg["Project"] = agg["idx"].map(src_df["Select Your Project"])
    return agg.sort_values("Last Seen", ascending=False, na_position="last").reset_index(drop=True)


BRAND_ALIASES = {
    "soft leaf": "Softleaf", "soft leave": "Softleaf", "softleaf": "Softleaf", "pure haven": "Pure Haven", "pure havens": "Pure Haven",
    "haven": "Pure Haven", "fields of africa": "Fields Of Africa", "tishu poa": "Tishu Poa", "quick choice": "Quickchoice",
    "quickchoice": "Quickchoice", "hanaa": "Hanan", "hanna": "Hanan", "statwrapp": "Statwrap", "posy": "Poshy", "petals": "Celine",
    "celine petals": "Celine", "floral": "Flora", "santa maria": "Santa Maria",
}


def brand_of(product):
    text = re.sub(r"[^a-z' ]+", " ", str(product).lower()).strip()
    for alias in sorted(BRAND_ALIASES, key=len, reverse=True):
        if text == alias or text.startswith(alias + " "):
            return BRAND_ALIASES[alias]
    words = re.findall(r"[A-Za-z][A-Za-z'&]*", str(product))
    return words[0].title() if words else "Other"


# --- LOAD ---
try:
    df, load_warnings, loaded_at = load_data()
except Exception as e:
    st.error(f"Failed to load live data: {e}")
    st.stop()

# --- SIDEBAR ---
st.sidebar.markdown("<div class='sb-brand'>Field Intelligence</div><div class='sb-brand-sub'>Chandaria · Smollan Kenya</div>", unsafe_allow_html=True)

valid_dates = df["Parsed_Date"].dropna()
start_d = end_d = None
if not valid_dates.empty:
    min_d, max_d = valid_dates.min().date(), valid_dates.max().date()
    def clamp(d):
        return max(min_d, min(max_d, d))

    if "date_from" not in st.session_state:
        st.session_state.date_from, st.session_state.date_to = min_d, max_d
    st.session_state.date_from = clamp(st.session_state.date_from)
    st.session_state.date_to = clamp(st.session_state.date_to)

    st.sidebar.date_input("From", key="date_from", min_value=min_d, max_value=max_d, format="DD/MM/YYYY")
    st.sidebar.date_input("To", key="date_to", min_value=min_d, max_value=max_d, format="DD/MM/YYYY")
    start_d, end_d = st.session_state.date_from, st.session_state.date_to
    if start_d > end_d:
        st.sidebar.warning("'From' is after 'To' — the dates have been swapped.")
        start_d, end_d = end_d, start_d

filtered_df = df.copy()
if start_d and end_d:
    d = filtered_df["Parsed_Date"].dt.date
    filtered_df = filtered_df[(d >= start_d) & (d <= end_d)]

st.sidebar.markdown("<div class='sb-title'>Project</div>", unsafe_allow_html=True)
project_list = ["All"] + sorted(df["Select Your Project"].dropna().unique().tolist())
selected_project = st.sidebar.segmented_control("Project", options=project_list, default="All", key="project", label_visibility="collapsed")
if selected_project and selected_project != "All":
    filtered_df = filtered_df[filtered_df["Select Your Project"] == selected_project]

st.sidebar.markdown("<div class='sb-title'>Field Rep</div>", unsafe_allow_html=True)
rep_options = sorted(filtered_df["Select Your Name"].dropna().unique().tolist())
selected_reps = st.sidebar.multiselect("Field Rep", rep_options, placeholder="All reps", key="reps", label_visibility="collapsed")
if selected_reps:
    filtered_df = filtered_df[filtered_df["Select Your Name"].isin(selected_reps)]

st.sidebar.markdown("<div class='sb-title'>Outlet</div>", unsafe_allow_html=True)
outlet_options = sorted(filtered_df["Type Outlet Name"].dropna().unique().tolist())
selected_outlets = st.sidebar.multiselect("Outlet", outlet_options, placeholder="All outlets", key="outlets", label_visibility="collapsed")
if selected_outlets:
    filtered_df = filtered_df[filtered_df["Type Outlet Name"].isin(selected_outlets)]

st.sidebar.divider()
if st.sidebar.button("↻  Refresh data", use_container_width=True):
    st.cache_data.clear()
    st.rerun()
st.sidebar.caption(f"Synced {loaded_at.strftime('%d %b %Y, %H:%M')} · auto-refresh every 5 min")
for w in load_warnings:
    st.sidebar.warning(w)

# --- DERIVED DATA ---
paid_mask = is_yes(filtered_df[PAID_COL])
unpaid_mask = is_yes(filtered_df[UNPAID_COL])
comp_mask = is_yes(filtered_df[COMP_COL])

comp_df = filtered_df[comp_mask]
paid_df = filtered_df[paid_mask]
unpaid_df = filtered_df[unpaid_mask]

promo_records = []
comp_parsed = {}
for idx, row in comp_df.iterrows():
    items = parse_price_discounts(row[DESC_COL], submission_date=row["Parsed_Date"]) if isinstance(row[DESC_COL], str) else []
    comp_parsed[idx] = items
    first_photo = next(iter(row_photos(row, [PHOTO_COMP, PHOTO_COMP_OTHERS])), None)
    for it in items:
        promo_records.append({
            "Date": row["Parsed_Date"], "Project": row["Select Your Project"], "Outlet": row["Type Outlet Name"],
            "Field Rep": row["Select Your Name"], "Product": it["product"],
            "Was (KSh)": it["old_price"], "Now (KSh)": it["new_price"],
            "Depth %": round(it["discount_pct"], 1) if it["discount_pct"] else None,
            "Saving (KSh)": it["discount_amount"],
            "Ends": (it["end_ts"].strftime("%d %b %Y") if it.get("end_ts") is not None else (it["end_date"] or "")),
            "Days Left": it["duration_days"], "Status": it["status"], "Check": it["reason"], "Note": it["raw_line"],
            "Brand": brand_of(it["product"]), "Photo": img_src(first_photo) if first_photo else None,
        })
promo_df = pd.DataFrame(promo_records)
priced_promos = promo_df[promo_df["Status"] == "ok"].copy() if not promo_df.empty else promo_df

PAID_DEFAULT = "Paid Asset (Type Not Recorded)"
UNPAID_DEFAULT = "Unpaid Asset (Type Not Recorded)"
paid_unique = unique_assets(paid_df, PAID_TYPES_COL, PAID_DEFAULT)
unpaid_unique = unique_assets(unpaid_df, UNPAID_TYPES_COL, UNPAID_DEFAULT)

paid_photos = {i: next(iter(row_photos(paid_df.loc[a["idx"]], [PHOTO_PAID])), None) for i, a in paid_unique.iterrows()}
unpaid_photos = {i: next(iter(row_photos(unpaid_df.loc[a["idx"]], [PHOTO_UNPAID])), None) for i, a in unpaid_unique.iterrows()}

# --- HEADER ---
if start_d and end_d:
    period_txt = f"{start_d.strftime('%d %b %Y')} – {end_d.strftime('%d %b %Y')}"
elif not filtered_df["Parsed_Date"].dropna().empty:
    p_from = filtered_df["Parsed_Date"].min().strftime("%d %b %Y")
    p_to = filtered_df["Parsed_Date"].max().strftime("%d %b %Y")
    period_txt = f"{p_from} – {p_to}"
else:
    period_txt = "No dated submissions"

st.markdown(f"""
<div class="dash-header">
  <div>
    <div class="dash-eyebrow">Chandaria Field Intelligence</div>
    <div class="dash-title">Competitor Activities &amp; Paid Assets</div>
  </div>
  <div class="dash-meta">
    Period <b>{esc(period_txt)}</b><br>
    Project <b>{esc(selected_project or 'All')}</b> · Reps <b>{len(selected_reps) or 'All'}</b> · Outlets <b>{len(selected_outlets) or 'All'}</b>
  </div>
</div>
""", unsafe_allow_html=True)

avg_depth = priced_promos["Depth %"].mean() if not priced_promos.empty else None
max_depth = priced_promos["Depth %"].max() if not priced_promos.empty else None
comp_outlets = comp_df["Type Outlet Name"].map(outlet_key).nunique()
n_brands = priced_promos["Brand"].nunique() if not priced_promos.empty else 0
top_brand = priced_promos["Brand"].value_counts().idxmax() if not priced_promos.empty else None

kpis = [
    ("Unique Paid Assets", f"{len(paid_unique):,}", f"across {paid_unique['key'].nunique():,} outlets", ACTIVITY_COLORS["Paid Assets"]),
    ("Unique Unpaid Assets", f"{len(unpaid_unique):,}", f"across {unpaid_unique['key'].nunique():,} outlets", ACTIVITY_COLORS["Unpaid Assets"]),
    ("Competitor Outlets", f"{comp_outlets:,}", f"{len(priced_promos):,} price cuts captured", ACTIVITY_COLORS["Competitor Activities"]),
    ("Avg Promo Depth", f"{avg_depth:.1f}%" if avg_depth is not None else "—", f"deepest {max_depth:.1f}%" if max_depth is not None else "—", "#F87171"),
    ("Competitor Brands", f"{n_brands:,}", f"most active: {top_brand}" if top_brand else "—", "#A78BFA"),
]
st.markdown(
    "<div class='kpi-row'>" + "".join(
        f"<div class='kpi' style='--kpi-accent:{c}'><div class='kpi-label'>{esc(l)}</div><div class='kpi-value'>{esc(v)}</div><div class='kpi-sub'>{esc(s)}</div></div>"
        for l, v, s, c in kpis
    ) + "</div>",
    unsafe_allow_html=True,
)

if filtered_df.empty:
    st.info("No submissions match the selected filters.")
    st.stop()

tab_overview, tab_comp, tab_assets, tab_gallery, tab_export = st.tabs(
    ["Overview", "Competitor Activity", "Paid & Unpaid Assets", "Gallery", "Exports"]
)

# ==========================================
# OVERVIEW
# ==========================================
with tab_overview:
    c1, c2 = st.columns(2, gap="large")
    with c1:
        if not priced_promos.empty:
            bstats = priced_promos.groupby("Brand").agg(**{"Price cuts": ("Product", "size"), "Avg depth %": ("Depth %", "mean")}).reset_index()
            top = bstats.sort_values("Price cuts", ascending=False).head(12).sort_values("Price cuts")
            fig = px.bar(top, x="Price cuts", y="Brand", orientation="h", title="Most active competitor brands",
                         text=top["Avg depth %"].map(lambda v: f"{v:.1f}% avg"), color_discrete_sequence=[ACTIVITY_COLORS["Competitor Activities"]],
                         hover_data={"Avg depth %": ":.1f"})
            fig.update_traces(textposition="outside", cliponaxis=False, textfont_size=11)
            fig.update_xaxes(range=[0, top["Price cuts"].max() * 1.3])
            st.plotly_chart(style_fig(fig, 400), use_container_width=True, config=PLOT_CONFIG)
        else:
            st.info("No priced competitor promotions for the selected filters.")
    with c2:
        if not priced_promos.empty:
            deep = bstats[bstats["Price cuts"] >= 3].sort_values("Avg depth %", ascending=False).head(12).sort_values("Avg depth %")
            if deep.empty:
                deep = bstats.sort_values("Avg depth %", ascending=False).head(12).sort_values("Avg depth %")
            fig = px.bar(deep, x="Avg depth %", y="Brand", orientation="h", title="Deepest average promo depth by brand",
                         text=deep["Avg depth %"].map(lambda v: f"{v:.1f}%"), color_discrete_sequence=["#FB923C"],
                         hover_data={"Price cuts": True})
            fig.update_traces(textposition="outside", cliponaxis=False, textfont_size=11)
            fig.update_xaxes(ticksuffix="%", range=[0, deep["Avg depth %"].max() * 1.2])
            st.plotly_chart(style_fig(fig, 400), use_container_width=True, config=PLOT_CONFIG)

    c3, c4 = st.columns(2, gap="large")
    with c3:
        if not paid_unique.empty:
            tc = paid_unique["Asset"].value_counts().sort_values().reset_index()
            tc.columns = ["Asset Type", "Unique assets"]
            fig = px.bar(tc, x="Unique assets", y="Asset Type", orientation="h", title="Unique paid assets by type",
                         text="Unique assets", color_discrete_sequence=[ACTIVITY_COLORS["Paid Assets"]])
            fig.update_traces(textposition="outside", cliponaxis=False)
            fig.update_xaxes(range=[0, tc["Unique assets"].max() * 1.15])
            st.plotly_chart(style_fig(fig, 360), use_container_width=True, config=PLOT_CONFIG)
        else:
            st.info("No paid assets for the selected filters.")
    with c4:
        mix = pd.concat([
            paid_unique.groupby("Project").size().rename("Paid Assets"),
            unpaid_unique.groupby("Project").size().rename("Unpaid Assets"),
            comp_df.assign(k=comp_df["Type Outlet Name"].map(outlet_key)).groupby("Select Your Project")["k"].nunique().rename("Competitor Activities"),
        ], axis=1).fillna(0).astype(int).rename_axis("Project").reset_index() \
            .melt(id_vars="Project", var_name="Measure", value_name="Count")
        if not mix.empty:
            fig = px.bar(mix, x="Project", y="Count", color="Measure", barmode="group", text="Count",
                         color_discrete_map=ACTIVITY_COLORS, title="Unique assets and competitor outlets by project")
            fig.update_traces(textposition="outside", cliponaxis=False)
            st.plotly_chart(style_fig(fig, 360), use_container_width=True, config=PLOT_CONFIG)

    st.markdown("<div class='section-title'>Deepest competitor price cuts</div><div class='section-sub'>Top 15 by promo depth for the selected period</div>", unsafe_allow_html=True)
    if not priced_promos.empty:
        top_cuts = priced_promos.sort_values("Depth %", ascending=False).head(15)[
            ["Date", "Outlet", "Brand", "Product", "Was (KSh)", "Now (KSh)", "Depth %", "Saving (KSh)", "Ends"]]
        st.dataframe(
            top_cuts, hide_index=True, use_container_width=True,
            column_config={
                "Date": st.column_config.DatetimeColumn("Date", format="DD MMM YYYY"),
                "Was (KSh)": st.column_config.NumberColumn("Was (KSh)", format="%,.0f"),
                "Now (KSh)": st.column_config.NumberColumn("Now (KSh)", format="%,.0f"),
                "Saving (KSh)": st.column_config.NumberColumn("Saving (KSh)", format="%,.0f"),
                "Depth %": st.column_config.NumberColumn("Depth %", format="%.1f%%"),
            },
        )

# ==========================================
# COMPETITOR ACTIVITY
# ==========================================
with tab_comp:
    view = st.segmented_control("View", ["Price cuts table", "Outlet reports"], default="Price cuts table", key="comp_view", label_visibility="collapsed")

    if view == "Outlet reports":
        st.markdown(f"<div class='section-title'>Competitor activity by outlet</div><div class='section-sub'>{len(comp_df):,} reports, newest first</div>", unsafe_allow_html=True)
        if comp_df.empty:
            st.info("No competitor activities reported for the selected filters.")
        else:
            s, e = paginate(len(comp_df), "comp_page", 20)
            for idx, row in comp_df.iloc[s:e].iterrows():
                items = comp_parsed.get(idx) or []
                body = promo_list_html(items) if items else "<div class='empty-note'>Competitor activity confirmed — no details entered.</div>"
                st.markdown(f"<div class='card'>{card_head(row)}{body}{photos_html(row_photos(row, [PHOTO_COMP, PHOTO_COMP_OTHERS]))}</div>", unsafe_allow_html=True)
    else:
        st.markdown(f"<div class='section-title'>Competitor price cuts</div><div class='section-sub'>{len(priced_promos):,} price cuts parsed from rep notes · sort any column</div>", unsafe_allow_html=True)
        if priced_promos.empty:
            st.info("No priced competitor promotions captured for the selected filters.")
        else:
            table_cols = ["Date", "Outlet", "Brand", "Product", "Was (KSh)", "Now (KSh)", "Depth %", "Saving (KSh)",
                          "Ends", "Days Left", "Project", "Field Rep"]
            st.dataframe(
                priced_promos.sort_values("Date", ascending=False)[table_cols], hide_index=True, use_container_width=True,
                height=min(600, 38 + 35 * len(priced_promos)),
                column_config={
                    "Date": st.column_config.DatetimeColumn("Date", format="DD MMM YYYY"),
                    "Was (KSh)": st.column_config.NumberColumn("Was (KSh)", format="%,.0f"),
                    "Now (KSh)": st.column_config.NumberColumn("Now (KSh)", format="%,.0f"),
                    "Saving (KSh)": st.column_config.NumberColumn("Saving (KSh)", format="%,.0f"),
                    "Depth %": st.column_config.NumberColumn("Depth %", format="%.1f%%"),
                    "Days Left": st.column_config.NumberColumn("Days Left", format="%d", help="Days from the visit to the offer end date"),
                },
            )
            unpriced = promo_df[promo_df["Status"] != "ok"] if not promo_df.empty else promo_df
            if not unpriced.empty:
                with st.expander(f"Lines not counted as price cuts ({len(unpriced):,}) — single prices, typos and notes"):
                    st.dataframe(unpriced[["Date", "Outlet", "Field Rep", "Note", "Was (KSh)", "Now (KSh)", "Check"]],
                                 hide_index=True, use_container_width=True,
                                 column_config={"Date": st.column_config.DatetimeColumn("Date", format="DD MMM YYYY"),
                                                "Was (KSh)": st.column_config.NumberColumn("Was (KSh)", format="%,.0f"),
                                                "Now (KSh)": st.column_config.NumberColumn("Now (KSh)", format="%,.0f")})

# ==========================================
# PAID & UNPAID ASSETS
# ==========================================
with tab_assets:
    asset_view = st.segmented_control("Assets", ["Paid assets", "Unpaid assets"], default="Paid assets", key="asset_view", label_visibility="collapsed")
    is_paid = asset_view != "Unpaid assets"
    src_df = paid_df if is_paid else unpaid_df
    uniq = paid_unique if is_paid else unpaid_unique
    photo_col = PHOTO_PAID if is_paid else PHOTO_UNPAID
    badge_cls = "badge-green" if is_paid else "badge-amber"
    label = "paid" if is_paid else "unpaid"

    n_outlets = uniq["key"].nunique() if not uniq.empty else 0
    st.markdown(f"<div class='section-title'>Unique {label} assets</div><div class='section-sub'>{len(uniq):,} unique assets across {n_outlets:,} outlets · repeat visits to the same asset are counted once ({len(src_df):,} reports)</div>", unsafe_allow_html=True)
    if uniq.empty:
        st.info(f"No outlets with {label} assets reported for the selected filters.")
    else:
        outlet_groups = list(uniq.groupby("key", sort=False))
        s_, e_ = paginate(len(outlet_groups), f"{label}_page", 20)
        for key, grp in outlet_groups[s_:e_]:
            latest = src_df.loc[grp.sort_values("Last Seen", ascending=False, na_position="last")["idx"].iloc[0]]
            badges = "".join(
                f"<span class='badge {badge_cls}'>{esc(a['Asset'])}</span>"
                f"<span class='badge badge-muted'>seen on {int(a['Visits'])} visit{'s' if a['Visits'] != 1 else ''}"
                + (f" · since {a['First Seen'].strftime('%d %b %Y')}" if pd.notna(a['First Seen']) else "") + "</span> "
                for _, a in grp.iterrows()
            )
            st.markdown(f"<div class='card'>{card_head(latest)}<div style='display:flex;flex-wrap:wrap;gap:6px'>{badges}</div>{photos_html(row_photos(latest, [photo_col]))}</div>", unsafe_allow_html=True)

# ==========================================
# GALLERY
# ==========================================
with tab_gallery:
    photo_type = st.segmented_control(
        "Photo type", ["All photos", "Paid assets", "Unpaid assets", "Competitor activities"],
        default="All photos", key="gallery_type", label_visibility="collapsed",
    )
    type_map = {
        "Paid assets": [(PHOTO_PAID, "Paid asset")],
        "Unpaid assets": [(PHOTO_UNPAID, "Unpaid asset")],
        "Competitor activities": [(PHOTO_COMP, "Competitor"), (PHOTO_COMP_OTHERS, "Competitor (others)")],
    }
    chosen = type_map.get(photo_type) or [p for v in type_map.values() for p in v]

    gallery = []
    for _, row in filtered_df.iterrows():
        for col, cat in chosen:
            for u in split_urls(row.get(col)):
                gallery.append((row["Parsed_Date"], u, cat, row))
    gallery = sorted(gallery, key=lambda g: g[0] if pd.notna(g[0]) else pd.Timestamp.min, reverse=True)

    st.markdown(f"<div class='section-title'>Photo gallery</div><div class='section-sub'>{len(gallery):,} photos, newest first · click a photo to open it</div>", unsafe_allow_html=True)
    if not gallery:
        st.info("No photos submitted for this category under the current filters.")
    else:
        s, e = paginate(len(gallery), "gallery_page", 24)
        tiles = "".join(
            f"<div class='g-item'><a href='{esc(img_src(u))}' target='_blank'><img src='{esc(img_src(u))}' loading='lazy'></a>"
            f"<div class='g-cap'><div class='g-outlet'>{esc(r['Type Outlet Name'])}</div>"
            f"<div class='g-meta'>{esc(cat)} · {esc(r['Select Your Name'])}</div>"
            f"<div class='g-meta'>{esc(r['Formatted_Timestamp'])}</div></div></div>"
            for _, u, cat, r in gallery[s:e]
        )
        st.markdown(f"<div class='gallery-grid'>{tiles}</div>", unsafe_allow_html=True)

# ==========================================
# EXPORTS
# ==========================================
with tab_export:
    st.markdown("<div class='section-title'>Report exports</div><div class='section-sub'>Exports use the filters currently applied in the sidebar</div>", unsafe_allow_html=True)
    export_ctx = {
        "meta_line": f"Chandaria · Smollan Kenya   |   Period: {period_txt}   |   Project: {selected_project or 'All'}   |   Generated: {datetime.now().strftime('%d %b %Y, %H:%M')}",
        "period_txt": period_txt, "project_label": selected_project if selected_project and selected_project != "All" else "All Projects",
        "priced": priced_promos, "promo": promo_df, "paid_unique": paid_unique, "unpaid_unique": unpaid_unique,
        "paid_df": paid_df, "unpaid_df": unpaid_df, "comp_df": comp_df, "comp_outlets": comp_outlets,
        "paid_photos": paid_photos, "unpaid_photos": unpaid_photos, "raw_df": filtered_df,
    }
    stamp = datetime.now().strftime("%Y%m%d")
    e1, e2 = st.columns(2, gap="large")
    with e1:
        st.markdown(f"<div class='card'><div class='card-outlet'>Excel report</div><div class='card-rep'>Summary, competitor price cuts, unique paid and unpaid assets with photos, and raw data · {len(promo_df):,} competitor lines · {len(paid_unique):,} unique paid assets</div></div>", unsafe_allow_html=True)
        if st.button("Build Excel report", key="excel_btn", use_container_width=True):
            with st.spinner("Downloading photos and building the report…"):
                excel_bytes = generate_excel_report(export_ctx)
            st.download_button("Download .xlsx", data=excel_bytes, file_name=f"Competitor_And_Paid_Assets_Report_{stamp}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="excel_dl", use_container_width=True, type="primary")
    with e2:
        tpl = find_template()
        tpl_note = ("<span class='badge badge-green'>Smollan template found</span>" if tpl else
                    "<span class='badge badge-amber'>Smollan template not found · upload Smollan_Core_Template.pptx to the GitHub repo</span>")
        st.markdown(f"<div class='card'><div class='card-outlet'>Paid assets deck</div><div class='card-rep'>Cover, summary, four photos per slide · {len(paid_unique):,} unique paid assets</div><div style='margin-top:10px'>{tpl_note}</div></div>", unsafe_allow_html=True)
        if st.button("Build PowerPoint deck", key="ppt_btn", use_container_width=True, disabled=paid_unique.empty):
            with st.spinner("Downloading photos and building the deck…"):
                ppt_bytes = generate_pptx_report(export_ctx)
            st.download_button("Download .pptx", data=ppt_bytes, file_name=f"Chandaria_Paid_Assets_Review_{stamp}.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               key="ppt_dl", use_container_width=True, type="primary")
