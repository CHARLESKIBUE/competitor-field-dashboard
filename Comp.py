import html
import math
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
from openpyxl.styles import Alignment, Font, PatternFill
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
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
.kpi-row { display: grid; grid-template-columns: repeat(6, minmax(0, 1fr)); gap: 12px; margin-bottom: 26px; }
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
def parse_price_discounts(text, submission_date=None):
    if not isinstance(text, str) or pd.isna(text):
        return []

    parsed_items = []
    current_year = datetime.now().year

    for line in re.split(r"[\r\n]+", text):
        line_clean = line.strip("-* ").strip()
        if not line_clean:
            continue

        start_date_str = None
        end_date_str = None
        duration_days = None

        range_match = re.search(r'(?:from\s*|start\s*)?(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\s*(?:to|-|until|–)\s*(\d{1,2}/\d{1,2}(?:/\d{2,4})?)', line_clean, re.IGNORECASE)

        if range_match:
            start_date_str = range_match.group(1).strip()
            end_date_str = range_match.group(2).strip()
            line_clean = re.sub(r'(?:from\s*|start\s*)?\d{1,2}/\d{1,2}(?:/\d{2,4})?\s*(?:to|-|until|–)\s*\d{1,2}/\d{1,2}(?:/\d{2,4})?', '', line_clean, flags=re.IGNORECASE).strip()
        else:
            start_match = re.search(r'(?:start\s*date\s*|start\s*|from\s*)(\d{1,2}/\d{1,2}(?:/\d{2,4})?)', line_clean, re.IGNORECASE)
            if start_match:
                start_date_str = start_match.group(1).strip()
                line_clean = re.sub(r'(?:start\s*date\s*|start\s*|from\s*)\d{1,2}/\d{1,2}(?:/\d{2,4})?', '', line_clean, flags=re.IGNORECASE).strip()

            end_match = re.search(r'(?:end\s*date\s*|ends\s*|\()(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\)?', line_clean, re.IGNORECASE)
            if end_match:
                end_date_str = end_match.group(1).strip()
                line_clean = re.sub(r'(?:end\s*date\s*|ends\s*|\()?\d{1,2}/\d{1,2}(?:/\d{2,4})?\)?', '', line_clean, flags=re.IGNORECASE).strip()

        if not start_date_str and submission_date is not None:
            if isinstance(submission_date, (pd.Timestamp, datetime)) and pd.notna(submission_date):
                start_date_str = submission_date.strftime("%d/%m/%Y")

        def parse_to_date(d_str):
            if not d_str:
                return None
            parts = str(d_str).split('/')
            if len(parts) == 2:
                d_str = f"{d_str}/{current_year}"
            elif len(parts) == 3 and len(parts[2]) == 2:
                parts[2] = f"20{parts[2]}"
                d_str = "/".join(parts)
            try:
                return datetime.strptime(d_str, "%d/%m/%Y").date()
            except Exception:
                return None

        d_start = parse_to_date(start_date_str)
        d_end = parse_to_date(end_date_str)
        if d_start and d_end and d_end >= d_start:
            duration_days = (d_end - d_start).days + 1

        found = False
        old_price, new_price, product = None, None, line_clean

        m1 = re.search(r"(.+?)(?:@|now\s*|KSh\s*)?\s*(\d+(?:\.\d+)?)\s*(?:from|was)\s*(?:KSh\s*)?(\d+(?:\.\d+)?)", line_clean, re.IGNORECASE)
        m2 = re.search(r"(.+?)(?:was|from)\s*(?:KSh\s*)?(\d+(?:\.\d+)?)\s*(?:now|@|to|mow|KSh)?\s*(\d+(?:\.\d+)?)", line_clean, re.IGNORECASE)
        m3 = re.search(r"(.+?)(?:KSh\s*)?(\d+(?:\.\d+)?)\s*[\-/]\s*(?:KSh\s*)?(\d+(?:\.\d+)?)", line_clean, re.IGNORECASE)

        if m1:
            product = m1.group(1).strip()
            new_price = float(m1.group(2))
            old_price = float(m1.group(3))
            found = True
        elif m2:
            product = m2.group(1).strip()
            old_price = float(m2.group(2))
            new_price = float(m2.group(3))
            found = True
        elif m3:
            product = m3.group(1).strip()
            p1, p2 = float(m3.group(2)), float(m3.group(3))
            if p1 > p2:
                old_price, new_price = p1, p2
                found = True

        product = product.rstrip("@:;- ").strip()

        if found and old_price and new_price and old_price > new_price:
            discount_amount = old_price - new_price
            parsed_items.append({
                "product": product, "old_price": old_price, "new_price": new_price,
                "discount_amount": discount_amount, "discount_pct": discount_amount / old_price * 100,
                "start_date": start_date_str, "end_date": end_date_str, "duration_days": duration_days,
                "raw_line": line_clean,
            })
        else:
            parsed_items.append({
                "product": line_clean, "old_price": None, "new_price": None,
                "discount_amount": None, "discount_pct": None,
                "start_date": start_date_str, "end_date": end_date_str, "duration_days": duration_days,
                "raw_line": line_clean,
            })

    return parsed_items


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


# --- EXCEL EXPORT ---
def _style_header(ws):
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="0F172A")
        cell.alignment = Alignment(vertical="center")
    ws.freeze_panes = "A2"


def generate_excel_with_images(comp_rows, paid_rows, raw_df):
    wb = openpyxl.Workbook()
    all_urls = [item.get("image_url") for item in comp_rows + paid_rows if item.get("image_url")]
    image_cache = fetch_all_images_parallel(all_urls)

    ws_comp = wb.active
    ws_comp.title = "Competitor Depth & Activity"
    ws_comp.append(["Date", "Project", "Outlet", "Field Rep", "Activity", "Was", "Now", "Depth", "Start Date", "Offer End Date", "Promo Duration", "Image"])
    for row_idx, item in enumerate(comp_rows, start=2):
        ws_comp.append([
            item["Date"], item["Project"], item["Outlet"], item["Field Rep"], item["Activity"],
            item["Was"], item["Now"], item["Depth"], item["Start Date"], item["End Date"], item["Duration Days"], "",
        ])
        ws_comp.row_dimensions[row_idx].height = 60
        img_url = item.get("image_url")
        if img_url in image_cache:
            try:
                xl_img = OpenPyxlImage(BytesIO(image_cache[img_url]))
                xl_img.width, xl_img.height = 70, 70
                ws_comp.add_image(xl_img, f"L{row_idx}")
            except Exception:
                ws_comp.cell(row=row_idx, column=12, value="[Image Error]")
        elif img_url:
            ws_comp.cell(row=row_idx, column=12, value="[Image Unavailable]")
    for col, width in zip("ABCDEFGHIJKL", [18, 9, 24, 26, 40, 11, 11, 26, 12, 14, 14, 14]):
        ws_comp.column_dimensions[col].width = width
    _style_header(ws_comp)

    ws_paid = wb.create_sheet(title="Paid Asset Executions")
    ws_paid.append(["Date", "Project", "Outlet", "Field Rep", "Paid Assets Executed", "Image"])
    for row_idx, item in enumerate(paid_rows, start=2):
        ws_paid.append([item["Date"], item["Project"], item["Outlet"], item["Field Rep"], item["Paid Assets Executed"], ""])
        ws_paid.row_dimensions[row_idx].height = 60
        img_url = item.get("image_url")
        if img_url in image_cache:
            try:
                xl_img = OpenPyxlImage(BytesIO(image_cache[img_url]))
                xl_img.width, xl_img.height = 70, 70
                ws_paid.add_image(xl_img, f"F{row_idx}")
            except Exception:
                ws_paid.cell(row=row_idx, column=6, value="[Image Error]")
        elif img_url:
            ws_paid.cell(row=row_idx, column=6, value="[Image Unavailable]")
    for col, width in zip("ABCDEF", [18, 9, 24, 26, 30, 14]):
        ws_paid.column_dimensions[col].width = width
    _style_header(ws_paid)

    ws_raw = wb.create_sheet(title="Raw Data")
    export_cols = [c for c in raw_df.columns if c not in ("Parsed_Date", "Formatted_Timestamp")]
    ws_raw.append(export_cols)
    for r in raw_df[export_cols].itertuples(index=False):
        ws_raw.append([None if (isinstance(v, float) and pd.isna(v)) else v for v in r])
    _style_header(ws_raw)

    output = BytesIO()
    wb.save(output)
    return output.getvalue()


# --- POWERPOINT EXPORT (PAID ASSETS) ---
def generate_pptx_report(paid_rows):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    image_cache = fetch_all_images_parallel([i.get("image_url") for i in paid_rows if i.get("image_url")])

    DARK_NAVY = RGBColor(15, 23, 42)
    SLATE_MUTED = RGBColor(100, 116, 139)
    WHITE = RGBColor(255, 255, 255)

    slide_title = prs.slides.add_slide(blank)
    tf_title = slide_title.shapes.add_textbox(Inches(0.8), Inches(2.5), Inches(11.733), Inches(3.0)).text_frame
    tf_title.word_wrap = True
    p1 = tf_title.paragraphs[0]
    p1.text = "IB WEEKLY FIELD FINDINGS"
    p1.font.size, p1.font.bold, p1.font.color.rgb, p1.font.name = Pt(38), True, DARK_NAVY, "Arial"
    p2 = tf_title.add_paragraph()
    p2.text = f"\nPaid Asset Executions Review | {datetime.now().strftime('%d %B %Y')}"
    p2.font.size, p2.font.color.rgb, p2.font.name = Pt(20), SLATE_MUTED, "Arial"
    bar = slide_title.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(6.8), Inches(13.333), Inches(0.7))
    bar.fill.solid()
    bar.fill.fore_color.rgb = DARK_NAVY
    bar.line.color.rgb = DARK_NAVY

    items = [i for i in paid_rows if i.get("image_url") in image_cache] or paid_rows
    grid = [
        (Inches(0.8), Inches(1.3)), (Inches(6.8), Inches(1.3)),
        (Inches(0.8), Inches(4.3)), (Inches(6.8), Inches(4.3)),
    ]
    w, h = Inches(5.6), Inches(2.7)

    for slide_num, start in enumerate(range(0, len(items), 4), start=1):
        slide = prs.slides.add_slide(blank)
        p_hdr = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.8)).text_frame.paragraphs[0]
        p_hdr.text = f"PAID ASSETS EXECUTIONS (PAGE {slide_num})"
        p_hdr.font.size, p_hdr.font.bold, p_hdr.font.color.rgb, p_hdr.font.name = Pt(22), True, DARK_NAVY, "Arial"

        for (left, top), item in zip(grid, items[start:start + 4]):
            if item.get("image_url") in image_cache:
                try:
                    slide.shapes.add_picture(BytesIO(image_cache[item["image_url"]]), left, top, w, h)
                except Exception:
                    pass
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top + h - Inches(0.4), w, Inches(0.4))
            bg.fill.solid()
            bg.fill.fore_color.rgb = DARK_NAVY
            bg.line.fill.background()
            p_cap = slide.shapes.add_textbox(left, top + h - Inches(0.4), w, Inches(0.4)).text_frame.paragraphs[0]
            p_cap.text = f" 📍 {item['Outlet']} | Rep: {item['Field Rep']}"
            p_cap.font.size, p_cap.font.bold, p_cap.font.color.rgb, p_cap.font.name = Pt(10), True, WHITE, "Arial"

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
        if it.get("end_date"):
            extras += f"<span class='badge badge-muted'>Ends {esc(it['end_date'])}</span>"
        if it.get("duration_days"):
            extras += f"<span class='badge badge-muted'>{it['duration_days']} days</span>"
        if it.get("discount_amount"):
            lis.append(
                f"<li><span class='p-name'>{esc(it['product'])}</span>"
                f"<span class='p-now'>KSh {it['new_price']:,.0f}</span>"
                f"<span class='p-was'>KSh {it['old_price']:,.0f}</span>"
                f"<span class='badge badge-red'>−{it['discount_pct']:.1f}% · KSh {it['discount_amount']:,.0f}</span>{extras}</li>"
            )
        else:
            lis.append(f"<li>{esc(it['raw_line'])}{extras}</li>")
    return f"<ul class='promo-list'>{''.join(lis)}</ul>"


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
    st.sidebar.markdown("<div class='sb-title'>Period</div>", unsafe_allow_html=True)
    period = st.sidebar.date_input("Date range", value=(min_d, max_d), min_value=min_d, max_value=max_d, key="period", label_visibility="collapsed")
    if isinstance(period, (list, tuple)) and len(period) == 2:
        start_d, end_d = period
    elif isinstance(period, (list, tuple)) and len(period) == 1:
        start_d = end_d = period[0]

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
    for it in items:
        promo_records.append({
            "Date": row["Parsed_Date"], "Project": row["Select Your Project"], "Outlet": row["Type Outlet Name"],
            "Field Rep": row["Select Your Name"], "Product": it["product"],
            "Was (KSh)": it["old_price"], "Now (KSh)": it["new_price"],
            "Depth %": round(it["discount_pct"], 1) if it["discount_pct"] else None,
            "Saving (KSh)": it["discount_amount"], "Ends": it["end_date"], "Duration (days)": it["duration_days"],
        })
promo_df = pd.DataFrame(promo_records)
priced_promos = promo_df[promo_df["Depth %"].notna()] if not promo_df.empty else promo_df

# Export rows
global_comp_rows = []
for idx, row in comp_df.iterrows():
    first_img = next(iter(row_photos(row, [PHOTO_COMP, PHOTO_COMP_OTHERS])), None)
    base = {"Date": row["Formatted_Timestamp"], "Project": row["Select Your Project"], "Outlet": row["Type Outlet Name"], "Field Rep": row["Select Your Name"], "image_url": first_img}
    items = comp_parsed.get(idx) or []
    if items:
        for it in items:
            global_comp_rows.append({**base,
                "Activity": it["product"],
                "Was": f"KSh {it['old_price']:.0f}" if it["old_price"] else "N/A",
                "Now": f"KSh {it['new_price']:.0f}" if it["new_price"] else "N/A",
                "Depth": f"Save KSh {it['discount_amount']:.0f} ({it['discount_pct']:.1f}% OFF)" if it["discount_amount"] else "N/A",
                "Start Date": it["start_date"] or "N/A", "End Date": it["end_date"] or "N/A",
                "Duration Days": f"{it['duration_days']} Days" if it["duration_days"] else "N/A",
            })
    else:
        global_comp_rows.append({**base, "Activity": "Competitor activity confirmed (no notes entered).",
            "Was": "N/A", "Now": "N/A", "Depth": "N/A", "Start Date": "N/A", "End Date": "N/A", "Duration Days": "N/A"})

global_paid_rows = []
for _, row in paid_df.iterrows():
    global_paid_rows.append({
        "Date": row["Formatted_Timestamp"], "Project": row["Select Your Project"], "Outlet": row["Type Outlet Name"],
        "Field Rep": row["Select Your Name"], "Paid Assets Executed": row[PAID_TYPES_COL] or "Paid asset confirmed",
        "image_url": next(iter(row_photos(row, [PHOTO_PAID])), None),
    })

# --- HEADER ---
if not filtered_df["Parsed_Date"].dropna().empty:
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

total = len(filtered_df)
outlets_visited = filtered_df["Type Outlet Name"].nunique()
avg_depth = priced_promos["Depth %"].mean() if not priced_promos.empty else None


def pct_of(n):
    return f"{n / total * 100:.0f}% of submissions" if total else "—"


kpis = [
    ("Submissions", f"{total:,}", f"{filtered_df['Select Your Name'].nunique()} field reps", "#38BDF8"),
    ("Outlets Visited", f"{outlets_visited:,}", "unique outlets", "#38BDF8"),
    ("Paid Assets", f"{int(paid_mask.sum()):,}", pct_of(paid_mask.sum()), ACTIVITY_COLORS["Paid Assets"]),
    ("Unpaid Assets", f"{int(unpaid_mask.sum()):,}", pct_of(unpaid_mask.sum()), ACTIVITY_COLORS["Unpaid Assets"]),
    ("Competitor", f"{int(comp_mask.sum()):,}", f"activity in {pct_of(comp_mask.sum())}" if total else "—", ACTIVITY_COLORS["Competitor Activities"]),
    ("Avg Promo Depth", f"{avg_depth:.1f}%" if avg_depth is not None else "—", f"{len(priced_promos):,} price cuts captured", "#F87171"),
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
    dated = filtered_df.dropna(subset=["Parsed_Date"])
    c1, c2 = st.columns([3, 2], gap="large")
    with c1:
        if not dated.empty:
            span_days = (dated["Parsed_Date"].max() - dated["Parsed_Date"].min()).days
            freq, label = ("D", "Daily") if span_days <= 45 else ("W-MON", "Weekly")
            trend = dated.assign(Period=dated["Parsed_Date"].dt.to_period(freq).dt.start_time) \
                .groupby(["Period", "Select Your Project"]).size().reset_index(name="Submissions")
            fig = px.bar(trend, x="Period", y="Submissions", color="Select Your Project",
                         color_discrete_map=PROJECT_COLORS, title=f"{label} submissions by project")
            st.plotly_chart(style_fig(fig), use_container_width=True, config=PLOT_CONFIG)
    with c2:
        mix = filtered_df.groupby("Select Your Project").agg(
            **{"Paid Assets": (PAID_COL, lambda s: is_yes(s).sum()),
               "Unpaid Assets": (UNPAID_COL, lambda s: is_yes(s).sum()),
               "Competitor Activities": (COMP_COL, lambda s: is_yes(s).sum())}
        ).reset_index().melt(id_vars="Select Your Project", var_name="Activity", value_name="Outlets")
        fig = px.bar(mix, x="Select Your Project", y="Outlets", color="Activity", barmode="group",
                     color_discrete_map=ACTIVITY_COLORS, title="Activity mix by project")
        st.plotly_chart(style_fig(fig), use_container_width=True, config=PLOT_CONFIG)

    c3, c4 = st.columns(2, gap="large")
    with c3:
        top_outlets = filtered_df["Type Outlet Name"].value_counts().head(10).sort_values().reset_index()
        top_outlets.columns = ["Outlet", "Submissions"]
        fig = px.bar(top_outlets, x="Submissions", y="Outlet", orientation="h", title="Most visited outlets",
                     color_discrete_sequence=["#38BDF8"])
        st.plotly_chart(style_fig(fig, 360), use_container_width=True, config=PLOT_CONFIG)
    with c4:
        types = filtered_df[PAID_TYPES_COL].dropna().astype(str).str.split(",").explode().str.strip()
        types = types[types != ""]
        if not types.empty:
            tc = types.str.title().value_counts().sort_values().reset_index()
            tc.columns = ["Asset Type", "Outlets"]
            fig = px.bar(tc, x="Outlets", y="Asset Type", orientation="h", title="Paid asset types executed",
                         color_discrete_sequence=[ACTIVITY_COLORS["Paid Assets"]])
            st.plotly_chart(style_fig(fig, 360), use_container_width=True, config=PLOT_CONFIG)
        elif not priced_promos.empty:
            depth_by_rep = priced_promos.groupby("Project")["Depth %"].mean().reset_index()
            fig = px.bar(depth_by_rep, x="Project", y="Depth %", title="Average competitor promo depth (%)",
                         color="Project", color_discrete_map=PROJECT_COLORS)
            st.plotly_chart(style_fig(fig, 360), use_container_width=True, config=PLOT_CONFIG)

    st.markdown("<div class='section-title'>Field rep activity</div><div class='section-sub'>Submissions and findings per rep for the selected period</div>", unsafe_allow_html=True)
    rep_tbl = filtered_df.groupby(["Select Your Name", "Select Your Project"]).agg(
        Submissions=("Type Outlet Name", "size"),
        Outlets=("Type Outlet Name", "nunique"),
        **{"Paid Assets": (PAID_COL, lambda s: int(is_yes(s).sum())),
           "Unpaid Assets": (UNPAID_COL, lambda s: int(is_yes(s).sum())),
           "Competitor Activities": (COMP_COL, lambda s: int(is_yes(s).sum()))},
        **{"Last Submission": ("Parsed_Date", "max")},
    ).reset_index().rename(columns={"Select Your Name": "Field Rep", "Select Your Project": "Project"}) \
        .sort_values("Submissions", ascending=False)
    st.dataframe(
        rep_tbl, hide_index=True, use_container_width=True,
        column_config={
            "Submissions": st.column_config.ProgressColumn("Submissions", format="%d", min_value=0, max_value=int(rep_tbl["Submissions"].max() or 1)),
            "Last Submission": st.column_config.DatetimeColumn("Last Submission", format="DD MMM YYYY, HH:mm"),
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
            st.dataframe(
                priced_promos.sort_values("Date", ascending=False), hide_index=True, use_container_width=True,
                height=min(600, 38 + 35 * len(priced_promos)),
                column_config={
                    "Date": st.column_config.DatetimeColumn("Date", format="DD MMM YYYY"),
                    "Was (KSh)": st.column_config.NumberColumn("Was (KSh)", format="%,.0f"),
                    "Now (KSh)": st.column_config.NumberColumn("Now (KSh)", format="%,.0f"),
                    "Saving (KSh)": st.column_config.NumberColumn("Saving (KSh)", format="%,.0f"),
                    "Depth %": st.column_config.ProgressColumn("Depth %", format="%.1f%%", min_value=0, max_value=float(max(50, priced_promos["Depth %"].max()))),
                    "Duration (days)": st.column_config.NumberColumn("Duration (days)", format="%d"),
                },
            )
            unpriced = promo_df[promo_df["Depth %"].isna()] if not promo_df.empty else promo_df
            if not unpriced.empty:
                with st.expander(f"Notes without a readable price ({len(unpriced):,})"):
                    st.dataframe(unpriced[["Date", "Project", "Outlet", "Field Rep", "Product", "Ends"]].rename(columns={"Product": "Note"}),
                                 hide_index=True, use_container_width=True,
                                 column_config={"Date": st.column_config.DatetimeColumn("Date", format="DD MMM YYYY")})

# ==========================================
# PAID & UNPAID ASSETS
# ==========================================
with tab_assets:
    asset_view = st.segmented_control("Assets", ["Paid assets", "Unpaid assets"], default="Paid assets", key="asset_view", label_visibility="collapsed")
    is_paid = asset_view != "Unpaid assets"
    src_df = paid_df if is_paid else unpaid_df
    photo_col = PHOTO_PAID if is_paid else PHOTO_UNPAID
    type_col = PAID_TYPES_COL if is_paid else UNPAID_TYPES_COL
    badge_cls = "badge-green" if is_paid else "badge-amber"
    label = "paid" if is_paid else "unpaid"

    st.markdown(f"<div class='section-title'>Outlets with {label} assets</div><div class='section-sub'>{len(src_df):,} outlet reports, newest first</div>", unsafe_allow_html=True)
    if src_df.empty:
        st.info(f"No outlets with {label} assets reported for the selected filters.")
    else:
        s, e = paginate(len(src_df), f"{label}_page", 20)
        for _, row in src_df.iloc[s:e].iterrows():
            types_txt = row[type_col] if isinstance(row[type_col], str) and row[type_col] else f"{label.title()} asset confirmed"
            badges = "".join(f"<span class='badge {badge_cls}'>{esc(t.strip())}</span> " for t in types_txt.split(",") if t.strip())
            st.markdown(f"<div class='card'>{card_head(row)}<div>{badges}</div>{photos_html(row_photos(row, [photo_col]))}</div>", unsafe_allow_html=True)

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
    e1, e2 = st.columns(2, gap="large")
    with e1:
        st.markdown(f"<div class='card'><div class='card-outlet'>Excel workbook</div><div class='card-rep'>Competitor depth, paid asset executions and raw data, with embedded photos · {len(global_comp_rows):,} activity rows · {len(global_paid_rows):,} paid asset rows</div></div>", unsafe_allow_html=True)
        if st.button("Build Excel workbook", key="excel_btn", use_container_width=True):
            with st.spinner("Downloading photos and building the workbook…"):
                excel_bytes = generate_excel_with_images(global_comp_rows, global_paid_rows, filtered_df)
            st.download_button("Download .xlsx", data=excel_bytes, file_name="Competitor_And_Paid_Assets_Report.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="excel_dl", use_container_width=True, type="primary")
    with e2:
        st.markdown(f"<div class='card'><div class='card-outlet'>Paid assets deck</div><div class='card-rep'>PowerPoint review with four photos per slide · {len(global_paid_rows):,} paid asset executions</div></div>", unsafe_allow_html=True)
        if st.button("Build PowerPoint deck", key="ppt_btn", use_container_width=True, disabled=not global_paid_rows):
            with st.spinner("Downloading photos and building the deck…"):
                ppt_bytes = generate_pptx_report(global_paid_rows)
            st.download_button("Download .pptx", data=ppt_bytes, file_name="Chandaria_Paid_Assets_Review.pptx",
                               mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                               key="ppt_dl", use_container_width=True, type="primary")
