import pandas as pd
import streamlit as st
import requests
import re
from io import BytesIO
from PIL import Image
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

# PPTX Library Imports
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# --- JOTFORM API KEY ---
JOTFORM_API_KEY = "6666042ab61ea98f9b378a4a7d"

# 1. Page Configuration
st.set_page_config(
    page_title="Competitor Activities and Paid Assets", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
    <style>
    div[data-testid="stSidebar"] {
        background-color: #f8f9fa;
        padding-top: 15px;
    }
    div[data-testid="stSidebar"] label {
        font-weight: 600 !important;
        color: #2c3e50 !important;
    }
    button[data-baseweb="tab"] {
        font-size: 16px !important;
        font-weight: 700 !important;
        color: #1e293b !important;
        padding: 10px 20px !important;
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #0f172a !important;
        border-bottom-color: #0f172a !important;
    }
    .timestamp-badge {
        background-color: #eef2f7;
        color: #475569;
        font-size: 12px;
        font-weight: 600;
        padding: 4px 10px;
        border-radius: 12px;
        border: 1px solid #cbd5e1;
    }
    .discount-badge {
        background-color: #f1f5f9;
        color: #0f172a;
        font-size: 12px;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 6px;
        border: 1px solid #cbd5e1;
        margin-left: 6px;
    }
    .enddate-badge {
        background-color: #fef3c7;
        color: #92400e;
        font-size: 12px;
        font-weight: 600;
        padding: 2px 8px;
        border-radius: 6px;
        border: 1px solid #fde68a;
        margin-left: 6px;
    }
    </style>
""", unsafe_allow_html=True)

# Helper function to parse pricing discount depth, promo start/end dates, AND duration
def parse_price_discounts(text, submission_date=None):
    if not isinstance(text, str) or pd.isna(text):
        return []
    
    parsed_items = []
    lines = text.split('\n')
    current_year = datetime.now().year
    
    for line in lines:
        line_clean = line.strip('-* ').strip()
        if not line_clean:
            continue
            
        start_date_str = None
        end_date_str = None
        duration_days = None
        
        # 1. Check for explicit Date Ranges (e.g., "01/09-30/09", "01/9 to 30/9")
        range_match = re.search(r'(?:from\s*|start\s*)?(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\s*(?:to|-|until|–)\s*(\d{1,2}/\d{1,2}(?:/\d{2,4})?)', line_clean, re.IGNORECASE)
        
        if range_match:
            start_date_str = range_match.group(1).strip()
            end_date_str = range_match.group(2).strip()
            line_clean = re.sub(r'(?:from\s*|start\s*)?\d{1,2}/\d{1,2}(?:/\d{2,4})?\s*(?:to|-|until|–)\s*\d{1,2}/\d{1,2}(?:/\d{2,4})?', '', line_clean, flags=re.IGNORECASE).strip()
        else:
            # 2. Check for single Start Date (e.g., "start 01/9", "from 01/09")
            start_match = re.search(r'(?:start\s*date\s*|start\s*|from\s*)(\d{1,2}/\d{1,2}(?:/\d{2,4})?)', line_clean, re.IGNORECASE)
            if start_match:
                start_date_str = start_match.group(1).strip()
                line_clean = re.sub(r'(?:start\s*date\s*|start\s*|from\s*)\d{1,2}/\d{1,2}(?:/\d{2,4})?', '', line_clean, flags=re.IGNORECASE).strip()

            # 3. Check for single End Date (e.g., "ends 31/10", "end date 31/10", "(31/10)")
            end_match = re.search(r'(?:end\s*date\s*|ends\s*|\()(\d{1,2}/\d{1,2}(?:/\d{2,4})?)\)?', line_clean, re.IGNORECASE)
            if end_match:
                end_date_str = end_match.group(1).strip()
                line_clean = re.sub(r'(?:end\s*date\s*|ends\s*|\()?\d{1,2}/\d{1,2}(?:/\d{2,4})?\)?', '', line_clean, flags=re.IGNORECASE).strip()

        # FALLBACK: If Start Date is not keyed in, pick the Key-In / Submission Date
        if not start_date_str and submission_date:
            if isinstance(submission_date, (pd.Timestamp, datetime)):
                start_date_str = submission_date.strftime("%d/%m/%Y")
            elif isinstance(submission_date, str) and submission_date != "Recent Entry":
                try:
                    dt_parsed = pd.to_datetime(submission_date, errors='coerce')
                    if pd.notna(dt_parsed):
                        start_date_str = dt_parsed.strftime("%d/%m/%Y")
                except Exception:
                    pass

        # Helper to convert date strings to datetime objects
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

        # Calculate duration in days
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
            p1 = float(m3.group(2))
            p2 = float(m3.group(3))
            if p1 > p2:
                old_price, new_price = p1, p2
                found = True

        if found and old_price and new_price and old_price > new_price:
            discount_amount = old_price - new_price
            discount_pct = (discount_amount / old_price) * 100
            parsed_items.append({
                "product": product,
                "old_price": old_price,
                "new_price": new_price,
                "discount_amount": discount_amount,
                "discount_pct": discount_pct,
                "start_date": start_date_str,
                "end_date": end_date_str,
                "duration_days": duration_days,
                "raw_line": line_clean
            })
        else:
            parsed_items.append({
                "product": line_clean,
                "old_price": None,
                "new_price": None,
                "discount_amount": None,
                "discount_pct": None,
                "start_date": start_date_str,
                "end_date": end_date_str,
                "duration_days": duration_days,
                "raw_line": line_clean
            })
            
    return parsed_items

# --- CACHED HIGH-SPEED IMAGE DOWNLOADER ---
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_and_resize_image(url):
    """Fetch an image and convert it to a resized byte buffer with caching."""
    if not url or pd.isna(url) or str(url).strip() in ["", "nan", "None"]:
        return None
        
    auth_url = f"{url}{'&' if '?' in url else '?'}apiKey={JOTFORM_API_KEY}" if "jotform.com" in url else url
    headers = {"User-Agent": "Mozilla/5.0"}
    
    try:
        res = requests.get(auth_url, headers=headers, timeout=5)
        if res.status_code == 200:
            img = Image.open(BytesIO(res.content))
            img.thumbnail((400, 400))
            img_bytes = BytesIO()
            img.save(img_bytes, format='PNG')
            img_bytes.seek(0)
            return img_bytes.getvalue()
    except Exception:
        pass
    return None

def fetch_all_images_parallel(urls):
    """Concurrently download and process images using thread pools."""
    unique_urls = list(set([u for u in urls if u]))
    url_to_bytes = {}
    
    with ThreadPoolExecutor(max_workers=10) as executor:
        future_to_url = {executor.submit(fetch_and_resize_image, url): url for url in unique_urls}
        for future in as_completed(future_to_url):
            url = future_to_url[future]
            try:
                img_data = future.result()
                if img_data:
                    url_to_bytes[url] = img_data
            except Exception:
                pass
    return url_to_bytes

# --- EXCEL IMAGE EMBEDDING GENERATOR ---
def generate_excel_with_images(comp_rows, paid_rows, raw_df):
    wb = openpyxl.Workbook()
    all_urls = [item.get("image_url") for item in comp_rows + paid_rows if item.get("image_url")]
    image_cache = fetch_all_images_parallel(all_urls)
    
    # 1. Competitor Activities Sheet
    ws_comp = wb.active
    ws_comp.title = "Competitor Depth & Activity"
    ws_comp.append(["Date", "Outlet", "Field Rep", "Activity", "Was", "Now", "Depth", "Start Date", "Offer End Date", "Promo Duration", "Embedded Image"])
    
    for row_idx, item in enumerate(comp_rows, start=2):
        ws_comp.append([
            item["Date"], item["Outlet"], item["Field Rep"], 
            item["Activity"], item["Was"], item["Now"], item["Depth"], 
            item.get("Start Date", "N/A"), item.get("End Date", "N/A"),
            item.get("Duration Days", "N/A"), ""
        ])
        ws_comp.row_dimensions[row_idx].height = 60
        
        img_url = item.get("image_url")
        if img_url in image_cache:
            try:
                img_bytes = BytesIO(image_cache[img_url])
                xl_img = OpenPyxlImage(img_bytes)
                ws_comp.add_image(xl_img, f"K{row_idx}")
            except Exception:
                ws_comp.cell(row=row_idx, column=11, value="[Image Error]")
        elif img_url:
            ws_comp.cell(row=row_idx, column=11, value="[Image Unavailable]")

    ws_comp.column_dimensions['K'].width = 15

    # 2. Paid Assets Sheet
    ws_paid = wb.create_sheet(title="Paid Asset Executions")
    ws_paid.append(["Date", "Outlet", "Field Rep", "Paid Assets Executed", "Embedded Image"])
    
    for row_idx, item in enumerate(paid_rows, start=2):
        ws_paid.append([
            item["Date"], item["Outlet"], item["Field Rep"], 
            item["Paid Assets Executed"], ""
        ])
        ws_paid.row_dimensions[row_idx].height = 60
        
        img_url = item.get("image_url")
        if img_url in image_cache:
            try:
                img_bytes = BytesIO(image_cache[img_url])
                xl_img = OpenPyxlImage(img_bytes)
                ws_paid.add_image(xl_img, f"E{row_idx}")
            except Exception:
                ws_paid.cell(row=row_idx, column=5, value="[Image Error]")
        elif img_url:
            ws_paid.cell(row=row_idx, column=5, value="[Image Unavailable]")

    ws_paid.column_dimensions['E'].width = 15

    # 3. Raw Data Sheet
    ws_raw = wb.create_sheet(title="Raw Data")
    ws_raw.append(list(raw_df.columns))
    for r in raw_df.itertuples(index=False):
        ws_raw.append(list(r))
        
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()

# --- POWERPOINT (.PPTX) GENERATOR MATCHING TEMPLATE_2 (PAID ASSETS ONLY - NO RED) ---
def generate_pptx_report(paid_rows):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_slide_layout = prs.slide_layouts[6]
    
    # Pre-fetch Images for Paid Assets
    all_urls = [item.get("image_url") for item in paid_rows if item.get("image_url")]
    image_cache = fetch_all_images_parallel(all_urls)
    
    # Custom Palette Extracted from PPTX Template 2 (Strictly No Red)
    DARK_NAVY = RGBColor(15, 23, 42)      # Primary Header & Text (#0F172A)
    SLATE_MUTED = RGBColor(100, 116, 139) # Secondary Details (#64748B)
    WHITE = RGBColor(255, 255, 255)

    # ----------------------------------------------------
    # SLIDE 1: TITLE SLIDE (Matching Slide 1 in Template 2)
    # ----------------------------------------------------
    slide_title = prs.slides.add_slide(blank_slide_layout)
    
    # Clean Title Text Box
    tb_title = slide_title.shapes.add_textbox(Inches(0.8), Inches(2.5), Inches(11.733), Inches(3.0))
    tf_title = tb_title.text_frame
    tf_title.word_wrap = True
    
    p1 = tf_title.paragraphs[0]
    p1.text = "IB WEEKLY FIELD FINDINGS"
    p1.font.size = Pt(38)
    p1.font.bold = True
    p1.font.color.rgb = DARK_NAVY
    p1.font.name = "Arial"
    
    p2 = tf_title.add_paragraph()
    p2.text = f"\nPaid Asset Executions Review | {datetime.now().strftime('%d %B %Y')}"
    p2.font.size = Pt(20)
    p2.font.color.rgb = SLATE_MUTED
    p2.font.name = "Arial"

    # Bottom Dark Accent Bar
    bottom_bar = slide_title.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(6.8), Inches(13.333), Inches(0.7))
    bottom_bar.fill.solid()
    bottom_bar.fill.fore_color.rgb = DARK_NAVY
    bottom_bar.line.color.rgb = DARK_NAVY

    # ----------------------------------------------------
    # SLIDE 2+: PAID ASSET PHOTO GRID SLIDES (Matching Template 2x2 Grid)
    # ----------------------------------------------------
    items_with_images = [item for item in paid_rows if item.get("image_url") in image_cache]
    if not items_with_images:
        items_with_images = paid_rows

    chunks = [items_with_images[i:i + 4] for i in range(0, len(items_with_images), 4)]
    
    grid_positions = [
        {"left": Inches(0.8), "top": Inches(1.3), "width": Inches(5.6), "height": Inches(2.7)},   # Top Left
        {"left": Inches(6.8), "top": Inches(1.3), "width": Inches(5.6), "height": Inches(2.7)},   # Top Right
        {"left": Inches(0.8), "top": Inches(4.3), "width": Inches(5.6), "height": Inches(2.7)},   # Bottom Left
        {"left": Inches(6.8), "top": Inches(4.3), "width": Inches(5.6), "height": Inches(2.7)}    # Bottom Right
    ]

    for slide_num, chunk in enumerate(chunks, start=1):
        slide = prs.slides.add_slide(blank_slide_layout)
        
        # Header Text Matching Template 2
        tb_hdr = slide.shapes.add_textbox(Inches(0.8), Inches(0.3), Inches(11.733), Inches(0.8))
        tf_hdr = tb_hdr.text_frame
        p_hdr = tf_hdr.paragraphs[0]
        p_hdr.text = f"PAID ASSETS EXECUTIONS (PAGE {slide_num})"
        p_hdr.font.size = Pt(22)
        p_hdr.font.bold = True
        p_hdr.font.color.rgb = DARK_NAVY
        p_hdr.font.name = "Arial"

        for idx, item in enumerate(chunk):
            pos = grid_positions[idx]
            img_url = item.get("image_url")
            
            if img_url in image_cache:
                try:
                    img_bytes = BytesIO(image_cache[img_url])
                    slide.shapes.add_picture(img_bytes, pos["left"], pos["top"], pos["width"], pos["height"])
                except Exception:
                    pass
            
            # Muted Overlay Label at bottom of image
            bg_rect = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, pos["left"], pos["top"] + pos["height"] - Inches(0.4), pos["width"], Inches(0.4))
            bg_rect.fill.solid()
            bg_rect.fill.fore_color.rgb = DARK_NAVY
            bg_rect.line.fill.background()

            cap_box = slide.shapes.add_textbox(pos["left"], pos["top"] + pos["height"] - Inches(0.4), pos["width"], Inches(0.4))
            tf_cap = cap_box.text_frame
            p_cap = tf_cap.paragraphs[0]
            p_cap.text = f" 📍 {item['Outlet']} | Rep: {item['Field Rep']}"
            p_cap.font.size = Pt(10)
            p_cap.font.bold = True
            p_cap.font.color.rgb = WHITE
            p_cap.font.name = "Arial"

    ppt_output = BytesIO()
    prs.save(ppt_output)
    ppt_output.seek(0)
    return ppt_output.getvalue()

# 2. Load Data Function
@st.cache_data(ttl=300)
def load_data():
    CSV_URL = "https://docs.google.com/spreadsheets/d/1XeWdaYVyLdqfgF00nKBOjZ0OlhtcqRHb6HO0b1lND0Y/export?format=csv&gid=284416453"
    df = pd.read_csv(CSV_URL)
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    
    date_col = "Submission Date" if "Submission Date" in df.columns else None
    if not date_col:
        for c in df.columns:
            if "submission date" in c.lower() or ("date" in c.lower() and "id" not in c.lower()):
                date_col = c
                break

    if date_col:
        df['Parsed_Date'] = pd.to_datetime(df[date_col], errors='coerce')
        valid_dates_mask = (df['Parsed_Date'].dt.year >= 2020) & (df['Parsed_Date'].dt.year <= 2030)
        df.loc[~valid_dates_mask, 'Parsed_Date'] = pd.NaT
        
        df['Formatted_Timestamp'] = df['Parsed_Date'].dt.strftime('%d %b %Y, %H:%M')
        df['Formatted_Timestamp'] = df['Formatted_Timestamp'].fillna("Recent Entry")
        df = df.sort_values(by='Parsed_Date', ascending=False, na_position='last').reset_index(drop=True)
    else:
        df = df.iloc[::-1].reset_index(drop=True)
        df['Parsed_Date'] = None
        df['Formatted_Timestamp'] = "Recent Entry"
        
    return df

try:
    df = load_data()
except Exception as e:
    st.error(f"Failed to load live data: {e}")
    st.stop()

# --- 3. SIDEBAR FILTERS ---
st.sidebar.title("📅 Date")

valid_dates = df['Parsed_Date'].dropna() if 'Parsed_Date' in df.columns else pd.Series()
if not valid_dates.empty:
    min_date = valid_dates.min().date()
    max_date = valid_dates.max().date()
    
    date_range = st.sidebar.date_input(
        "Select Date Range",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
else:
    date_range = None

st.sidebar.divider()

filtered_df = df.copy()

if 'Select Your Project' in df.columns:
    project_list = ["All"] + sorted(df['Select Your Project'].dropna().unique().tolist())
    st.sidebar.markdown("### 📂 Project")
    selected_project_button = st.sidebar.segmented_control(
        "Filter by Project:",
        options=project_list,
        default="All"
    )
    if selected_project_button and selected_project_button != "All":
        filtered_df = filtered_df[filtered_df['Select Your Project'] == selected_project_button]

if 'Select Your Name' in df.columns:
    rep_list = ["All"] + sorted(df['Select Your Name'].dropna().unique().tolist())
    st.sidebar.markdown("### 👤 Field Rep")
    selected_rep_button = st.sidebar.segmented_control(
        "Filter by Representative:",
        options=rep_list,
        default="All"
    )
    if selected_rep_button and selected_rep_button != "All":
        filtered_df = filtered_df[filtered_df['Select Your Name'] == selected_rep_button]

if 'Type Outlet Name' in df.columns:
    st.sidebar.markdown("### 🏬 Outlets")
    outlet_list = sorted(filtered_df['Type Outlet Name'].dropna().unique().tolist())
    selected_outlets = st.sidebar.multiselect("Filter Outlets:", outlet_list, default=outlet_list)
    if selected_outlets:
        filtered_df = filtered_df[filtered_df['Type Outlet Name'].isin(selected_outlets)]

if date_range and isinstance(date_range, tuple) and len(date_range) == 2 and not valid_dates.empty:
    start_date, end_date = date_range
    filtered_df = filtered_df[
        (filtered_df['Parsed_Date'].dt.date >= start_date) & 
        (filtered_df['Parsed_Date'].dt.date <= end_date)
    ]

st.sidebar.divider()
st.sidebar.caption("⚡ Live sync active (Google Sheets & Jotform API)")

# --- DASHBOARD HEADER & HIGH-LEVEL METRICS ---
st.title("📊 Competitor Activities and Paid Assets")

total_submissions = len(filtered_df)
paid_col = "Does the store have any paid assets"
comp_col = "Are there any Competitor activities in the outlet"

paid_assets_count = (filtered_df[filtered_df[paid_col].astype(str).str.strip().str.lower() == "yes"].shape[0] if paid_col in filtered_df.columns else 0)
comp_activities_count = (filtered_df[filtered_df[comp_col].astype(str).str.strip().str.lower() == "yes"].shape[0] if comp_col in filtered_df.columns else 0)

col1, col2, col3 = st.columns(3)
col1.metric("Total Submissions", total_submissions)
col2.metric("Outlets with Paid Assets", paid_assets_count)
col3.metric("Competitor Activities Reported", comp_activities_count)

st.divider()

# --- PRE-PROCESS EXPORT ROWS ---
all_image_columns = [col for col in filtered_df.columns if "Take Photo" in str(col)]
comp_image_cols = [c for c in all_image_columns if "competitor" in str(c).lower()]
paid_image_cols = [c for c in all_image_columns if "paid asset" in str(c).lower()]
desc_col = "Describe the competitor activities"
asset_select_col = "Select Paid assets in the outlet"

global_comp_rows = []
if comp_col in filtered_df.columns and desc_col in filtered_df.columns:
    comp_df_export = filtered_df[filtered_df[comp_col].astype(str).str.strip().str.lower() == "yes"]
    for _, row in comp_df_export.iterrows():
        outlet = row.get("Type Outlet Name", "Unknown Outlet")
        rep = row.get("Select Your Name", "Unknown Rep")
        activity = row.get(desc_col, "No description provided.")
        timestamp_str = row.get('Formatted_Timestamp', 'Recent Entry')
        
        parsed_discounts = parse_price_discounts(
            str(activity), 
            submission_date=row.get('Parsed_Date', timestamp_str)
        ) if pd.notna(activity) else []
        
        first_img = None
        for img_col in comp_image_cols:
            img_val = row.get(img_col)
            if pd.notna(img_val) and str(img_val).strip() != "" and str(img_val).strip().lower() != "nan":
                first_img = str(img_val).split(",")[0].strip()
                break

        if parsed_discounts:
            for item in parsed_discounts:
                global_comp_rows.append({
                    "Date": timestamp_str, "Outlet": outlet, "Field Rep": rep,
                    "Activity": item["product"],
                    "Was": f"KSh {item['old_price']:.0f}" if item["old_price"] else "N/A",
                    "Now": f"KSh {item['new_price']:.0f}" if item["new_price"] else "N/A",
                    "Depth": f"Save KSh {item['discount_amount']:.0f} ({item['discount_pct']:.1f}% OFF)" if item["discount_amount"] else "N/A",
                    "Start Date": item["start_date"] if item["start_date"] else "N/A",
                    "End Date": item["end_date"] if item["end_date"] else "N/A",
                    "Duration Days": f"{item['duration_days']} Days" if item["duration_days"] else "N/A",
                    "image_url": first_img
                })
        else:
            global_comp_rows.append({
                "Date": timestamp_str, "Outlet": outlet, "Field Rep": rep,
                "Activity": "Competitor activity confirmed (no notes entered).",
                "Was": "N/A", "Now": "N/A", "Depth": "N/A", 
                "Start Date": "N/A", "End Date": "N/A", "Duration Days": "N/A",
                "image_url": first_img
            })

global_paid_rows = []
if paid_col in filtered_df.columns and asset_select_col in filtered_df.columns:
    paid_df_export = filtered_df[filtered_df[paid_col].astype(str).str.strip().str.lower() == "yes"]
    for _, row in paid_df_export.iterrows():
        outlet = row.get("Type Outlet Name", "Unknown Outlet")
        rep = row.get("Select Your Name", "Unknown Rep")
        assets = row.get(asset_select_col, "Paid asset confirmed.")
        timestamp_str = row.get('Formatted_Timestamp', 'Recent Entry')
        
        first_img = None
        for img_col in paid_image_cols:
            img_val = row.get(img_col)
            if pd.notna(img_val) and str(img_val).strip() != "" and str(img_val).strip().lower() != "nan":
                first_img = str(img_val).split(",")[0].strip()
                break

        global_paid_rows.append({
            "Date": timestamp_str, "Outlet": outlet, "Field Rep": rep,
            "Paid Assets Executed": assets, "image_url": first_img
        })

# --- TOP NAVIGATION TABS ---
nav_tab1, nav_tab2, nav_tab3 = st.tabs([
    "📦 Summary", 
    "🚨 Competitor Activities & Paid Assets", 
    "📸 Gallery"
])

# ==========================================
# PAGE 1: SUMMARY
# ==========================================
with nav_tab1:
    st.markdown("### 📦 Paid Asset Distribution Breakdown")
    if asset_select_col in filtered_df.columns:
        assets_series = filtered_df[asset_select_col].dropna().astype(str)
        assets_exploded = assets_series.str.split(",").explode().str.strip()
        assets_exploded = assets_exploded[assets_exploded != ""]

        if not assets_exploded.empty:
            asset_counts = assets_exploded.value_counts().reset_index()
            asset_counts.columns = ["Asset Type", "Count"]

            col_chart, col_data = st.columns([2, 1])
            with col_chart:
                st.bar_chart(asset_counts.set_index("Asset Type"))
            with col_data:
                st.dataframe(asset_counts, hide_index=True, use_container_width=True)
        else:
            st.info("No paid assets recorded for the selected filters.")

# ==========================================
# PAGE 2: COMPETITOR ACTIVITIES & PAID ASSETS
# ==========================================
with nav_tab2:
    tab_comp, tab_paid = st.tabs(["🚨 Competitor Activity & Price Depth", "📦 Paid Asset Executions"])

    # SUB-TAB 1: COMPETITOR ACTIVITY
    with tab_comp:
        if comp_col in filtered_df.columns and desc_col in filtered_df.columns:
            comp_df = filtered_df[filtered_df[comp_col].astype(str).str.strip().str.lower() == "yes"]
            
            if not comp_df.empty:
                comp_summary_col1, comp_summary_col2 = st.columns([2.2, 1])
                
                with comp_summary_col1:
                    st.markdown("##### Detailed Competitor Price Cuts & Depth Analysis")
                    for _, row in comp_df.iterrows():
                        outlet = row.get("Type Outlet Name", "Unknown Outlet")
                        rep = row.get("Select Your Name", "Unknown Rep")
                        activity = row.get(desc_col, "No description provided.")
                        timestamp_str = row.get('Formatted_Timestamp', 'Recent Entry')
                        
                        parsed_discounts = parse_price_discounts(
                            str(activity), 
                            submission_date=row.get('Parsed_Date', timestamp_str)
                        ) if pd.notna(activity) else []

                        with st.container():
                            header_col1, header_col2 = st.columns([3, 1])
                            with header_col1:
                                st.markdown(f"📍 **{outlet}** &nbsp;*(Rep: {rep})*", unsafe_allow_html=True)
                            with header_col2:
                                st.markdown(f"<div style='text-align: right;'><span class='timestamp-badge'>⏱️ {timestamp_str}</span></div>", unsafe_allow_html=True)
                            
                            if parsed_discounts:
                                st.markdown("**Reported Pricing / Activities:**")
                                for item in parsed_discounts:
                                    start_date_html = f"<span class='enddate-badge'>🚀 Starts: {item['start_date']}</span>" if item.get("start_date") else ""
                                    end_date_html = f"<span class='enddate-badge'>⏳ Ends: {item['end_date']}</span>" if item.get("end_date") else ""
                                    duration_html = f"<span class='discount-badge'>📅 Duration: {item['duration_days']} Days</span>" if item.get("duration_days") else ""
                                    
                                    if item.get("discount_amount"):
                                        st.markdown(
                                            f"- **{item['product']}**: Now **KSh {item['new_price']:.0f}** "
                                            f"<span style='text-decoration: line-through; color: #888;'>from KSh {item['old_price']:.0f}</span> "
                                            f"<span class='discount-badge'>Save KSh {item['discount_amount']:.0f} ({item['discount_pct']:.1f}% OFF)</span> "
                                            f"{start_date_html} {end_date_html} {duration_html}", 
                                            unsafe_allow_html=True
                                        )
                                    else:
                                        st.markdown(f"- {item['raw_line']} {start_date_html} {end_date_html} {duration_html}", unsafe_allow_html=True)
                            else:
                                st.info("Competitor activity confirmed (no text notes entered).")
                            
                            comp_images = []
                            for img_col in comp_image_cols:
                                img_val = row.get(img_col)
                                if pd.notna(img_val) and str(img_val).strip() != "" and str(img_val).strip().lower() != "nan":
                                    for u in str(img_val).split(","):
                                        if u.strip():
                                            title = str(img_col).replace("Take Photo of competitor activites", "Competitor Activity").strip()
                                            comp_images.append((title, u.strip()))
                            
                            if comp_images:
                                st.markdown("<div style='margin-top: 8px;'></div>", unsafe_allow_html=True)
                                img_cols = st.columns(min(len(comp_images), 4))
                                for idx, (img_title, raw_url) in enumerate(comp_images):
                                    auth_url = f"{raw_url}{'&' if '?' in raw_url else '?'}apiKey={JOTFORM_API_KEY}" if "jotform.com" in raw_url else raw_url
                                    with img_cols[idx % 4]:
                                        st.markdown(f'''
                                            <div style="margin-bottom: 8px;">
                                                <span style="font-size: 11px; font-weight: 600; color: #0f172a;">{img_title}</span>
                                                <a href="{auth_url}" target="_blank">
                                                    <img src="{auth_url}" width="100%" style="border-radius: 6px; border: 1px solid #cbd5e1; margin-top: 2px;">
                                                </a>
                                            </div>
                                        ''', unsafe_allow_html=True)
                            st.divider()
                    
                with comp_summary_col2:
                    st.markdown("##### 📥 Export Package")
                    st.info("Click below to compile and download the `.xlsx` report package with embedded photos.")
                    
                    if st.button("📊 Generate Excel Package", key="comp_excel_btn"):
                        with st.spinner("Compiling images concurrently and building Excel workbook..."):
                            excel_bytes = generate_excel_with_images(global_comp_rows, global_paid_rows, filtered_df)
                            st.download_button(
                                label="💾 Download Excel Workbook (.xlsx)",
                                data=excel_bytes,
                                file_name="Competitor_And_Paid_Assets_Report.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="comp_download_btn"
                            )
            else:
                st.info("No competitor activities reported for the selected filters.")

    # SUB-TAB 2: PAID ASSETS EXECUTIONS
    with tab_paid:
        if paid_col in filtered_df.columns:
            paid_df = filtered_df[filtered_df[paid_col].astype(str).str.strip().str.lower() == "yes"]
            if asset_select_col in paid_df.columns:
                paid_df = paid_df[paid_df[asset_select_col].notna() & (paid_df[asset_select_col].astype(str).str.strip() != "")]
            
            if not paid_df.empty:
                paid_col_left, paid_col_right = st.columns([2.2, 1])
                
                with paid_col_left:
                    for _, row in paid_df.iterrows():
                        outlet = row.get("Type Outlet Name", "Unknown Outlet")
                        rep = row.get("Select Your Name", "Unknown Rep")
                        assets = row.get(asset_select_col, "Paid asset confirmed.")
                        timestamp_str = row.get('Formatted_Timestamp', 'Recent Entry')

                        with st.container():
                            header_col1, header_col2 = st.columns([3, 1])
                            with header_col1:
                                st.markdown(f"📍 **{outlet}** &nbsp;*(Rep: {rep})*", unsafe_allow_html=True)
                            with header_col2:
                                st.markdown(f"<div style='text-align: right;'><span class='timestamp-badge'>⏱️ {timestamp_str}</span></div>", unsafe_allow_html=True)
                            
                            st.markdown(f"**Asset Types Executed:** `{assets}`")
                            
                            paid_images = []
                            for img_col in paid_image_cols:
                                img_val = row.get(img_col)
                                if pd.notna(img_val) and str(img_val).strip() != "" and str(img_val).strip().lower() != "nan":
                                    for u in str(img_val).split(","):
                                        if u.strip():
                                            title = str(img_col).replace("Take Photo of paid assets", "Paid Asset").replace("(", "").replace(")", "").strip()
                                            paid_images.append((title, u.strip()))
                            
                            if paid_images:
                                st.markdown("<div style='margin-top: 8px;'></div>", unsafe_allow_html=True)
                                img_cols = st.columns(min(len(paid_images), 4))
                                for idx, (img_title, raw_url) in enumerate(paid_images):
                                    auth_url = f"{raw_url}{'&' if '?' in raw_url else '?'}apiKey={JOTFORM_API_KEY}" if "jotform.com" in raw_url else raw_url
                                    with img_cols[idx % 4]:
                                        st.markdown(f'''
                                            <div style="margin-bottom: 8px;">
                                                <span style="font-size: 11px; font-weight: 600; color: #0f172a;">{img_title}</span>
                                                <a href="{auth_url}" target="_blank">
                                                    <img src="{auth_url}" width="100%" style="border-radius: 6px; border: 1px solid #cbd5e1; margin-top: 2px;">
                                                </a>
                                            </div>
                                        ''', unsafe_allow_html=True)
                            st.divider()

                with paid_col_right:
                    st.markdown("##### 📥 Export Package")
                    st.info("Click below to compile Excel or PowerPoint review decks with embedded photos.")
                    
                    if st.button("📊 Generate Excel Package", key="paid_excel_btn"):
                        with st.spinner("Compiling images concurrently and building Excel workbook..."):
                            excel_bytes = generate_excel_with_images(global_comp_rows, global_paid_rows, filtered_df)
                            st.download_button(
                                label="💾 Download Excel Workbook (.xlsx)",
                                data=excel_bytes,
                                file_name="Competitor_And_Paid_Assets_Report.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="paid_download_btn"
                            )

                    if st.button("📈 Generate Paid Assets PPT Review", key="paid_ppt_btn"):
                        with st.spinner("Building CIL branded PowerPoint deck..."):
                            ppt_bytes = generate_pptx_report(global_paid_rows)
                            st.download_button(
                                label="💾 Download Presentation Deck (.pptx)",
                                data=ppt_bytes,
                                file_name="Chandaria_Paid_Assets_Review.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                key="paid_ppt_download_btn"
                            )
            else:
                st.info("No outlets with paid assets reported for the selected filters.")

# ==========================================
# PAGE 3: GALLERY
# ==========================================
with nav_tab3:
    st.markdown("### 📸 Gallery")
    image_type_choice = st.segmented_control(
        "Select Photo Type:",
        options=["All Images", "Paid Assets", "Competitor Activities"],
        default="All Images",
        key="gallery_photo_type"
    )

    if image_type_choice == "Paid Assets":
        active_image_cols = [c for c in all_image_columns if "paid asset" in str(c).lower()]
    elif image_type_choice == "Competitor Activities":
        active_image_cols = [c for c in all_image_columns if "competitor" in str(c).lower()]
    else:
        active_image_cols = all_image_columns

    if active_image_cols:
        selected_img_col = st.selectbox("Select Category Column:", active_image_cols)
        df_images = filtered_df.dropna(subset=[selected_img_col])

        if df_images.empty:
            st.info("No images submitted for this category under current filters.")
        else:
            cols = st.columns(4)
            col_index = 0 
            
            for _, row in df_images.iterrows():
                img_url_raw = row[selected_img_col]
                outlet_name = row.get("Type Outlet Name", "Unknown Outlet")
                timestamp_str = row.get('Formatted_Timestamp', 'Recent Entry')
                
                if pd.notna(img_url_raw) and str(img_url_raw).strip() != "" and str(img_url_raw).strip().lower() != "nan":
                    first_url = str(img_url_raw).split(",")[0].strip()
                    auth_url = f"{first_url}{'&' if '?' in first_url else '?'}apiKey={JOTFORM_API_KEY}" if "jotform.com" in first_url else first_url

                    with cols[col_index % 4]:
                        st.markdown(f"**{outlet_name}**")
                        st.caption(f"⏱️ {timestamp_str}")
                        st.markdown(f'''
                            <a href="{auth_url}" target="_blank">
                                <img src="{auth_url}" width="100%" style="border-radius: 6px; border: 1px solid #ced4da;">
                            </a>
                            <br>
                            <a href="{auth_url}" target="_blank" style="font-size: 12px; text-decoration: none; color: #0d6efd;">🔗 View Original</a>
                        ''', unsafe_allow_html=True)
                    col_index += 1