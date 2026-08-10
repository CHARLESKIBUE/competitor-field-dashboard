import pandas as pd
import streamlit as st
import re

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
    /* Styled Sidebar */
    div[data-testid="stSidebar"] {
        background-color: #f8f9fa;
        padding-top: 15px;
    }
    div[data-testid="stSidebar"] label {
        font-weight: 600 !important;
        color: #2c3e50 !important;
    }
    
    /* Prominent, Bold, Custom Styled Tabs */
    button[data-baseweb="tab"] {
        font-size: 18px !important;
        font-weight: 800 !important;
        color: #1e293b !important;
        padding: 10px 20px !important;
        font-family: 'Segoe UI', Roboto, Helvetica, Arial, sans-serif !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        color: #e74c3c !important;
        border-bottom-color: #e74c3c !important;
    }

    /* Cards Styling */
    .comp-card {
        background-color: #ffffff;
        border-left: 4px solid #e74c3c;
        padding: 16px 20px;
        margin-bottom: 16px;
        border-radius: 8px;
        box-shadow: 0 2px 5px rgba(0,0,0,0.06);
        border: 1px solid #f1f3f5;
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
        background-color: #fee2e2;
        color: #991b1b;
        font-size: 12px;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 6px;
        border: 1px solid #fca5a5;
        margin-left: 6px;
    }
    </style>
""", unsafe_allow_html=True)

# Helper function to parse pricing discount depth
def parse_price_discounts(text):
    if not isinstance(text, str) or pd.isna(text):
        return []
    
    parsed_items = []
    lines = text.split('\n')
    
    for line in lines:
        line_clean = line.strip('-* ').strip()
        if not line_clean:
            continue
            
        found = False
        # Pattern 1: @NEW from OLD or now NEW from OLD
        match1 = re.search(r"(.+?)(?:@|now\s*)\s*(\d+)\s*(?:from|was)\s*(\d+)", line_clean, re.IGNORECASE)
        if match1:
            product = match1.group(1).strip()
            new_price = float(match1.group(2))
            old_price = float(match1.group(3))
            found = True
        else:
            # Pattern 2: was OLD now NEW
            match2 = re.search(r"(.+?)(?:was\s*)\s*(\d+)\s*(?:now\s*|@\s*)\s*(\d+)", line_clean, re.IGNORECASE)
            if match2:
                product = match2.group(1).strip()
                old_price = float(match2.group(2))
                new_price = float(match2.group(3))
                found = True
                
        if found and old_price > new_price:
            discount_amount = old_price - new_price
            discount_pct = (discount_amount / old_price) * 100
            parsed_items.append({
                "product": product,
                "old_price": old_price,
                "new_price": new_price,
                "discount_amount": discount_amount,
                "discount_pct": discount_pct,
                "raw_line": line_clean
            })
        elif not found:
            parsed_items.append({"raw_line": line_clean})
            
    return parsed_items

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
    
    # Label cleaned up (removed "Submissions Window:")
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

# --- DASHBOARD HEADER ---
st.title("📊 Competitor Activities and Paid Assets")
st.markdown("Real-time operational field reports sorted chronologically (latest entries first).")

# --- 4. HIGH-LEVEL METRICS ---
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

# --- 5. BOLD TAB HEADERS (Store Reports Segregation Header Removed) ---
tab_comp, tab_paid = st.tabs(["🚨 Competitor Activity & Price Depth", "📦 Paid Asset Executions"])

all_image_columns = [col for col in filtered_df.columns if "Take Photo" in str(col)]

# --- TAB 1: COMPETITOR ACTIVITY & DISCOUNT DEPTH ---
with tab_comp:
    desc_col = "Describe the competitor activities"
    comp_image_cols = [c for c in all_image_columns if "competitor" in str(c).lower()]
    
    if comp_col in filtered_df.columns and desc_col in filtered_df.columns:
        # Strictly exempt outlets that answered 'No' to competitor activities
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

                    with st.container():
                        header_col1, header_col2 = st.columns([3, 1])
                        with header_col1:
                            st.markdown(f"📍 **{outlet}** &nbsp;*(Rep: {rep})*", unsafe_allow_html=True)
                        with header_col2:
                            st.markdown(f"<div style='text-align: right;'><span class='timestamp-badge'>⏱️ {timestamp_str}</span></div>", unsafe_allow_html=True)
                        
                        # Price Depth Parsing
                        if pd.notna(activity) and str(activity).strip().lower() != "nan":
                            parsed_discounts = parse_price_discounts(str(activity))
                            
                            st.markdown("**Reported Pricing / Activities:**")
                            for item in parsed_discounts:
                                if "discount_amount" in item:
                                    st.markdown(
                                        f"- **{item['product']}**: Now **KSh {item['new_price']:.0f}** "
                                        f"<span style='text-decoration: line-through; color: #888;'>from KSh {item['old_price']:.0f}</span> "
                                        f"<span class='discount-badge'>Save KSh {item['discount_amount']:.0f} ({item['discount_pct']:.1f}% OFF)</span>", 
                                        unsafe_allow_html=True
                                    )
                                else:
                                    st.markdown(f"- {item['raw_line']}")
                        else:
                            st.info("Competitor activity confirmed (no text notes entered).")
                        
                        # Only show Competitor photos here
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
                                            <span style="font-size: 11px; font-weight: 600; color: #e74c3c;">{img_title}</span>
                                            <a href="{auth_url}" target="_blank">
                                                <img src="{auth_url}" width="100%" style="border-radius: 6px; border: 1px solid #fca5a5; margin-top: 2px;">
                                            </a>
                                        </div>
                                    ''', unsafe_allow_html=True)
                        st.divider()
                    
            with comp_summary_col2:
                st.markdown("##### Export Options")
                st.info("Export competitor activity logs directly to CSV.")
                
                export_df = comp_df[['Type Outlet Name', 'Select Your Name', desc_col, 'Formatted_Timestamp']].rename(columns={
                    'Type Outlet Name': 'Outlet',
                    'Select Your Name': 'Field Rep',
                    desc_col: 'Competitor Activity Details',
                    'Formatted_Timestamp': 'Timestamp'
                })
                csv_data = export_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Competitor Summary (CSV)",
                    data=csv_data,
                    file_name="Competitor_Activity_Report.csv",
                    mime="text/csv",
                )
        else:
            st.info("No competitor activities reported for the selected filters.")

# --- TAB 2: PAID ASSET EXECUTIONS ---
with tab_paid:
    asset_select_col = "Select Paid assets in the outlet"
    paid_image_cols = [c for c in all_image_columns if "paid asset" in str(c).lower()]
    
    if paid_col in filtered_df.columns:
        # Strictly exempt outlets that answered 'No' to paid assets or have empty asset data
        paid_df = filtered_df[filtered_df[paid_col].astype(str).str.strip().str.lower() == "yes"]
        
        # Exclude outlets where asset text is blank/NaN
        if asset_select_col in paid_df.columns:
            paid_df = paid_df[paid_df[asset_select_col].notna() & (paid_df[asset_select_col].astype(str).str.strip() != "")]
        
        if not paid_df.empty:
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
                    
                    # Only show Paid Asset photos here
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
                                        <span style="font-size: 11px; font-weight: 600; color: #2ecc71;">{img_title}</span>
                                        <a href="{auth_url}" target="_blank">
                                            <img src="{auth_url}" width="100%" style="border-radius: 6px; border: 1px solid #86efac; margin-top: 2px;">
                                        </a>
                                    </div>
                                ''', unsafe_allow_html=True)
                    st.divider()
        else:
            st.info("No outlets with paid assets reported for the selected filters.")

st.divider()

# --- 6. TYPES OF PAID ASSETS CHARTS ---
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
            st.dataframe(asset_counts, hide_index=True, width='stretch')
    else:
        st.info("No paid assets recorded for the selected filters.")

st.divider()

# --- 7. FIELD GALLERY BY CATEGORY ---
st.markdown("### 📸 Authenticated Field Gallery")
image_type_choice = st.segmented_control(
    "Select Photo Type:",
    options=["All Images", "Paid Assets", "Competitor Activities"],
    default="All Images"
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