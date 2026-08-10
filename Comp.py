import pandas as pd
import streamlit as st

# --- JOTFORM API KEY ---
JOTFORM_API_KEY = "6666042ab61ea98f9b378a4a7d"

# 1. Page Configuration
st.set_page_config(
    page_title="Competitor Activities and Paid Assets", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# Custom CSS for UI styling
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
    </style>
""", unsafe_allow_html=True)

# 2. Parse Data Function using exact "Submission Date" Column & CSV Export URL
@st.cache_data(ttl=300)
def load_data():
    # Correct CSV export link format for Google Sheets
    CSV_URL = "https://docs.google.com/spreadsheets/d/1XeWdaYVyLdqfgF00nKBOjZ0OlhtcqRHb6HO0b1lND0Y/export?format=csv&gid=284416453"
    df = pd.read_csv(CSV_URL)
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    
    # Target exact "Submission Date" column
    date_col = "Submission Date" if "Submission Date" in df.columns else None
    
    if not date_col:
        for c in df.columns:
            if "submission date" in c.lower() or ("date" in c.lower() and "id" not in c.lower()):
                date_col = c
                break

    if date_col:
        df['Parsed_Date'] = pd.to_datetime(df[date_col], errors='coerce')
        # Filter out invalid dates
        valid_dates_mask = (df['Parsed_Date'].dt.year >= 2020) & (df['Parsed_Date'].dt.year <= 2030)
        df.loc[~valid_dates_mask, 'Parsed_Date'] = pd.NaT
        
        df['Formatted_Timestamp'] = df['Parsed_Date'].dt.strftime('%d %b %Y, %H:%M')
        df['Formatted_Timestamp'] = df['Formatted_Timestamp'].fillna("Recent Entry")
        
        # Sort latest entries first
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

# Date Range Filter
valid_dates = df['Parsed_Date'].dropna() if 'Parsed_Date' in df.columns else pd.Series()
if not valid_dates.empty:
    min_date = valid_dates.min().date()
    max_date = valid_dates.max().date()
    
    st.sidebar.markdown("### Select Date Range")
    date_range = st.sidebar.date_input(
        "Submissions Window:",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
else:
    date_range = None

st.sidebar.divider()

filtered_df = df.copy()

# Project Filter Button
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

# Field Rep Filter Button
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

# Dropdown Multi-Select for Outlets
if 'Type Outlet Name' in df.columns:
    st.sidebar.markdown("### 🏬 Outlets")
    outlet_list = sorted(filtered_df['Type Outlet Name'].dropna().unique().tolist())
    selected_outlets = st.sidebar.multiselect("Filter Outlets:", outlet_list, default=outlet_list)
    if selected_outlets:
        filtered_df = filtered_df[filtered_df['Type Outlet Name'].isin(selected_outlets)]

# Apply Date Filter
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

# --- 5. COMPETITOR INTELLIGENCE SUMMARY PANEL ---
st.markdown("### 🚨 Live Competitor Activities Summary")
st.write("Real-time pricing adjustments, promotions, and Brand Ambassador activations.")

desc_col = "Describe the competitor activities"
all_image_columns = [col for col in filtered_df.columns if "Take Photo" in str(col)]

if comp_col in filtered_df.columns and desc_col in filtered_df.columns:
    comp_df = filtered_df[filtered_df[comp_col].astype(str).str.strip().str.lower() == "yes"]
    
    if not comp_df.empty:
        comp_summary_col1, comp_summary_col2 = st.columns([2.2, 1])
        
        with comp_summary_col1:
            st.markdown("##### Detailed Store Reports")
            for _, row in comp_df.iterrows():
                outlet = row.get("Type Outlet Name", "Unknown Outlet")
                rep = row.get("Select Your Name", "Unknown Rep")
                activity = row.get(desc_col, "No description provided.")
                if pd.isna(activity) or str(activity).strip().lower() == "nan":
                    activity = "Competitor activity confirmed (no notes entered)."
                    
                timestamp_str = row.get('Formatted_Timestamp', 'Recent Entry')

                # Store Container Card
                with st.container():
                    header_col1, header_col2 = st.columns([3, 1])
                    with header_col1:
                        st.markdown(f"📍 **{outlet}** &nbsp;*(Rep: {rep})*", unsafe_allow_html=True)
                    with header_col2:
                        st.markdown(f"<div style='text-align: right;'><span class='timestamp-badge'>⏱️ {timestamp_str}</span></div>", unsafe_allow_html=True)
                    
                    st.markdown(f"<div style='margin-top: 4px; margin-bottom: 12px; font-size: 14px; color: #34495e; white-space: pre-line;'>{activity}</div>", unsafe_allow_html=True)
                    
                    # Render attached photos inside store report card
                    attached_images = []
                    for img_col in all_image_columns:
                        img_val = row.get(img_col)
                        if pd.notna(img_val) and str(img_val).strip() != "" and str(img_val).strip().lower() != "nan":
                            urls = str(img_val).split(",")
                            for u in urls:
                                u_clean = u.strip()
                                if u_clean:
                                    clean_title = str(img_col).replace("Take Photo of competitor activites", "Competitor Activity")\
                                                              .replace("Take Photo of paid assets", "Paid Asset")\
                                                              .replace("(", "").replace(")", "").strip()
                                    attached_images.append((clean_title, u_clean))
                    
                    if attached_images:
                        img_cols = st.columns(min(len(attached_images), 4))
                        for idx, (img_title, raw_url) in enumerate(attached_images):
                            if "jotform.com" in raw_url:
                                separator = "&" if "?" in raw_url else "?"
                                auth_url = f"{raw_url}{separator}apiKey={JOTFORM_API_KEY}"
                            else:
                                auth_url = raw_url

                            with img_cols[idx % 4]:
                                html_code = f'''
                                    <div style="margin-bottom: 8px;">
                                        <span style="font-size: 11px; font-weight: 600; color: #64748b;">{img_title}</span>
                                        <a href="{auth_url}" target="_blank">
                                            <img src="{auth_url}" width="100%" style="border-radius: 6px; border: 1px solid #cbd5e1; margin-top: 2px;">
                                        </a>
                                    </div>
                                '''
                                st.markdown(html_code, unsafe_allow_html=True)
                    
                    st.divider()
                
        with comp_summary_col2:
            st.markdown("##### Export Options")
            st.info("Export filtered competitor activity logs directly to CSV.")
            
            export_df = comp_df[['Type Outlet Name', 'Select Your Name', desc_col, 'Formatted_Timestamp']].rename(columns={
                'Type Outlet Name': 'Outlet',
                'Select Your Name': 'Field Rep',
                desc_col: 'Competitor Activity Details',
                'Formatted_Timestamp': 'Timestamp'
            })
            csv_data = export_df.to_csv(index=False).encode('utf-8')
            
            st.download_button(
                label="📥 Download Summary (CSV)",
                data=csv_data,
                file_name="Competitor_Activity_Report.csv",
                mime="text/csv",
            )
    else:
        st.info("No competitor activities reported for the selected filters.")
else:
    st.info("Competitor activity columns are missing from dataset.")

st.divider()

# --- 6. TYPES OF PAID ASSETS PANEL ---
st.markdown("### 📦 Paid Asset Distribution")
asset_col = "Select Paid assets in the outlet"

if asset_col in filtered_df.columns:
    assets_series = filtered_df[asset_col].dropna().astype(str)
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
else:
    st.info("Paid assets column missing from dataset.")

st.divider()

# --- 7. FIELD GALLERY WITH IMAGE TYPE TOGGLE BUTTONS ---
st.markdown("### 📸 Authenticated Field Gallery")
st.write("Toggle between Paid Asset photos and Competitor Activity photos below.")

paid_image_cols = [c for c in all_image_columns if "paid asset" in str(c).lower()]
comp_image_cols = [c for c in all_image_columns if "competitor" in str(c).lower()]

# Button Distinguisher Slicer
image_type_choice = st.segmented_control(
    "Select Photo Type:",
    options=["All Images", "Paid Assets", "Competitor Activities"],
    default="All Images"
)

if image_type_choice == "Paid Assets":
    active_image_cols = paid_image_cols
elif image_type_choice == "Competitor Activities":
    active_image_cols = comp_image_cols
else:
    active_image_cols = all_image_columns

if active_image_cols:
    selected_img_col = st.selectbox("Select Specific Category Column:", active_image_cols)
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
                
                if "jotform.com" in first_url:
                    separator = "&" if "?" in first_url else "?"
                    auth_url = f"{first_url}{separator}apiKey={JOTFORM_API_KEY}"
                else:
                    auth_url = first_url

                with cols[col_index % 4]:
                    st.markdown(f"**{outlet_name}**")
                    st.caption(f"⏱️ {timestamp_str}")
                    html_code = f'''
                        <a href="{auth_url}" target="_blank">
                            <img src="{auth_url}" width="100%" style="border-radius: 6px; border: 1px solid #ced4da;">
                        </a>
                        <br>
                        <a href="{auth_url}" target="_blank" style="font-size: 12px; text-decoration: none; color: #0d6efd;">🔗 View Original</a>
                    '''
                    st.markdown(html_code, unsafe_allow_html=True)
                
                col_index += 1
else:
    st.info("No image columns match the selected photo type.")