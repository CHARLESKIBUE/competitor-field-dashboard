import pandas as pd
import streamlit as st

# --- JOTFORM API KEY ---
JOTFORM_API_KEY = "6666042ab61ea98f9b378a4a7d"

# 1. Page Configuration
st.set_page_config(
    page_title="Executive Field & Competitor Dashboard", 
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
        padding: 14px 18px;
        margin-bottom: 12px;
        border-radius: 6px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.08);
        border-top: 1px solid #f1f3f5;
        border-right: 1px solid #f1f3f5;
        border-bottom: 1px solid #f1f3f5;
    }
    </style>
""", unsafe_allow_html=True)

# 2. Load & Parse Data Function
@st.cache_data(ttl=300)
def load_data():
    CSV_URL = "https://docs.google.com/spreadsheets/d/1OoAeqfCO4k-9zm75_FsAT0PSDySqD_GR8RGfdIxjgTA/export?format=csv&gid=102621635"
    df = pd.read_csv(CSV_URL)
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    
    # Try parsing date column if present (e.g., Submission Date or Timestamp)
    date_cols = [c for c in df.columns if "date" in c.lower() or "time" in c.lower()]
    if date_cols:
        target_date_col = date_cols[0]
        df['Parsed_Date'] = pd.to_datetime(df[target_date_col], errors='coerce')
        # Sort so latest entries are at the top
        df = df.sort_values(by='Parsed_Date', ascending=False).reset_index(drop=True)
    else:
        # If no explicit date column, reverse index order to put newest entries at the top
        df = df.iloc[::-1].reset_index(drop=True)
        df['Parsed_Date'] = None
        
    return df

try:
    df = load_data()
except Exception as e:
    st.error("Failed to load live data. Ensure your Google Sheet is accessible.")
    st.stop()

# --- 3. SIDEBAR FILTERS (BUTTON SLICERS & DATE PICKER) ---
st.sidebar.title("🎛️ Report Slicers")

# Date Selector Filter
if df['Parsed_Date'].notna().any():
    min_date = df['Parsed_Date'].min().date()
    max_date = df['Parsed_Date'].max().date()
    
    st.sidebar.markdown("### 📅 Date Range")
    date_range = st.sidebar.date_input(
        "Select Submissions Window:",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
else:
    date_range = None

st.sidebar.divider()

# Button Slicer for Project Selection
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

# Button Slicer for Field Representative Selection
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

# Apply Date Filter if selected
if date_range and len(date_range) == 2 and filtered_df['Parsed_Date'].notna().any():
    start_date, end_date = date_range
    filtered_df = filtered_df[
        (filtered_df['Parsed_Date'].dt.date >= start_date) & 
        (filtered_df['Parsed_Date'].dt.date <= end_date)
    ]

st.sidebar.divider()
st.sidebar.caption("⚡ Live sync enabled (Google Sheets / Jotform API)")

# --- DASHBOARD HEADER ---
st.title("📊 Weekly Field Dashboard: Paid Assets & Competitor Intelligence")
st.markdown("Real-time operational field updates sorted chronologically (newest first).")

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
st.markdown("### 🚨 Live Competitor Activities Summary (Latest Entries First)")
st.write("Real-time pricing adjustments, promotions, and Brand Ambassador activations.")

desc_col = "Describe the competitor activities"
if comp_col in filtered_df.columns and desc_col in filtered_df.columns:
    comp_df = filtered_df[filtered_df[comp_col].astype(str).str.strip().str.lower() == "yes"]
    
    if not comp_df.empty:
        comp_summary_col1, comp_summary_col2 = st.columns([2, 1])
        
        with comp_summary_col1:
            st.markdown("##### Detailed Store Reports")
            for _, row in comp_df.iterrows():
                outlet = row.get("Type Outlet Name", "Unknown Outlet")
                rep = row.get("Select Your Name", "Unknown Rep")
                activity = row.get(desc_col, "No description provided.")
                date_str = f" | {row['Parsed_Date'].strftime('%d %b %Y')}" if pd.notna(row.get('Parsed_Date')) else ""
                
                st.markdown(f"""
                <div class="comp-card">
                    <strong style="color: #2c3e50; font-size: 16px;">📍 {outlet}</strong> 
                    <span style="font-size: 12px; color: #7f8c8d;"> (Rep: {rep}{date_str})</span><br>
                    <div style="margin-top: 6px; font-size: 14px; color: #34495e; white-space: pre-line;">{activity}</div>
                </div>
                """, unsafe_allow_html=True)
                
        with comp_summary_col2:
            st.markdown("##### Export Report")
            st.info("Export filtered competitor activity logs directly to CSV.")
            
            export_df = comp_df[['Type Outlet Name', 'Select Your Name', desc_col]].rename(columns={
                'Type Outlet Name': 'Outlet',
                'Select Your Name': 'Field Rep',
                desc_col: 'Competitor Activity Details'
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
    st.info("Competitor activity columns are missing from the dataset.")

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

# --- 7. FIELD IMAGES VIEW GALLERY ---
st.markdown("### 📸 Authenticated Field Gallery")
st.write("Browse field images chronologically across categories.")

image_columns = [col for col in filtered_df.columns if "Take Photo" in str(col)]

if image_columns:
    selected_img_col = st.selectbox("Select Photo Category:", image_columns)
    df_images = filtered_df.dropna(subset=[selected_img_col])

    if df_images.empty:
        st.info("No images submitted for this category.")
    else:
        cols = st.columns(4)
        col_index = 0 
        
        for _, row in df_images.iterrows():
            img_url_raw = row[selected_img_col]
            outlet_name = row.get("Type Outlet Name", "Unknown Outlet")
            
            if pd.notna(img_url_raw) and str(img_url_raw).strip() != "":
                first_url = str(img_url_raw).split(",")[0].strip()
                
                # Append Jotform API Key
                if "jotform.com" in first_url:
                    separator = "&" if "?" in first_url else "?"
                    auth_url = f"{first_url}{separator}apiKey={JOTFORM_API_KEY}"
                else:
                    auth_url = first_url

                with cols[col_index % 4]:
                    st.markdown(f"**{outlet_name}**")
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
    st.info("No image columns detected in dataset.")