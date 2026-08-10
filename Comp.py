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

# Custom Styling for Dashboard UI & Slicers
st.markdown("""
    <style>
    /* Styled Sidebar Filters */
    div[data-testid="stSidebar"] {
        background-color: #f8f9fa;
        padding-top: 15px;
    }
    div[data-testid="stSidebar"] label {
        font-weight: 600 !important;
        color: #2c3e50 !important;
    }
    /* Summary Cards */
    .metric-box {
        background-color: #ffffff;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border: 1px solid #e9ecef;
    }
    .comp-card {
        background-color: #fdfdfe;
        border-left: 4px solid #e74c3c;
        padding: 12px 16px;
        margin-bottom: 12px;
        border-radius: 4px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

# 2. Load Data Function
@st.cache_data(ttl=300)
def load_data():
    CSV_URL = "https://docs.google.com/spreadsheets/d/1OoAeqfCO4k-9zm75_FsAT0PSDySqD_GR8RGfdIxjgTA/export?format=csv&gid=102621635"
    df = pd.read_csv(CSV_URL)
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    return df

try:
    df = load_data()
except Exception as e:
    st.error("Failed to load live data. Ensure your Google Sheet is accessible.")
    st.stop()

# --- 3. PRESENTABLE SIDEBAR SLICERS ---
st.sidebar.title("🎛️ Report Slicers")
st.sidebar.markdown("Filter live operational metrics by region or outlet.")

# Project Filter
if 'Select Your Project' in df.columns:
    project_list = sorted(df['Select Your Project'].dropna().unique().tolist())
    selected_projects = st.sidebar.multiselect("Select Project:", project_list, default=project_list)
else:
    selected_projects = []

# Field Rep Filter
if 'Select Your Name' in df.columns:
    rep_list = sorted(df['Select Your Name'].dropna().unique().tolist())
    selected_reps = st.sidebar.multiselect("Field Representative:", rep_list, default=rep_list)
else:
    selected_reps = []

# Outlet Filter
if 'Type Outlet Name' in df.columns:
    outlet_list = sorted(df['Type Outlet Name'].dropna().unique().tolist())
    selected_outlets = st.sidebar.multiselect("Outlet Name:", outlet_list, default=outlet_list)
else:
    selected_outlets = []

# Apply Slicers
filtered_df = df.copy()
if selected_projects:
    filtered_df = filtered_df[filtered_df['Select Your Project'].isin(selected_projects)]
if selected_reps:
    filtered_df = filtered_df[filtered_df['Select Your Name'].isin(selected_reps)]
if selected_outlets:
    filtered_df = filtered_df[filtered_df['Type Outlet Name'].isin(selected_outlets)]

st.sidebar.divider()
st.sidebar.caption("Data syncs automatically with Google Sheets every 5 minutes.")

# --- DASHBOARD HEADER ---
st.title("📊 Weekly Field Dashboard: Paid Assets & Competitor Intelligence")
st.markdown("Live operational updates from field representatives across retail outlets.")

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
st.write("Real-time pricing adjustments, promotions, and Brand Ambassador activations reported from stores.")

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
                
                st.markdown(f"""
                <div class="comp-card">
                    <strong style="color: #2c3e50; font-size: 16px;">📍 {outlet}</strong> 
                    <span style="font-size: 12px; color: #7f8c8d;"> (Rep: {rep})</span><br>
                    <div style="margin-top: 6px; font-size: 14px; color: #34495e; white-space: pre-line;">{activity}</div>
                </div>
                """, unsafe_allow_html=True)
                
        with comp_summary_col2:
            st.markdown("##### Export Options")
            st.info("Download the structured competitor activity logs for client distribution.")
            
            # Export CSV Button
            export_df = comp_df[['Type Outlet Name', 'Select Your Name', desc_col]].rename(columns={
                'Type Outlet Name': 'Outlet',
                'Select Your Name': 'Field Rep',
                desc_col: 'Competitor Activity Details'
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

# --- 7. FIELD IMAGES VIEW BOX ---
st.markdown("### 📸 Authenticated Field Gallery")
st.write("Select a category to review field asset executions and competitor activations.")

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