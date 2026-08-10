import pandas as pd
import streamlit as st

# --- JOTFORM API KEY ---
JOTFORM_API_KEY = "6666042ab61ea98f9b378a4a7d"

# 1. Page Configuration
st.set_page_config(page_title="Weekly Field Dashboard", layout="wide", initial_sidebar_state="expanded")
st.title("Weekly Dashboard: Paid Assets & Competitor Activities")

# 2. Load Data Function
@st.cache_data(ttl=300)
def load_data():
    # Your specific Google Sheet converted to a CSV export link
    CSV_URL = "https://docs.google.com/spreadsheets/d/1OoAeqfCO4k-9zm75_FsAT0PSDySqD_GR8RGfdIxjgTA/export?format=csv&gid=102621635"
    
    df = pd.read_csv(CSV_URL)
    df.columns = df.columns.str.strip()
    df = df.dropna(how="all")
    return df

try:
    df = load_data()
except Exception as e:
    st.error("Failed to load live data. Ensure your Google Sheet is set to 'Anyone with the link' can view, or that it is Published to the Web.")
    st.stop()

# --- 3. SLICERS (SIDEBAR FILTERS) ---
st.sidebar.header("Dashboard Filters 🎛️")
st.sidebar.write("Use these filters to slice the data.")

if 'Select Your Project' in df.columns:
    project_list = df['Select Your Project'].dropna().unique().tolist()
    selected_projects = st.sidebar.multiselect("Filter by Project:", project_list, default=project_list)
else:
    selected_projects = []

if 'Type Outlet Name' in df.columns:
    outlet_list = df['Type Outlet Name'].dropna().unique().tolist()
    selected_outlets = st.sidebar.multiselect("Filter by Outlet Name:", outlet_list, default=outlet_list)
else:
    selected_outlets = []

filtered_df = df.copy()
if selected_projects:
    filtered_df = filtered_df[filtered_df['Select Your Project'].isin(selected_projects)]
if selected_outlets:
    filtered_df = filtered_df[filtered_df['Type Outlet Name'].isin(selected_outlets)]

# --- 4. HIGH-LEVEL METRICS ---
st.markdown("### High-Level Metrics")
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

# --- 5. TYPES OF PAID ASSETS PANEL ---
st.markdown("### Types of Paid Assets")
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
        st.info("No paid assets recorded in the filtered data.")
else:
    st.info("Paid assets column not found in dataset.")

st.divider()

# --- 6. AUTHENTICATED FRONTEND BROWSER IMAGE VIEW ---
st.markdown("### 📸 Field Images View")
st.write("Select a photo category below to view the uploaded images from the field.")

image_columns = [col for col in filtered_df.columns if "Take Photo" in str(col)]

if image_columns:
    selected_img_col = st.selectbox("Select Image Category:", image_columns)
    df_images = filtered_df.dropna(subset=[selected_img_col])

    if df_images.empty:
        st.info("No images were submitted for this category.")
    else:
        cols = st.columns(4)
        col_index = 0 
        
        for _, row in df_images.iterrows():
            img_url_raw = row[selected_img_col]
            outlet_name = row.get("Type Outlet Name", "Unknown Outlet")
            
            if pd.notna(img_url_raw) and str(img_url_raw).strip() != "":
                first_url = str(img_url_raw).split(",")[0].strip()
                
                # --- Inject the API Key into the HTML URL ---
                if "jotform.com" in first_url:
                    separator = "&" if "?" in first_url else "?"
                    auth_url = f"{first_url}{separator}apiKey={JOTFORM_API_KEY}"
                else:
                    auth_url = first_url

                with cols[col_index % 4]:
                    st.write(f"**{outlet_name}**")
                    
                    # HTML forces the browser to load the image, API key provides authorization!
                    html_code = f'''
                        <a href="{auth_url}" target="_blank">
                            <img src="{auth_url}" width="100%" style="border-radius: 8px; border: 1px solid #ddd;">
                        </a>
                        <br>
                        <a href="{auth_url}" target="_blank" style="font-size: 14px; text-decoration: none; color: #1f77b4;">🔗 Click to open original</a>
                    '''
                    st.markdown(html_code, unsafe_allow_html=True)
                
                col_index += 1
else:
    st.info("No image columns found in dataset.")