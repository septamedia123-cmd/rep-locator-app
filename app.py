import streamlit as st
import pandas as pd
import gspread
import folium
from streamlit_folium import st_folium
from google.oauth2.service_account import Credentials
from datetime import datetime

st.set_page_config(page_title="NuLife Rep Locator", page_icon="📍", layout="wide")

GSHEET_ID = st.secrets["GSHEET_ID"]
APP_PASSWORD = st.secrets["APP_PASSWORD"]

REP_HEADERS = [
    "RepID", "Active", "Manager", "Region", "MarketTerritory", "State", "City",
    "FirstName", "LastName", "FullName", "PhoneNumber", "PersonalEmail",
    "NuLifeEmail", "LinksHandles", "BusinessName", "Address", "Latitude",
    "Longitude", "Notes", "StartDate", "LastUpdated"
]

def get_gsheet_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )
    return gspread.authorize(creds)

@st.cache_data(ttl=300)
def load_reps():
    try:
        gc = get_gsheet_client()
        sheet = gc.open_by_key(GSHEET_ID)
        ws = sheet.worksheet("rep_profiles")
        data = ws.get_all_records()
        df = pd.DataFrame(data)

        for col in REP_HEADERS:
            if col not in df.columns:
                df[col] = ""

        return df[REP_HEADERS]

    except Exception as e:
        st.error("Google Sheets connection failed.")
        st.write("Error type:", type(e).__name__)
        st.write("Error details:", str(e))
        st.stop()

def save_reps(df):
    try:
        gc = get_gsheet_client()
        sheet = gc.open_by_key(GSHEET_ID)
        ws = sheet.worksheet("rep_profiles")

        clean_df = df.copy()
        for col in REP_HEADERS:
            if col not in clean_df.columns:
                clean_df[col] = ""

        clean_df = clean_df[REP_HEADERS].fillna("")
        ws.clear()
        ws.update([REP_HEADERS] + clean_df.astype(str).values.tolist())
        st.cache_data.clear()
        return True

    except Exception as e:
        st.error("Could not save data to Google Sheets.")
        st.write("Error type:", type(e).__name__)
        st.write("Error details:", str(e))
        return False

def stable_offset(index):
    offsets = [
        (0.0000, 0.0000),
        (0.0080, 0.0080),
        (-0.0080, -0.0080),
        (0.0080, -0.0080),
        (-0.0080, 0.0080),
    ]
    return offsets[index % len(offsets)]

def login():
    st.title("NuLife Rep Locator")
    st.caption("Secure access required")
    pw = st.text_input("Password", type="password")

    if st.button("Login"):
        if pw == APP_PASSWORD:
            st.session_state.auth = True
            st.rerun()
        else:
            st.error("Wrong password")

def metric_card(label, value):
    st.metric(label, value)

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    login()
    st.stop()

df = load_reps()

st.sidebar.title("NuLife Rep Locator")
page = st.sidebar.radio(
    "Navigation",
    ["Dashboard", "Map", "Rep Directory", "Manage Reps"]
)

if st.sidebar.button("Log out"):
    st.session_state.auth = False
    st.rerun()

if st.sidebar.button("Refresh Data"):
    st.cache_data.clear()
    st.rerun()

# =========================
# DASHBOARD
# =========================
if page == "Dashboard":
    st.title("Dashboard")

    working_df = df.copy()
    working_df["Latitude"] = pd.to_numeric(working_df["Latitude"], errors="coerce")
    working_df["Longitude"] = pd.to_numeric(working_df["Longitude"], errors="coerce")

    active_df = working_df[working_df["Active"].astype(str).str.lower() == "yes"]
    missing_coords = working_df[
        working_df["Latitude"].isna() | working_df["Longitude"].isna()
    ]

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Reps", len(working_df))
    c2.metric("Active Reps", len(active_df))
    c3.metric("Markets", working_df["MarketTerritory"].replace("", pd.NA).dropna().nunique())
    c4.metric("States", working_df["State"].replace("", pd.NA).dropna().nunique())
    c5.metric("Missing Coordinates", len(missing_coords))

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Reps by Manager")
        if "Manager" in working_df.columns and not working_df.empty:
            manager_counts = working_df["Manager"].replace("", "Unassigned").value_counts()
            st.bar_chart(manager_counts)

    with col2:
        st.subheader("Reps by State")
        if "State" in working_df.columns and not working_df.empty:
            state_counts = working_df["State"].replace("", "Unknown").value_counts()
            st.bar_chart(state_counts)

    st.markdown("---")
    st.subheader("Data Alerts")

    if missing_coords.empty:
        st.success("All reps have map coordinates.")
    else:
        st.warning(f"{len(missing_coords)} rep(s) are missing Latitude/Longitude.")
        st.dataframe(
            missing_coords[["RepID", "FullName", "MarketTerritory", "State", "City", "Latitude", "Longitude"]],
            use_container_width=True
        )

# =========================
# MAP
# =========================
elif page == "Map":
    st.title("NuLife Rep Map")

    map_df = df.copy()
    map_df["Latitude"] = pd.to_numeric(map_df["Latitude"], errors="coerce")
    map_df["Longitude"] = pd.to_numeric(map_df["Longitude"], errors="coerce")
    map_df = map_df.dropna(subset=["Latitude", "Longitude"]).reset_index(drop=True)

    st.subheader("Filters")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        states = ["All"] + sorted(map_df["State"].dropna().astype(str).unique().tolist())
        selected_state = st.selectbox("State", states)

    with col2:
        managers = ["All"] + sorted(map_df["Manager"].dropna().astype(str).unique().tolist())
        selected_manager = st.selectbox("Manager", managers)

    with col3:
        regions = ["All"] + sorted(map_df["Region"].dropna().astype(str).unique().tolist())
        selected_region = st.selectbox("Region", regions)

    with col4:
        search = st.text_input("Search")

    filtered_df = map_df.copy()

    if selected_state != "All":
        filtered_df = filtered_df[filtered_df["State"].astype(str) == selected_state]

    if selected_manager != "All":
        filtered_df = filtered_df[filtered_df["Manager"].astype(str) == selected_manager]

    if selected_region != "All":
        filtered_df = filtered_df[filtered_df["Region"].astype(str) == selected_region]

    if search:
        mask = filtered_df.astype(str).apply(
            lambda row: row.str.contains(search, case=False, na=False).any(),
            axis=1
        )
        filtered_df = filtered_df[mask]

    filtered_df = filtered_df.reset_index(drop=True)

    st.markdown(f"### Showing {len(filtered_df)} Rep(s)")

    m = folium.Map(location=[39.5, -98.35], zoom_start=4, tiles="OpenStreetMap")

    for i, (_, row) in enumerate(filtered_df.iterrows()):
        offset_lat, offset_lng = stable_offset(i)
        lat = float(row["Latitude"]) + offset_lat
        lng = float(row["Longitude"]) + offset_lng

        popup_html = f"""
        <div style="width:270px; font-family: Arial, sans-serif;">
            <h4 style="margin-bottom:6px;">{row.get('FullName', '')}</h4>
            <b>Rep ID:</b> {row.get('RepID', '')}<br>
            <b>Territory:</b> {row.get('MarketTerritory', '')}<br>
            <b>City/State:</b> {row.get('City', '')}, {row.get('State', '')}<br>
            <b>Manager:</b> {row.get('Manager', '')}<br>
            <b>Region:</b> {row.get('Region', '')}<br><br>

            <b>Phone:</b><br>{row.get('PhoneNumber', '')}<br><br>
            <b>Email:</b><br>{row.get('PersonalEmail', '')}<br><br>
            <b>NuLife Email:</b><br>{row.get('NuLifeEmail', '')}<br><br>
            <b>Business:</b><br>{row.get('BusinessName', '')}<br><br>
            <b>Notes:</b><br>{row.get('Notes', '')}
        </div>
        """

        folium.Marker(
            [lat, lng],
            popup=folium.Popup(popup_html, max_width=330),
            tooltip=row.get("FullName", "Rep"),
            icon=folium.Icon(color="blue", icon="flag")
        ).add_to(m)

    st_folium(m, width=1150, height=650, returned_objects=[], key="rep_map")

    st.markdown("---")
    st.subheader("Rep Profiles")

    if filtered_df.empty:
        st.info("No reps match the selected filters.")
    else:
        for _, row in filtered_df.iterrows():
            with st.expander(f"{row.get('FullName', '')} — {row.get('MarketTerritory', '')}"):
                c1, c2 = st.columns(2)

                with c1:
                    st.write(f"**Rep ID:** {row.get('RepID', '')}")
                    st.write(f"**Active:** {row.get('Active', '')}")
                    st.write(f"**Manager:** {row.get('Manager', '')}")
                    st.write(f"**Region:** {row.get('Region', '')}")
                    st.write(f"**Territory:** {row.get('MarketTerritory', '')}")
                    st.write(f"**Location:** {row.get('City', '')}, {row.get('State', '')}")

                with c2:
                    st.write(f"**Phone:** {row.get('PhoneNumber', '')}")
                    st.write(f"**Personal Email:** {row.get('PersonalEmail', '')}")
                    st.write(f"**NuLife Email:** {row.get('NuLifeEmail', '')}")
                    st.write(f"**Business:** {row.get('BusinessName', '')}")
                    st.write(f"**Links/Handles:** {row.get('LinksHandles', '')}")

                st.write("**Notes:**")
                st.write(row.get("Notes", ""))

# =========================
# REP DIRECTORY
# =========================
elif page == "Rep Directory":
    st.title("Rep Directory")

    search_dir = st.text_input("Search reps, markets, managers, states")

    directory_df = df.copy()

    if search_dir:
        mask = directory_df.astype(str).apply(
            lambda row: row.str.contains(search_dir, case=False, na=False).any(),
            axis=1
        )
        directory_df = directory_df[mask]

    st.markdown(f"### {len(directory_df)} Rep(s)")

    for _, row in directory_df.iterrows():
        with st.container():
            st.markdown(
                f"""
                <div style="
                    background:#ffffff;
                    border:1px solid #e5e7eb;
                    border-radius:16px;
                    padding:18px;
                    margin-bottom:12px;
                    box-shadow:0 4px 10px rgba(0,0,0,0.04);
                ">
                    <div style="font-size:22px; font-weight:800;">
                        {row.get('FullName', '')}
                    </div>
                    <div style="font-size:14px; color:#6b7280; margin-top:4px;">
                        {row.get('MarketTerritory', '')} • {row.get('City', '')}, {row.get('State', '')} • Manager: {row.get('Manager', '')}
                    </div>
                    <div style="margin-top:10px;">
                        <b>Phone:</b> {row.get('PhoneNumber', '')}<br>
                        <b>Email:</b> {row.get('PersonalEmail', '')}<br>
                        <b>NuLife:</b> {row.get('NuLifeEmail', '')}
                    </div>
                    <div style="margin-top:10px; color:#374151;">
                        {row.get('Notes', '')}
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

# =========================
# MANAGE REPS
# =========================
elif page == "Manage Reps":
    st.title("Manage Reps")

    st.info("Edit reps below, then click Save Changes to update Google Sheets.")

    editable_df = df.copy()

    edited_df = st.data_editor(
        editable_df,
        use_container_width=True,
        num_rows="dynamic",
        key="rep_editor"
    )

    c1, c2 = st.columns(2)

    with c1:
        if st.button("Save Changes", type="primary", use_container_width=True):
            edited_df["LastUpdated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if save_reps(edited_df):
                st.success("Rep profiles saved successfully.")
                st.rerun()

    with c2:
        if st.button("Discard Changes / Refresh", use_container_width=True):
            st.cache_data.clear()
            st.rerun()
