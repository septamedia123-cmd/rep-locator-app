import streamlit as st
import pandas as pd
import gspread
import folium
from streamlit_folium import st_folium
from google.oauth2.service_account import Credentials

st.set_page_config(page_title="NuLife Rep Locator", layout="wide")

GSHEET_ID = st.secrets["GSHEET_ID"]
APP_PASSWORD = st.secrets["APP_PASSWORD"]

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
def load_data():
    try:
        gc = get_gsheet_client()
        sheet = gc.open_by_key(GSHEET_ID)
        ws = sheet.worksheet("rep_profiles")
        data = ws.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error("Google Sheets connection failed.")
        st.write("Error type:", type(e).__name__)
        st.write("Error details:", str(e))
        st.stop()

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

if "auth" not in st.session_state:
    st.session_state.auth = False

if not st.session_state.auth:
    login()
    st.stop()

df = load_data()

st.sidebar.title("Navigation")
page = st.sidebar.radio("Go to", ["Map", "Data"])

if st.sidebar.button("Log out"):
    st.session_state.auth = False
    st.rerun()

if page == "Map":
    st.title("NuLife Rep Map")

    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    df = df.dropna(subset=["Latitude", "Longitude"]).reset_index(drop=True)

    st.subheader("Filters")

    col1, col2, col3 = st.columns(3)

    with col1:
        states = ["All"] + sorted(df["State"].dropna().astype(str).unique().tolist())
        selected_state = st.selectbox("Filter by State", states)

    with col2:
        managers = ["All"] + sorted(df["Manager"].dropna().astype(str).unique().tolist())
        selected_manager = st.selectbox("Filter by Manager", managers)

    with col3:
        search = st.text_input("Search Rep / Territory")

    filtered_df = df.copy()

    if selected_state != "All":
        filtered_df = filtered_df[filtered_df["State"].astype(str) == selected_state]

    if selected_manager != "All":
        filtered_df = filtered_df[filtered_df["Manager"].astype(str) == selected_manager]

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
        <div style="width:260px; font-family: Arial, sans-serif;">
            <h4 style="margin-bottom:6px;">{row.get('FullName', '')}</h4>

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
            popup=folium.Popup(popup_html, max_width=320),
            tooltip=row.get("FullName", "Rep"),
            icon=folium.Icon(color="blue", icon="flag")
        ).add_to(m)

    st_folium(
        m,
        width=1100,
        height=650,
        returned_objects=[],
        key="rep_map"
    )

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

if page == "Data":
    st.title("Rep Data")

    if st.button("Refresh Google Sheet Data"):
        st.cache_data.clear()
        st.rerun()

    search_data = st.text_input("Search table")

    data_df = df.copy()

    if search_data:
        mask = data_df.astype(str).apply(
            lambda row: row.str.contains(search_data, case=False, na=False).any(),
            axis=1
        )
        data_df = data_df[mask]

    st.dataframe(data_df, use_container_width=True)
