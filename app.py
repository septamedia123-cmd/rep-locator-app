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

def load_data():
    gc = get_gsheet_client()
    sheet = gc.open_by_key(GSHEET_ID)
    ws = sheet.worksheet("rep_profiles")
    data = ws.get_all_records()
    return pd.DataFrame(data)

def login():
    st.title("NuLife Rep Locator")
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

if page == "Map":
    st.title("Rep Map")

    df["Latitude"] = pd.to_numeric(df["Latitude"], errors="coerce")
    df["Longitude"] = pd.to_numeric(df["Longitude"], errors="coerce")
    df = df.dropna(subset=["Latitude", "Longitude"])

    m = folium.Map(location=[39.5, -98.35], zoom_start=4)

    for _, row in df.iterrows():
        popup = f"{row['FullName']}<br>{row['MarketTerritory']}<br>{row['PhoneNumber']}"
        folium.Marker(
            [row["Latitude"], row["Longitude"]],
            popup=popup
        ).add_to(m)

    st_folium(m, width=1000, height=600)

if page == "Data":
    st.title("Rep Data")
    st.dataframe(df)
