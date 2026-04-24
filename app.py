import streamlit as st
import pandas as pd
import gspread
import folium
from streamlit_folium import st_folium
from google.oauth2.service_account import Credentials
from datetime import datetime

st.set_page_config(page_title="NuLife Rep Locator", page_icon="📍", layout="wide")

col1, col2, col3 = st.columns([1,2,1])

with col2:
    st.image("logo.png", width=200)

st.markdown("<h2 style='text-align: center;'>NuLife Rep Locator</h2>", unsafe_allow_html=True)
st.markdown("---")

GSHEET_ID = st.secrets["GSHEET_ID"]
APP_PASSWORD = st.secrets["APP_PASSWORD"]

REP_HEADERS = [
    "RepID", "Active", "Manager", "Region", "MarketTerritory", "State", "City",
    "FirstName", "LastName", "FullName", "PhoneNumber", "PersonalEmail",
    "NuLifeEmail", "LinksHandles", "BusinessName", "Address", "Latitude",
    "Longitude", "Notes", "StartDate", "LastUpdated"
]

SALES_HEADERS = [
    "Date", "RepID", "FullName", "MarketTerritory", "State", "Orders",
    "Revenue", "Providers", "TopProduct", "LastOrderDate", "AverageOrderValue"
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
        st.error("Google Sheets rep_profiles connection failed.")
        st.write("Error type:", type(e).__name__)
        st.write("Error details:", str(e))
        st.stop()

@st.cache_data(ttl=300)
def load_sales():
    try:
        gc = get_gsheet_client()
        sheet = gc.open_by_key(GSHEET_ID)
        ws = sheet.worksheet("rep_sales")
        data = ws.get_all_records()
        df = pd.DataFrame(data)

        for col in SALES_HEADERS:
            if col not in df.columns:
                df[col] = ""

        return df[SALES_HEADERS]

    except Exception:
        return pd.DataFrame(columns=SALES_HEADERS)

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

def clean_sales_df(sales_df):
    df = sales_df.copy()

    if df.empty:
        return df

    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df["LastOrderDate"] = pd.to_datetime(df["LastOrderDate"], errors="coerce")
    df["Orders"] = pd.to_numeric(df["Orders"], errors="coerce").fillna(0)
    df["Revenue"] = pd.to_numeric(df["Revenue"], errors="coerce").fillna(0)
    df["Providers"] = pd.to_numeric(df["Providers"], errors="coerce").fillna(0)
    df["AverageOrderValue"] = pd.to_numeric(df["AverageOrderValue"], errors="coerce").fillna(0)

    return df

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

reps_df = load_reps()
sales_df = clean_sales_df(load_sales())

st.sidebar.title("NuLife Rep Locator")
page = st.sidebar.radio(
    "Navigation",
    ["Dashboard", "Map", "Rep Directory", "Sales Dashboard", "Manage Reps"]
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

    working_df = reps_df.copy()
    working_df["Latitude"] = pd.to_numeric(working_df["Latitude"], errors="coerce")
    working_df["Longitude"] = pd.to_numeric(working_df["Longitude"], errors="coerce")

    active_df = working_df[working_df["Active"].astype(str).str.lower() == "yes"]
    missing_coords = working_df[
        working_df["Latitude"].isna() | working_df["Longitude"].isna()
    ]

    total_revenue = sales_df["Revenue"].sum() if not sales_df.empty else 0
    total_orders = sales_df["Orders"].sum() if not sales_df.empty else 0

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Reps", len(working_df))
    c2.metric("Active Reps", len(active_df))
    c3.metric("Markets", working_df["MarketTerritory"].replace("", pd.NA).dropna().nunique())
    c4.metric("Total Revenue", f"${total_revenue:,.0f}")
    c5.metric("Total Orders", f"{int(total_orders):,}")

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Reps by Manager")
        manager_counts = working_df["Manager"].replace("", "Unassigned").value_counts()
        st.bar_chart(manager_counts)

    with col2:
        st.subheader("Revenue by Rep")
        if not sales_df.empty:
            rev_by_rep = sales_df.groupby("FullName")["Revenue"].sum().sort_values(ascending=False)
            st.bar_chart(rev_by_rep)
        else:
            st.info("No sales data yet.")

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

    map_df = reps_df.copy()
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

        rep_sales = sales_df[sales_df["RepID"].astype(str) == str(row.get("RepID", ""))]
        rep_revenue = rep_sales["Revenue"].sum() if not rep_sales.empty else 0
        rep_orders = rep_sales["Orders"].sum() if not rep_sales.empty else 0

        popup_html = f"""
        <div style="width:280px; font-family: Arial, sans-serif;">
            <h4 style="margin-bottom:6px;">{row.get('FullName', '')}</h4>
            <b>Rep ID:</b> {row.get('RepID', '')}<br>
            <b>Territory:</b> {row.get('MarketTerritory', '')}<br>
            <b>City/State:</b> {row.get('City', '')}, {row.get('State', '')}<br>
            <b>Manager:</b> {row.get('Manager', '')}<br>
            <b>Region:</b> {row.get('Region', '')}<br><br>

            <b>Total Revenue:</b> ${rep_revenue:,.0f}<br>
            <b>Total Orders:</b> {int(rep_orders)}<br><br>

            <b>Phone:</b><br>{row.get('PhoneNumber', '')}<br><br>
            <b>Email:</b><br>{row.get('PersonalEmail', '')}<br><br>
            <b>NuLife Email:</b><br>{row.get('NuLifeEmail', '')}<br><br>
            <b>Business:</b><br>{row.get('BusinessName', '')}<br><br>
            <b>Notes:</b><br>{row.get('Notes', '')}
        </div>
        """

        folium.Marker(
            [lat, lng],
            popup=folium.Popup(popup_html, max_width=340),
            tooltip=row.get("FullName", "Rep"),
            icon=folium.Icon(color="blue", icon="flag")
        ).add_to(m)

    st_folium(m, width=1150, height=650, returned_objects=[], key="rep_map")

# =========================
# REP DIRECTORY
# =========================
elif page == "Rep Directory":
    st.title("Rep Directory")

    search_dir = st.text_input("Search reps, markets, managers, states")

    directory_df = reps_df.copy()

    if search_dir:
        mask = directory_df.astype(str).apply(
            lambda row: row.str.contains(search_dir, case=False, na=False).any(),
            axis=1
        )
        directory_df = directory_df[mask]

    st.markdown(f"### {len(directory_df)} Rep(s)")

    for _, row in directory_df.iterrows():
        rep_sales = sales_df[sales_df["RepID"].astype(str) == str(row.get("RepID", ""))]
        rep_revenue = rep_sales["Revenue"].sum() if not rep_sales.empty else 0
        rep_orders = rep_sales["Orders"].sum() if not rep_sales.empty else 0

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
                        <b>Revenue:</b> ${rep_revenue:,.0f} &nbsp; | &nbsp;
                        <b>Orders:</b> {int(rep_orders)}
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
# SALES DASHBOARD
# =========================
elif page == "Sales Dashboard":
    st.title("Sales Dashboard")

    if sales_df.empty:
        st.warning("No sales data found in rep_sales.")
        st.stop()

    total_revenue = sales_df["Revenue"].sum()
    total_orders = sales_df["Orders"].sum()
    total_providers = sales_df["Providers"].sum()
    avg_order_value = total_revenue / total_orders if total_orders else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Revenue", f"${total_revenue:,.0f}")
    c2.metric("Total Orders", f"{int(total_orders):,}")
    c3.metric("Providers", f"{int(total_providers):,}")
    c4.metric("Avg Order Value", f"${avg_order_value:,.0f}")

    st.markdown("---")

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Rep Leaderboard")
        leaderboard = sales_df.groupby("FullName", as_index=False).agg({
            "Revenue": "sum",
            "Orders": "sum",
            "Providers": "sum"
        }).sort_values("Revenue", ascending=False)

        st.dataframe(leaderboard, use_container_width=True)

    with col2:
        st.subheader("Revenue by Territory")
        territory_sales = sales_df.groupby("MarketTerritory")["Revenue"].sum().sort_values(ascending=False)
        st.bar_chart(territory_sales)

    st.markdown("---")

    col3, col4 = st.columns(2)

    with col3:
        st.subheader("Orders by Rep")
        orders_by_rep = sales_df.groupby("FullName")["Orders"].sum().sort_values(ascending=False)
        st.bar_chart(orders_by_rep)

    with col4:
        st.subheader("Top Products")
        top_products = sales_df.groupby("TopProduct")["Revenue"].sum().sort_values(ascending=False)
        st.bar_chart(top_products)

    st.markdown("---")
    st.subheader("Raw Sales Data")
    st.dataframe(sales_df, use_container_width=True)

# =========================
# MANAGE REPS
# =========================
elif page == "Manage Reps":
    st.title("Manage Reps")

    def generate_next_rep_id(existing_df):
        existing_ids = existing_df["RepID"].dropna().astype(str).tolist()
        numbers = []

        for rep_id in existing_ids:
            if rep_id.startswith("REP-"):
                try:
                    numbers.append(int(rep_id.replace("REP-", "")))
                except:
                    pass

        next_number = max(numbers) + 1 if numbers else 1
        return f"REP-{next_number:03d}"

    st.subheader("Add New Rep")

    with st.form("add_rep_form"):
        c1, c2, c3 = st.columns(3)

        with c1:
            first_name = st.text_input("First Name")
            last_name = st.text_input("Last Name")
            active = st.selectbox("Active", ["Yes", "No"], index=0)
            manager = st.text_input("Manager")

        with c2:
            region = st.text_input("Region")
            market = st.text_input("Market / Territory")
            state = st.text_input("State")
            city = st.text_input("City")

        with c3:
            phone = st.text_input("Phone Number")
            personal_email = st.text_input("Personal Email")
            nulife_email = st.text_input("NuLife Email")
            links = st.text_input("Links / Handles")

        business = st.text_input("Business Name")
        address = st.text_input("Address")

        c4, c5 = st.columns(2)
        with c4:
            latitude = st.text_input("Latitude")
        with c5:
            longitude = st.text_input("Longitude")

        notes = st.text_area("Notes")

        submitted = st.form_submit_button("Add Rep")

        if submitted:
            if not first_name.strip() or not last_name.strip():
                st.error("First Name and Last Name are required.")
            else:
                new_rep_id = generate_next_rep_id(reps_df)
                full_name = f"{first_name.strip()} {last_name.strip()}"
                now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                new_row = {
                    "RepID": new_rep_id,
                    "Active": active,
                    "Manager": manager,
                    "Region": region,
                    "MarketTerritory": market,
                    "State": state,
                    "City": city,
                    "FirstName": first_name,
                    "LastName": last_name,
                    "FullName": full_name,
                    "PhoneNumber": phone,
                    "PersonalEmail": personal_email,
                    "NuLifeEmail": nulife_email,
                    "LinksHandles": links,
                    "BusinessName": business,
                    "Address": address,
                    "Latitude": latitude,
                    "Longitude": longitude,
                    "Notes": notes,
                    "StartDate": now,
                    "LastUpdated": now
                }

                updated_df = pd.concat(
                    [reps_df, pd.DataFrame([new_row])],
                    ignore_index=True
                )

                if save_reps(updated_df):
                    st.success(f"Added {full_name} as {new_rep_id}.")
                    st.rerun()

    st.markdown("---")

    st.subheader("Edit Existing Reps")
    st.info("Edit reps below, then click Save Changes to update Google Sheets.")

    editable_df = reps_df.copy()

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
