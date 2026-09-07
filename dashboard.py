"""
AI Sales Force Tracking Dashboard
----------------------------------
"SO Activity" o "Attendance" — duita sheet theke Sales Officer der
performance, attendance, o outlet visit report dekhay.

No-upload workflow (director/boss just opens the link):
1. Notun daily Excel file ke ei repo-r `data/SO_Activity_Attendance.xlsx`
   path-e rekhe (same filename-e overwrite kore) GitHub-e push korun.
2. Streamlit Community Cloud automatically redeploy hoye notun data show
   korbe — kaউকে kichu upload korte hobe na, shudhu link open korlei
   sob theke shesh data dekha jabe.

Run locally:
    streamlit run dashboard.py
"""

import os
import re
from datetime import datetime

import pandas as pd
import plotly.express as px
import streamlit as st

# ------------------------------------------------------------------
# ====================  EASY-TO-EDIT BRANDING  ======================
# Company naam / title / tagline ekhane change korun — code-er baki
# jaiga touch korar dorkar nai.
# ------------------------------------------------------------------
COMPANY_NAME = "ACI Premio Plastics"
DASHBOARD_TITLE = "AI Sales Force Tracking Dashboard"
DASHBOARD_TAGLINE = f"{COMPANY_NAME} — Sales Officer Activity, Attendance & Productivity Analytics"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
ASSETS_DIR = os.path.join(BASE_DIR, "assets")
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(ASSETS_DIR, exist_ok=True)

# Ei fixed file-ta protidin replace kore GitHub-e push korben.
DATA_FILE_PATH = os.path.join(DATA_DIR, "SO_Activity_Attendance.xlsx")

# Optional branding images — thakle automatically use hobe, na thakle
# CSS-based placeholder logo dekhabe.
LOGO_PATH = os.path.join(ASSETS_DIR, "logo.png")       # sidebar top card
BADGE_PATH = os.path.join(ASSETS_DIR, "badge.png")     # circular badge

DAILY_OUTLET_TARGET = 20      # "good performer" target per active day
BAD_PERFORMER_OUTLET_LIMIT = 15   # below this = bad performer
ACHIEVEMENT_RED_THRESHOLD = 60    # % niche hole red mark

# ------------------------------------------------------------------
# Page config
# ------------------------------------------------------------------
st.set_page_config(
    page_title=DASHBOARD_TITLE,
    page_icon="📊",
    layout="wide",
)

# ------------------------------------------------------------------
# Global CSS — dark navy sidebar + gradient KPI cards + gradient title
# ------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* ---- Sidebar ---- */
    section[data-testid="stSidebar"] {
        background-color: #0f1729;
    }
    section[data-testid="stSidebar"] * {
        color: #e8eaf2 !important;
    }
    section[data-testid="stSidebar"] .stDateInput input {
        color: #1f2333 !important;
    }
    section[data-testid="stSidebar"] label {
        font-weight: 600 !important;
    }

    .sidebar-logo-card {
        background: #ffffff;
        border-radius: 16px;
        padding: 22px 16px;
        text-align: center;
        margin-bottom: 18px;
        box-shadow: 0 6px 18px rgba(0,0,0,0.25);
    }
    .sidebar-logo-card img { max-width: 100%; max-height: 90px; }
    .sidebar-logo-placeholder {
        font-weight: 800;
        font-size: 20px;
        letter-spacing: 1px;
        background: linear-gradient(135deg, #6C5DD3, #ec4899);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }

    .sidebar-badge {
        width: 78px; height: 78px;
        border-radius: 50%;
        background: #ffffff;
        margin: 0 auto 22px auto;
        display: flex; align-items: center; justify-content: center;
        box-shadow: 0 6px 16px rgba(0,0,0,0.3);
        font-weight: 800; color: #16a34a; font-size: 13px;
        overflow: hidden;
    }
    .sidebar-badge img { width: 100%; height: 100%; object-fit: cover; }

    .sidebar-filters-title {
        font-size: 22px;
        font-weight: 800;
        margin-bottom: 14px;
        color: #ffffff !important;
    }

    /* ---- Header ---- */
    .dash-header-title {
        font-size: 30px;
        font-weight: 800;
        background: linear-gradient(90deg, #6C5DD3, #ec4899);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 2px;
    }
    .dash-header-sub { color: #6b7280; font-size: 15px; margin-top: 0; }
    .dash-header-badge {
        width: 64px; height: 64px; border-radius: 50%;
        background: #ffffff; box-shadow: 0 6px 16px rgba(0,0,0,0.18);
        display: flex; align-items: center; justify-content: center;
        margin-left: auto; overflow: hidden;
        font-weight: 800; color: #16a34a; font-size: 12px;
    }
    .dash-header-badge img { width: 100%; height: 100%; object-fit: cover; }

    /* ---- KPI cards ---- */
    .kpi-card {
        border-radius: 18px;
        padding: 22px 10px 26px 10px;
        text-align: center;
        color: #ffffff;
        min-height: 118px;
    }
    .kpi-label { font-weight: 700; font-size: 14px; opacity: 0.95; }
    .kpi-value { font-weight: 800; font-size: 32px; margin-top: 10px; }
    .kpi-purple { background: linear-gradient(135deg,#8b7cf6,#6c5ce7); box-shadow: 0 10px 22px rgba(108,92,231,0.35); }
    .kpi-green  { background: linear-gradient(135deg,#34d399,#10b981); box-shadow: 0 10px 22px rgba(16,185,129,0.35); }
    .kpi-blue   { background: linear-gradient(135deg,#38bdf8,#0ea5e9); box-shadow: 0 10px 22px rgba(14,165,233,0.35); }
    .kpi-orange { background: linear-gradient(135deg,#fbbf24,#f59e0b); box-shadow: 0 10px 22px rgba(245,158,11,0.35); }
    .kpi-pink   { background: linear-gradient(135deg,#f472b6,#ec4899); box-shadow: 0 10px 22px rgba(236,72,153,0.35); }

    .section-title { font-size: 24px; font-weight: 800; margin-top: 6px; }
    </style>
    """,
    unsafe_allow_html=True,
)


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def parse_total_hours_to_minutes(value):
    """Convert strings like '10h 34m' / '0h 0m' -> total minutes (int)."""
    if pd.isna(value):
        return None
    match = re.match(r"\s*(\d+)h\s*(\d+)m", str(value))
    if not match:
        return None
    hours, minutes = int(match.group(1)), int(match.group(2))
    return hours * 60 + minutes


def minutes_to_hm(total_minutes):
    """Convert minutes (float) -> 'Xh Ym' display string."""
    if total_minutes is None or pd.isna(total_minutes):
        return "-"
    total_minutes = int(round(total_minutes))
    h, m = divmod(total_minutes, 60)
    return f"{h}h {m}m"


def time_to_minutes(t):
    """Convert a datetime.time / string 'HH:MM:SS' -> minutes since midnight."""
    if pd.isna(t):
        return None
    if hasattr(t, "hour"):
        return t.hour * 60 + t.minute + t.second / 60
    return None


def minutes_to_hhmm(m):
    if m is None or pd.isna(m):
        return "-"
    m = int(round(m))
    h, mm = divmod(m, 60)
    return f"{h:02d}:{mm:02d}"


def style_red_flag(val, threshold=ACHIEVEMENT_RED_THRESHOLD):
    """Return background style for a percentage value under threshold."""
    try:
        if pd.isna(val):
            return ""
        return "background-color: #ffcccc; color: #7a0000; font-weight: 600;" if val < threshold else ""
    except Exception:
        return ""


# ------------------------------------------------------------------
# Data loading
# ------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_workbook(file_bytes, file_key):
    """Read both sheets and do basic cleaning. Cached on file bytes hash (file_key)."""
    so = pd.read_excel(file_bytes, sheet_name="SO Activity")
    att = pd.read_excel(file_bytes, sheet_name="Attendance")

    # ---- SO Activity cleaning ----
    so["Visit Date"] = pd.to_datetime(so["Visit Date"]).dt.normalize()
    so["Staff ID"] = so["Staff ID"].astype(str).str.strip()
    for col in ["Number Of Outlet Visit", "Number Of Ordered Shop"]:
        so[col] = pd.to_numeric(so[col], errors="coerce").fillna(0)

    # ---- Attendance cleaning ----
    att["Attendance Date"] = pd.to_datetime(att["Attendance Date"]).dt.normalize()
    att["Staff Id"] = att["Staff Id"].astype(str).str.strip()
    att["In Time (min)"] = pd.to_datetime(
        att["In Time Only"], format="%H:%M:%S", errors="coerce"
    ).dt.time.map(time_to_minutes)
    att["Total Hours (min)"] = att["Total Hours"].map(parse_total_hours_to_minutes)

    return so, att


def build_so_master(so, att):
    """One row per SO Staff ID with basic identity info (name/group/region/zone/RSM/base)."""
    so_info = (
        so.sort_values("Visit Date")
        .groupby("Staff ID")
        .agg(
            SR_Name=("SR Name", "last"),
            SubBusiness=("SubBusiness", "last"),
            Region=("Region Name", "last"),
            Zone=("Zone Name", "last"),
            RSM=("RSM", "last"),
            Base=("Base Name", "last"),
        )
        .reset_index()
        .rename(columns={"Staff ID": "StaffID"})
    )

    att_info = (
        att.sort_values("Attendance Date")
        .groupby("Staff Id")
        .agg(
            Employee_Name=("Employee Name", "last"),
            Group=("Group Name", "last"),
        )
        .reset_index()
        .rename(columns={"Staff Id": "StaffID"})
    )

    master = pd.merge(so_info, att_info, on="StaffID", how="outer")
    master["SR_Name"] = master["SR_Name"].fillna(master["Employee_Name"])
    master["SubBusiness"] = master["SubBusiness"].fillna(master["Group"])
    master = master.drop(columns=["Employee_Name", "Group"])
    master["SR_Name"] = master["SR_Name"].fillna(master["StaffID"])
    return master


def build_attendance_summary(att):
    """Per-SO attendance / late summary for the selected period."""
    valid = att[att["Attendance Status"].isin(["Present", "Late"])].copy()

    # Market-stay time only makes sense on days a checkout was actually recorded
    # (no checkout -> Total Hours shows as 0h 0m, which would understate the average).
    valid["Has_Checkout"] = valid["Check Out Status"].notna()
    checkout_rows = valid[valid["Has_Checkout"]]

    summary = (
        valid.groupby("Staff Id")
        .agg(
            Active_Days=("Attendance Date", "nunique"),
            Late_Days=("Attendance Status", lambda s: (s == "Late").sum()),
            Avg_InTime_min=("In Time (min)", "mean"),
        )
        .reset_index()
        .rename(columns={"Staff Id": "StaffID"})
    )

    stay = (
        checkout_rows.groupby("Staff Id")
        .agg(
            Avg_MarketStay_min=("Total Hours (min)", "mean"),
            Checkout_Days=("Attendance Date", "nunique"),
        )
        .reset_index()
        .rename(columns={"Staff Id": "StaffID"})
    )
    summary = summary.merge(stay, on="StaffID", how="left")
    summary["Present_Days"] = summary["Active_Days"] - summary["Late_Days"]
    summary["Late_%"] = (summary["Late_Days"] / summary["Active_Days"] * 100).round(1)
    summary["Avg_InTime"] = summary["Avg_InTime_min"].map(minutes_to_hhmm)
    summary["Avg_MarketStay"] = summary["Avg_MarketStay_min"].map(minutes_to_hm)

    absent = (
        att[att["Attendance Status"] == "Absent"]
        .groupby("Staff Id")
        .size()
        .rename("Absent_Days")
        .reset_index()
        .rename(columns={"Staff Id": "StaffID"})
    )
    summary = pd.merge(summary, absent, on="StaffID", how="left")
    summary["Absent_Days"] = summary["Absent_Days"].fillna(0).astype(int)
    return summary


def build_activity_summary(so):
    """Per-SO outlet-visit / target-achievement summary for the selected period."""
    summary = (
        so.groupby("Staff ID")
        .agg(
            Visit_Active_Days=("Visit Date", "nunique"),
            Total_Outlet_Visit=("Number Of Outlet Visit", "sum"),
            Total_Ordered_Shop=("Number Of Ordered Shop", "sum"),
        )
        .reset_index()
        .rename(columns={"Staff ID": "StaffID"})
    )
    summary["Avg_Visit_Per_Day"] = (
        summary["Total_Outlet_Visit"] / summary["Visit_Active_Days"]
    ).round(1)
    summary["Target_Visits"] = summary["Visit_Active_Days"] * DAILY_OUTLET_TARGET
    summary["Achievement_%"] = (
        summary["Total_Outlet_Visit"] / summary["Target_Visits"] * 100
    ).round(1)
    return summary


# ------------------------------------------------------------------
# Sidebar — file upload / persistence
# ------------------------------------------------------------------
st.sidebar.title("⚙️ Data Upload")
# ------------------------------------------------------------------
# Sidebar — logo card + circular badge (image if present, else placeholder)
# ------------------------------------------------------------------
if os.path.exists(LOGO_PATH):
    st.sidebar.markdown('<div class="sidebar-logo-card">', unsafe_allow_html=True)
    st.sidebar.image(LOGO_PATH, use_container_width=True)
    st.sidebar.markdown("</div>", unsafe_allow_html=True)
else:
    st.sidebar.markdown(
        f"""<div class="sidebar-logo-card">
                <div class="sidebar-logo-placeholder">{COMPANY_NAME}</div>
            </div>""",
        unsafe_allow_html=True,
    )

if os.path.exists(BADGE_PATH):
    st.sidebar.markdown(
        f'<div class="sidebar-badge"><img src="data:image/png;base64,'
        f'{__import__("base64").b64encode(open(BADGE_PATH, "rb").read()).decode()}"></div>',
        unsafe_allow_html=True,
    )
else:
    st.sidebar.markdown('<div class="sidebar-badge">LOGO</div>', unsafe_allow_html=True)

st.sidebar.markdown('<div class="sidebar-filters-title">Dashboard Filters</div>', unsafe_allow_html=True)

# ------------------------------------------------------------------
# Load the fixed data file (no upload — replace the file in the repo
# and push to GitHub to update).
# ------------------------------------------------------------------
if not os.path.exists(DATA_FILE_PATH):
    st.markdown(f'<div class="dash-header-title">{DASHBOARD_TITLE}</div>', unsafe_allow_html=True)
    st.markdown(f'<p class="dash-header-sub">{DASHBOARD_TAGLINE}</p>', unsafe_allow_html=True)
    st.info(
        f"⚠️ Data file paওয়া jayni. Excel file-ta ei path-e rakhun:\n\n"
        f"`{DATA_FILE_PATH}`\n\n"
        "(sheets: 'SO Activity' o 'Attendance') — tarpor GitHub-e push korun, "
        "app automatically notun data load korbe."
    )
    st.stop()

file_key = f"{os.path.getmtime(DATA_FILE_PATH)}"
so_raw, att_raw = load_workbook(DATA_FILE_PATH, file_key)
last_modified = datetime.fromtimestamp(os.path.getmtime(DATA_FILE_PATH))

# ------------------------------------------------------------------
# Sidebar — filters (Start / End Date shown as separate boxes)
# ------------------------------------------------------------------
min_date = min(so_raw["Visit Date"].min(), att_raw["Attendance Date"].min())
max_date = max(so_raw["Visit Date"].max(), att_raw["Attendance Date"].max())

start_date = st.sidebar.date_input(
    "Start Date", value=min_date, min_value=min_date, max_value=max_date
)
end_date = st.sidebar.date_input(
    "End Date", value=max_date, min_value=min_date, max_value=max_date
)
start_date, end_date = pd.to_datetime(start_date), pd.to_datetime(end_date)
if start_date > end_date:
    st.sidebar.warning("Start Date, End Date-er por hote parbe na — date range shodhon kora hoyeche.")
    start_date, end_date = end_date, start_date

with st.sidebar.expander("More Filters (Group / Region / Zone / RSM)", expanded=False):
    groups = sorted(so_raw["SubBusiness"].dropna().unique().tolist())
    sel_groups = st.multiselect("Group (SubBusiness)", groups, default=groups)

    regions = sorted(so_raw["Region Name"].dropna().unique().tolist())
    sel_regions = st.multiselect("Region", regions, default=regions)

    zones_available = sorted(
        so_raw[so_raw["Region Name"].isin(sel_regions)]["Zone Name"].dropna().unique().tolist()
    )
    sel_zones = st.multiselect("Zone", zones_available, default=zones_available)

    rsms_available = sorted(
        so_raw[so_raw["Zone Name"].isin(sel_zones)]["RSM"].dropna().unique().tolist()
    )
    sel_rsms = st.multiselect("RSM", rsms_available, default=rsms_available)

st.sidebar.caption(f"Last data update: {last_modified.strftime('%d %b %Y, %I:%M %p')}")

# ------------------------------------------------------------------
# Apply filters
# ------------------------------------------------------------------
so = so_raw[
    (so_raw["Visit Date"] >= start_date)
    & (so_raw["Visit Date"] <= end_date)
    & (so_raw["SubBusiness"].isin(sel_groups))
    & (so_raw["Region Name"].isin(sel_regions))
    & (so_raw["Zone Name"].isin(sel_zones))
    & (so_raw["RSM"].isin(sel_rsms))
].copy()

filtered_staff_ids = set(so["Staff ID"].unique())

att = att_raw[
    (att_raw["Attendance Date"] >= start_date)
    & (att_raw["Attendance Date"] <= end_date)
    & (att_raw["Group Name"].isin(sel_groups) if sel_groups else True)
].copy()
# Keep only attendance rows whose staff also appear in the filtered SO activity
# (falls back to full attendance set if the SO filter removes everyone —
# e.g. Region/Zone/RSM filters don't exist in the Attendance sheet).
if filtered_staff_ids:
    att = att[att["Staff Id"].isin(filtered_staff_ids) | att["Staff Id"].notna()]

master = build_so_master(so_raw, att_raw)
att_summary = build_attendance_summary(att)
act_summary = build_activity_summary(so)

report = master.merge(att_summary, on="StaffID", how="left").merge(
    act_summary, on="StaffID", how="left"
)
report = report[report["StaffID"].isin(filtered_staff_ids)].reset_index(drop=True)


# ------------------------------------------------------------------
# Header (gradient title + tagline + circular badge on the right)
# ------------------------------------------------------------------
h_left, h_right = st.columns([6, 1])
with h_left:
    st.markdown(f'<div class="dash-header-title">{DASHBOARD_TITLE}</div>', unsafe_allow_html=True)
    st.markdown(f'<p class="dash-header-sub">{DASHBOARD_TAGLINE}</p>', unsafe_allow_html=True)
with h_right:
    if os.path.exists(BADGE_PATH):
        st.markdown(
            f'<div class="dash-header-badge"><img src="data:image/png;base64,'
            f'{__import__("base64").b64encode(open(BADGE_PATH, "rb").read()).decode()}"></div>',
            unsafe_allow_html=True,
        )
    else:
        st.markdown('<div class="dash-header-badge">LOGO</div>', unsafe_allow_html=True)

st.caption(
    f"Report Period: **{start_date.strftime('%d %b %Y')} → {end_date.strftime('%d %b %Y')}**"
)
st.markdown("<hr style='margin-top:6px;'>", unsafe_allow_html=True)

# ------------------------------------------------------------------
# KPI cards (gradient, matches the reference design)
# ------------------------------------------------------------------
kpi_total_so = int(report["StaffID"].nunique())
kpi_late_pct = f"{report['Late_%'].mean():.1f}%" if report["Late_%"].notna().any() else "-"
kpi_achv_pct = (
    f"{report['Achievement_%'].mean():.1f}%" if report["Achievement_%"].notna().any() else "-"
)
kpi_total_visit = int(report["Total_Outlet_Visit"].fillna(0).sum())
kpi_red_flag = int((report["Achievement_%"] < ACHIEVEMENT_RED_THRESHOLD).sum())

kpi_defs = [
    ("Sales Officer", kpi_total_so, "kpi-purple"),
    ("Total Outlet Visit", kpi_total_visit, "kpi-green"),
    ("Avg Late %", kpi_late_pct, "kpi-blue"),
    ("Avg Achievement %", kpi_achv_pct, "kpi-orange"),
    ("Red-Flag SO (<60%)", kpi_red_flag, "kpi-pink"),
]

kpi_cols = st.columns(5)
for col, (label, value, css_class) in zip(kpi_cols, kpi_defs):
    with col:
        st.markdown(
            f"""
            <div class="kpi-card {css_class}">
                <div class="kpi-label">{label}</div>
                <div class="kpi-value">{value}</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

st.markdown("<hr>", unsafe_allow_html=True)
st.markdown('<div class="section-title">Overall Summary</div>', unsafe_allow_html=True)

tabs = st.tabs(
    [
        "🕒 Attendance & Late",
        "🚶 First Visit & Market Stay",
        "🏆 Performance (Outlet Visit)",
        "📈 Group & RSM Charts",
        "🗺️ Route & Repeat-Market Analysis",
        "📋 Full Data Table",
    ]
)

# ------------------------------------------------------------------
# TAB 1 — Attendance & Late
# ------------------------------------------------------------------
with tabs[0]:
    st.subheader("Daily Late Count & Total Late %")
    st.caption(
        "Je koidin attendance data ache, shei koidin dhore hoy Late na hoy "
        "Present (on time) dhora hoyeche. Achievement % attendance er upor na."
    )

    daily_table = report[
        [
            "StaffID",
            "SR_Name",
            "SubBusiness",
            "RSM",
            "Active_Days",
            "Present_Days",
            "Late_Days",
            "Absent_Days",
            "Late_%",
        ]
    ].sort_values("Late_%", ascending=False)

    styled = daily_table.style.map(
        lambda v: "background-color:#ffcccc; color:#7a0000; font-weight:600;"
        if isinstance(v, (int, float)) and not pd.isna(v) and v >= 30
        else "",
        subset=["Late_%"],
    ).format({"Late_%": "{:.1f}%"})
    st.dataframe(styled, width='stretch', height=420)

    c1, c2 = st.columns(2)
    with c1:
        top_late = daily_table.head(15)
        fig = px.bar(
            top_late,
            x="Late_%",
            y="SR_Name",
            orientation="h",
            color="Late_%",
            color_continuous_scale="Reds",
            title="Top 15 — Highest Late %",
        )
        fig.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, width='stretch')
    with c2:
        grp_late = report.groupby("SubBusiness", as_index=False)["Late_%"].mean()
        fig2 = px.bar(
            grp_late,
            x="SubBusiness",
            y="Late_%",
            color="SubBusiness",
            title="Group-wise Average Late %",
            text_auto=".1f",
        )
        st.plotly_chart(fig2, width='stretch')

# ------------------------------------------------------------------
# TAB 2 — First Visit & Market Stay
# ------------------------------------------------------------------
with tabs[1]:
    st.subheader("First Visit (Check-in) Time & Market Stay / Idle Time")
    st.info(
        "ℹ️ SO Activity sheet e kono 'visit time' na thakay, protidin er "
        "**first check-in time** ke 'First Visit' hisebe o **Total Hours "
        "(field time, checkout recorded thakle)** ke 'Market Stay' hisebe "
        "dhora hoyeche. Onno kono time-stamped field thakle janaben, oi "
        "onujayi update kore deওয়া jabe.\n\n"
        "⚠️ Note: Onek din checkout record thake na (app e out-time deনি), "
        "shei din gulo 'Market Stay' calculation e bad deওয়া hoyeche — tai "
        "**Checkout Days** column dekhe bujhben koto din er data die average ta hoyeche."
    )

    visit_table = report[
        [
            "StaffID",
            "SR_Name",
            "SubBusiness",
            "RSM",
            "Avg_InTime",
            "Avg_MarketStay",
            "Checkout_Days",
        ]
    ].sort_values("SR_Name")
    st.dataframe(visit_table, width='stretch', height=380)

    c1, c2 = st.columns(2)
    with c1:
        plot_df = report.dropna(subset=["Avg_InTime_min"]).sort_values("Avg_InTime_min")
        fig = px.bar(
            plot_df.head(20),
            x="Avg_InTime_min",
            y="SR_Name",
            orientation="h",
            title="Earliest Average First-Visit / Check-in (Top 20)",
            labels={"Avg_InTime_min": "Minutes after midnight"},
        )
        fig.update_layout(yaxis={"categoryorder": "total descending"})
        st.plotly_chart(fig, width='stretch')
    with c2:
        plot_df2 = report.dropna(subset=["Avg_MarketStay_min"]).sort_values(
            "Avg_MarketStay_min"
        )
        fig2 = px.bar(
            plot_df2.head(20),
            x="Avg_MarketStay_min",
            y="SR_Name",
            orientation="h",
            title="Lowest Average Market Stay Time (possible idle time)",
            labels={"Avg_MarketStay_min": "Minutes"},
            color_discrete_sequence=["#e67e22"],
        )
        fig2.update_layout(yaxis={"categoryorder": "total descending"})
        st.plotly_chart(fig2, width='stretch')

# ------------------------------------------------------------------
# TAB 3 — Performance (Outlet Visit)
# ------------------------------------------------------------------
with tabs[2]:
    st.subheader("Outlet Visit Target Achievement")
    st.caption(
        f"Target = Active Days × {DAILY_OUTLET_TARGET} outlets/day. "
        f"Achievement % {ACHIEVEMENT_RED_THRESHOLD}% er niche hole 🔴 Red mark."
    )

    perf_table = report[
        [
            "StaffID",
            "SR_Name",
            "SubBusiness",
            "RSM",
            "Visit_Active_Days",
            "Total_Outlet_Visit",
            "Avg_Visit_Per_Day",
            "Target_Visits",
            "Achievement_%",
        ]
    ].sort_values("Achievement_%")

    styled_perf = perf_table.style.map(
        style_red_flag, subset=["Achievement_%"]
    ).format({"Achievement_%": "{:.1f}%"})
    st.dataframe(styled_perf, width='stretch', height=420)

    st.divider()
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 🏆 Top 10 Good Performer")
        st.caption(f"Criteria: Avg Outlet Visit/day ≥ {DAILY_OUTLET_TARGET}")
        good = report[report["Avg_Visit_Per_Day"] >= DAILY_OUTLET_TARGET].sort_values(
            "Avg_Visit_Per_Day", ascending=False
        ).head(10)
        if good.empty:
            good = report.sort_values("Avg_Visit_Per_Day", ascending=False).head(10)
            st.caption("(Kew exact target chhuiynai — tai overall top 10 dekhano hocche)")
        fig = px.bar(
            good,
            x="Avg_Visit_Per_Day",
            y="SR_Name",
            orientation="h",
            color="Avg_Visit_Per_Day",
            color_continuous_scale="Greens",
        )
        fig.update_layout(yaxis={"categoryorder": "total ascending"})
        st.plotly_chart(fig, width='stretch')

    with c2:
        st.markdown("### ⚠️ Top 10 Bad Performer")
        st.caption(f"Criteria: Avg Outlet Visit/day < {BAD_PERFORMER_OUTLET_LIMIT}")
        bad = report[report["Avg_Visit_Per_Day"] < BAD_PERFORMER_OUTLET_LIMIT].sort_values(
            "Avg_Visit_Per_Day"
        ).head(10)
        if bad.empty:
            bad = report.sort_values("Avg_Visit_Per_Day").head(10)
            st.caption("(Kew threshold er niche na — tai overall bottom 10 dekhano hocche)")
        fig2 = px.bar(
            bad,
            x="Avg_Visit_Per_Day",
            y="SR_Name",
            orientation="h",
            color="Avg_Visit_Per_Day",
            color_continuous_scale="Reds_r",
        )
        fig2.update_layout(yaxis={"categoryorder": "total descending"})
        st.plotly_chart(fig2, width='stretch')

# ------------------------------------------------------------------
# TAB 4 — Group & RSM Charts
# ------------------------------------------------------------------
with tabs[3]:
    st.subheader("Group-wise Charts")
    c1, c2 = st.columns(2)
    with c1:
        grp_visit = so.groupby("SubBusiness", as_index=False)["Number Of Outlet Visit"].sum()
        fig = px.pie(
            grp_visit,
            names="SubBusiness",
            values="Number Of Outlet Visit",
            title="Total Outlet Visit share by Group",
            hole=0.4,
        )
        st.plotly_chart(fig, width='stretch')
    with c2:
        grp_ach = report.groupby("SubBusiness", as_index=False)["Achievement_%"].mean()
        fig2 = px.bar(
            grp_ach,
            x="SubBusiness",
            y="Achievement_%",
            color="SubBusiness",
            title="Group-wise Avg Achievement %",
            text_auto=".1f",
        )
        st.plotly_chart(fig2, width='stretch')

    st.markdown("#### Group-wise Top / Bad Performer (Column Chart)")
    grp_pick = st.selectbox("Group select korun", groups, key="grp_chart_select")
    grp_data = report[report["SubBusiness"] == grp_pick].sort_values(
        "Avg_Visit_Per_Day", ascending=False
    )
    c3, c4 = st.columns(2)
    with c3:
        fig3 = px.bar(
            grp_data.head(10),
            x="SR_Name",
            y="Avg_Visit_Per_Day",
            title=f"{grp_pick} — Top 10 Performer",
            color_discrete_sequence=["#2ecc71"],
        )
        st.plotly_chart(fig3, width='stretch')
    with c4:
        fig4 = px.bar(
            grp_data.tail(10),
            x="SR_Name",
            y="Avg_Visit_Per_Day",
            title=f"{grp_pick} — Bottom 10 Performer",
            color_discrete_sequence=["#e74c3c"],
        )
        st.plotly_chart(fig4, width='stretch')

    st.divider()
    st.subheader("RSM-wise Top / Low Performer")
    rsm_agg = (
        report.groupby(["RSM", "SR_Name"], as_index=False)["Avg_Visit_Per_Day"]
        .mean()
        .dropna()
    )
    rsm_top = rsm_agg.sort_values(["RSM", "Avg_Visit_Per_Day"], ascending=[True, False]).groupby(
        "RSM"
    ).head(1)
    rsm_low = rsm_agg.sort_values(["RSM", "Avg_Visit_Per_Day"], ascending=[True, True]).groupby(
        "RSM"
    ).head(1)

    c5, c6 = st.columns(2)
    with c5:
        fig5 = px.bar(
            rsm_top.sort_values("Avg_Visit_Per_Day", ascending=False),
            x="RSM",
            y="Avg_Visit_Per_Day",
            color="RSM",
            text="SR_Name",
            title="RSM-wise Top Sales Officer",
        )
        fig5.update_traces(textposition="outside")
        st.plotly_chart(fig5, width='stretch')
    with c6:
        fig6 = px.bar(
            rsm_low.sort_values("Avg_Visit_Per_Day"),
            x="RSM",
            y="Avg_Visit_Per_Day",
            color="RSM",
            text="SR_Name",
            title="RSM-wise Lowest Performing Sales Officer",
        )
        fig6.update_traces(textposition="outside")
        st.plotly_chart(fig6, width='stretch')

# ------------------------------------------------------------------
# TAB 5 — Route & Repeat-Market Analysis
# ------------------------------------------------------------------
with tabs[4]:
    st.subheader("SO Day-wise Most-visited Route")
    st.caption(
        "Protidin (weekday) SO shob cheye beshi outlet visit korche kon Route e — "
        "shei route highlight kora holo."
    )
    so["Weekday"] = so["Visit Date"].dt.day_name()
    route_day = (
        so.groupby(["Staff ID", "Weekday", "Route Name"], as_index=False)[
            "Number Of Outlet Visit"
        ].sum()
    )
    idx = route_day.groupby(["Staff ID", "Weekday"])["Number Of Outlet Visit"].idxmax()
    top_route_per_day = route_day.loc[idx].merge(
        master[["StaffID", "SR_Name"]], left_on="Staff ID", right_on="StaffID", how="left"
    )
    top_route_per_day = top_route_per_day[
        ["SR_Name", "Weekday", "Route Name", "Number Of Outlet Visit"]
    ].sort_values(["SR_Name", "Weekday"])
    st.dataframe(top_route_per_day, width='stretch', height=380)

    st.divider()
    st.subheader("Repeat Market Visit Check")
    st.caption(
        "Protyek Market ID-r nirdishto (assigned) Route Day thake. Ei report e "
        "dekhano hocche kon SO kon Market ekbar-er beshi (different date-e) "
        "visit korche, ebong actual visit-day assigned Route Day theke match "
        "kore kina."
    )
    market_visits = (
        so.groupby(["Staff ID", "Market ID", "Market Name", "Route Day"])
        .agg(
            Visit_Dates=("Visit Date", lambda s: ", ".join(sorted(d.strftime("%d-%b") for d in s.unique()))),
            Visit_Count=("Visit Date", "nunique"),
            Actual_Weekdays=("Weekday", lambda s: ", ".join(sorted(s.unique()))),
        )
        .reset_index()
    )
    repeat_markets = market_visits[market_visits["Visit_Count"] > 1].copy()

    if repeat_markets.empty:
        st.info("Selected filter-e kono repeat market visit paওয়া jayni.")
    else:
        def _check_mismatch(row):
            weekdays = set(row["Actual_Weekdays"].split(", "))
            return "⚠️ Yes" if (row["Route Day"] not in weekdays) or (len(weekdays) > 1) else "No"

        repeat_markets["Day_Mismatch"] = repeat_markets.apply(_check_mismatch, axis=1)
        repeat_markets = repeat_markets.merge(
            master[["StaffID", "SR_Name"]], left_on="Staff ID", right_on="StaffID", how="left"
        )
        repeat_markets = repeat_markets[
            [
                "SR_Name",
                "Market Name",
                "Route Day",
                "Actual_Weekdays",
                "Visit_Count",
                "Visit_Dates",
                "Day_Mismatch",
            ]
        ].sort_values("Visit_Count", ascending=False)
        st.dataframe(repeat_markets, width='stretch', height=380)

# ------------------------------------------------------------------
# TAB 6 — Full Data Table (export-friendly)
# ------------------------------------------------------------------
with tabs[5]:
    st.subheader("Full Consolidated SO Report")
    export_cols = [
        "StaffID",
        "SR_Name",
        "SubBusiness",
        "Region",
        "Zone",
        "RSM",
        "Base",
        "Active_Days",
        "Present_Days",
        "Late_Days",
        "Absent_Days",
        "Late_%",
        "Avg_InTime",
        "Avg_MarketStay",
        "Checkout_Days",
        "Visit_Active_Days",
        "Total_Outlet_Visit",
        "Avg_Visit_Per_Day",
        "Target_Visits",
        "Achievement_%",
    ]
    final_table = report[export_cols].sort_values("SR_Name")
    st.dataframe(final_table, width='stretch', height=500)
    st.download_button(
        "⬇️ Download as CSV",
        data=final_table.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"SO_Report_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}.csv",
        mime="text/csv",
    )
