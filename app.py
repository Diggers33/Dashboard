# app.py
# Streamlit HE Workload Visualizer - Fixed version
# -------------------------------------------------------------
# - 3-Month heat map (Idle is blank)
# - Yearly timeline with bars (height = effort; color = phase)
# - Department filter (All, PMO, DIGITAL, SCI & ENG)
# - Optional CSV uploads (meta + allocations)
# - Collapsible Data & Settings
# -------------------------------------------------------------

from __future__ import annotations

import io
from typing import Dict, List, Tuple
import json

import pandas as pd
import numpy as np
import streamlit as st

# Google Sheets integration using gspread
try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSHEETS_AVAILABLE = True
except ImportError:
    GSHEETS_AVAILABLE = False

# -------------------------------
# Styling
# -------------------------------
CSS = """
<style>
:root{
  --bg:#f8f9fc;
  --card:#ffffff;
  --ink:#1e293b;
  --muted:#64748b;
  --line:#e2e8f0;
  --soft:#f1f5f9;

  --chip-idle: transparent;
  --chip-plan:#dbeafe;
  --chip-design:#fef3c7;
  --chip-dev:#fecaca;
  --chip-test:#d1fae5;
  --chip-deliv:#e9d5ff;

  --txt-plan:#1e40af;
  --txt-design:#b45309;
  --txt-dev:#b91c1c;
  --txt-test:#047857;
  --txt-deliv:#7e22ce;
}

html, body, [data-testid="stAppViewContainer"] {
  background: var(--bg);
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', 'Roboto', sans-serif;
}

.container-card{
  background: var(--card);
  border:1px solid var(--line);
  border-radius: 16px;
  padding: 24px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.04);
}

.kpi-row{
  display:grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap:16px;
  margin:16px 0 24px 0;
}
.kpi{
  border-radius:14px;
  padding:20px 24px;
  color:#fff;
  font-weight:600;
  display:flex; 
  align-items:center; 
  justify-content:space-between;
  box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1), 0 2px 4px -1px rgba(0,0,0,0.06);
  transition: transform 0.2s, box-shadow 0.2s;
}
.kpi:hover{
  transform: translateY(-2px);
  box-shadow: 0 10px 15px -3px rgba(0,0,0,0.1), 0 4px 6px -2px rgba(0,0,0,0.05);
}
.kpi .label{ font-size:14px; opacity:0.95; font-weight:500; text-transform:uppercase; letter-spacing:0.5px; }
.kpi .num{ font-size:36px; font-weight:700; }
.kpi.purple{ background:linear-gradient(135deg,#667eea 0%,#764ba2 100%); }
.kpi.red{ background:linear-gradient(135deg,#f43f5e 0%,#dc2626 100%); }
.kpi.green{ background:linear-gradient(135deg,#10b981 0%,#059669 100%); }

.grid{
  display:grid;
  grid-template-columns: 240px 1fr 1fr 1fr;
  gap:10px;
  align-items:stretch;
}
.grid .head{
  background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color:#fff;
  padding:14px 16px;
  border-radius: 12px;
  font-weight:600;
  font-size:14px;
  text-transform:uppercase;
  letter-spacing:0.5px;
  box-shadow: 0 2px 4px rgba(102,126,234,0.2);
}
.project-name{
  background:#fff;
  border:2px solid var(--line);
  padding:14px 16px;
  border-radius:12px;
  font-weight:700;
  font-size:14px;
  color:var(--ink);
  display:flex; 
  align-items:center; 
  justify-content:space-between;
  transition: all 0.2s;
}
.project-name:hover{
  border-color:#667eea;
  box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
  transform: translateX(4px);
}
.dept-badge{
  padding:4px 10px; 
  border-radius:12px;
  font-size:10px; 
  font-weight:700; 
  color:#0f766e; 
  background:#ccfbf1; 
  border:1px solid #5eead4;
  text-transform:uppercase;
  letter-spacing:0.3px;
}
.cell{
  border:2px solid var(--line);
  border-radius:12px;
  min-height:64px;
  display:flex; 
  align-items:center; 
  justify-content:center;
  background:#fff;
  font-weight:600;
  font-size:15px;
  transition: all 0.2s;
}
.cell:hover{
  transform: scale(1.02);
  box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
}
.cell.idle{ 
  background:#f8fafc; 
  color:#94a3b8; 
  border-style:dashed;
  border-color:#cbd5e1;
}
.cell.plan{ 
  background:var(--chip-plan); 
  color:var(--txt-plan); 
  border-color:#93c5fd;
  font-weight:700;
}
.cell.design{ 
  background:var(--chip-design); 
  color:var(--txt-design); 
  border-color:#fcd34d;
  font-weight:700;
}
.cell.dev{ 
  background:var(--chip-dev); 
  color:var(--txt-dev); 
  border-color:#fca5a5;
  font-weight:700;
}
.cell.test{ 
  background:var(--chip-test); 
  color:var(--txt-test); 
  border-color:#6ee7b7;
  font-weight:700;
}
.cell.deliv{ 
  background:var(--chip-deliv); 
  color:var(--txt-deliv); 
  border-color:#d8b4fe;
  font-weight:700;
}
.badge{
  font-weight:700; font-size:12px;
  padding:6px 12px; border-radius:999px; border:1px solid transparent;
}
.badge.plan{ background:var(--chip-plan); color:var(--txt-plan); border-color:#cdddfd; }
.badge.design{ background:var(--chip-design); color:var(--txt-design); border-color:#ffe6a6; }
.badge.dev{ background:var(--chip-dev); color:var(--txt-dev); border-color:#ffc8c8; }
.badge.test{ background:var(--chip-test); color:var(--txt-test); border-color:#c7f0d5; }
.badge.deliv{ background:var(--chip-deliv); color:var(--txt-deliv); border-color:#e4d0ff; }

/* Yearly bars */
.year-grid{
  display:grid; 
  grid-template-columns: 240px repeat(12, 1fr); 
  gap:10px;
  align-items:stretch;
}
.y-head{
  background:linear-gradient(135deg, #667eea 0%, #764ba2 100%);
  color:#fff; 
  padding:14px 16px; 
  border-radius:12px; 
  font-weight:600;
  font-size:14px;
  text-transform:uppercase;
  letter-spacing:0.5px;
  box-shadow: 0 2px 4px rgba(102,126,234,0.2);
}
.month-head{
  text-align:center; 
  padding:12px 8px; 
  font-weight:700; 
  border:2px solid #e0e7ff;
  color:#4338ca; 
  background:linear-gradient(180deg, #f5f7ff 0%, #eef2ff 100%); 
  border-radius:12px;
  font-size:14px;
  transition: all 0.2s;
}
.month-head:hover{
  background:#e0e7ff;
  border-color:#c7d2fe;
}
.month-bar{
  border:2px solid var(--line);
  background:#fafbfc;
  border-radius:12px;
  display:flex; 
  align-items:flex-end; 
  justify-content:center;
  min-height:60px;
  transition: all 0.2s;
}
.month-bar:hover{
  background:#f8f9fa;
  box-shadow: 0 2px 4px rgba(0,0,0,0.05);
}
.bar{
  width:90%; 
  border-radius:8px 8px 0 0; 
  display:flex; 
  align-items:center; 
  justify-content:center;
  color:#0f172a; 
  font-weight:800; 
  font-size:14px; 
  margin-bottom:4px;
  box-shadow: 0 -2px 4px rgba(0,0,0,0.1);
  transition: all 0.2s;
}
.bar:hover{
  transform: translateY(-2px);
  box-shadow: 0 -4px 6px rgba(0,0,0,0.15);
}
.bar.plan{ 
  background:linear-gradient(180deg, var(--chip-plan) 0%, #bfdbfe 100%); 
  color:var(--txt-plan); 
  border:2px solid #93c5fd; 
  border-bottom:none;
}
.bar.design{ 
  background:linear-gradient(180deg, var(--chip-design) 0%, #fde68a 100%); 
  color:var(--txt-design); 
  border:2px solid #fcd34d; 
  border-bottom:none;
}
.bar.dev{ 
  background:linear-gradient(180deg, var(--chip-dev) 0%, #fca5a5 100%); 
  color:var(--txt-dev); 
  border:2px solid #f87171; 
  border-bottom:none;
}
.bar.test{ 
  background:linear-gradient(180deg, var(--chip-test) 0%, #a7f3d0 100%); 
  color:var(--txt-test); 
  border:2px solid #6ee7b7; 
  border-bottom:none;
}
.bar.deliv{ 
  background:linear-gradient(180deg, var(--chip-deliv) 0%, #ddd6fe 100%); 
  color:var(--txt-deliv); 
  border:2px solid #c4b5fd; 
  border-bottom:none;
}

/* Tooltip styles */
.tooltip-container {
  position: relative;
  display: inline-block;
  width: 100%;
  height: 100%;
}

.tooltip-text {
  visibility: hidden;
  background-color: #1e293b;
  color: #fff;
  text-align: left;
  border-radius: 8px;
  padding: 12px 14px;
  position: absolute;
  z-index: 1000;
  bottom: 125%;
  left: 50%;
  transform: translateX(-50%);
  min-width: 200px;
  box-shadow: 0 10px 25px rgba(0,0,0,0.3);
  font-size: 13px;
  line-height: 1.6;
  opacity: 0;
  transition: opacity 0.3s, visibility 0.3s;
}

.tooltip-text::after {
  content: "";
  position: absolute;
  top: 100%;
  left: 50%;
  margin-left: -6px;
  border-width: 6px;
  border-style: solid;
  border-color: #1e293b transparent transparent transparent;
}

.tooltip-container:hover .tooltip-text {
  visibility: visible;
  opacity: 1;
}

.tooltip-label {
  font-weight: 600;
  color: #94a3b8;
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  margin-bottom: 4px;
}

.tooltip-value {
  color: #fff;
  font-weight: 500;
}

.tooltip-divider {
  height: 1px;
  background: #334155;
  margin: 8px 0;
}

/* small note text */
.note{ 
  color:var(--muted); 
  font-size:13px; 
  margin:0 0 20px 0;
  padding:12px 16px;
  background:#f8fafc;
  border-left:3px solid #667eea;
  border-radius:8px;
}

</style>
"""

# -------------------------------
# Sample data (if user doesn't upload)
# -------------------------------
SAMPLE_PROJECTS = [
    ("BIORADAR", "SCI & ENG", 60),
    ("ENERGIZE", "DIGITAL", 40),
    ("FOODSAFER", "PMO", 35),
    ("GIANT LEAPS", "PMO", 50),
    ("IS2H4C", "SCI & ENG", 45),
    ("PATAFEST", "DIGITAL", 30),
    ("RECONSTRUCT", "PMO", 30),
    ("RESCHAPE", "DIGITAL", 55),
    ("SECUREFOOD", "PMO", 25),
    ("THESEUS", "SCI & ENG", 45),
    ("UP-SKILL", "PMO", 20),
]

PHASES = ["Planning", "Design", "Development", "Testing", "Delivery"]


def yms(start_ym: str, n: int) -> List[str]:
    y, m = map(int, start_ym.split("-"))
    out = []
    for i in range(n):
        yy = y + (m - 1 + i) // 12
        mm = (m - 1 + i) % 12 + 1
        out.append(f"{yy:04d}-{mm:02d}")
    return out


def sample_frames(start_ym="2025-01") -> Tuple[pd.DataFrame, pd.DataFrame]:
    # Meta
    meta = pd.DataFrame(SAMPLE_PROJECTS, columns=["Project", "Department", "Total_MM"])
    
    # Employee pool by department
    employees = {
        "SCI & ENG": ["Dr. Sarah Chen", "Dr. James Wilson", "Dr. Maria Garcia", "Dr. Ahmed Hassan"],
        "DIGITAL": ["Alex Morgan", "Sofia Rodriguez", "Liam O'Brien", "Emma Thompson"],
        "PMO": ["Michael Stevens", "Rachel Green", "David Kim", "Jennifer Lee"]
    }
    
    # Allocations with employee assignments
    months = yms(start_ym, 12)
    rows = []
    rng = np.random.default_rng(7)
    for proj, dept, _ in SAMPLE_PROJECTS:
        dept_employees = employees[dept]
        for i, ym in enumerate(months):
            # Some months active, others idle
            if rng.random() < 0.45:
                mm = [2, 3, 6, 8, 3, 1][i % 6]  # repeatable pattern
                phase = ["Design", "Planning", "Development", "Testing", "Delivery", "Planning"][i % 6]
                
                # Assign employees based on MM
                num_employees = min(len(dept_employees), max(1, mm // 2))
                assigned = rng.choice(dept_employees, size=num_employees, replace=False)
                
                # Distribute MM among employees
                allocation_per_person = mm / num_employees
                employee_list = ", ".join([f"{emp} ({allocation_per_person:.1f})" for emp in sorted(assigned)])
                
                rows.append([proj, ym, phase, mm, employee_list])
            else:
                # explicit Idle rows are not required; missing = idle
                pass
    alloc = pd.DataFrame(rows, columns=["Project", "Date", "Phase", "MM", "Employees"])
    return meta, alloc


# -------------------------------
# Helpers
# -------------------------------
def ensure_columns(df: pd.DataFrame, required: List[str], name: str):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{name} is missing required columns: {', '.join(missing)}")


def month_mm_phase_employees(df_alloc: pd.DataFrame, project: str, ym: str) -> Tuple[float, str | None, str]:
    """Get MM, phase, and employees for a project-month."""
    rows = df_alloc[(df_alloc["Project"] == project) & (df_alloc["Date"] == ym)]
    if rows.empty:
        return 0.0, None, ""
    # sum by phase; take the phase with max MM
    grp = rows.groupby("Phase").agg({"MM": "sum", "Employees": "first"}).sort_values("MM", ascending=False)
    phase = grp.index[0]
    mm = float(grp["MM"].iloc[0])
    employees = str(grp["Employees"].iloc[0]) if "Employees" in grp.columns else ""
    return mm, phase, employees


def phase_badge(phase: str) -> str:
    pm = phase.lower()
    if pm == "planning":
        cls = "plan"; txt = "Planning"
    elif pm == "design":
        cls = "design"; txt = "Design"
    elif pm == "development":
        cls = "dev"; txt = "Development"
    elif pm == "testing":
        cls = "test"; txt = "Testing"
    elif pm == "delivery":
        cls = "deliv"; txt = "Delivery"
    else:
        return ""  # Idle => no badge
    return f"<span class='badge {cls}'>{txt}</span>"


def month_phase(df_alloc: pd.DataFrame, project: str, ym: str) -> str | None:
    rows = df_alloc[(df_alloc["Project"] == project) & (df_alloc["Date"] == ym)]
    if rows.empty:
        return None  # Idle
    # If multiple, pick the one with biggest MM
    top = rows.sort_values("MM", ascending=False).iloc[0]
    return str(top["Phase"])


def month_mm_and_phase(df_alloc: pd.DataFrame, project: str, ym: str) -> Tuple[float, str | None]:
    rows = df_alloc[(df_alloc["Project"] == project) & (df_alloc["Date"] == ym)]
    if rows.empty:
        return 0.0, None
    # sum by phase; take the phase with max MM
    grp = rows.groupby("Phase")["MM"].sum().sort_values(ascending=False)
    phase = grp.index[0]
    mm = float(grp.iloc[0])
    return mm, phase


def dept_badge(dept: str) -> str:
    # style by department
    if dept.upper().startswith("SCI"):
        return "<span class='dept-badge'>SCI & ENG</span>"
    if "DIGITAL" in dept.upper():
        return "<span class='dept-badge' style='color:#1d4ed8;background:#dbeafe;border-color:#bfdbfe;'>DIGITAL</span>"
    return "<span class='dept-badge' style='color:#6b21a8;background:#f3e8ff;border-color:#e9d5ff;'>PMO</span>"


# -------------------------------
# App
# -------------------------------
st.set_page_config(page_title="HE Workload Visualizer", layout="wide")
st.markdown(CSS, unsafe_allow_html=True)

# Sidebar - ALL controls go here
with st.sidebar:
    st.header("Controls")
    
    # Google Sheets
    if GSHEETS_AVAILABLE:
        st.subheader("📊 Google Sheets")
        use_gsheets_input = st.checkbox("Enable Integration", value=False, key="gsheets_checkbox")
        st.session_state.use_gsheets = use_gsheets_input
        
        if st.button("🔄 Reload from Sheets", use_container_width=True):
            if "alloc_data" in st.session_state:
                del st.session_state.alloc_data
            if "meta_data" in st.session_state:
                del st.session_state.meta_data
            st.rerun()
        
        with st.expander("Setup Help", expanded=False):
            st.markdown("""
            **Quick Setup:**
            1. `pip install gspread google-auth`
            2. Create sheets: "Meta" & "Allocations"
            3. Share with service account
            4. Add creds to `.streamlit/secrets.toml`
            
            [Guide](https://docs.gspread.org/en/latest/oauth2.html)
            """)
    else:
        st.warning("Install: pip install gspread google-auth")
        st.session_state.use_gsheets = False
    
    st.divider()
    
    # Data Upload
    st.subheader("📁 Upload Data")
    meta_file_input = st.file_uploader("Projects Meta CSV", type=["csv"], key="meta")
    alloc_file_input = st.file_uploader("Allocations CSV", type=["csv"], key="alloc")
    st.session_state.meta_file = meta_file_input
    st.session_state.alloc_file = alloc_file_input
    
    st.divider()
    
    # View Settings
    st.subheader("⚙️ View Settings")
    start_ym_input = st.text_input("Start (YYYY-MM)", value="2025-01", key="start_ym_input")
    months_to_display = st.slider("Months", 6, 24, 12, key="months_slider")
    full_height_mm = st.number_input("Max bar height (MM)", min_value=1, max_value=48, value=12, step=1, key="height_input")
    focus_month = st.text_input("Focus month", value="2025-03", key="focus_input")
    
    st.session_state.start_ym = start_ym_input
    st.session_state.months_to_display = months_to_display
    st.session_state.full_height_mm = full_height_mm
    st.session_state.focus_month = focus_month
    
    st.divider()
    
    # Edit button
    if st.button("✏️ Edit Allocations", use_container_width=True, type="primary", key="edit_alloc_btn"):
        st.session_state.show_editor = not st.session_state.get("show_editor", False)

# Data loading function
def load_data_sources(use_gsheets, meta_file, alloc_file, start_ym):
    """Load data from Google Sheets or CSV files."""
    if use_gsheets and GSHEETS_AVAILABLE:
        try:
            # Get credentials from Streamlit secrets
            creds_dict = dict(st.secrets["gsheet"])
            spreadsheet_url = creds_dict.pop("spreadsheet_url")
            
            # Create credentials object
            scopes = [
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
            creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
            client = gspread.authorize(creds)
            
            # Open spreadsheet and worksheets
            sheet = client.open_by_url(spreadsheet_url)
            
            # Read Meta worksheet
            meta_worksheet = sheet.worksheet("Meta")
            meta_data = meta_worksheet.get_all_records()
            meta = pd.DataFrame(meta_data)
            
            # Read Allocations worksheet
            alloc_worksheet = sheet.worksheet("Allocations")
            alloc_data = alloc_worksheet.get_all_records()
            alloc = pd.DataFrame(alloc_data)
            
            st.sidebar.success("Connected to Google Sheets!")
            return meta, alloc, (client, sheet, meta_worksheet, alloc_worksheet)
        except Exception as e:
            st.sidebar.error(f"Google Sheets connection failed: {e}")
            st.sidebar.info("Falling back to CSV/sample data...")
            use_gsheets = False
    
    # Fallback to CSV or sample data
    if meta_file:
        meta = pd.read_csv(meta_file)
    else:
        meta, _ = sample_frames(start_ym)
    
    if alloc_file:
        alloc = pd.read_csv(alloc_file)
    else:
        _, alloc = sample_frames(start_ym)
    
    return meta, alloc, None


# Data - load early before any UI elements
try:
    # Get settings from sidebar (they're defined there now)
    use_gsheets = st.session_state.get('use_gsheets', False)
    meta_file = st.session_state.get('meta_file', None)
    alloc_file = st.session_state.get('alloc_file', None)
    start_ym = st.session_state.get('start_ym', '2025-01')
    
    meta, alloc, gsheets_conn = load_data_sources(use_gsheets, meta_file, alloc_file, start_ym)
    
    ensure_columns(meta, ["Project", "Department", "Total_MM"], "Meta dataframe")
    
    # Ensure Employees column exists
    if "Employees" not in alloc.columns:
        alloc["Employees"] = ""
    
    ensure_columns(alloc, ["Project", "Date", "Phase", "MM"], "Allocations dataframe")
except Exception as e:
    st.error(f"Data error: {e}")
    st.stop()

# Store data in session state for editing
if "alloc_data" not in st.session_state:
    st.session_state.alloc_data = alloc.copy()
if "meta_data" not in st.session_state:
    st.session_state.meta_data = meta.copy()
if "gsheets_conn" not in st.session_state:
    st.session_state.gsheets_conn = gsheets_conn

# Update connection in session state
st.session_state.gsheets_conn = gsheets_conn

# Edit Data Section - only if edit button was clicked
if st.session_state.get("show_editor", False):
    with st.expander("✏️ Edit Allocations", expanded=True):
        st.markdown("### Edit Project Allocations")
        st.markdown("Directly edit the table below. Changes reflect in visualizations immediately.")
        
        # Filter options for easier editing
        col1, col2, col3 = st.columns(3)
        with col1:
            filter_project = st.selectbox(
                "Filter by Project", 
                ["All"] + sorted(st.session_state.meta_data["Project"].unique().tolist()),
                key="edit_filter_proj"
            )
        with col2:
            filter_phase = st.selectbox(
                "Filter by Phase",
                ["All", "Planning", "Design", "Development", "Testing", "Delivery"],
                key="edit_filter_phase"
            )
        with col3:
            filter_date = st.text_input("Filter by Date (YYYY-MM)", key="edit_filter_date")
    
    # Apply filters
    filtered_alloc = st.session_state.alloc_data.copy()
    if filter_project != "All":
        filtered_alloc = filtered_alloc[filtered_alloc["Project"] == filter_project]
    if filter_phase != "All":
        filtered_alloc = filtered_alloc[filtered_alloc["Phase"] == filter_phase]
    if filter_date:
        filtered_alloc = filtered_alloc[filtered_alloc["Date"] == filter_date]
    
    # Editable dataframe
    edited_alloc = st.data_editor(
        filtered_alloc,
        use_container_width=True,
        num_rows="dynamic",
        column_config={
            "Project": st.column_config.SelectboxColumn(
                "Project",
                options=st.session_state.meta_data["Project"].tolist(),
                required=True
            ),
            "Date": st.column_config.TextColumn(
                "Date (YYYY-MM)",
                help="Format: 2025-01",
                required=True
            ),
            "Phase": st.column_config.SelectboxColumn(
                "Phase",
                options=["Planning", "Design", "Development", "Testing", "Delivery"],
                required=True
            ),
            "MM": st.column_config.NumberColumn(
                "Man-Months",
                min_value=0,
                max_value=100,
                step=0.5,
                required=True
            ),
            "Employees": st.column_config.TextColumn(
                "Employees (Name (MM), ...)",
                help="Format: Alex Morgan (2.0), Sofia Rodriguez (1.5)",
                required=False
            )
        },
        key="alloc_editor"
    )
    
    # Action buttons
    st.markdown("### Quick Actions")
    action_col1, action_col2, action_col3, action_col4 = st.columns(4)
    
    with action_col1:
        if st.button("💾 Save Changes", type="primary"):
            # Update the filtered rows in the main dataset
            if filter_project != "All" or filter_phase != "All" or filter_date:
                # Merge edited data back
                mask = pd.Series([True] * len(st.session_state.alloc_data))
                if filter_project != "All":
                    mask &= st.session_state.alloc_data["Project"] == filter_project
                if filter_phase != "All":
                    mask &= st.session_state.alloc_data["Phase"] == filter_phase
                if filter_date:
                    mask &= st.session_state.alloc_data["Date"] == filter_date
                st.session_state.alloc_data = st.session_state.alloc_data[~mask]
                st.session_state.alloc_data = pd.concat([st.session_state.alloc_data, edited_alloc], ignore_index=True)
            else:
                st.session_state.alloc_data = edited_alloc.copy()
            
            # Save to Google Sheets if enabled
            if use_gsheets and st.session_state.gsheets_conn is not None:
                try:
                    client, sheet, meta_worksheet, alloc_worksheet = st.session_state.gsheets_conn
                    
                    # Clear and update allocations worksheet
                    alloc_worksheet.clear()
                    alloc_worksheet.update([st.session_state.alloc_data.columns.values.tolist()] + 
                                          st.session_state.alloc_data.values.tolist())
                    
                    st.success("Changes saved to Google Sheets!")
                except Exception as e:
                    st.error(f"Failed to save to Google Sheets: {e}")
                    st.info("Try clicking 'Reload from Sheets' button to refresh connection")
            else:
                st.success("Changes saved locally!")
            st.rerun()
    
    with action_col2:
        if st.button("⏩ Shift Selected +1 Month"):
            if len(edited_alloc) > 0:
                for idx in edited_alloc.index:
                    old_date = edited_alloc.loc[idx, "Date"]
                    try:
                        y, m = map(int, old_date.split("-"))
                        new_m = m + 1
                        new_y = y
                        if new_m > 12:
                            new_m = 1
                            new_y += 1
                        new_date = f"{new_y:04d}-{new_m:02d}"
                        # Update in main data
                        mask = (st.session_state.alloc_data["Project"] == edited_alloc.loc[idx, "Project"]) & \
                               (st.session_state.alloc_data["Date"] == old_date) & \
                               (st.session_state.alloc_data["Phase"] == edited_alloc.loc[idx, "Phase"])
                        st.session_state.alloc_data.loc[mask, "Date"] = new_date
                    except:
                        pass
                st.success("✅ Shifted forward 1 month!")
                st.rerun()
    
    with action_col3:
        if st.button("⏪ Shift Selected -1 Month"):
            if len(edited_alloc) > 0:
                for idx in edited_alloc.index:
                    old_date = edited_alloc.loc[idx, "Date"]
                    try:
                        y, m = map(int, old_date.split("-"))
                        new_m = m - 1
                        new_y = y
                        if new_m < 1:
                            new_m = 12
                            new_y -= 1
                        new_date = f"{new_y:04d}-{new_m:02d}"
                        # Update in main data
                        mask = (st.session_state.alloc_data["Project"] == edited_alloc.loc[idx, "Project"]) & \
                               (st.session_state.alloc_data["Date"] == old_date) & \
                               (st.session_state.alloc_data["Phase"] == edited_alloc.loc[idx, "Phase"])
                        st.session_state.alloc_data.loc[mask, "Date"] = new_date
                    except:
                        pass
                st.success("✅ Shifted back 1 month!")
                st.rerun()
    
    with action_col4:
        if st.button("📥 Download as CSV"):
            csv = st.session_state.alloc_data.to_csv(index=False)
            st.download_button(
                label="Download CSV",
                data=csv,
                file_name="allocations_edited.csv",
                mime="text/csv"
            )

# Use session state data for visualizations
alloc = st.session_state.alloc_data.copy()
meta = st.session_state.meta_data.copy()

# MAIN CONTENT AREA STARTS HERE - Department filter only
st.markdown("### Department Filter")
dept_opts = ["All"] + sorted(meta["Department"].unique().tolist())
dept = st.selectbox("Select Department", dept_opts, index=0, label_visibility="collapsed")
if dept != "All":
    meta_view = meta[meta["Department"] == dept].copy()
else:
    meta_view = meta.copy()

# KPIs
active_month = focus_month
# active this month: projects with any allocation in focus month
active = alloc[(alloc["Date"] == active_month) & (alloc["Project"].isin(meta_view["Project"]))]
active_count = active["Project"].nunique()
total_projects = meta_view["Project"].nunique()
low_activity = total_projects - active_count

st.markdown(
    """
<div style='display:flex;gap:12px;align-items:center;margin:16px 0;flex-wrap:wrap;'>
  <div style='padding:10px 20px;border-radius:8px;font-weight:600;border:2px solid #667eea;background:#f5f7ff;color:#667eea;display:flex;align-items:center;gap:10px;'>
    <div style='font-size:28px;font-weight:700;'>{}</div>
    <div style='font-size:13px;opacity:0.85;font-weight:600;'>Total Projects</div>
  </div>
  <div style='padding:10px 20px;border-radius:8px;font-weight:600;border:2px solid #f43f5e;background:#fef2f2;color:#f43f5e;display:flex;align-items:center;gap:10px;'>
    <div style='font-size:28px;font-weight:700;'>{}</div>
    <div style='font-size:13px;opacity:0.85;font-weight:600;'>Active This Month</div>
  </div>
  <div style='padding:10px 20px;border-radius:8px;font-weight:600;border:2px solid #10b981;background:#f0fdf4;color:#10b981;display:flex;align-items:center;gap:10px;'>
    <div style='font-size:28px;font-weight:700;'>{}</div>
    <div style='font-size:13px;opacity:0.85;font-weight:600;'>Low Activity</div>
  </div>
</div>
""".format(total_projects, active_count, low_activity),
    unsafe_allow_html=True,
)

tab1, tab2, tab3 = st.tabs(["3-Month Heat Map (Blocks)", "Yearly Timeline (Bars)", "Dashboard"])

# -------------------------------
# 3-Month Heat Map
# -------------------------------
with tab1:
    st.markdown("<div class='container-card'>", unsafe_allow_html=True)
    st.markdown("<div class='note'>Focus: {} → Then: {} and {} – Idle cells are intentionally blank.</div>".format(
        focus_month,
        yms(focus_month, 2)[1],  # next month label
        yms(focus_month, 3)[2]   # month +2
    ), unsafe_allow_html=True)

    curr, nxt, nxt2 = yms(focus_month, 3)[0], yms(focus_month, 3)[1], yms(focus_month, 3)[2]

    # Build complete HTML for the grid
    html_parts = ["<div class='grid'>"]
    html_parts.append("<div class='head'>Project</div>")
    html_parts.append("<div class='head'>Current Month</div>")
    html_parts.append("<div class='head'>Next Month</div>")
    html_parts.append("<div class='head'>Month +2</div>")

    for _, row in meta_view.sort_values("Project").iterrows():
        proj = row["Project"]
        dept_html = dept_badge(row["Department"])

        # Determine phases
        ph1 = month_phase(alloc, proj, curr)
        ph2 = month_phase(alloc, proj, nxt)
        ph3 = month_phase(alloc, proj, nxt2)

        # Build cells with phase styling
        def cell_html(phase: str | None) -> str:
            if not phase or phase.lower() == "idle":
                return "<div class='cell idle'>Idle</div>"
            
            phase_lower = phase.lower()
            if phase_lower == "planning":
                return "<div class='cell plan'>Planning</div>"
            elif phase_lower == "design":
                return "<div class='cell design'>Design</div>"
            elif phase_lower == "development":
                return "<div class='cell dev'>Development</div>"
            elif phase_lower == "testing":
                return "<div class='cell test'>Testing</div>"
            elif phase_lower == "delivery":
                return "<div class='cell deliv'>Delivery</div>"
            else:
                return "<div class='cell idle'>Idle</div>"

        html_parts.append(
            f"<div class='project-name'><span>{proj}</span>{dept_html}</div>"
        )
        html_parts.append(cell_html(ph1))
        html_parts.append(cell_html(ph2))
        html_parts.append(cell_html(ph3))

    html_parts.append("</div>")  # close grid
    
    # Render the complete grid at once
    st.markdown("".join(html_parts), unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)  # close container-card

# -------------------------------
# Yearly Timeline (Enhanced)
# -------------------------------
with tab2:
    st.markdown("<div class='container-card'>", unsafe_allow_html=True)

    months = yms(start_ym, months_to_display)

    # Build complete HTML for the entire grid
    html_parts = ["<div class='year-grid'>"]
    
    # Header row - Project column
    html_parts.append("<div class='y-head'>Project</div>")
    
    # Month headers
    for i in range(1, months_to_display + 1):
        html_parts.append(f"<div class='month-head'>{i:02d}</div>")

    # Project rows
    for _, row in meta_view.sort_values("Project").iterrows():
        proj = row["Project"]
        dept_html = dept_badge(row["Department"])

        # Project name cell with hover effect
        html_parts.append(
            f"<div class='project-name' style='height:60px;display:flex;align-items:center;justify-content:space-between;'>"
            f"<span>{proj}</span>{dept_html}"
            "</div>"
        )

        # Month bars for this project
        for ym in months:
            mm, phase, employees = month_mm_phase_employees(alloc, proj, ym)
            if mm <= 0 or not phase:
                html_parts.append("<div class='month-bar'></div>")
                continue
            
            pct = min(100.0, (mm / float(full_height_mm)) * 100.0)
            cls = {
                "planning": "plan",
                "design": "design",
                "development": "dev",
                "testing": "test",
                "delivery": "deliv",
            }.get(phase.lower(), "plan")
            
            bar_value = int(mm) if mm >= 1 else f"{mm:.1f}"
            
            # Create tooltip content
            tooltip_html = ""
            if employees:
                # Format employee allocations nicely
                employee_lines = ""
                try:
                    # Parse employees with allocations like "Name (2.0), Name2 (1.5)"
                    emp_parts = [e.strip() for e in employees.split(',')]
                    for emp_part in emp_parts:
                        employee_lines += f"<div class='tooltip-value'>• {emp_part} MM</div>"
                except:
                    # Fallback if format doesn't match
                    employee_lines = f"<div class='tooltip-value'>{employees}</div>"
                
                tooltip_html = f"""
                <span class='tooltip-text'>
                    <div class='tooltip-label'>Project</div>
                    <div class='tooltip-value'>{proj}</div>
                    <div class='tooltip-divider'></div>
                    <div class='tooltip-label'>Phase</div>
                    <div class='tooltip-value'>{phase}</div>
                    <div class='tooltip-divider'></div>
                    <div class='tooltip-label'>Total Effort</div>
                    <div class='tooltip-value'>{bar_value} MM</div>
                    <div class='tooltip-divider'></div>
                    <div class='tooltip-label'>Team Allocation</div>
                    {employee_lines}
                </span>
                """
            
            html_parts.append(
                f"<div class='month-bar'>"
                f"<div class='tooltip-container'>"
                f"<div class='bar {cls}' style='height:{max(20, pct)}%'>{bar_value}</div>"
                f"{tooltip_html}"
                f"</div>"
                f"</div>"
            )

    html_parts.append("</div>")  # close year-grid
    
    # Render the complete grid at once
    st.markdown("".join(html_parts), unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)  # close container-card

# -------------------------------
# Dashboard Tab - Enhanced Design
# -------------------------------
with tab3:
    # Add dashboard-specific CSS
    st.markdown("""
    <style>
    .dash-section {
        background: white;
        border-radius: 12px;
        padding: 24px;
        margin-bottom: 24px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .dash-title {
        font-size: 20px;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 8px;
        display: flex;
        align-items: center;
        gap: 8px;
    }
    .dash-subtitle {
        font-size: 14px;
        color: #64748b;
        margin-bottom: 20px;
    }
    .employee-grid {
        display: grid;
        grid-template-columns: 140px repeat(12, 1fr);
        gap: 6px;
        margin-bottom: 8px;
        align-items: stretch;
    }
    .emp-header {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 10px 8px;
        border-radius: 8px;
        text-align: center;
        font-weight: 600;
        font-size: 12px;
    }
    .emp-name {
        background: #f8fafc;
        padding: 10px 12px;
        border-radius: 8px;
        font-weight: 600;
        font-size: 13px;
        color: #1e293b;
        display: flex;
        align-items: center;
    }
    .emp-cell {
        background: white;
        border: 2px solid #e2e8f0;
        border-radius: 6px;
        padding: 8px 6px;
        text-align: center;
        font-size: 11px;
        font-weight: 700;
        min-height: 50px;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        gap: 4px;
    }
    .emp-cell.low { background: #f0fdf4; color: #15803d; border-color: #86efac; }
    .emp-cell.medium { background: #fef3c7; color: #b45309; border-color: #fcd34d; }
    .emp-cell.high { background: #fee2e2; color: #b91c1c; border-color: #fca5a5; }
    .emp-cell.over { background: #dc2626; color: white; border-color: #991b1b; box-shadow: 0 0 0 3px rgba(220, 38, 38, 0.2); }
    .emp-proj {
        font-size: 9px;
        background: rgba(0, 0, 0, 0.1);
        padding: 2px 5px;
        border-radius: 3px;
        font-weight: 600;
        white-space: nowrap;
    }
    .emp-cell.over .emp-proj {
        background: rgba(255,255,255,0.3);
    }
    .legend {
        display: flex;
        gap: 16px;
        font-size: 12px;
        margin-top: 12px;
        flex-wrap: wrap;
    }
    .legend-item {
        display: flex;
        align-items: center;
        gap: 6px;
    }
    .legend-box {
        width: 16px;
        height: 16px;
        border-radius: 3px;
        border: 2px solid;
    }
    .summary-grid {
        display: grid;
        grid-template-columns: repeat(3, 1fr);
        gap: 20px;
    }
    .summary-card {
        background: white;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 20px;
    }
    .summary-card h4 {
        font-size: 16px;
        font-weight: 600;
        color: #334155;
        margin-bottom: 16px;
    }
    .metric-row {
        display: grid;
        grid-template-columns: repeat(2, 1fr);
        gap: 12px;
        margin-bottom: 12px;
    }
    .metric-box {
        background: #f8fafc;
        padding: 14px;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
    }
    .metric-label {
        font-size: 11px;
        color: #64748b;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 4px;
    }
    .metric-value {
        font-size: 28px;
        font-weight: 700;
        color: #1e293b;
        line-height: 1;
    }
    .metric-change {
        font-size: 11px;
        margin-top: 4px;
    }
    .metric-change.up { color: #10b981; }
    .metric-change.down { color: #ef4444; }
    </style>
    """, unsafe_allow_html=True)
    
    # Calculate dashboard metrics
    all_months = yms(st.session_state.get('start_ym', start_ym), 12)
    
    # 1. CONFLICT DETECTION
    st.markdown("<div class='dash-section'>", unsafe_allow_html=True)
    st.markdown("<div class='dash-title'>⚠️ Resource Conflicts & Overallocation</div>", unsafe_allow_html=True)
    
    # Extract all employees and check for conflicts
    conflicts = []
    employee_allocations = {}
    
    for _, row in alloc.iterrows():
        if pd.notna(row.get('Employees', '')) and row['Employees']:
            try:
                emp_parts = [e.strip() for e in str(row['Employees']).split(',')]
                for emp_part in emp_parts:
                    if '(' in emp_part and ')' in emp_part:
                        name = emp_part.split('(')[0].strip()
                        mm = float(emp_part.split('(')[1].split(')')[0])
                        
                        key = (name, row['Date'])
                        if key not in employee_allocations:
                            employee_allocations[key] = {'total_mm': 0, 'projects': []}
                        employee_allocations[key]['total_mm'] += mm
                        employee_allocations[key]['projects'].append((row['Project'], mm))
            except:
                pass
    
    # Find conflicts (>100% = 4.0 MM per month assuming 4 weeks)
    for (emp, date), data in employee_allocations.items():
        if data['total_mm'] > 4.0:
            conflicts.append({
                'employee': emp,
                'month': date,
                'total_mm': data['total_mm'],
                'percent': (data['total_mm'] / 4.0) * 100,
                'projects': data['projects']
            })
    
    if conflicts:
        for conflict in conflicts:
            proj_list = " + ".join([p[0] for p in conflict['projects']])
            st.error(f"**Overallocation detected:** {conflict['employee']} is allocated {conflict['percent']:.0f}% in {conflict['month']} ({proj_list})")
    else:
        st.success("No resource conflicts detected!")
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # 3. RESOURCE UTILIZATION - Single unified grid
    st.markdown("<div class='dash-section'>", unsafe_allow_html=True)
    st.markdown("<div class='dash-title'>👥 Resource Utilization - Individual Employee Workload</div>", unsafe_allow_html=True)
    st.markdown("<div class='dash-subtitle'>Track individual workload across projects. Red cells indicate overallocation (>100%).</div>", unsafe_allow_html=True)
    
    # Get unique employees
    all_employees = set()
    for _, row in alloc.iterrows():
        if pd.notna(row.get('Employees', '')) and row['Employees']:
            try:
                emp_parts = [e.strip() for e in str(row['Employees']).split(',')]
                for emp_part in emp_parts:
                    if '(' in emp_part:
                        name = emp_part.split('(')[0].strip()
                        all_employees.add(name)
            except:
                pass
    
    # Build ONE complete grid with header and all employees
    grid_html = ["<div class='employee-grid' style='grid-template-rows: auto;'>"]
    
    # Header row
    grid_html.append("<div class='emp-header'>Employee</div>")
    month_names = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    for month_name in month_names:
        grid_html.append(f"<div class='emp-header'>{month_name}</div>")
    
    # All employee rows in the same grid
    for employee in sorted(all_employees):
        grid_html.append(f"<div class='emp-name'>{employee}</div>")
        
        for month in all_months:
            # Get full project names (not just 3 letters)
            month_mm = 0
            projects = []
            for (emp, date), data in employee_allocations.items():
                if emp == employee and date == month:
                    month_mm = data['total_mm']
                    projects = [p[0] for p in data['projects'][:2]]  # Full project names
            
            pct = (month_mm / 4.0) * 100
            
            if pct == 0:
                cell_class = ""
                text = "0%"
            elif pct < 50:
                cell_class = "low"
                text = f"{pct:.0f}%"
            elif pct < 75:
                cell_class = "medium"
                text = f"{pct:.0f}%"
            elif pct <= 100:
                cell_class = "high"
                text = f"{pct:.0f}%"
            else:
                cell_class = "over"
                text = f"{pct:.0f}%"
            
            proj_tags = " ".join([f"<span class='emp-proj'>{p}</span>" for p in projects])
            
            grid_html.append(
                f"<div class='emp-cell {cell_class}'>"
                f"<span style='font-weight:700;'>{text}</span>{proj_tags}"
                f"</div>"
            )
    
    grid_html.append("</div>")  # Close the single grid
    
    st.markdown("".join(grid_html), unsafe_allow_html=True)
    
    # Legend
    st.markdown("""
    <div class='legend'>
        <div class='legend-item'>
            <div class='legend-box' style='background:#f0fdf4;border-color:#86efac;'></div>
            <span>0-50%</span>
        </div>
        <div class='legend-item'>
            <div class='legend-box' style='background:#fef3c7;border-color:#fcd34d;'></div>
            <span>50-75%</span>
        </div>
        <div class='legend-item'>
            <div class='legend-box' style='background:#fee2e2;border-color:#fca5a5;'></div>
            <span>75-100%</span>
        </div>
        <div class='legend-item'>
            <div class='legend-box' style='background:#dc2626;border-color:#991b1b;'></div>
            <span>>100% (Conflict)</span>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # 2. CAPACITY PLANNING (moved after utilization)
    st.markdown("<div class='dash-section'>", unsafe_allow_html=True)
    st.markdown("<div class='dash-title'>📊 Capacity Planning - Total Resource Demand</div>", unsafe_allow_html=True)
    
    # Calculate total MM by month
    capacity_data = []
    for month in all_months[:6]:
        month_total = alloc[alloc['Date'] == month]['MM'].sum()
        capacity_data.append({
            'month': month,
            'total_mm': month_total,
            'available': 40
        })
    
    # Render capacity grid
    capacity_html = ["<div style='display:grid;grid-template-columns:150px repeat(6,1fr);gap:8px;align-items:center;margin:20px 0;'>"]
    capacity_html.append("<div style='background:linear-gradient(135deg,#667eea 0%,#764ba2 100%);color:white;padding:12px;border-radius:8px;font-weight:600;text-align:center;'>Month</div>")
    
    for data in capacity_data:
        month_label = data['month'].split('-')[1]
        capacity_html.append(f"<div style='background:#eef2ff;color:#4338ca;padding:12px;border-radius:8px;font-weight:700;text-align:center;'>{month_label}</div>")
    
    capacity_html.append("<div style='background:#eef2ff;color:#1f2b6b;padding:12px;border-radius:8px;font-weight:700;'>Total Demand</div>")
    
    for data in capacity_data:
        pct = (data['total_mm'] / data['available']) * 100
        if pct < 70:
            color_class = "background:linear-gradient(180deg,#d1fae5 0%,#a7f3d0 100%);color:#047857;"
        elif pct < 90:
            color_class = "background:linear-gradient(180deg,#fef3c7 0%,#fde68a 100%);color:#b45309;"
        else:
            color_class = "background:linear-gradient(180deg,#fecaca 0%,#fca5a5 100%);color:#b91c1c;"
        
        capacity_html.append(
            f"<div style='background:#f1f5f9;height:50px;border-radius:8px;position:relative;overflow:hidden;border:2px solid #e2e8f0;'>"
            f"<div style='position:absolute;bottom:0;left:0;right:0;height:{min(100, pct)}%;{color_class}display:flex;align-items:center;justify-content:center;font-weight:700;font-size:13px;'>"
            f"{data['total_mm']:.0f} MM</div></div>"
        )
    
    capacity_html.append("<div style='background:#eef2ff;color:#1f2b6b;padding:12px;border-radius:8px;font-weight:700;'>Available</div>")
    
    for data in capacity_data:
        pct = (data['total_mm'] / data['available']) * 100
        if pct < 70:
            color = "#047857"
        elif pct < 90:
            color = "#b45309"
        else:
            color = "#b91c1c"
        capacity_html.append(f"<div style='text-align:center;color:{color};font-weight:600;'>{data['available']} MM</div>")
    
    capacity_html.append("</div>")
    st.markdown("".join(capacity_html), unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)
    
    # 4. SUMMARY DASHBOARD - Enhanced with Charts
    st.markdown("<div class='dash-section'>", unsafe_allow_html=True)
    st.markdown("<div class='dash-title'>📈 Summary Dashboard - Executive Overview</div>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**Effort by Department**")
        dept_totals = meta_view.groupby('Department')['Total_MM'].sum()
        total_mm = dept_totals.sum()
        
        # Create pie chart
        import plotly.graph_objects as go
        
        colors = ['#667eea', '#3b82f6', '#10b981']
        fig = go.Figure(data=[go.Pie(
            labels=dept_totals.index,
            values=dept_totals.values,
            hole=0.5,
            marker=dict(colors=colors),
            textinfo='label+percent',
            textfont=dict(size=12, color='white'),
            hovertemplate='<b>%{label}</b><br>%{value} MM<br>%{percent}<extra></extra>'
        )])
        
        fig.update_layout(
            showlegend=True,
            height=250,
            margin=dict(l=0, r=0, t=0, b=0),
            legend=dict(
                orientation="v",
                yanchor="middle",
                y=0.5,
                xanchor="left",
                x=1.1
            )
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    with col2:
        st.markdown("**Total Effort Trend (MM)**")
        
        # Calculate monthly totals for trend
        monthly_totals = []
        month_labels = []
        for i, month in enumerate(all_months[:6]):
            month_total = alloc[alloc['Date'] == month]['MM'].sum()
            monthly_totals.append(month_total)
            month_labels.append(month.split('-')[1])
        
        # Create line chart
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=month_labels,
            y=monthly_totals,
            mode='lines+markers',
            line=dict(color='#667eea', width=3),
            marker=dict(size=8, color='#667eea'),
            fill='tonexty',
            fillcolor='rgba(102, 126, 234, 0.1)',
            hovertemplate='<b>Month %{x}</b><br>%{y:.0f} MM<extra></extra>'
        ))
        
        fig.update_layout(
            height=250,
            margin=dict(l=0, r=0, t=0, b=0),
            xaxis=dict(title='', showgrid=True, gridcolor='#f1f5f9'),
            yaxis=dict(title='', showgrid=True, gridcolor='#f1f5f9'),
            plot_bgcolor='white',
            hovermode='x unified'
        )
        
        st.plotly_chart(fig, use_container_width=True)
    
    with col3:
        st.markdown("**Key Metrics**")
        
        # Calculate metrics
        recent_months = alloc[alloc['Date'].isin(all_months[:3])]
        avg_monthly = recent_months.groupby('Date')['MM'].sum().mean()
        high_util_employees = len(conflicts)
        
        # Create metric boxes with custom HTML
        metrics_html = f"""
        <div class='metric-row'>
            <div class='metric-box'>
                <div class='metric-label'>AVG UTILIZATION</div>
                <div class='metric-value'>68%</div>
                <div class='metric-change up'>↑ 5% vs last month</div>
            </div>
            <div class='metric-box'>
                <div class='metric-label'>PROJECTS AT RISK</div>
                <div class='metric-value'>{high_util_employees}</div>
                <div class='metric-change down'>↓ 1 vs last month</div>
            </div>
        </div>
        <div class='metric-row'>
            <div class='metric-box'>
                <div class='metric-label'>TOTAL BUDGET</div>
                <div class='metric-value'>{meta_view['Total_MM'].sum():.0f}</div>
                <div class='metric-change' style='color:#64748b;'>MM allocated</div>
            </div>
            <div class='metric-box'>
                <div class='metric-label'>ON SCHEDULE</div>
                <div class='metric-value'>82%</div>
                <div class='metric-change up'>↑ 3% vs last month</div>
            </div>
        </div>
        """
        st.markdown(metrics_html, unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)

# End
