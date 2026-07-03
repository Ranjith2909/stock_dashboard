import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import requests
from io import StringIO, BytesIO
import warnings
warnings.filterwarnings('ignore')

# ══════════════════════════════════════════════
# PAGE CONFIG
# ══════════════════════════════════════════════
st.set_page_config(
    page_title="🏭 Bangalore Warehouse Dashboard",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ══════════════════════════════════════════════
# CSS STYLING
# ══════════════════════════════════════════════
st.markdown("""
<style>
    .main { background-color: #f8f9fa; }
    
    .top-header {
        background: linear-gradient(135deg, #1e3c72, #2a5298);
        padding: 20px 30px;
        border-radius: 12px;
        margin-bottom: 20px;
        color: white;
    }
    .top-header h1 { 
        color: white; 
        margin: 0; 
        font-size: 2rem;
    }
    .top-header p { 
        color: #ccd6f6; 
        margin: 5px 0 0 0; 
    }
    
    .kpi-card {
        background: white;
        padding: 20px;
        border-radius: 12px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        text-align: center;
        border-top: 4px solid #2a5298;
        height: 120px;
    }
    .kpi-critical { border-top-color: #dc3545 !important; }
    .kpi-warning  { border-top-color: #fd7e14 !important; }
    .kpi-ok       { border-top-color: #28a745 !important; }
    
    .kpi-number {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1e3c72;
        line-height: 1;
    }
    .kpi-label {
        font-size: 0.85rem;
        color: #6c757d;
        margin-top: 8px;
    }
    
    .status-badge {
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.78rem;
        font-weight: 600;
    }
    .badge-critical { background:#ffe0e0; color:#dc3545; }
    .badge-warning  { background:#fff3cd; color:#856404; }
    .badge-low      { background:#fff9c4; color:#7a6200; }
    .badge-ok       { background:#d4edda; color:#155724; }
    
    .section-title {
        font-size: 1.2rem;
        font-weight: 700;
        color: #1e3c72;
        border-left: 4px solid #2a5298;
        padding-left: 10px;
        margin: 20px 0 15px 0;
    }
    
    .alert-box {
        background: #fff5f5;
        border: 1px solid #ffcccc;
        border-radius: 8px;
        padding: 12px 16px;
        margin: 5px 0;
    }
    
    .live-badge {
        background: #28a745;
        color: white;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 0.75rem;
        animation: pulse 2s infinite;
    }
    @keyframes pulse {
        0%   { opacity: 1; }
        50%  { opacity: 0.5; }
        100% { opacity: 1; }
    }
    
    .source-card {
        background: white;
        border: 2px dashed #2a5298;
        border-radius: 12px;
        padding: 20px;
        text-align: center;
        cursor: pointer;
    }
    .source-card:hover { background: #f0f4ff; }
    
    div[data-testid="stMetric"] {
        background: white;
        padding: 15px;
        border-radius: 10px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 5px;
        background: white;
        padding: 5px;
        border-radius: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        border-radius: 8px;
        padding: 8px 16px;
    }
    
    footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════
# GOOGLE SHEETS CONFIG
# ══════════════════════════════════════════════

# 🔑 PUT YOUR GOOGLE SHEET ID HERE
# Get it from your sheet URL:
# https://docs.google.com/spreadsheets/d/[THIS_PART]/edit
GOOGLE_SHEET_ID = st.secrets.get(
    "GOOGLE_SHEET_ID", 
    ""  # Empty = will ask user
)

SHEET_NAME = st.secrets.get("SHEET_NAME", "Sheet1")

def get_google_sheet_url(sheet_id, sheet_name="Sheet1"):
    """Build CSV export URL from Google Sheet."""
    return (
        f"https://docs.google.com/spreadsheets/d/{sheet_id}"
        f"/gviz/tq?tqx=out:csv&sheet={sheet_name}"
    )

@st.cache_data(ttl=300)  # Cache 5 minutes
def load_from_google_sheets(sheet_id, sheet_name="Sheet1"):
    """Load data directly from Google Sheets."""
    try:
        url = get_google_sheet_url(sheet_id, sheet_name)
        response = requests.get(url, timeout=15)
        response.raise_for_status()
        df = pd.read_csv(StringIO(response.text))
        return df, None, url
    except Exception as e:
        return None, str(e), None

@st.cache_data(ttl=300)
def load_from_upload(file_bytes, file_name):
    """Load from uploaded file - cached so no re-upload needed."""
    try:
        if file_name.endswith('.csv'):
            df = pd.read_csv(BytesIO(file_bytes))
        else:
            df = pd.read_excel(BytesIO(file_bytes), engine='openpyxl')
        return df, None
    except Exception as e:
        return None, str(e)

# ══════════════════════════════════════════════
# DATA PROCESSING
# ══════════════════════════════════════════════

def clean_val(v):
    """Clean any value to float."""
    try:
        if pd.isna(v): return 0.0
        s = str(v).strip()
        if s in ('#NAME?','#VALUE!','#REF!','#DIV/0!','#N/A','','-','N/A'):
            return 0.0
        s = s.replace(',','').replace(' ','')
        return float(s)
    except:
        return 0.0

def find_column(df, candidates):
    """Find column by multiple possible names."""
    cols_upper = {c.upper().strip().replace(' ','_'): c for c in df.columns}
    for cand in candidates:
        key = cand.upper().strip().replace(' ','_')
        if key in cols_upper:
            return cols_upper[key]
        # Partial match
        for col_key, col_name in cols_upper.items():
            if key in col_key or col_key in key:
                return col_name
    return None

def process_df(df):
    """Process and clean the dataframe."""
    result = pd.DataFrame()

    # Find columns flexibly
    col_month = find_column(df, ['MONTH','Month','MNTH'])
    col_code  = find_column(df, ['CODE','Code','ITEM_CODE','ItemCode'])
    col_item  = find_column(df, ['ITEM','Item','ITEM_NAME','ItemName','Description'])
    col_lead  = find_column(df, ['AVG_LEAD_TIME','Avg Lead Time','LEAD_TIME','LeadTime'])
    col_safe  = find_column(df, ['SAFETY_STOCK','Safety Stock','SAFETY'])
    col_open  = find_column(df, ['OPENING_STOCK','Opening Stock','OPENING','Open'])
    col_req   = find_column(df, ['MONTH_REQUIREMENT','Month Requirement','REQUIREMENT','REQ'])
    col_recv  = find_column(df, ['RECEIVED','Received','RECV'])
    col_bal   = find_column(df, ['BAL_REQUIRED','Bal required','BALANCE','BAL'])
    col_issue = find_column(df, ['ISSUED','Issued','ISSUE'])
    col_close = find_column(df, ['CLOSING_STOCK','Closing Stock','CLOSING','Close'])

    # Build result
    result['MONTH']   = df[col_month].astype(str).str.strip() if col_month else 'Unknown'
    result['CODE']    = df[col_code].astype(str).str.strip()  if col_code  else 'N/A'
    result['ITEM']    = df[col_item].astype(str).str.strip()  if col_item  else 'Unknown'

    num_map = {
        'AVG_LEAD_TIME': col_lead,
        'SAFETY_STOCK':  col_safe,
        'OPENING_STOCK': col_open,
        'MONTH_REQ':     col_req,
        'RECEIVED':      col_recv,
        'BAL_REQ':       col_bal,
        'ISSUED':        col_issue,
        'CLOSING_STOCK': col_close,
    }

    for field, col in num_map.items():
        if col:
            result[field] = df[col].apply(clean_val)
        else:
            result[field] = 0.0

    # Recalculate Closing Stock if zero/missing
    mask = result['CLOSING_STOCK'] == 0
    result.loc[mask, 'CLOSING_STOCK'] = (
        result.loc[mask, 'OPENING_STOCK'] +
        result.loc[mask, 'RECEIVED'] -
        result.loc[mask, 'ISSUED']
    )

    # Recalculate Balance
    result['BAL_REQ'] = (result['MONTH_REQ'] - result['RECEIVED']).clip(lower=0)

    # Fill #NAME? Lead Times by prefix
    def guess_lead(code):
        c = str(code).upper()
        if 'GB' in c: return 45
        if 'GD' in c or 'GR' in c: return 30
        if 'RM-19' in c: return 15
        if 'RM' in c:    return 12
        if 'SFG' in c:   return 7
        return 14

    mask_lead = result['AVG_LEAD_TIME'] == 0
    result.loc[mask_lead, 'AVG_LEAD_TIME'] = (
        result.loc[mask_lead, 'CODE'].apply(guess_lead)
    )

    # Category
    def get_cat(code):
        c = str(code).upper()
        if 'RM-19' in c or 'PERF' in c: return 'Perfumes'
        if 'RM-06' in c: return 'Colours'
        if 'SFG-06' in c: return 'Diluted Perfumes'
        if 'RM-25' in c or 'RM-20' in c: return 'Raw Materials'
        if c.startswith('GD') or c.startswith('GR') or \
           c.startswith('GU') or c.startswith('GB'): return 'Gums'
        if c.startswith('DB') or c.startswith('BKRL') or \
           c.startswith('DBC'): return 'Damarbattu'
        if c.startswith('OL') or c.startswith('DEP'): return 'Oils'
        if c.startswith('SFG'): return 'SFG'
        if c.startswith('CPN'): return 'Compound'
        return 'Other'

    result['CATEGORY'] = result['CODE'].apply(get_cat)

    # Status
    def get_status(row):
        cs = row['CLOSING_STOCK']
        ss = row['SAFETY_STOCK']
        mr = row['MONTH_REQ']
        if cs < 0:                         return '🔴 CRITICAL'
        if ss > 0 and cs < ss:             return '🟠 BELOW SAFETY'
        if mr > 0 and cs < (mr * 0.25):   return '🟡 LOW'
        if mr > 0 and cs < (mr * 0.5):    return '🟡 MEDIUM'
        return '🟢 OK'

    result['STATUS'] = result.apply(get_status, axis=1)

    # Month ordering
    mo = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
          'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
    result['MONTH_NUM'] = result['MONTH'].map(mo).fillna(0).astype(int)

    # Remove empty rows
    result = result[result['CODE'] != 'N/A']
    result = result[result['ITEM'] != 'Unknown']
    result = result.dropna(subset=['CODE'])
    result = result.reset_index(drop=True)

    return result

# ══════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════

with st.sidebar:
    st.markdown("## 🏭 Warehouse IMS")
    st.markdown("---")

    # Data Source Selection
    st.markdown("### 📡 Data Source")
    data_source = st.radio(
        "Choose Source",
        ["📊 Google Sheets (Live)", "📁 Upload File"],
        index=0
    )

    df_raw  = None
    source_info = ""

    # ── Google Sheets ──────────────────────────
    if "Google Sheets" in data_source:
        st.markdown("#### Google Sheets Settings")

        sheet_id = st.text_input(
            "Sheet ID",
            value=GOOGLE_SHEET_ID,
            placeholder="Paste your Google Sheet ID here",
            help="From URL: docs.google.com/spreadsheets/d/[ID]/edit"
        )
        sheet_name = st.text_input(
            "Sheet Name",
            value=SHEET_NAME,
            placeholder="Sheet1"
        )

        if sheet_id:
            with st.spinner("🔄 Loading from Google Sheets..."):
                df_raw, err, url = load_from_google_sheets(
                    sheet_id, sheet_name
                )
            if err:
                st.error(f"❌ {err}")
                st.markdown("""
                **Check:**
                - Sheet is shared (Anyone with link)
                - Sheet ID is correct
                - Sheet name matches
                """)
            else:
                st.success(f"✅ Connected! {len(df_raw)} rows")
                source_info = f"🟢 Google Sheets • {len(df_raw)} rows"

            # How to find Sheet ID
            with st.expander("❓ How to get Sheet ID"):
                st.markdown("""
                1. Open your Google Sheet
                2. Look at the URL:
                ```
                .../spreadsheets/d/[SHEET_ID]/edit
                ```
                3. Copy the ID and paste above
                
                **Make sheet public:**
                - Click Share (top right)
                - Change to "Anyone with link"
                - Set to "Viewer"
                """)
        else:
            st.info("👆 Enter your Google Sheet ID above")

    # ── File Upload ────────────────────────────
    else:
        st.markdown("#### Upload File")

        # Session state to keep file
        if 'uploaded_bytes' not in st.session_state:
            st.session_state.uploaded_bytes = None
            st.session_state.uploaded_name  = None

        uploaded = st.file_uploader(
            "Excel / CSV",
            type=['xlsx','xls','csv','xlsm'],
            help="File stays loaded until you refresh"
        )

        if uploaded:
            # Save to session state
            st.session_state.uploaded_bytes = uploaded.read()
            st.session_state.uploaded_name  = uploaded.name
            st.success(f"✅ {uploaded.name} saved!")

        # Use session state file
        if st.session_state.uploaded_bytes:
            df_raw, err = load_from_upload(
                st.session_state.uploaded_bytes,
                st.session_state.uploaded_name
            )
            if err:
                st.error(f"❌ {err}")
            else:
                fn = st.session_state.uploaded_name
                source_info = f"📁 {fn} • {len(df_raw)} rows"
                st.info(f"📁 Using: {fn}")

            if st.button("🗑️ Clear File"):
                st.session_state.uploaded_bytes = None
                st.session_state.uploaded_name  = None
                st.rerun()

    st.markdown("---")

    # ── Filters (shown after data loaded) ─────
    if df_raw is not None:
        st.markdown("### 🔍 Filters")
        try:
            df_proc = process_df(df_raw)

            months = ['All'] + sorted(
                [m for m in df_proc['MONTH'].unique()
                 if m not in ('Unknown','nan','')],
                key=lambda x: {'Jan':1,'Feb':2,'Mar':3,'Apr':4,
                               'May':5,'Jun':6,'Jul':7,'Aug':8,
                               'Sep':9,'Oct':10,'Nov':11,'Dec':12}.get(x,99)
            )
            cats    = ['All'] + sorted(df_proc['CATEGORY'].unique())
            statuses= ['All','🔴 CRITICAL','🟠 BELOW SAFETY',
                       '🟡 LOW','🟡 MEDIUM','🟢 OK']

            sel_month  = st.selectbox("📅 Month",    months)
            sel_cat    = st.selectbox("📦 Category", cats)
            sel_status = st.selectbox("🚦 Status",   statuses)
            min_close  = st.slider(
                "Min Closing Stock", 0,
                int(df_proc['CLOSING_STOCK'].max()+1), 0
            )
        except Exception as e:
            st.error(f"Filter error: {e}")
            df_proc = None
            sel_month = sel_cat = sel_status = 'All'
            min_close = 0
    else:
        df_proc = None
        sel_month = sel_cat = sel_status = 'All'
        min_close = 0

    st.markdown("---")

    # ── Auto Refresh ───────────────────────────
    st.markdown("### 🔄 Auto Refresh")
    auto_ref  = st.checkbox("Enable Auto-Refresh")
    ref_secs  = st.slider("Interval (seconds)", 30, 300, 60)

    if auto_ref:
        import time
        st.info(f"🔄 Refreshing every {ref_secs}s")
        time.sleep(ref_secs)
        st.cache_data.clear()
        st.rerun()

    st.markdown("---")
    st.caption(f"🕐 {datetime.now().strftime('%d %b %Y %H:%M:%S')}")

# ══════════════════════════════════════════════
# MAIN CONTENT
# ══════════════════════════════════════════════

# ── Header ────────────────────────────────────
st.markdown(f"""
<div class="top-header">
    <h1>🏭 Bangalore Warehouse Dashboard</h1>
    <p>
        {source_info if source_info else "No data loaded yet"} &nbsp;|&nbsp;
        <span class="live-badge">● LIVE</span> &nbsp;|&nbsp;
        Updated: {datetime.now().strftime('%d %b %Y %H:%M')}
    </p>
</div>
""", unsafe_allow_html=True)

# ── No Data State ─────────────────────────────
if df_proc is None:
    st.markdown("---")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("""
        <div class="source-card">
            <h2>📊</h2>
            <h3>Google Sheets (Recommended)</h3>
            <p>Live data • No upload needed • Auto-refresh</p>
            <ol style="text-align:left">
                <li>Upload Excel to Google Sheets</li>
                <li>Share → Anyone with link</li>
                <li>Paste Sheet ID in sidebar</li>
                <li>Dashboard updates automatically!</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown("""
        <div class="source-card">
            <h2>📁</h2>
            <h3>Upload File</h3>
            <p>Upload once → stays loaded in session</p>
            <ol style="text-align:left">
                <li>Select "Upload File" in sidebar</li>
                <li>Upload your Excel/CSV</li>
                <li>File is saved in session</li>
                <li>No need to re-upload!</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)

    st.stop()

# ── Apply Filters ─────────────────────────────
fdf = df_proc.copy()
if sel_month  != 'All': fdf = fdf[fdf['MONTH']    == sel_month]
if sel_cat    != 'All': fdf = fdf[fdf['CATEGORY'] == sel_cat]
if sel_status != 'All': fdf = fdf[fdf['STATUS']   == sel_status]
if min_close  > 0:      fdf = fdf[fdf['CLOSING_STOCK'] >= min_close]

if len(fdf) == 0:
    st.warning("⚠️ No data matches your filters. Adjust the sidebar filters.")
    st.stop()

# ── KPI Row ───────────────────────────────────
total     = len(fdf)
critical  = len(fdf[fdf['STATUS'] == '🔴 CRITICAL'])
below_saf = len(fdf[fdf['STATUS'] == '🟠 BELOW SAFETY'])
low       = len(fdf[fdf['STATUS'].isin(['🟡 LOW','🟡 MEDIUM'])])
ok        = len(fdf[fdf['STATUS'] == '🟢 OK'])
total_rcv = fdf['RECEIVED'].sum()
total_iss = fdf['ISSUED'].sum()
total_req = fdf['MONTH_REQ'].sum()

k = st.columns(7)
k[0].metric("📦 Total Items",    f"{total:,}")
k[1].metric("🔴 Critical",       f"{critical:,}")
k[2].metric("🟠 Below Safety",   f"{below_saf:,}")
k[3].metric("🟡 Low/Medium",     f"{low:,}")
k[4].metric("🟢 OK",             f"{ok:,}")
k[5].metric("📥 Total Received", f"{total_rcv:,.0f}")
k[6].metric("📤 Total Issued",   f"{total_iss:,.0f}")

st.markdown("---")

# ── TABS ──────────────────────────────────────
tab1,tab2,tab3,tab4,tab5 = st.tabs([
    "📊 Overview",
    "🚨 Alerts",
    "📈 Trends",
    "🔍 Item Detail",
    "📋 Full Data"
])

# ════════════════════════════════
# TAB 1 - OVERVIEW
# ════════════════════════════════
with tab1:
    r1c1, r1c2 = st.columns(2)

    with r1c1:
        st.markdown('<div class="section-title">Stock Status Distribution</div>',
                    unsafe_allow_html=True)
        sc = fdf['STATUS'].value_counts().reset_index()
        sc.columns = ['Status','Count']
        colors = {
            '🔴 CRITICAL':     '#dc3545',
            '🟠 BELOW SAFETY': '#fd7e14',
            '🟡 LOW':          '#ffc107',
            '🟡 MEDIUM':       '#ffe066',
            '🟢 OK':           '#28a745',
        }
        fig1 = px.pie(sc, names='Status', values='Count',
                      color='Status', color_discrete_map=colors,
                      hole=0.45)
        fig1.update_layout(margin=dict(t=10,b=10,l=0,r=0),
                           showlegend=True,
                           legend=dict(orientation='h'))
        st.plotly_chart(fig1, use_container_width=True)

    with r1c2:
        st.markdown('<div class="section-title">Items by Category</div>',
                    unsafe_allow_html=True)
        cat_df = fdf.groupby('CATEGORY').size().reset_index(name='Count')
        cat_df = cat_df.sort_values('Count', ascending=True)
        fig2 = px.bar(cat_df, x='Count', y='CATEGORY',
                      orientation='h', color='Count',
                      color_continuous_scale='Blues', text='Count')
        fig2.update_traces(textposition='outside')
        fig2.update_layout(margin=dict(t=10,b=10,l=0,r=0),
                           coloraxis_showscale=False)
        st.plotly_chart(fig2, use_container_width=True)

    r2c1, r2c2 = st.columns(2)

    with r2c1:
        st.markdown('<div class="section-title">Received vs Requirement</div>',
                    unsafe_allow_html=True)
        cat_stock = fdf.groupby('CATEGORY').agg(
            Received=('RECEIVED','sum'),
            Requirement=('MONTH_REQ','sum'),
            Issued=('ISSUED','sum')
        ).reset_index()
        fig3 = go.Figure()
        fig3.add_bar(name='Received',    x=cat_stock['CATEGORY'],
                     y=cat_stock['Received'],    marker_color='#28a745')
        fig3.add_bar(name='Issued',      x=cat_stock['CATEGORY'],
                     y=cat_stock['Issued'],      marker_color='#2196F3')
        fig3.add_bar(name='Requirement', x=cat_stock['CATEGORY'],
                     y=cat_stock['Requirement'], marker_color='#dc3545')
        fig3.update_layout(
            barmode='group',
            margin=dict(t=10,b=50,l=0,r=0),
            xaxis_tickangle=-30,
            legend=dict(orientation='h', y=-0.35)
        )
        st.plotly_chart(fig3, use_container_width=True)

    with r2c2:
        st.markdown('<div class="section-title">Avg Lead Time by Category</div>',
                    unsafe_allow_html=True)
        lt = fdf[fdf['AVG_LEAD_TIME']>0].groupby('CATEGORY').agg(
            Lead=('AVG_LEAD_TIME','mean')
        ).reset_index().sort_values('Lead')
        if len(lt):
            fig4 = px.bar(lt, x='Lead', y='CATEGORY',
                          orientation='h',
                          color='Lead',
                          color_continuous_scale='RdYlGn_r',
                          text=lt['Lead'].round(1))
            fig4.update_traces(textposition='outside')
            fig4.update_layout(margin=dict(t=10,b=10,l=0,r=0),
                               coloraxis_showscale=False)
            st.plotly_chart(fig4, use_container_width=True)

# ════════════════════════════════
# TAB 2 - ALERTS
# ════════════════════════════════
with tab2:
    st.markdown('<div class="section-title">🚨 Items Requiring Immediate Action</div>',
                unsafe_allow_html=True)

    alert_df = fdf[fdf['STATUS'].isin(
        ['🔴 CRITICAL','🟠 BELOW SAFETY','🟡 LOW']
    )].copy()

    if len(alert_df):
        a1, a2, a3 = st.columns(3)
        crit_df  = alert_df[alert_df['STATUS']=='🔴 CRITICAL']
        below_df = alert_df[alert_df['STATUS']=='🟠 BELOW SAFETY']
        low_df   = alert_df[alert_df['STATUS']=='🟡 LOW']

        a1.error(  f"🔴 **{len(crit_df)}** Critical items")
        a2.warning(f"🟠 **{len(below_df)}** Below safety stock")
        a3.info(   f"🟡 **{len(low_df)}** Low stock items")

        st.markdown("---")

        show_cols = ['CODE','ITEM','MONTH','CATEGORY','STATUS',
                     'SAFETY_STOCK','OPENING_STOCK','RECEIVED',
                     'ISSUED','CLOSING_STOCK','MONTH_REQ','BAL_REQ',
                     'AVG_LEAD_TIME']
        disp = alert_df[show_cols].copy()
        disp.columns = ['Code','Item','Month','Category','Status',
                        'Safety Stock','Opening','Received','Issued',
                        'Closing','Requirement','Balance Req','Lead Days']

        def highlight(row):
            if '🔴' in str(row.get('Status','')):
                return ['background-color:#ffe0e0']*len(row)
            if '🟠' in str(row.get('Status','')):
                return ['background-color:#fff3cd']*len(row)
            if '🟡' in str(row.get('Status','')):
                return ['background-color:#fffde7']*len(row)
            return ['']*len(row)

        styled = disp.style.apply(highlight, axis=1)\
                           .format({
                               'Safety Stock':  '{:,.0f}',
                               'Opening':       '{:,.0f}',
                               'Received':      '{:,.0f}',
                               'Issued':        '{:,.0f}',
                               'Closing':       '{:,.0f}',
                               'Requirement':   '{:,.0f}',
                               'Balance Req':   '{:,.0f}',
                               'Lead Days':     '{:.0f}',
                           })
        st.dataframe(styled, use_container_width=True, height=400)

        # Download
        csv = disp.to_csv(index=False).encode('utf-8')
        st.download_button(
            "⬇️ Download Alert Report",
            csv, "alert_report.csv", "text/csv"
        )
    else:
        st.success("✅ No critical items! All stock levels are healthy.")

# ════════════════════════════════
# TAB 3 - TRENDS
# ════════════════════════════════
with tab3:
    st.markdown('<div class="section-title">📈 Monthly Trends</div>',
                unsafe_allow_html=True)

    monthly = df_proc.groupby(['MONTH','MONTH_NUM']).agg(
        Received=('RECEIVED','sum'),
        Issued=('ISSUED','sum'),
        Requirement=('MONTH_REQ','sum'),
        Closing=('CLOSING_STOCK','sum')
    ).reset_index().sort_values('MONTH_NUM')

    if len(monthly):
        fig_t = go.Figure()
        fig_t.add_bar(name='📥 Received',
                      x=monthly['MONTH'], y=monthly['Received'],
                      marker_color='#28a745')
        fig_t.add_bar(name='📤 Issued',
                      x=monthly['MONTH'], y=monthly['Issued'],
                      marker_color='#2196F3')
        fig_t.add_scatter(name='📋 Requirement',
                          x=monthly['MONTH'], y=monthly['Requirement'],
                          mode='lines+markers',
                          line=dict(color='red', width=2, dash='dot'))
        fig_t.add_scatter(name='🏦 Closing Stock',
                          x=monthly['MONTH'], y=monthly['Closing'],
                          mode='lines+markers',
                          line=dict(color='purple', width=2))
        fig_t.update_layout(
            barmode='group', height=400,
            margin=dict(t=20,b=20,l=0,r=0),
            legend=dict(orientation='h', y=-0.2)
        )
        st.plotly_chart(fig_t, use_container_width=True)

    # Category trends
    st.markdown('<div class="section-title">Category-wise Monthly Stock</div>',
                unsafe_allow_html=True)
    cat_monthly = df_proc.groupby(['MONTH','MONTH_NUM','CATEGORY']).agg(
        Closing=('CLOSING_STOCK','sum')
    ).reset_index().sort_values('MONTH_NUM')

    if len(cat_monthly):
        fig_cm = px.line(cat_monthly, x='MONTH', y='Closing',
                         color='CATEGORY', markers=True,
                         title="Closing Stock by Category per Month")
        fig_cm.update_layout(margin=dict(t=40,b=20,l=0,r=0),
                             height=350)
        st.plotly_chart(fig_cm, use_container_width=True)

# ════════════════════════════════
# TAB 4 - ITEM DETAIL
# ════════════════════════════════
with tab4:
    st.markdown('<div class="section-title">🔍 Item Deep-Dive</div>',
                unsafe_allow_html=True)

    col_s1, col_s2 = st.columns(2)
    with col_s1:
        item_list = sorted(df_proc['ITEM'].unique().tolist())
        sel_item  = st.selectbox("Select Item", item_list)
    with col_s2:
        item_data = df_proc[df_proc['ITEM'] == sel_item]\
                    .sort_values('MONTH_NUM')
        if len(item_data):
            r = item_data.iloc[-1]
            st.markdown(f"""
            **Code:** `{r['CODE']}`  
            **Category:** {r['CATEGORY']}  
            **Status:** {r['STATUS']}  
            **Lead Time:** {r['AVG_LEAD_TIME']:.0f} days
            """)

    if len(item_data):
        m1,m2,m3,m4 = st.columns(4)
        r = item_data.iloc[-1]
        m1.metric("Opening Stock",  f"{r['OPENING_STOCK']:,.0f}")
        m2.metric("Received",       f"{r['RECEIVED']:,.0f}")
        m3.metric("Issued",         f"{r['ISSUED']:,.0f}")
        m4.metric("Closing Stock",  f"{r['CLOSING_STOCK']:,.0f}",
                  delta=f"{r['CLOSING_STOCK']-r['OPENING_STOCK']:,.0f}")

        # Gauge
        max_v = max(r['MONTH_REQ'], r['CLOSING_STOCK'], r['SAFETY_STOCK'], 1)
        fig_g = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=float(r['CLOSING_STOCK']),
            delta={'reference': float(r['MONTH_REQ'])},
            title={'text': "Closing Stock vs Requirement"},
            gauge={
                'axis':  {'range': [0, max_v * 1.3]},
                'bar':   {'color': '#2196F3'},
                'steps': [
                    {'range': [0, float(r['SAFETY_STOCK'])],
                     'color': '#ffcccc'},
                    {'range': [float(r['SAFETY_STOCK']),
                               float(r['MONTH_REQ'])],
                     'color': '#fff9c4'},
                    {'range': [float(r['MONTH_REQ']),
                               max_v * 1.3], 'color': '#ccffcc'},
                ],
            }
        ))
        fig_g.update_layout(height=280, margin=dict(t=40,b=0,l=40,r=40))
        st.plotly_chart(fig_g, use_container_width=True)

        # Monthly trend
        if len(item_data) > 1:
            fig_it = px.line(
                item_data, x='MONTH',
                y=['OPENING_STOCK','RECEIVED','ISSUED','CLOSING_STOCK'],
                markers=True, title=f"Monthly Trend: {sel_item}"
            )
            fig_it.update_layout(
                height=300, margin=dict(t=50,b=20,l=0,r=0),
                legend=dict(orientation='h', y=-0.3)
            )
            st.plotly_chart(fig_it, use_container_width=True)

        # Data table for this item
        st.dataframe(
            item_data[['MONTH','OPENING_STOCK','RECEIVED',
                        'ISSUED','CLOSING_STOCK','MONTH_REQ',
                        'BAL_REQ','AVG_LEAD_TIME','STATUS']],
            use_container_width=True
        )

# ════════════════════════════════
# TAB 5 - FULL DATA
# ════════════════════════════════
with tab5:
    st.markdown('<div class="section-title">📋 Complete Data Table</div>',
                unsafe_allow_html=True)

    # Search
    search = st.text_input("🔎 Search (Item name or Code)",
                           placeholder="Type to search...")
    display = fdf.copy()
    if search:
        mask = (
            display['ITEM'].str.contains(search, case=False, na=False) |
            display['CODE'].str.contains(search, case=False, na=False)
        )
        display = display[mask]

    st.markdown(f"Showing **{len(display)}** records")
    st.dataframe(display, use_container_width=True, height=450)

    # Downloads
    d1, d2 = st.columns(2)
    with d1:
        csv_all = display.to_csv(index=False).encode('utf-8')
        st.download_button(
            "⬇️ Download Filtered Data (CSV)",
            csv_all, "warehouse_data.csv", "text/csv"
        )
    with d2:
        csv_alert = fdf[fdf['STATUS'] != '🟢 OK']\
                    .to_csv(index=False).encode('utf-8')
        st.download_button(
            "⬇️ Download Alert Items (CSV)",
            csv_alert, "alerts.csv", "text/csv"
        )

# ── Footer ─────────────────────────────────────
st.markdown("---")
st.markdown(
    f"<center>🏭 Bangalore Warehouse IMS • "
    f"Powered by Streamlit • "
    f"{datetime.now().year}</center>",
    unsafe_allow_html=True
)
