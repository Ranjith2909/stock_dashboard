import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')

# ── Page Config ──────────────────────────────────────────────
st.set_page_config(
    page_title="Warehouse Stock Dashboard",
    page_icon="🏭",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background-color: #f0f2f6; }
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        text-align: center;
    }
    .critical { border-left: 5px solid #ff4444; }
    .warning  { border-left: 5px solid #ffaa00; }
    .good     { border-left: 5px solid #00cc44; }
    .header-title {
        font-size: 2rem;
        font-weight: bold;
        color: #1f2937;
    }
    .stMetric { background: white; padding: 10px; border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ── Helper: Clean numeric values ──────────────────────────────
def clean_numeric(val):
    """Convert any value to float, return 0 for errors."""
    try:
        if pd.isna(val):
            return 0.0
        if isinstance(val, str):
            val = val.replace(',', '').replace('#NAME?', '0').strip()
            if val in ('', '-', 'N/A', '#VALUE!', '#REF!', '#NAME?'):
                return 0.0
        return float(val)
    except:
        return 0.0

# ── Data Loader ───────────────────────────────────────────────
@st.cache_data
def load_data(uploaded_file):
    """Load and clean uploaded Excel/CSV file."""
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            # Try reading with header detection
            xl = pd.ExcelFile(uploaded_file)
            sheet = xl.sheet_names[0]
            df = pd.read_excel(uploaded_file, sheet_name=sheet, header=0)

        return df, None
    except Exception as e:
        return None, str(e)

# ── Build Clean DataFrame from raw upload ─────────────────────
def process_dataframe(df):
    """
    Expects columns (flexible naming):
    S.No | Month | CODE | ITEM | Avg Lead Time |
    Safety Stock | Opening Stock | Month Requirement |
    Received | Bal required | Issued | Closing Stock
    """
    # Normalise column names
    df.columns = [str(c).strip().upper().replace(' ', '_') for c in df.columns]

    col_map = {
        'S.NO': 'SNO', 'S_NO': 'SNO', 'SNO': 'SNO',
        'MONTH': 'MONTH',
        'CODE': 'CODE',
        'ITEM': 'ITEM',
        'AVG_LEAD_TIME': 'AVG_LEAD_TIME', 'AVG_LEAD': 'AVG_LEAD_TIME',
        'SAFETY_STOCK': 'SAFETY_STOCK',
        'OPENING_STOCK': 'OPENING_STOCK',
        'MONTH_REQUIREMENT': 'MONTH_REQ', 'MONTH_REQUIREMEN': 'MONTH_REQ',
        'RECEIVED': 'RECEIVED',
        'BAL_REQUIRED': 'BAL_REQ', 'BAL_REQUIRED': 'BAL_REQ',
        'ISSUED': 'ISSUED',
        'CLOSING_STOCK': 'CLOSING_STOCK',
    }

    df.rename(columns=col_map, inplace=True)

    # Required columns with defaults
    needed = {
        'SNO': range(1, len(df)+1),
        'MONTH': 'Unknown',
        'CODE': 'N/A',
        'ITEM': 'Unknown Item',
        'AVG_LEAD_TIME': 0,
        'SAFETY_STOCK': 0,
        'OPENING_STOCK': 0,
        'MONTH_REQ': 0,
        'RECEIVED': 0,
        'BAL_REQ': 0,
        'ISSUED': 0,
        'CLOSING_STOCK': 0,
    }

    for col, default in needed.items():
        if col not in df.columns:
            df[col] = default

    # Clean numeric columns
    num_cols = ['AVG_LEAD_TIME','SAFETY_STOCK','OPENING_STOCK',
                'MONTH_REQ','RECEIVED','BAL_REQ','ISSUED','CLOSING_STOCK']
    for c in num_cols:
        df[c] = df[c].apply(clean_numeric)

    # Fix Avg Lead Time #NAME? → compute from data
    mask = df['AVG_LEAD_TIME'] == 0
    if mask.any() and 'MONTH_REQ' in df.columns:
        pass  # Keep 0; user can update

    # Recalculate Closing Stock where missing
    recalc = df['CLOSING_STOCK'] == 0
    df.loc[recalc, 'CLOSING_STOCK'] = (
        df.loc[recalc, 'OPENING_STOCK'] +
        df.loc[recalc, 'RECEIVED'] -
        df.loc[recalc, 'ISSUED']
    )

    # Recalculate Balance Required
    df['BAL_REQ'] = df['MONTH_REQ'] - df['RECEIVED']

    # Stock Status
    def status(row):
        cs = row['CLOSING_STOCK']
        ss = row['SAFETY_STOCK']
        mr = row['MONTH_REQ']
        if cs < 0:             return '🔴 CRITICAL'
        if ss > 0 and cs < ss: return '🟠 BELOW SAFETY'
        if mr > 0 and cs < mr * 0.3: return '🟡 LOW'
        return '🟢 OK'

    df['STATUS'] = df.apply(status, axis=1)

    # Category from CODE prefix
    def categorise(code):
        code = str(code).upper()
        if code.startswith('RM'):   return 'Raw Material'
        if code.startswith('GD') or code.startswith('GR') or code.startswith('GU') or code.startswith('GB'):
            return 'Gum'
        if code.startswith('DB') or code.startswith('BKRL') or code.startswith('DBC'):
            return 'Damarbattu'
        if code.startswith('OL'):   return 'Oils'
        if code.startswith('SFG'):  return 'SFG / Diluted'
        if code.startswith('DEP'): return 'DEP Oil'
        if code.startswith('CPN'): return 'Compound'
        if code.startswith('RM-19') or 'PERF' in code.upper():
            return 'Perfumes'
        return 'Other'

    df['CATEGORY'] = df['CODE'].apply(categorise)

    # Month order
    month_order = ['Jan','Feb','Mar','Apr','May','Jun',
                   'Jul','Aug','Sep','Oct','Nov','Dec']
    df['MONTH'] = df['MONTH'].astype(str).str.strip()
    df['MONTH_NUM'] = df['MONTH'].apply(
        lambda m: month_order.index(m)+1
        if m in month_order else 0
    )

    return df

# ── Sample data (fallback) ────────────────────────────────────
def get_sample_data():
    """Return sample data based on your PDF."""
    data = {
        'MONTH': ['Apr','Apr','Apr','Apr','Apr','May','May','May','Jun','Jun'],
        'CODE':  ['RM-25-02','RM-25-27','DB-03-04','GD-03-17','GR-02-04',
                  'RM-25-02','DB-03-04','GR-02-04','RM-25-02','DB-03-04'],
        'ITEM':  ['RAW MATERIAL (Charcoal)','RAW MATERIAL (SAW DUST)',
                  'DAMARBATTU (MIX, 50 KGS)','GUM DAMAR (ABX, 20KGS)',
                  'GUM ROSIN (240KGS, PHT)','RAW MATERIAL (Charcoal)',
                  'DAMARBATTU (MIX, 50 KGS)','GUM ROSIN (240KGS, PHT)',
                  'RAW MATERIAL (Charcoal)','DAMARBATTU (MIX, 50 KGS)'],
        'AVG_LEAD_TIME': [7,5,14,12,10,7,14,10,7,14],
        'SAFETY_STOCK':  [10000,0,0,0,0,10000,0,0,10000,0],
        'OPENING_STOCK': [4338,7195,8200,2200,1980,
                          13003,46814,1685,8662,46814],
        'MONTH_REQ':     [52000,0,52000,0,0,32000,43350,2400,15580,16000],
        'RECEIVED':      [50400,10000,96187,4000,2400,
                          0,0,0,0,0],
        'ISSUED':        [41735,8532,57573,2774,2694,
                          36870,55587,3834,20453,23755],
    }
    df = pd.DataFrame(data)
    df['CLOSING_STOCK'] = df['OPENING_STOCK'] + df['RECEIVED'] - df['ISSUED']
    df['BAL_REQ'] = df['MONTH_REQ'] - df['RECEIVED']
    df['SNO'] = range(1, len(df)+1)

    def status(row):
        cs = row['CLOSING_STOCK']
        ss = row['SAFETY_STOCK']
        if cs < 0:             return '🔴 CRITICAL'
        if ss > 0 and cs < ss: return '🟠 BELOW SAFETY'
        if cs < row['MONTH_REQ'] * 0.3: return '🟡 LOW'
        return '🟢 OK'

    df['STATUS'] = df.apply(status, axis=1)
    df['CATEGORY'] = ['Raw Material','Raw Material','Damarbattu',
                       'Gum','Gum','Raw Material','Damarbattu',
                       'Gum','Raw Material','Damarbattu']
    df['MONTH_NUM'] = [4,4,4,4,4,5,5,5,6,6]
    return df

# ══════════════════════════════════════════════════════════════
#  MAIN APP
# ══════════════════════════════════════════════════════════════
def main():
    # ── Sidebar ───────────────────────────────────────────────
    with st.sidebar:
        st.image("https://via.placeholder.com/200x60?text=Warehouse+IMS",
                 use_column_width=True)
        st.markdown("---")
        st.markdown("### 📂 Upload Data")

        uploaded = st.file_uploader(
            "Upload Excel / CSV",
            type=['xlsx','xls','csv'],
            help="Upload your Fac Outlet Report"
        )

        st.markdown("---")
        st.markdown("### 🔧 Settings")

        auto_refresh = st.checkbox("Auto Refresh (30s)", value=False)
        show_raw    = st.checkbox("Show Raw Data",    value=False)

        if auto_refresh:
            import time
            time.sleep(30)
            st.rerun()

        st.markdown("---")
        st.markdown("### 📅 Last Updated")
        st.info(datetime.now().strftime("%d %b %Y  %H:%M:%S"))

    # ── Load data ─────────────────────────────────────────────
    if uploaded:
        raw_df, err = load_data(uploaded)
        if err:
            st.error(f"❌ File load error: {err}")
            st.info("Using sample data instead.")
            df = get_sample_data()
        else:
            try:
                df = process_dataframe(raw_df)
                st.sidebar.success(f"✅ Loaded {len(df)} rows")
            except Exception as e:
                st.sidebar.error(f"Processing error: {e}")
                df = get_sample_data()
    else:
        df = get_sample_data()
        st.sidebar.warning("⚠️ Using sample data. Upload your file above.")

    # ── Filters ───────────────────────────────────────────────
    st.markdown('<div class="header-title">🏭 Bangalore Warehouse Dashboard</div>',
                unsafe_allow_html=True)
    st.markdown("---")

    f1, f2, f3 = st.columns(3)
    with f1:
        months = ['All'] + sorted(df['MONTH'].unique().tolist())
        sel_month = st.selectbox("📅 Month", months)
    with f2:
        cats = ['All'] + sorted(df['CATEGORY'].unique().tolist())
        sel_cat = st.selectbox("📦 Category", cats)
    with f3:
        statuses = ['All','🔴 CRITICAL','🟠 BELOW SAFETY','🟡 LOW','🟢 OK']
        sel_status = st.selectbox("🚦 Status", statuses)

    # Apply filters
    fdf = df.copy()
    if sel_month  != 'All': fdf = fdf[fdf['MONTH']    == sel_month]
    if sel_cat    != 'All': fdf = fdf[fdf['CATEGORY'] == sel_cat]
    if sel_status != 'All': fdf = fdf[fdf['STATUS']   == sel_status]

    # ── KPI Cards ─────────────────────────────────────────────
    st.markdown("### 📊 Key Metrics")
    k1, k2, k3, k4, k5 = st.columns(5)

    total        = len(fdf)
    critical     = len(fdf[fdf['STATUS'] == '🔴 CRITICAL'])
    below_safety = len(fdf[fdf['STATUS'] == '🟠 BELOW SAFETY'])
    low_stock    = len(fdf[fdf['STATUS'] == '🟡 LOW'])
    ok_stock     = len(fdf[fdf['STATUS'] == '🟢 OK'])

    k1.metric("📦 Total Items",       total)
    k2.metric("🔴 Critical",          critical,
              delta=f"-{critical}" if critical else None,
              delta_color="inverse")
    k3.metric("🟠 Below Safety",      below_safety)
    k4.metric("🟡 Low Stock",         low_stock)
    k5.metric("🟢 OK",                ok_stock)

    st.markdown("---")

    # ── Row 1: Charts ─────────────────────────────────────────
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("#### 📈 Stock Status Distribution")
        status_counts = fdf['STATUS'].value_counts().reset_index()
        status_counts.columns = ['Status','Count']
        color_map = {
            '🔴 CRITICAL':      '#ff4444',
            '🟠 BELOW SAFETY':  '#ff8800',
            '🟡 LOW':           '#ffcc00',
            '🟢 OK':            '#00cc44',
        }
        fig_pie = px.pie(
            status_counts, names='Status', values='Count',
            color='Status', color_discrete_map=color_map,
            hole=0.4
        )
        fig_pie.update_layout(margin=dict(t=20,b=20,l=0,r=0))
        st.plotly_chart(fig_pie, use_container_width=True)

    with c2:
        st.markdown("#### 📦 Items by Category")
        cat_counts = fdf.groupby('CATEGORY').agg(
            Count=('CODE','count'),
            Avg_Closing=('CLOSING_STOCK','mean')
        ).reset_index()
        fig_bar = px.bar(
            cat_counts, x='CATEGORY', y='Count',
            color='Count', color_continuous_scale='Blues',
            text='Count'
        )
        fig_bar.update_traces(textposition='outside')
        fig_bar.update_layout(
            margin=dict(t=20,b=60,l=0,r=0),
            xaxis_tickangle=-30,
            coloraxis_showscale=False
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    # ── Row 2: Trend + Lead Time ──────────────────────────────
    c3, c4 = st.columns(2)

    with c3:
        st.markdown("#### 📅 Monthly Received vs Issued")
        monthly = df.groupby(['MONTH','MONTH_NUM']).agg(
            Received=('RECEIVED','sum'),
            Issued=('ISSUED','sum'),
            Requirement=('MONTH_REQ','sum')
        ).reset_index().sort_values('MONTH_NUM')

        fig_trend = go.Figure()
        fig_trend.add_trace(go.Bar(
            name='Received', x=monthly['MONTH'],
            y=monthly['Received'], marker_color='#4CAF50'))
        fig_trend.add_trace(go.Bar(
            name='Issued', x=monthly['MONTH'],
            y=monthly['Issued'], marker_color='#2196F3'))
        fig_trend.add_trace(go.Scatter(
            name='Requirement', x=monthly['MONTH'],
            y=monthly['Requirement'],
            mode='lines+markers',
            line=dict(color='red', width=2, dash='dot')))
        fig_trend.update_layout(
            barmode='group',
            margin=dict(t=20,b=20,l=0,r=0),
            legend=dict(orientation='h', y=-0.2)
        )
        st.plotly_chart(fig_trend, use_container_width=True)

    with c4:
        st.markdown("#### ⏱️ Avg Lead Time by Category")
        lt_df = fdf[fdf['AVG_LEAD_TIME'] > 0].groupby('CATEGORY').agg(
            Avg_Lead=('AVG_LEAD_TIME','mean')
        ).reset_index().sort_values('Avg_Lead', ascending=True)

        if len(lt_df):
            fig_lt = px.bar(
                lt_df, x='Avg_Lead', y='CATEGORY',
                orientation='h', color='Avg_Lead',
                color_continuous_scale='RdYlGn_r',
                text=lt_df['Avg_Lead'].round(1)
            )
            fig_lt.update_traces(textposition='outside')
            fig_lt.update_layout(
                margin=dict(t=20,b=20,l=0,r=0),
                coloraxis_showscale=False
            )
            st.plotly_chart(fig_lt, use_container_width=True)
        else:
            st.info("No lead time data available")

    # ── Row 3: Critical Items Table ───────────────────────────
    st.markdown("---")
    st.markdown("### 🚨 Items Needing Attention")

    alert_df = fdf[fdf['STATUS'].isin(
        ['🔴 CRITICAL','🟠 BELOW SAFETY','🟡 LOW']
    )][['CODE','ITEM','MONTH','CATEGORY','STATUS',
        'OPENING_STOCK','RECEIVED','ISSUED',
        'CLOSING_STOCK','MONTH_REQ','AVG_LEAD_TIME']].copy()

    alert_df.columns = ['Code','Item','Month','Category','Status',
                         'Opening','Received','Issued',
                         'Closing Stock','Requirement','Lead Time(days)']

    if len(alert_df):
        def color_status(val):
            if '🔴' in str(val): return 'background-color:#ffcccc'
            if '🟠' in str(val): return 'background-color:#ffe0b2'
            if '🟡' in str(val): return 'background-color:#fff9c4'
            return ''

        styled = alert_df.style.applymap(
            color_status, subset=['Status']
        ).format({
            'Opening':       '{:,.1f}',
            'Received':      '{:,.1f}',
            'Issued':        '{:,.1f}',
            'Closing Stock': '{:,.1f}',
            'Requirement':   '{:,.1f}',
        })
        st.dataframe(styled, use_container_width=True, height=350)
    else:
        st.success("✅ No critical items in current selection!")

    # ── Row 4: Full data table ────────────────────────────────
    if show_raw:
        st.markdown("---")
        st.markdown("### 📋 Full Data Table")
        display_cols = ['CODE','ITEM','MONTH','CATEGORY','STATUS',
                        'SAFETY_STOCK','OPENING_STOCK','MONTH_REQ',
                        'RECEIVED','ISSUED','CLOSING_STOCK',
                        'BAL_REQ','AVG_LEAD_TIME']
        existing = [c for c in display_cols if c in fdf.columns]
        st.dataframe(fdf[existing], use_container_width=True)

    # ── Row 5: Stock Level Gauge ──────────────────────────────
    st.markdown("---")
    st.markdown("### 🔍 Item Deep-Dive")

    item_list = fdf['ITEM'].unique().tolist()
    sel_item  = st.selectbox("Select Item", item_list)

    item_data = fdf[fdf['ITEM'] == sel_item]

    if len(item_data):
        row = item_data.iloc[-1]
        d1, d2, d3, d4 = st.columns(4)
        d1.metric("Opening Stock",  f"{row['OPENING_STOCK']:,.0f}")
        d2.metric("Received",       f"{row['RECEIVED']:,.0f}")
        d3.metric("Issued",         f"{row['ISSUED']:,.0f}")
        d4.metric("Closing Stock",  f"{row['CLOSING_STOCK']:,.0f}",
                  delta=f"{row['CLOSING_STOCK']-row['OPENING_STOCK']:,.0f}")

        # Gauge
        max_val = max(row['MONTH_REQ'], row['CLOSING_STOCK'], 1)
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=row['CLOSING_STOCK'],
            delta={'reference': row['MONTH_REQ'],
                   'valueformat': ',.0f'},
            title={'text': f"Closing Stock vs Requirement"},
            gauge={
                'axis': {'range': [0, max_val * 1.2]},
                'bar':  {'color': '#2196F3'},
                'steps': [
                    {'range': [0, row['SAFETY_STOCK']],
                     'color': '#ffcccc'},
                    {'range': [row['SAFETY_STOCK'], row['MONTH_REQ']],
                     'color': '#fff9c4'},
                    {'range': [row['MONTH_REQ'], max_val*1.2],
                     'color': '#ccffcc'},
                ],
                'threshold': {
                    'line': {'color': 'red', 'width': 4},
                    'thickness': 0.75,
                    'value': row['SAFETY_STOCK']
                }
            }
        ))
        fig_gauge.update_layout(height=300, margin=dict(t=30,b=0,l=0,r=0))
        st.plotly_chart(fig_gauge, use_container_width=True)

        # Monthly trend for this item
        item_monthly = df[df['ITEM'] == sel_item].sort_values('MONTH_NUM')
        if len(item_monthly) > 1:
            fig_item = px.line(
                item_monthly, x='MONTH',
                y=['OPENING_STOCK','RECEIVED','ISSUED','CLOSING_STOCK'],
                markers=True,
                title=f"Monthly Trend: {sel_item}"
            )
            st.plotly_chart(fig_item, use_container_width=True)

    # ── Download ──────────────────────────────────────────────
    st.markdown("---")
    col_dl1, col_dl2 = st.columns(2)

    with col_dl1:
        csv = fdf.to_csv(index=False).encode('utf-8')
        st.download_button(
            "⬇️ Download Filtered Data (CSV)",
            csv, "filtered_stock.csv", "text/csv"
        )

    with col_dl2:
        alert_csv = alert_df.to_csv(index=False).encode('utf-8') \
                    if len(alert_df) else b"No alerts"
        st.download_button(
            "⬇️ Download Alert Items (CSV)",
            alert_csv, "alert_items.csv", "text/csv"
        )

    # ── Footer ────────────────────────────────────────────────
    st.markdown("---")
    st.markdown(
        "<center>🏭 Bangalore Warehouse IMS • "
        f"Built with Streamlit • {datetime.now().year}</center>",
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
