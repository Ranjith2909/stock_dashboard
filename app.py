import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
import time
import io

# ==================== PAGE CONFIG ====================
st.set_page_config(
    page_title="Ultimate Pro Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== SESSION STATE ====================
if 'data' not in st.session_state:
    st.session_state.data = None
if 'clean_data' not in st.session_state:
    st.session_state.clean_data = None
if 'refresh_count' not in st.session_state:
    st.session_state.refresh_count = 0

# ==================== SIDEBAR ====================
with st.sidebar:
    st.title("⚙️ Control Panel")
    st.markdown("---")
    
    # THEME SELECTION
    st.subheader("🎨 Theme & Colors")
    theme = st.selectbox("Choose Theme", ["Professional Blue", "Dark Mode", "Ocean", "Sunset", "Forest", "Custom"])
    
    if theme == "Custom":
        primary_color = st.color_picker("Primary Color", "#667eea")
        secondary_color = st.color_picker("Secondary Color", "#764ba2")
        accent_color = st.color_picker("Accent Color", "#00cc00")
    else:
        color_themes = {
            "Professional Blue": ("#667eea", "#764ba2", "#00cc00"),
            "Dark Mode": ("#1e1e1e", "#333333", "#4CAF50"),
            "Ocean": ("#0077be", "#00a8cc", "#7fdbff"),
            "Sunset": ("#ff6b6b", "#feca57", "#ff9ff3"),
            "Forest": ("#2d5016", "#73a942", "#aad576"),
        }
        primary_color, secondary_color, accent_color = color_themes[theme]
    
    st.markdown("---")
    
    # AUTO REFRESH
    st.subheader("🔄 Auto Refresh")
    auto_refresh = st.checkbox("Enable Auto-Refresh", value=False)
    refresh_seconds = st.slider("Refresh every (seconds)", 5, 300, 30)
    
    st.markdown("---")
    
    # DATA SOURCE
    st.subheader("📁 Data Source")
    data_source = st.radio("Choose Source", ["Upload File", "Live URL (Google Sheets)", "Sample Data"])
    
    st.markdown("---")
    
    # AUTO CLEAN OPTIONS
    st.subheader("🧹 Auto-Clean Options")
    remove_duplicates = st.checkbox("Remove Duplicates", value=True)
    fill_missing = st.checkbox("Fill Missing Values", value=True)
    trim_spaces = st.checkbox("Trim Spaces", value=True)
    standardize_case = st.checkbox("Standardize Columns", value=True)
    remove_special = st.checkbox("Remove Special Characters", value=False)
    
    st.markdown("---")
    
    # DISPLAY OPTIONS
    st.subheader("🖥️ Display")
    show_kpis = st.checkbox("Show KPIs", value=True)
    show_alerts = st.checkbox("Show Alerts", value=True)
    show_charts = st.checkbox("Show Charts", value=True)
    show_pivot = st.checkbox("Show Pivot", value=True)
    
    st.markdown("---")
    st.info(f"🔄 Refresh Count: {st.session_state.refresh_count}")

# ==================== CUSTOM CSS ====================
st.markdown(f"""
    <style>
    .stApp {{
        background: linear-gradient(135deg, {primary_color}15 0%, {secondary_color}15 100%);
    }}
    .main-header {{
        background: linear-gradient(135deg, {primary_color} 0%, {secondary_color} 100%);
        padding: 30px;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin-bottom: 30px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.1);
    }}
    .kpi-card {{
        background: white;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid {primary_color};
        box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        margin-bottom: 15px;
    }}
    .alert-critical {{
        background: linear-gradient(135deg, #ff6b6b 0%, #ff4757 100%);
        color: white;
        padding: 15px;
        border-radius: 10px;
        margin: 10px 0;
    }}
    .alert-warning {{
        background: linear-gradient(135deg, #feca57 0%, #ff9ff3 100%);
        color: white;
        padding: 15px;
        border-radius: 10px;
        margin: 10px 0;
    }}
    .stTabs [data-baseweb="tab-list"] {{
        gap: 10px;
    }}
    .stTabs [data-baseweb="tab"] {{
        background-color: white;
        border-radius: 10px;
        padding: 10px 20px;
        font-weight: 600;
    }}
    .stTabs [aria-selected="true"] {{
        background: linear-gradient(135deg, {primary_color} 0%, {secondary_color} 100%);
        color: white;
    }}
    </style>
""", unsafe_allow_html=True)

# ==================== HEADER ====================
st.markdown("""
<div class="main-header">
    <h1>📊 Ultimate Professional Dashboard</h1>
    <p>Auto-Clean | Live Data | Advanced Analytics | Real-time Updates</p>
</div>
""", unsafe_allow_html=True)

# ==================== DATA LOADING ====================
def clean_data(df, options):
    """Auto-clean the dataframe"""
    original_rows = len(df)
    
    # Standardize columns
    if options['standardize']:
        df.columns = df.columns.str.strip().str.upper()
    
    # Trim spaces
    if options['trim']:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
    
    # Remove duplicates
    duplicates_removed = 0
    if options['duplicates']:
        before = len(df)
        df = df.drop_duplicates()
        duplicates_removed = before - len(df)
    
    # Fill missing
    missing_filled = 0
    if options['missing']:
        missing_filled = df.isnull().sum().sum()
        for col in df.select_dtypes(include=[np.number]).columns:
            df[col] = df[col].fillna(0)
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].fillna('N/A')
    
    # Remove special characters
    if options['special']:
        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.replace(r'[^\w\s]', '', regex=True)
    
    # Convert numeric strings to numbers
    for col in df.columns:
        try:
            df[col] = pd.to_numeric(df[col])
        except:
            pass
    
    return df, {
        'original_rows': original_rows,
        'final_rows': len(df),
        'duplicates_removed': duplicates_removed,
        'missing_filled': missing_filled
    }

# DATA SOURCE HANDLING
df = None

if data_source == "Upload File":
    uploaded_file = st.file_uploader(
        "📁 Upload Excel/CSV File",
        type=['csv', 'xlsx', 'xlsm', 'xls'],
        help="Supports Excel and CSV formats"
    )
    
    if uploaded_file:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            st.session_state.data = df
        except Exception as e:
            st.error(f"Error reading file: {e}")

elif data_source == "Live URL (Google Sheets)":
    url = st.text_input(
        "📎 Paste Google Sheets URL (or CSV URL)",
        help="Make sheet public: File → Share → Anyone with link → Viewer"
    )
    st.info("💡 For Google Sheets: Replace '/edit' with '/export?format=csv' in URL")
    
    if url:
        try:
            if 'docs.google.com/spreadsheets' in url:
                if '/edit' in url:
                    url = url.replace('/edit#gid=', '/export?format=csv&gid=')
                    url = url.replace('/edit?usp=sharing', '/export?format=csv')
                    url = url.replace('/edit', '/export?format=csv')
            
            df = pd.read_csv(url)
            st.session_state.data = df
            st.success("✅ Live data loaded from URL!")
        except Exception as e:
            st.error(f"Error loading URL: {e}")

elif data_source == "Sample Data":
    # Create sample stock data
    np.random.seed(42)
    sample_data = {
        'ITEM': [f'Product_{i}' for i in range(1, 51)],
        'CODE': [f'PRD-{i:03d}' for i in range(1, 51)],
        'CATEGORY': np.random.choice(['Raw Material', 'Finished Goods', 'Semi-Finished'], 50),
        'SAFETY STOCK': np.random.randint(100, 1000, 50),
        'OPENING STOCK': np.random.randint(500, 5000, 50),
        'RECEIVED': np.random.randint(0, 2000, 50),
        'ISSUED': np.random.randint(0, 3000, 50),
        'CLOSING STOCK': np.random.randint(0, 5000, 50),
        'PRICE': np.random.uniform(10, 500, 50).round(2)
    }
    df = pd.DataFrame(sample_data)
    st.session_state.data = df
    st.success("✅ Sample data loaded!")

# ==================== MAIN DASHBOARD ====================
if st.session_state.data is not None:
    df_raw = st.session_state.data
    
    # AUTO CLEAN
    clean_options = {
        'duplicates': remove_duplicates,
        'missing': fill_missing,
        'trim': trim_spaces,
        'standardize': standardize_case,
        'special': remove_special
    }
    
    df, clean_stats = clean_data(df_raw.copy(), clean_options)
    st.session_state.clean_data = df
    
    # Show cleaning stats
    with st.expander("🧹 Data Cleaning Report"):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Original Rows", clean_stats['original_rows'])
        with col2:
            st.metric("Cleaned Rows", clean_stats['final_rows'])
        with col3:
            st.metric("Duplicates Removed", clean_stats['duplicates_removed'])
        with col4:
            st.metric("Missing Filled", clean_stats['missing_filled'])
    
    # Get column types
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    text_cols = df.select_dtypes(include=['object']).columns.tolist()
    
    # ==================== TABS ====================
    tabs = st.tabs(["📈 Live KPIs", "🚨 Alerts", "📊 Advanced Charts", "🔄 Pivot Analysis", 
                     "📉 Statistics", "🔍 Filter & Query", "📋 Raw Data", "📥 Export"])
    
    # ========== TAB 1: KPIs ==========
    with tabs[0]:
        if show_kpis:
            st.subheader("📈 Real-Time Key Performance Indicators")
            
            # Auto-generate KPIs from data
            st.markdown("### 📊 Overview Metrics")
            cols = st.columns(4)
            
            with cols[0]:
                st.metric("📦 Total Records", f"{len(df):,}")
            with cols[1]:
                st.metric("📊 Columns", len(df.columns))
            with cols[2]:
                st.metric("🔢 Numeric Fields", len(numeric_cols))
            with cols[3]:
                st.metric("📝 Text Fields", len(text_cols))
            
            st.markdown("---")
            
            # Business KPIs
            st.markdown("### 💼 Business Metrics")
            
            business_kpis = []
            if 'CLOSING STOCK' in df.columns:
                business_kpis.append(('Total Closing Stock', df['CLOSING STOCK'].sum(), '📦'))
            if 'OPENING STOCK' in df.columns:
                business_kpis.append(('Total Opening Stock', df['OPENING STOCK'].sum(), '📥'))
            if 'RECEIVED' in df.columns:
                business_kpis.append(('Total Received', df['RECEIVED'].sum(), '⬇️'))
            if 'ISSUED' in df.columns:
                business_kpis.append(('Total Issued', df['ISSUED'].sum(), '⬆️'))
            if 'PRICE' in df.columns and 'CLOSING STOCK' in df.columns:
                total_value = (df['PRICE'] * df['CLOSING STOCK']).sum()
                business_kpis.append(('Total Value', f"₹{total_value:,.0f}", '💰'))
            
            if business_kpis:
                cols = st.columns(len(business_kpis))
                for i, (label, value, icon) in enumerate(business_kpis):
                    with cols[i]:
                        if isinstance(value, str):
                            st.metric(f"{icon} {label}", value)
                        else:
                            st.metric(f"{icon} {label}", f"{value:,.0f}")
            
            st.markdown("---")
            
            # Auto-generated KPIs for all numeric columns
            st.markdown("### 🔢 Numeric Column Analysis")
            
            for col in numeric_cols[:6]:  # Show first 6
                with st.expander(f"📊 {col}"):
                    c1, c2, c3, c4, c5 = st.columns(5)
                    with c1:
                        st.metric("Sum", f"{df[col].sum():,.2f}")
                    with c2:
                        st.metric("Avg", f"{df[col].mean():,.2f}")
                    with c3:
                        st.metric("Max", f"{df[col].max():,.2f}")
                    with c4:
                        st.metric("Min", f"{df[col].min():,.2f}")
                    with c5:
                        st.metric("Median", f"{df[col].median():,.2f}")
    
    # ========== TAB 2: ALERTS ==========
    with tabs[1]:
        if show_alerts:
            st.subheader("🚨 Alerts & Warnings")
            
            alerts_found = False
            
            # Low Stock Alert
            if 'CLOSING STOCK' in df.columns and 'SAFETY STOCK' in df.columns:
                critical = df[df['CLOSING STOCK'] < df['SAFETY STOCK']]
                if len(critical) > 0:
                    alerts_found = True
                    st.markdown(f"""
                    <div class="alert-critical">
                        <h3>🔴 CRITICAL: {len(critical)} items below safety stock!</h3>
                    </div>
                    """, unsafe_allow_html=True)
                    st.dataframe(critical, use_container_width=True)
            
            # Dead Stock Alert
            if 'ISSUED' in df.columns and 'RECEIVED' in df.columns:
                dead = df[(df['ISSUED'] == 0) & (df['RECEIVED'] == 0)]
                if len(dead) > 0:
                    alerts_found = True
                    st.markdown(f"""
                    <div class="alert-warning">
                        <h3>⚠️ WARNING: {len(dead)} items with no movement (Dead Stock)</h3>
                    </div>
                    """, unsafe_allow_html=True)
                    st.dataframe(dead, use_container_width=True)
            
            # Out of Stock
            if 'CLOSING STOCK' in df.columns:
                out_stock = df[df['CLOSING STOCK'] == 0]
                if len(out_stock) > 0:
                    alerts_found = True
                    st.error(f"❌ {len(out_stock)} items are OUT OF STOCK")
                    st.dataframe(out_stock, use_container_width=True)
            
            if not alerts_found:
                st.success("✅ No alerts! Everything looks good!")
    
    # ========== TAB 3: CHARTS ==========
    with tabs[2]:
        if show_charts:
            st.subheader("📊 Advanced Visualizations")
            
            # Chart selection
            col1, col2 = st.columns([1, 3])
            
            with col1:
                chart_type = st.selectbox("Chart Type", [
                    "Bar Chart", "Horizontal Bar", "Pie Chart", "Donut Chart",
                    "Line Chart", "Area Chart", "Scatter Plot", "Bubble Chart",
                    "Histogram", "Box Plot", "Violin Plot", "Heatmap",
                    "Treemap", "Sunburst", "Funnel Chart"
                ])
                
                if numeric_cols:
                    x_axis = st.selectbox("X-Axis", df.columns.tolist())
                    y_axis = st.selectbox("Y-Axis", numeric_cols)
                    
                    color_by = st.selectbox("Color By", [None] + df.columns.tolist())
                    
                    if chart_type in ["Bar Chart", "Horizontal Bar", "Pie Chart", "Donut Chart"]:
                        top_n = st.slider("Top N", 5, 50, 10)
            
            with col2:
                try:
                    color_scheme = px.colors.sequential.Viridis
                    
                    if chart_type == "Bar Chart":
                        data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(top_n)
                        fig = px.bar(data, x=x_axis, y=y_axis, color=color_by or y_axis,
                                    color_continuous_scale='Viridis', title=f"{y_axis} by {x_axis}")
                    
                    elif chart_type == "Horizontal Bar":
                        data = df.nlargest(top_n, y_axis)
                        fig = px.bar(data, x=y_axis, y=x_axis, orientation='h', color=color_by or y_axis,
                                    color_continuous_scale='Viridis')
                    
                    elif chart_type == "Pie Chart":
                        data = df.nlargest(top_n, y_axis)
                        fig = px.pie(data, names=x_axis, values=y_axis, title=f"Distribution")
                    
                    elif chart_type == "Donut Chart":
                        data = df.nlargest(top_n, y_axis)
                        fig = px.pie(data, names=x_axis, values=y_axis, hole=0.5)
                    
                    elif chart_type == "Line Chart":
                        fig = px.line(df.head(50), x=x_axis, y=y_axis, color=color_by, markers=True)
                    
                    elif chart_type == "Area Chart":
                        fig = px.area(df.head(50), x=x_axis, y=y_axis, color=color_by)
                    
                    elif chart_type == "Scatter Plot":
                        fig = px.scatter(df, x=x_axis, y=y_axis, color=color_by, size=y_axis)
                    
                    elif chart_type == "Bubble Chart":
                        fig = px.scatter(df, x=x_axis, y=y_axis, size=y_axis, color=color_by)
                    
                    elif chart_type == "Histogram":
                        fig = px.histogram(df, x=y_axis, nbins=30, color=color_by)
                    
                    elif chart_type == "Box Plot":
                        fig = px.box(df, y=y_axis, color=color_by)
                    
                    elif chart_type == "Violin Plot":
                        fig = px.violin(df, y=y_axis, color=color_by, box=True)
                    
                    elif chart_type == "Heatmap":
                        if len(numeric_cols) > 1:
                            corr = df[numeric_cols].corr()
                            fig = px.imshow(corr, text_auto=True, color_continuous_scale='RdBu_r')
                    
                    elif chart_type == "Treemap":
                        data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(20)
                        fig = px.treemap(data, path=[x_axis], values=y_axis)
                    
                    elif chart_type == "Sunburst":
                        if len(text_cols) >= 2:
                            fig = px.sunburst(df.head(30), path=text_cols[:2], values=y_axis)
                        else:
                            fig = px.sunburst(df.head(30), path=[x_axis], values=y_axis)
                    
                    elif chart_type == "Funnel Chart":
                        data = df.nlargest(top_n, y_axis)
                        fig = px.funnel(data, x=y_axis, y=x_axis)
                    
                    # Apply theme
                    fig.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)',
                        paper_bgcolor='rgba(0,0,0,0)',
                        height=500
                    )
                    st.plotly_chart(fig, use_container_width=True)
                
                except Exception as e:
                    st.error(f"Chart error: {e}")
            
            st.markdown("---")
            
            # Auto-generated dashboard
            st.subheader("📊 Auto-Generated Dashboard")
            
            if len(numeric_cols) >= 1 and len(text_cols) >= 1:
                col1, col2 = st.columns(2)
                
                with col1:
                    top10 = df.nlargest(10, numeric_cols[0])
                    fig1 = px.bar(top10, x=text_cols[0], y=numeric_cols[0],
                                 color=numeric_cols[0], color_continuous_scale='Viridis',
                                 title=f"Top 10 by {numeric_cols[0]}")
                    st.plotly_chart(fig1, use_container_width=True)
                
                with col2:
                    fig2 = px.pie(top10, names=text_cols[0], values=numeric_cols[0],
                                 title=f"Distribution")
                    st.plotly_chart(fig2, use_container_width=True)
                
                if len(numeric_cols) >= 2:
                    col1, col2 = st.columns(2)
                    with col1:
                        fig3 = px.scatter(df, x=numeric_cols[0], y=numeric_cols[1],
                                         title=f"{numeric_cols[0]} vs {numeric_cols[1]}")
                        st.plotly_chart(fig3, use_container_width=True)
                    
                    with col2:
                        fig4 = px.histogram(df, x=numeric_cols[0], nbins=20,
                                           title=f"Distribution of {numeric_cols[0]}")
                        st.plotly_chart(fig4, use_container_width=True)
    
    # ========== TAB 4: PIVOT ==========
    with tabs[3]:
        if show_pivot:
            st.subheader("🔄 Advanced Pivot Table (Excel-like)")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                pivot_rows = st.multiselect("📊 Rows", text_cols, default=text_cols[:1] if text_cols else [])
            
            with col2:
                pivot_cols = st.multiselect("📋 Columns", text_cols)
            
            with col3:
                pivot_values = st.multiselect("🔢 Values", numeric_cols, default=numeric_cols[:1] if numeric_cols else [])
            
            with col4:
                pivot_agg = st.selectbox("⚡ Function", ["sum", "mean", "count", "max", "min", "median", "std"])
            
            if pivot_rows and pivot_values:
                try:
                    pivot_table = pd.pivot_table(
                        df,
                        index=pivot_rows,
                        columns=pivot_cols if pivot_cols else None,
                        values=pivot_values,
                        aggfunc=pivot_agg,
                        fill_value=0,
                        margins=True,
                        margins_name='Total'
                    )
                    
                    st.markdown("### 📊 Pivot Result")
                    st.dataframe(pivot_table, use_container_width=True, height=400)
                    
                    # Download
                    pivot_csv = pivot_table.to_csv()
                    st.download_button("📥 Download Pivot", pivot_csv, "pivot_table.csv")
                    
                    # Visualize
                    st.markdown("### 📈 Pivot Visualization")
                    
                    try:
                        # Flatten for viz
                        pivot_reset = pivot_table.reset_index()
                        
                        chart_col1, chart_col2 = st.columns(2)
                        with chart_col1:
                            fig = px.bar(pivot_reset.head(20), 
                                        x=pivot_rows[0], 
                                        y=pivot_values[0] if isinstance(pivot_values[0], str) else pivot_values[0][0],
                                        title="Pivot Bar Chart")
                            st.plotly_chart(fig, use_container_width=True)
                        
                        with chart_col2:
                            fig = px.line(pivot_reset.head(20), 
                                         x=pivot_rows[0], 
                                         y=pivot_values[0] if isinstance(pivot_values[0], str) else pivot_values[0][0],
                                         markers=True, title="Pivot Line Chart")
                            st.plotly_chart(fig, use_container_width=True)
                    except:
                        pass
                
                except Exception as e:
                    st.error(f"Pivot error: {e}")
            else:
                st.info("👆 Select at least Rows and Values to create pivot")
    
    # ========== TAB 5: STATISTICS ==========
    with tabs[4]:
        st.subheader("📉 Statistical Analysis")
        
        st.markdown("### 📊 Descriptive Statistics")
        st.dataframe(df.describe(include='all'), use_container_width=True)
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### 🔢 Column Statistics")
            if numeric_cols:
                selected_col = st.selectbox("Select Column", numeric_cols)
                
                stats_data = {
                    'Metric': ['Count', 'Sum', 'Mean', 'Median', 'Mode', 'Std Dev', 'Variance', 'Min', 'Max', 'Range', 'Q1', 'Q3', 'IQR'],
                    'Value': [
                        df[selected_col].count(),
                        df[selected_col].sum(),
                        df[selected_col].mean(),
                        df[selected_col].median(),
                        df[selected_col].mode()[0] if not df[selected_col].mode().empty else 0,
                        df[selected_col].std(),
                        df[selected_col].var(),
                        df[selected_col].min(),
                        df[selected_col].max(),
                        df[selected_col].max() - df[selected_col].min(),
                        df[selected_col].quantile(0.25),
                        df[selected_col].quantile(0.75),
                        df[selected_col].quantile(0.75) - df[selected_col].quantile(0.25)
                    ]
                }
                stats_df = pd.DataFrame(stats_data)
                stats_df['Value'] = stats_df['Value'].round(2)
                st.dataframe(stats_df, use_container_width=True, hide_index=True)
        
        with col2:
            st.markdown("### 📈 Distribution")
            if numeric_cols:
                fig = px.histogram(df, x=selected_col, nbins=30, marginal="box", title=f"Distribution of {selected_col}")
                st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        st.markdown("### 🔗 Correlation Analysis")
        if len(numeric_cols) > 1:
            corr = df[numeric_cols].corr()
            fig = px.imshow(corr, text_auto=True, aspect="auto", color_continuous_scale='RdBu_r')
            st.plotly_chart(fig, use_container_width=True)
    
    # ========== TAB 6: FILTER & QUERY ==========
    with tabs[5]:
        st.subheader("🔍 Advanced Filter & Query")
        
        # Global search
        search = st.text_input("🔍 Global Search (across all columns)")
        
        # Multiple column filters
        st.markdown("### 🎯 Column Filters")
        
        num_filters = st.slider("Number of filters", 1, 5, 1)
        
        filters = {}
        for i in range(num_filters):
            col1, col2 = st.columns(2)
            with col1:
                filter_col = st.selectbox(f"Column {i+1}", df.columns.tolist(), key=f"col_{i}")
            with col2:
                if filter_col in numeric_cols:
                    min_v = float(df[filter_col].min())
                    max_v = float(df[filter_col].max())
                    filters[filter_col] = st.slider(f"Range", min_v, max_v, (min_v, max_v), key=f"range_{i}")
                else:
                    unique_vals = df[filter_col].unique().tolist()
                    filters[filter_col] = st.multiselect(f"Select", unique_vals, default=unique_vals[:5] if len(unique_vals) > 5 else unique_vals, key=f"select_{i}")
        
        # Apply filters
        filtered_df = df.copy()
        for col, val in filters.items():
            if col in numeric_cols:
                filtered_df = filtered_df[(filtered_df[col] >= val[0]) & (filtered_df[col] <= val[1])]
            else:
                if val:
                    filtered_df = filtered_df[filtered_df[col].isin(val)]
        
        if search:
            filtered_df = filtered_df[filtered_df.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]
        
        st.markdown(f"### 📋 Results: {len(filtered_df)} rows")
        st.dataframe(filtered_df, use_container_width=True, height=500)
        
        csv_filtered = filtered_df.to_csv(index=False)
        st.download_button("📥 Download Filtered", csv_filtered, "filtered.csv")
    
    # ========== TAB 7: RAW DATA ==========
    with tabs[6]:
        st.subheader("📋 Complete Raw Data")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Total Rows", len(df))
        with col2:
            st.metric("Columns", len(df.columns))
        with col3:
            st.metric("Missing", df.isnull().sum().sum())
        with col4:
            st.metric("Duplicates", df.duplicated().sum())
        
        st.dataframe(df, use_container_width=True, height=600)
        
        st.markdown("### 📊 Column Info")
        info_df = pd.DataFrame({
            'Column': df.columns,
            'Type': df.dtypes.astype(str),
            'Non-Null': df.count().values,
            'Null': df.isnull().sum().values,
            'Unique': [df[col].nunique() for col in df.columns]
        })
        st.dataframe(info_df, use_container_width=True, hide_index=True)
    
    # ========== TAB 8: EXPORT ==========
    with tabs[7]:
        st.subheader("📥 Export Options")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            csv = df.to_csv(index=False)
            st.download_button("📄 Download CSV", csv, f"data_{datetime.now():%Y%m%d_%H%M%S}.csv", "text/csv")
        
        with col2:
            json_data = df.to_json(orient='records', indent=2)
            st.download_button("📋 Download JSON", json_data, f"data_{datetime.now():%Y%m%d_%H%M%S}.json", "application/json")
        
        with col3:
            # Excel export
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Data', index=False)
            excel_data = output.getvalue()
            st.download_button("📊 Download Excel", excel_data, f"data_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # ==================== AUTO REFRESH ====================
    if auto_refresh:
        st.session_state.refresh_count += 1
        placeholder = st.empty()
        with placeholder:
            with st.spinner(f"⏳ Auto-refresh in {refresh_seconds} seconds..."):
                time.sleep(refresh_seconds)
        st.rerun()

else:
    # Landing page
    st.info("👆 Upload a file, provide a URL, or use sample data to begin!")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("### 🚀 Features")
        st.markdown("✅ Auto-Clean Data")
        st.markdown("✅ Live Updates")
        st.markdown("✅ 15+ Chart Types")
        st.markdown("✅ Excel-like Pivot")
    
    with col2:
        st.markdown("### 🎨 Customization")
        st.markdown("✅ Multiple Themes")
        st.markdown("✅ Custom Colors")
        st.markdown("✅ Refresh Control")
        st.markdown("✅ Display Options")
    
    with col3:
        st.markdown("### 📊 Analytics")
        st.markdown("✅ Real-time KPIs")
        st.markdown("✅ Statistics")
        st.markdown("✅ Advanced Filters")
        st.markdown("✅ Multi-Export")
