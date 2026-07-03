import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
import io
import time
import warnings
import re
from datetime import datetime
import pdfplumber
import tabula
import PyPDF2

warnings.filterwarnings('ignore')

# ==================== PAGE CONFIG ====================
st.set_page_config(
    page_title="Ultimate Pro Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==================== SESSION STATE ====================
for key in ['data', 'clean_data', 'refresh_count', 'file_name']:
    if key not in st.session_state:
        st.session_state[key] = None if key != 'refresh_count' else 0

# ==================== ULTRA SMART CLEAN DATA ====================
def ultra_clean_data(df, options):
    """Ultimate data cleaning - handles any messy data"""
    if df is None or df.empty:
        return df, {'error': 'Empty dataframe'}

    original_rows = len(df)
    original_cols = len(df.columns)
    duplicates_removed = 0
    missing_filled = 0
    cols_dropped = 0

    try:
        # ===== STEP 1: Fix Column Names =====
        if options.get('standardize', True):
            new_cols = []
            for col in df.columns:
                col_str = str(col).strip()
                col_str = re.sub(r'[^\w\s]', '_', col_str)
                col_str = re.sub(r'\s+', '_', col_str)
                col_str = re.sub(r'_+', '_', col_str)
                col_str = col_str.strip('_').upper()
                if not col_str or col_str.isdigit():
                    col_str = f'COL_{col_str}'
                new_cols.append(col_str)
            
            # Handle duplicate column names
            seen = {}
            final_cols = []
            for c in new_cols:
                if c in seen:
                    seen[c] += 1
                    final_cols.append(f'{c}_{seen[c]}')
                else:
                    seen[c] = 0
                    final_cols.append(c)
            df.columns = final_cols

        # ===== STEP 2: Drop Empty Rows and Columns =====
        df = df.dropna(how='all')
        df = df.dropna(axis=1, how='all')
        
        # Drop columns where 90%+ values are missing
        threshold = 0.9
        cols_before = len(df.columns)
        df = df.dropna(axis=1, thresh=int(len(df) * (1 - threshold)))
        cols_dropped = cols_before - len(df.columns)

        # ===== STEP 3: Trim Whitespace =====
        if options.get('trim', True):
            for col in df.columns:
                if col in df.columns:
                    try:
                        df[col] = df[col].astype(str).str.strip()
                        df[col] = df[col].replace(['nan', 'None', 'NaN', 'NULL', 'null', 'NA', '#N/A', '#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#NUM!', '#NULL!'], np.nan)
                    except Exception:
                        pass

        # ===== STEP 4: Remove Duplicates =====
        if options.get('duplicates', True):
            before = len(df)
            df = df.drop_duplicates()
            duplicates_removed = before - len(df)

        # ===== STEP 5: Smart Type Conversion =====
        for col in df.columns:
            if col not in df.columns:
                continue
            try:
                # Try numeric conversion
                numeric_vals = pd.to_numeric(df[col], errors='coerce')
                non_null_ratio = numeric_vals.notna().sum() / max(len(df), 1)
                if non_null_ratio > 0.5:
                    df[col] = numeric_vals
                    continue
                
                # Try datetime conversion
                if df[col].dtype == 'object':
                    try:
                        date_vals = pd.to_datetime(df[col], errors='coerce', infer_datetime_format=True)
                        date_ratio = date_vals.notna().sum() / max(len(df), 1)
                        if date_ratio > 0.5:
                            df[col] = date_vals
                            continue
                    except Exception:
                        pass
            except Exception:
                pass

        # ===== STEP 6: Fill Missing Values =====
        if options.get('missing', True):
            missing_filled = int(df.isnull().sum().sum())
            for col in df.columns:
                if col not in df.columns:
                    continue
                try:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        df[col] = df[col].fillna(0)
                    else:
                        df[col] = df[col].fillna('N/A')
                except Exception:
                    pass

        # ===== STEP 7: Remove Special Characters =====
        if options.get('special', False):
            for col in df.columns:
                if col not in df.columns:
                    continue
                try:
                    if df[col].dtype == 'object':
                        df[col] = df[col].astype(str).str.replace(r'[^\w\s\.\,\-\/\(\)]', '', regex=True)
                except Exception:
                    pass

        # ===== STEP 8: Reset Index =====
        df = df.reset_index(drop=True)

    except Exception as e:
        st.warning(f"⚠️ Some cleaning steps had issues: {str(e)[:100]}")

    return df, {
        'original_rows': original_rows,
        'final_rows': len(df),
        'original_cols': original_cols,
        'final_cols': len(df.columns),
        'duplicates_removed': duplicates_removed,
        'missing_filled': missing_filled,
        'cols_dropped': cols_dropped
    }

# ==================== PDF READER ====================
def read_pdf_tables(uploaded_file):
    """Extract tables from PDF using multiple methods"""
    dfs = []
    errors = []

    try:
        file_bytes = uploaded_file.read()
        uploaded_file.seek(0)
        
        # Method 1: pdfplumber
        try:
            with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
                total_pages = len(pdf.pages)
                st.info(f"📄 PDF has {total_pages} pages. Extracting tables...")
                
                progress = st.progress(0)
                
                for i, page in enumerate(pdf.pages):
                    try:
                        tables = page.extract_tables()
                        for table in tables:
                            if table and len(table) > 1:
                                try:
                                    df = pd.DataFrame(table[1:], columns=table[0])
                                    if not df.empty and len(df.columns) > 1:
                                        dfs.append(df)
                                except Exception:
                                    df = pd.DataFrame(table)
                                    if not df.empty:
                                        dfs.append(df)
                        
                        # Also try text extraction as table
                        if not tables:
                            text = page.extract_text()
                            if text:
                                lines = [line.split() for line in text.split('\n') if line.strip()]
                                if lines and len(lines) > 2:
                                    max_cols = max(len(line) for line in lines)
                                    padded = [line + [''] * (max_cols - len(line)) for line in lines]
                                    df = pd.DataFrame(padded[1:], columns=[f'Col_{j+1}' for j in range(max_cols)])
                                    if not df.empty:
                                        dfs.append(df)
                        
                        progress.progress((i + 1) / total_pages)
                    except Exception as e:
                        errors.append(f"Page {i+1}: {str(e)[:50]}")
                
                progress.empty()
                
        except Exception as e:
            errors.append(f"pdfplumber: {str(e)[:100]}")

        # If no tables found, try text-based extraction
        if not dfs:
            try:
                reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
                all_text = []
                for page in reader.pages:
                    text = page.extract_text()
                    if text:
                        all_text.append(text)
                
                if all_text:
                    full_text = '\n'.join(all_text)
                    lines = full_text.split('\n')
                    cleaned_lines = []
                    for line in lines:
                        parts = line.split()
                        if len(parts) >= 2:
                            cleaned_lines.append(parts)
                    
                    if cleaned_lines:
                        max_cols = max(len(l) for l in cleaned_lines)
                        padded = [l + [''] * (max_cols - len(l)) for l in cleaned_lines]
                        df = pd.DataFrame(padded, columns=[f'Column_{i+1}' for i in range(max_cols)])
                        dfs.append(df)
                        
            except Exception as e:
                errors.append(f"PyPDF2: {str(e)[:100]}")
    
    except Exception as e:
        st.error(f"❌ Error reading PDF: {str(e)}")
        return None

    if errors:
        st.warning(f"⚠️ Some pages had issues: {len(errors)} errors (data still extracted from others)")

    if dfs:
        # Combine all dataframes
        if len(dfs) == 1:
            return dfs[0]
        
        # Try to merge with same columns
        try:
            common_cols = dfs[0].columns.tolist()
            matching_dfs = [dfs[0]]
            other_dfs = []
            
            for df in dfs[1:]:
                if list(df.columns) == common_cols or len(df.columns) == len(common_cols):
                    df.columns = common_cols
                    matching_dfs.append(df)
                else:
                    other_dfs.append(df)
            
            result = pd.concat(matching_dfs, ignore_index=True)
            return result
            
        except Exception:
            return pd.concat(dfs, ignore_index=True)
    
    return None

# ==================== FILE LOADER ====================
@st.cache_data
def load_excel_file(file_bytes, file_name):
    """Load Excel file"""
    try:
        ext = file_name.lower().split('.')[-1]
        df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl' if ext in ['xlsx', 'xlsm'] else 'xlrd')
        return df
    except Exception as e:
        try:
            df = pd.read_excel(io.BytesIO(file_bytes))
            return df
        except Exception as e2:
            st.error(f"❌ Excel error: {str(e2)}")
            return None

@st.cache_data
def load_csv_file(file_bytes, file_name):
    """Load CSV with auto-encoding detection"""
    encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252', 'utf-16']
    separators = [',', ';', '\t', '|']
    
    for encoding in encodings:
        for sep in separators:
            try:
                df = pd.read_csv(
                    io.BytesIO(file_bytes),
                    encoding=encoding,
                    sep=sep,
                    on_bad_lines='skip',
                    low_memory=False
                )
                if len(df.columns) > 1:
                    return df
            except Exception:
                continue
    
    try:
        df = pd.read_csv(io.BytesIO(file_bytes), on_bad_lines='skip')
        return df
    except Exception as e:
        st.error(f"❌ CSV error: {str(e)}")
        return None

def create_chart_safely(chart_type, df, x_axis, y_axis, color_by, numeric_cols, text_cols, top_n=10):
    """Create charts with comprehensive error handling"""
    try:
        if x_axis not in df.columns:
            return None
        if y_axis and y_axis not in df.columns:
            return None

        fig = None
        
        if chart_type == "Bar Chart":
            data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(top_n)
            fig = px.bar(data, x=x_axis, y=y_axis, color=y_axis, color_continuous_scale='Viridis', title=f"{y_axis} by {x_axis}")
        
        elif chart_type == "Horizontal Bar":
            data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(top_n)
            fig = px.bar(data, x=y_axis, y=x_axis, orientation='h', color=y_axis, color_continuous_scale='Blues', title=f"Top {top_n}")
        
        elif chart_type == "Pie Chart":
            data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(top_n)
            fig = px.pie(data, names=x_axis, values=y_axis, title=f"Distribution of {y_axis}")
        
        elif chart_type == "Donut Chart":
            data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(top_n)
            fig = px.pie(data, names=x_axis, values=y_axis, hole=0.5, title=f"Distribution of {y_axis}")
        
        elif chart_type == "Line Chart":
            fig = px.line(df.head(50), x=x_axis, y=y_axis, markers=True, title=f"{y_axis} Trend")
        
        elif chart_type == "Area Chart":
            fig = px.area(df.head(50), x=x_axis, y=y_axis, title=f"{y_axis} Area")
        
        elif chart_type == "Scatter Plot":
            fig = px.scatter(df, x=x_axis, y=y_axis, color=color_by if color_by and color_by in df.columns else None, title=f"{x_axis} vs {y_axis}")
        
        elif chart_type == "Bubble Chart":
            fig = px.scatter(df, x=x_axis, y=y_axis, size=y_axis, color=color_by if color_by and color_by in df.columns else None, title=f"Bubble: {y_axis}")
        
        elif chart_type == "Histogram":
            fig = px.histogram(df, x=y_axis, nbins=30, title=f"Distribution of {y_axis}")
        
        elif chart_type == "Box Plot":
            fig = px.box(df, y=y_axis, color=color_by if color_by and color_by in df.columns else None, title=f"Box Plot: {y_axis}")
        
        elif chart_type == "Violin Plot":
            fig = px.violin(df, y=y_axis, box=True, title=f"Violin: {y_axis}")
        
        elif chart_type == "Heatmap":
            nc = df.select_dtypes(include=[np.number]).columns.tolist()
            if len(nc) > 1:
                corr = df[nc].corr()
                fig = px.imshow(corr, text_auto=True, color_continuous_scale='RdBu_r', title="Correlation Heatmap")
            else:
                st.warning("Need 2+ numeric columns for heatmap")
                return None
        
        elif chart_type == "Treemap":
            data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(top_n)
            fig = px.treemap(data, path=[x_axis], values=y_axis, title=f"Treemap: {y_axis}")
        
        elif chart_type == "Sunburst":
            paths = text_cols[:2] if len(text_cols) >= 2 else [x_axis]
            fig = px.sunburst(df.head(30), path=paths, values=y_axis, title="Sunburst Chart")
        
        elif chart_type == "Funnel Chart":
            data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(top_n)
            fig = px.funnel(data, x=y_axis, y=x_axis, title=f"Funnel: {y_axis}")
        
        if fig:
            fig.update_layout(
                plot_bgcolor='rgba(0,0,0,0)',
                paper_bgcolor='rgba(0,0,0,0)',
                height=500,
                font=dict(size=11)
            )
        return fig
    
    except Exception as e:
        st.error(f"❌ Chart error: {str(e)[:100]}")
        return None

def create_formatted_excel(df):
    """Create Excel download"""
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Data', index=False)
            ws = writer.sheets['Data']
            for col in ws.columns:
                max_len = 0
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_len:
                            max_len = len(str(cell.value))
                    except Exception:
                        pass
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)
        output.seek(0)
        return output.getvalue()
    except Exception as e:
        st.error(f"Excel error: {e}")
        return None

def generate_sample_data():
    np.random.seed(42)
    return pd.DataFrame({
        'ITEM': [f'Product_{i}' for i in range(1, 51)],
        'CODE': [f'PRD-{i:03d}' for i in range(1, 51)],
        'CATEGORY': np.random.choice(['Raw Material', 'Finished Goods', 'Semi-Finished', 'WIP'], 50),
        'SUPPLIER': np.random.choice(['Supplier A', 'Supplier B', 'Supplier C', 'Supplier D'], 50),
        'SAFETY_STOCK': np.random.randint(100, 1000, 50),
        'OPENING_STOCK': np.random.randint(500, 5000, 50),
        'RECEIVED': np.random.randint(0, 2000, 50),
        'ISSUED': np.random.randint(0, 3000, 50),
        'CLOSING_STOCK': np.random.randint(0, 5000, 50),
        'PRICE': np.random.uniform(10, 500, 50).round(2),
        'STATUS': np.random.choice(['Active', 'Inactive', 'Discontinued'], 50)
    })

# ==================== SIDEBAR ====================
with st.sidebar:
    st.title("⚙️ Control Panel")
    st.markdown("---")
    
    st.subheader("🎨 Theme")
    theme = st.selectbox("Choose Theme", [
        "Professional Blue", "Dark Mode", "Ocean", "Sunset", "Forest", "Purple", "Mint", "Custom"
    ])
    
    color_themes = {
        "Professional Blue": ("#667eea", "#764ba2"),
        "Dark Mode": ("#1e1e1e", "#333333"),
        "Ocean": ("#0077be", "#00a8cc"),
        "Sunset": ("#ff6b6b", "#feca57"),
        "Forest": ("#2d5016", "#73a942"),
        "Purple": ("#9b59b6", "#8e44ad"),
        "Mint": ("#1abc9c", "#16a085"),
    }
    
    if theme == "Custom":
        primary_color = st.color_picker("Primary", "#667eea")
        secondary_color = st.color_picker("Secondary", "#764ba2")
    else:
        primary_color, secondary_color = color_themes.get(theme, ("#667eea", "#764ba2"))
    
    st.markdown("---")
    st.subheader("🔄 Auto Refresh")
    auto_refresh = st.checkbox("Enable Auto-Refresh", False)
    refresh_seconds = st.slider("Interval (seconds)", 5, 300, 30)
    
    st.markdown("---")
    st.subheader("📁 Data Source")
    data_source = st.radio("Choose Source", [
        "📄 Upload File (CSV/Excel/PDF)",
        "🔗 Google Sheets URL",
        "📊 Sample Data"
    ])
    
    st.markdown("---")
    st.subheader("🧹 Auto-Clean Options")
    remove_duplicates = st.checkbox("Remove Duplicates", True)
    fill_missing = st.checkbox("Fill Missing Values", True)
    trim_spaces = st.checkbox("Trim & Fix Spaces", True)
    standardize_case = st.checkbox("Standardize Columns", True)
    remove_special = st.checkbox("Remove Special Chars", False)
    
    st.markdown("---")
    st.subheader("🖥️ Display")
    show_kpis = st.checkbox("Show KPIs", True)
    show_alerts = st.checkbox("Show Alerts", True)
    show_charts = st.checkbox("Show Charts", True)
    show_pivot = st.checkbox("Show Pivot", True)
    show_stats = st.checkbox("Show Statistics", True)
    show_raw = st.checkbox("Show Raw Data", True)
    
    st.markdown("---")
    st.info(f"🔄 Refresh: {st.session_state.refresh_count}")

# ==================== CSS ====================
st.markdown(f"""
    <style>
    .stApp {{
        background: linear-gradient(135deg, {primary_color}10 0%, {secondary_color}10 100%);
    }}
    .main-header {{
        background: linear-gradient(135deg, {primary_color} 0%, {secondary_color} 100%);
        padding: 25px 30px;
        border-radius: 15px;
        color: white;
        text-align: center;
        margin-bottom: 25px;
        box-shadow: 0 10px 30px rgba(0,0,0,0.15);
    }}
    .kpi-card {{
        background: white;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid {primary_color};
        box-shadow: 0 5px 15px rgba(0,0,0,0.08);
        margin-bottom: 15px;
    }}
    .alert-critical {{
        background: linear-gradient(135deg, #ff6b6b, #ff4757);
        color: white;
        padding: 15px;
        border-radius: 10px;
        margin: 8px 0;
        font-weight: bold;
    }}
    .alert-warning {{
        background: linear-gradient(135deg, #feca57, #ff9f43);
        color: white;
        padding: 15px;
        border-radius: 10px;
        margin: 8px 0;
        font-weight: bold;
    }}
    .alert-success {{
        background: linear-gradient(135deg, #1abc9c, #16a085);
        color: white;
        padding: 15px;
        border-radius: 10px;
        margin: 8px 0;
        font-weight: bold;
    }}
    .info-box {{
        background: linear-gradient(135deg, {primary_color}20, {secondary_color}20);
        padding: 12px;
        border-radius: 8px;
        border: 1px solid {primary_color}40;
        margin: 8px 0;
    }}
    .stTabs [data-baseweb="tab"] {{
        background: white;
        border-radius: 8px;
        padding: 8px 16px;
        font-weight: 600;
        font-size: 13px;
    }}
    .stTabs [aria-selected="true"] {{
        background: linear-gradient(135deg, {primary_color}, {secondary_color}) !important;
        color: white !important;
    }}
    </style>
""", unsafe_allow_html=True)

# ==================== HEADER ====================
st.markdown("""
<div class="main-header">
    <h1 style="margin:0; font-size:28px;">📊 Ultimate Professional Dashboard</h1>
    <p style="margin:8px 0 0 0; opacity:0.9; font-size:14px;">
        Auto-Clean • PDF Support • Live Analytics • Multi-Format Export
    </p>
</div>
""", unsafe_allow_html=True)

# ==================== DATA LOADING ====================
df_loaded = None

clean_options = {
    'duplicates': remove_duplicates,
    'missing': fill_missing,
    'trim': trim_spaces,
    'standardize': standardize_case,
    'special': remove_special
}

if "📄 Upload File" in data_source:
    uploaded_file = st.file_uploader(
        "📁 Upload File (CSV, Excel, or PDF)",
        type=['csv', 'xlsx', 'xlsm', 'xls', 'pdf'],
        help="Supports CSV, Excel (.xlsx/.xls), and PDF files"
    )
    
    if uploaded_file:
        file_name = uploaded_file.name
        file_size = uploaded_file.size
        st.session_state.file_name = file_name
        
        st.markdown(f"""
        <div class="info-box">
            📄 <b>{file_name}</b> &nbsp;|&nbsp; 
            📦 Size: {file_size/1024:.1f} KB &nbsp;|&nbsp;
            📅 {datetime.now().strftime('%Y-%m-%d %H:%M')}
        </div>
        """, unsafe_allow_html=True)
        
        with st.spinner(f"🔄 Loading {file_name}..."):
            try:
                file_ext = file_name.lower().split('.')[-1]
                
                if file_ext == 'pdf':
                    st.info("📄 PDF detected - Extracting tables automatically...")
                    df_loaded = read_pdf_tables(uploaded_file)
                    if df_loaded is not None:
                        st.success(f"✅ PDF extracted! {len(df_loaded)} rows × {len(df_loaded.columns)} columns")
                    else:
                        st.error("❌ Could not extract tables from PDF. Try a different PDF.")
                
                elif file_ext == 'csv':
                    file_bytes = uploaded_file.read()
                    df_loaded = load_csv_file(file_bytes, file_name)
                    if df_loaded is not None:
                        st.success(f"✅ CSV loaded! {len(df_loaded)} rows × {len(df_loaded.columns)} columns")
                
                else:  # Excel
                    file_bytes = uploaded_file.read()
                    df_loaded = load_excel_file(file_bytes, file_name)
                    if df_loaded is not None:
                        st.success(f"✅ Excel loaded! {len(df_loaded)} rows × {len(df_loaded.columns)} columns")
                
                if df_loaded is not None:
                    st.session_state.data = df_loaded
                    
            except Exception as e:
                st.error(f"❌ Error loading file: {str(e)}")

elif "🔗 Google Sheets" in data_source:
    url = st.text_input(
        "📎 Paste Google Sheets or CSV URL",
        placeholder="https://docs.google.com/spreadsheets/..."
    )
    st.info("💡 For Google Sheets: File → Share → Anyone with link → Replace '/edit' with '/export?format=csv'")
    
    if url and st.button("🔄 Load from URL"):
        with st.spinner("Loading..."):
            try:
                if 'docs.google.com/spreadsheets' in url:
                    url = url.replace('/edit#gid=', '/export?format=csv&gid=')
                    url = url.replace('/edit?usp=sharing', '/export?format=csv')
                    url = url.replace('/edit', '/export?format=csv')
                
                df_loaded = pd.read_csv(url, on_bad_lines='skip')
                st.session_state.data = df_loaded
                st.success(f"✅ Loaded {len(df_loaded)} rows!")
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")

elif "📊 Sample Data" in data_source:
    if st.button("📊 Load Sample Data"):
        df_loaded = generate_sample_data()
        st.session_state.data = df_loaded
        st.success(f"✅ Sample data loaded! {len(df_loaded)} rows")

# ==================== MAIN DASHBOARD ====================
if st.session_state.data is not None:
    df_raw = st.session_state.data.copy()
    
    # AUTO CLEAN with error handling
    with st.spinner("🧹 Auto-cleaning data..."):
        try:
            df, clean_stats = ultra_clean_data(df_raw.copy(), clean_options)
            st.session_state.clean_data = df
        except Exception as e:
            st.warning(f"⚠️ Cleaning had issues, using original: {str(e)[:80]}")
            df = df_raw.copy()
            clean_stats = {
                'original_rows': len(df_raw),
                'final_rows': len(df),
                'original_cols': len(df_raw.columns),
                'final_cols': len(df.columns),
                'duplicates_removed': 0,
                'missing_filled': 0,
                'cols_dropped': 0
            }
    
    # Show cleaning report
    with st.expander("🧹 Data Cleaning Report", expanded=False):
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            st.metric("Original Rows", f"{clean_stats.get('original_rows', 0):,}")
        with c2:
            st.metric("Cleaned Rows", f"{clean_stats.get('final_rows', 0):,}")
        with c3:
            st.metric("Duplicates Removed", clean_stats.get('duplicates_removed', 0))
        with c4:
            st.metric("Missing Filled", f"{clean_stats.get('missing_filled', 0):,}")
        with c5:
            st.metric("Original Cols", clean_stats.get('original_cols', 0))
        with c6:
            st.metric("Cols Dropped", clean_stats.get('cols_dropped', 0))
        
        # Data quality score
        total = clean_stats.get('original_rows', 1)
        cleaned = clean_stats.get('final_rows', 0)
        quality = min(100, int((cleaned / max(total, 1)) * 100))
        st.progress(quality / 100)
        st.caption(f"Data Quality Score: {quality}%")
    
    # Safe column type detection
    try:
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        text_cols = df.select_dtypes(include=['object']).columns.tolist()
        date_cols = df.select_dtypes(include=['datetime64']).columns.tolist()
    except Exception:
        numeric_cols = []
        text_cols = df.columns.tolist()
        date_cols = []
    
    # ==================== TABS ====================
    tabs = st.tabs([
        "📈 KPIs",
        "🚨 Alerts",
        "📊 Charts",
        "🔄 Pivot",
        "📉 Statistics",
        "🔍 Filter",
        "📋 Raw Data",
        "📥 Export"
    ])
    
    # ===== TAB 1: KPIs =====
    with tabs[0]:
        if show_kpis:
            st.subheader("📈 Key Performance Indicators")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📦 Total Records", f"{len(df):,}")
            with col2:
                st.metric("📊 Total Columns", len(df.columns))
            with col3:
                st.metric("🔢 Numeric Cols", len(numeric_cols))
            with col4:
                st.metric("📝 Text Cols", len(text_cols))
            
            st.markdown("---")
            
            # Numeric KPIs
            if numeric_cols:
                st.subheader("🔢 Numeric Summary")
                num_tabs = st.tabs([f"📊 {col[:15]}" for col in numeric_cols[:8]])
                
                for i, nt in enumerate(num_tabs):
                    if i < len(numeric_cols):
                        col_name = numeric_cols[i]
                        with nt:
                            try:
                                c1, c2, c3, c4, c5 = st.columns(5)
                                with c1:
                                    st.metric("Sum", f"{df[col_name].sum():,.2f}")
                                with c2:
                                    st.metric("Avg", f"{df[col_name].mean():,.2f}")
                                with c3:
                                    st.metric("Max", f"{df[col_name].max():,.2f}")
                                with c4:
                                    st.metric("Min", f"{df[col_name].min():,.2f}")
                                with c5:
                                    st.metric("Median", f"{df[col_name].median():,.2f}")
                                
                                fig = px.histogram(df, x=col_name, nbins=20, title=f"Distribution: {col_name}")
                                st.plotly_chart(fig, use_container_width=True, height=250)
                            except Exception:
                                st.warning(f"Cannot display KPI for {col_name}")
            
            # Text column analysis
            if text_cols:
                st.markdown("---")
                st.subheader("📝 Category Analysis")
                for col in text_cols[:3]:
                    try:
                        with st.expander(f"📊 {col} - Distribution"):
                            vc = df[col].value_counts().head(10)
                            c1, c2 = st.columns([1, 2])
                            with c1:
                                st.dataframe(vc, use_container_width=True)
                            with c2:
                                fig = px.bar(x=vc.index, y=vc.values, title=f"Top 10: {col}",
                                           labels={'x': col, 'y': 'Count'})
                                st.plotly_chart(fig, use_container_width=True, height=300)
                    except Exception:
                        pass
    
    # ===== TAB 2: ALERTS =====
    with tabs[1]:
        if show_alerts:
            st.subheader("🚨 Smart Alert System")
            alerts_found = False
            
            # Check numeric columns for issues
            for col in numeric_cols:
                try:
                    zero_count = int((df[col] == 0).sum())
                    zero_pct = zero_count / max(len(df), 1) * 100
                    
                    if zero_pct > 50:
                        alerts_found = True
                        st.markdown(f"""
                        <div class="alert-warning">
                            ⚠️ HIGH ZERO VALUES: <b>{col}</b> has {zero_pct:.1f}% zero values ({zero_count} rows)
                        </div>
                        """, unsafe_allow_html=True)
                    
                    neg_count = int((df[col] < 0).sum())
                    if neg_count > 0:
                        alerts_found = True
                        st.markdown(f"""
                        <div class="alert-warning">
                            ⚠️ NEGATIVE VALUES: <b>{col}</b> has {neg_count} negative values
                        </div>
                        """, unsafe_allow_html=True)
                except Exception:
                    pass
            
            # Inventory-specific alerts
            try:
                if 'CLOSING_STOCK' in df.columns and 'SAFETY_STOCK' in df.columns:
                    critical = df[df['CLOSING_STOCK'] < df['SAFETY_STOCK']]
                    if len(critical) > 0:
                        alerts_found = True
                        st.markdown(f"""
                        <div class="alert-critical">
                            🔴 CRITICAL: {len(critical)} items below safety stock!
                        </div>
                        """, unsafe_allow_html=True)
                        st.dataframe(critical.head(20), use_container_width=True)
                
                if 'CLOSING_STOCK' in df.columns:
                    out_stock = df[df['CLOSING_STOCK'] <= 0]
                    if len(out_stock) > 0:
                        alerts_found = True
                        st.markdown(f"""
                        <div class="alert-critical">
                            🔴 OUT OF STOCK: {len(out_stock)} items!
                        </div>
                        """, unsafe_allow_html=True)
                        st.dataframe(out_stock.head(10), use_container_width=True)
                
                if 'ISSUED' in df.columns and 'RECEIVED' in df.columns:
                    dead = df[(df['ISSUED'] == 0) & (df['RECEIVED'] == 0)]
                    if len(dead) > 0:
                        alerts_found = True
                        st.markdown(f"""
                        <div class="alert-warning">
                            ⚠️ DEAD STOCK: {len(dead)} items with no movement
                        </div>
                        """, unsafe_allow_html=True)
                        st.dataframe(dead.head(10), use_container_width=True)
            except Exception:
                pass
            
            # Missing data alert
            try:
                missing_pct = (df.isnull().sum() / len(df) * 100)
                high_missing = missing_pct[missing_pct > 20]
                if len(high_missing) > 0:
                    alerts_found = True
                    st.markdown(f"""
                    <div class="alert-warning">
                        ⚠️ HIGH MISSING DATA: {len(high_missing)} columns have 20%+ missing values
                    </div>
                    """, unsafe_allow_html=True)
                    st.dataframe(high_missing.to_frame('Missing %').round(1), use_container_width=True)
            except Exception:
                pass
            
            if not alerts_found:
                st.markdown("""
                <div class="alert-success">
                    ✅ All Clear! No critical issues found in your data.
                </div>
                """, unsafe_allow_html=True)
    
    # ===== TAB 3: CHARTS =====
    with tabs[2]:
        if show_charts:
            st.subheader("📊 Advanced Visualizations")
            
            col1, col2 = st.columns([1, 3])
            
            with col1:
                chart_type = st.selectbox("Chart Type", [
                    "Bar Chart", "Horizontal Bar", "Pie Chart", "Donut Chart",
                    "Line Chart", "Area Chart", "Scatter Plot", "Bubble Chart",
                    "Histogram", "Box Plot", "Violin Plot", "Heatmap",
                    "Treemap", "Sunburst", "Funnel Chart"
                ])
                
                all_cols = df.columns.tolist()
                x_axis = st.selectbox("X-Axis", all_cols, key="chart_x")
                y_axis = st.selectbox("Y-Axis", numeric_cols if numeric_cols else all_cols, key="chart_y")
                color_by = st.selectbox("Color By", [None] + all_cols, key="chart_c")
                top_n = st.slider("Top N items", 5, 50, 15)
            
            with col2:
                if x_axis and y_axis:
                    fig = create_chart_safely(chart_type, df, x_axis, y_axis, color_by, numeric_cols, text_cols, top_n)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("Select X and Y axes to generate chart")
            
            st.markdown("---")
            st.subheader("📊 Auto-Generated Charts")
            
            if numeric_cols and text_cols:
                cc1, cc2, cc3 = st.columns(3)
                
                with cc1:
                    try:
                        f1 = create_chart_safely("Bar Chart", df, text_cols[0], numeric_cols[0], None, numeric_cols, text_cols, 10)
                        if f1:
                            st.plotly_chart(f1, use_container_width=True)
                    except Exception:
                        pass
                
                with cc2:
                    try:
                        f2 = create_chart_safely("Pie Chart", df, text_cols[0], numeric_cols[0], None, numeric_cols, text_cols, 8)
                        if f2:
                            st.plotly_chart(f2, use_container_width=True)
                    except Exception:
                        pass
                
                with cc3:
                    try:
                        if len(numeric_cols) >= 2:
                            f3 = create_chart_safely("Scatter Plot", df, numeric_cols[0], numeric_cols[1], None, numeric_cols, text_cols)
                            if f3:
                                st.plotly_chart(f3, use_container_width=True)
                    except Exception:
                        pass
                
                if len(numeric_cols) > 1:
                    try:
                        cc1, cc2 = st.columns(2)
                        with cc1:
                            f4 = create_chart_safely("Histogram", df, numeric_cols[0], numeric_cols[0], None, numeric_cols, text_cols)
                            if f4:
                                st.plotly_chart(f4, use_container_width=True)
                        with cc2:
                            f5 = create_chart_safely("Heatmap", df, None, numeric_cols[0], None, numeric_cols, text_cols)
                            if f5:
                                st.plotly_chart(f5, use_container_width=True)
                    except Exception:
                        pass
    
    # ===== TAB 4: PIVOT =====
    with tabs[3]:
        if show_pivot:
            st.subheader("🔄 Pivot Table Analysis")
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                pivot_rows = st.multiselect("Rows", text_cols, default=text_cols[:1] if text_cols else [], key="p_rows")
            with c2:
                pivot_cols_sel = st.multiselect("Columns", text_cols, key="p_cols")
            with c3:
                pivot_vals = st.multiselect("Values", numeric_cols, default=numeric_cols[:1] if numeric_cols else [], key="p_vals")
            with c4:
                pivot_agg = st.selectbox("Aggregation", ["sum", "mean", "count", "max", "min", "median"])
            
            if pivot_rows and pivot_vals:
                try:
                    pt = pd.pivot_table(
                        df,
                        index=pivot_rows,
                        columns=pivot_cols_sel if pivot_cols_sel else None,
                        values=pivot_vals,
                        aggfunc=pivot_agg,
                        fill_value=0,
                        margins=True,
                        margins_name='TOTAL'
                    )
                    
                    st.dataframe(pt, use_container_width=True, height=400)
                    
                    pivot_csv = pt.to_csv()
                    st.download_button("📥 Download Pivot", pivot_csv,
                                      f"pivot_{datetime.now():%Y%m%d_%H%M%S}.csv")
                    
                    try:
                        pt_reset = pt.drop('TOTAL', errors='ignore').reset_index()
                        if pivot_rows[0] in pt_reset.columns:
                            vc1, vc2 = st.columns(2)
                            y_col = pivot_vals[0] if pivot_vals[0] in pt_reset.columns else pt_reset.columns[-1]
                            with vc1:
                                fig = px.bar(pt_reset.head(20), x=pivot_rows[0], y=y_col, title="Pivot Bar")
                                st.plotly_chart(fig, use_container_width=True)
                            with vc2:
                                fig = px.line(pt_reset.head(20), x=pivot_rows[0], y=y_col, title="Pivot Trend", markers=True)
                                st.plotly_chart(fig, use_container_width=True)
                    except Exception:
                        pass
                        
                except Exception as e:
                    st.error(f"❌ Pivot error: {str(e)[:100]}")
            else:
                st.info("👆 Select Rows and Values to create pivot table")
    
    # ===== TAB 5: STATISTICS =====
    with tabs[4]:
        if show_stats:
            st.subheader("📉 Statistical Analysis")
            
            try:
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown("### 📊 Descriptive Statistics")
                    st.dataframe(df.describe(include='all'), use_container_width=True)
                with c2:
                    st.markdown("### 📈 Data Types")
                    dt_df = pd.DataFrame({
                        'Column': df.columns,
                        'Type': df.dtypes.astype(str),
                        'Non-Null': df.count().values,
                        'Null%': (df.isnull().sum() / len(df) * 100).round(1).values,
                        'Unique': [df[c].nunique() for c in df.columns]
                    })
                    st.dataframe(dt_df, use_container_width=True, hide_index=True)
            except Exception:
                pass
            
            if numeric_cols:
                st.markdown("---")
                sc1, sc2 = st.columns(2)
                
                with sc1:
                    st.markdown("### 🔢 Column Deep Dive")
                    try:
                        selected_col = st.selectbox("Select Column", numeric_cols, key="stat_col")
                        
                        stats_data = {
                            'Metric': ['Count', 'Sum', 'Mean', 'Median', 'Std Dev', 'Variance',
                                      'Min', 'Q1', 'Q3', 'Max', 'Range', 'IQR', 'Skewness', 'Kurtosis'],
                            'Value': [
                                f"{df[selected_col].count():,}",
                                f"{df[selected_col].sum():,.2f}",
                                f"{df[selected_col].mean():,.2f}",
                                f"{df[selected_col].median():,.2f}",
                                f"{df[selected_col].std():,.2f}",
                                f"{df[selected_col].var():,.2f}",
                                f"{df[selected_col].min():,.2f}",
                                f"{df[selected_col].quantile(0.25):,.2f}",
                                f"{df[selected_col].quantile(0.75):,.2f}",
                                f"{df[selected_col].max():,.2f}",
                                f"{df[selected_col].max() - df[selected_col].min():,.2f}",
                                f"{df[selected_col].quantile(0.75) - df[selected_col].quantile(0.25):,.2f}",
                                f"{df[selected_col].skew():,.3f}",
                                f"{df[selected_col].kurtosis():,.3f}"
                            ]
                        }
                        st.dataframe(pd.DataFrame(stats_data), use_container_width=True, hide_index=True)
                    except Exception as e:
                        st.warning(f"Stats error: {str(e)[:50]}")
                
                with sc2:
                    st.markdown("### 📈 Distribution")
                    try:
                        fig = px.histogram(df, x=selected_col, nbins=30, marginal="box",
                                         title=f"Distribution: {selected_col}")
                        st.plotly_chart(fig, use_container_width=True)
                    except Exception:
                        pass
                
                if len(numeric_cols) > 1:
                    st.markdown("---")
                    st.markdown("### 🔗 Correlation Matrix")
                    try:
                        corr = df[numeric_cols].corr()
                        fig = px.imshow(corr, text_auto=True, aspect="auto",
                                       color_continuous_scale='RdBu_r', title="Correlation Heatmap")
                        st.plotly_chart(fig, use_container_width=True)
                    except Exception:
                        pass
    
    # ===== TAB 6: FILTER =====
    with tabs[5]:
        st.subheader("🔍 Advanced Filter & Search")
        
        search = st.text_input("🔍 Global Search", placeholder="Search across all columns...")
        
        st.markdown("### 🎯 Column Filters")
        num_filters = st.slider("Number of filters", 1, 5, 2)
        
        filters = {}
        filter_cols = st.columns(2)
        
        for i in range(num_filters):
            with filter_cols[i % 2]:
                st.markdown(f"**Filter {i+1}**")
                fc1, fc2 = st.columns(2)
                with fc1:
                    filter_col = st.selectbox(f"Column", df.columns.tolist(), key=f"fc_{i}")
                with fc2:
                    try:
                        if filter_col in numeric_cols:
                            min_v = float(df[filter_col].min())
                            max_v = float(df[filter_col].max())
                            if min_v != max_v:
                                filters[filter_col] = st.slider(f"Range", min_v, max_v, (min_v, max_v), key=f"fr_{i}")
                        else:
                            uv = sorted(df[filter_col].astype(str).unique().tolist())
                            dv = uv[:5] if len(uv) > 5 else uv
                            filters[filter_col] = st.multiselect("Values", uv, default=dv, key=f"fm_{i}")
                    except Exception:
                        pass
        
        # Apply filters
        filtered_df = df.copy()
        
        for col, val in filters.items():
            try:
                if col in filtered_df.columns:
                    if col in numeric_cols and isinstance(val, (list, tuple)) and len(val) == 2:
                        filtered_df = filtered_df[
                            (filtered_df[col] >= val[0]) & (filtered_df[col] <= val[1])
                        ]
                    elif isinstance(val, list) and val:
                        filtered_df = filtered_df[filtered_df[col].astype(str).isin(val)]
            except Exception:
                pass
        
        if search:
            try:
                mask = filtered_df.astype(str).apply(
                    lambda x: x.str.contains(search, case=False, na=False)
                ).any(axis=1)
                filtered_df = filtered_df[mask]
            except Exception:
                pass
        
        st.markdown(f"### 📋 Results: {len(filtered_df):,} rows ({len(filtered_df)/max(len(df),1)*100:.1f}%)")
        
        col1, col2, col3 = st.columns([2, 1, 1])
        with col2:
            st.info(f"Original: {len(df):,}")
        with col3:
            st.info(f"Filtered: {len(filtered_df):,}")
        
        st.dataframe(filtered_df, use_container_width=True, height=400)
        
        fc1, fc2, fc3 = st.columns(3)
        with fc1:
            st.download_button("📄 CSV", filtered_df.to_csv(index=False),
                              f"filtered_{datetime.now():%Y%m%d_%H%M%S}.csv")
        with fc2:
            st.download_button("📋 JSON", filtered_df.to_json(orient='records', indent=2),
                              f"filtered_{datetime.now():%Y%m%d_%H%M%S}.json")
        with fc3:
            excel_bytes = create_formatted_excel(filtered_df)
            if excel_bytes:
                st.download_button("📊 Excel", excel_bytes,
                                  f"filtered_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    
    # ===== TAB 7: RAW DATA =====
    with tabs[6]:
        if show_raw:
            st.subheader("📋 Complete Data View")
            
            c1, c2, c3, c4 = st.columns(4)
            with c1:
                st.metric("Rows", f"{len(df):,}")
            with c2:
                st.metric("Columns", len(df.columns))
            with c3:
                st.metric("Missing", int(df.isnull().sum().sum()))
            with c4:
                st.metric("Duplicates", int(df.duplicated().sum()))
            
            search_raw = st.text_input("🔍 Search in data", key="raw_search")
            
            display_df = df.copy()
            if search_raw:
                try:
                    mask = df.astype(str).apply(
                        lambda x: x.str.contains(search_raw, case=False, na=False)
                    ).any(axis=1)
                    display_df = df[mask]
                except Exception:
                    pass
            
            st.dataframe(display_df, use_container_width=True, height=600)
            
            st.markdown("### 📊 Column Information")
            try:
                info_df = pd.DataFrame({
                    'Column': df.columns,
                    'Type': df.dtypes.astype(str),
                    'Non-Null Count': df.count().values,
                    'Null Count': df.isnull().sum().values,
                    'Null %': (df.isnull().sum() / len(df) * 100).round(2).values,
                    'Unique Values': [df[col].nunique() for col in df.columns],
                    'Sample Value': [str(df[col].dropna().iloc[0]) if not df[col].dropna().empty else 'N/A' for col in df.columns]
                })
                st.dataframe(info_df, use_container_width=True, hide_index=True)
            except Exception:
                pass
    
    # ===== TAB 8: EXPORT =====
    with tabs[7]:
        st.subheader("📥 Export Options")
        
        st.markdown("### 📊 Export Full Dataset")
        c1, c2, c3, c4 = st.columns(4)
        
        with c1:
            st.download_button(
                "📄 Download CSV",
                df.to_csv(index=False),
                f"data_{datetime.now():%Y%m%d_%H%M%S}.csv",
                "text/csv",
                use_container_width=True
            )
        
        with c2:
            st.download_button(
                "📋 Download JSON",
                df.to_json(orient='records', indent=2),
                f"data_{datetime.now():%Y%m%d_%H%M%S}.json",
                "application/json",
                use_container_width=True
            )
        
        with c3:
            excel_data = create_formatted_excel(df)
            if excel_data:
                st.download_button(
                    "📊 Download Excel",
                    excel_data,
                    f"data_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        with c4:
            try:
                parquet_buf = io.BytesIO()
                df.to_parquet(parquet_buf, index=False)
                st.download_button(
                    "⚡ Download Parquet",
                    parquet_buf.getvalue(),
                    f"data_{datetime.now():%Y%m%d_%H%M%S}.parquet",
                    use_container_width=True
                )
            except Exception:
                pass
        
        st.markdown("---")
        st.markdown("### 📋 Dataset Summary")
        c1, c2 = st.columns(2)
        with c1:
            st.markdown(f"""
            **📊 Dataset Info**
            - **File:** {st.session_state.file_name or 'Unknown'}
            - **Rows:** {len(df):,}
            - **Columns:** {len(df.columns)}
            - **Numeric Cols:** {len(numeric_cols)}
            - **Text Cols:** {len(text_cols)}
            - **Missing Values:** {int(df.isnull().sum().sum()):,}
            - **Memory:** {df.memory_usage(deep=True).sum() / 1024:.2f} KB
            - **Export Time:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            """)
        with c2:
            st.markdown("**📊 Column Types**")
            type_counts = df.dtypes.value_counts()
            for dtype, count in type_counts.items():
                st.markdown(f"- **{dtype}:** {count} columns")
    
    # ==================== AUTO REFRESH ====================
    if auto_refresh:
        st.session_state.refresh_count += 1
        st.info(f"🔄 Auto-refreshing in {refresh_seconds} seconds...")
        time.sleep(refresh_seconds)
        st.rerun()

else:
    # Landing Page
    st.markdown("""
    <div class="main-header" style="margin-top:20px;">
        <h2>👋 Welcome! Upload your data to get started</h2>
        <p>Supports CSV • Excel • PDF • Google Sheets</p>
    </div>
    """, unsafe_allow_html=True)
    
    c1, c2, c3 = st.columns(3)
    
    with c1:
        st.markdown("""
        <div class="kpi-card">
            <h3>🚀 Features</h3>
            <ul>
                <li>✅ PDF Table Extraction</li>
                <li>✅ Auto-Clean Any Data</li>
                <li>✅ 15+ Chart Types</li>
                <li>✅ Smart Alerts</li>
                <li>✅ Pivot Tables</li>
                <li>✅ Multi-Format Export</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with c2:
        st.markdown("""
        <div class="kpi-card">
            <h3>📁 Supported Files</h3>
            <ul>
                <li>📄 PDF (auto extract tables)</li>
                <li>📊 Excel (.xlsx, .xls)</li>
                <li>📝 CSV (any encoding)</li>
                <li>🔗 Google Sheets URL</li>
                <li>📋 Any messy data</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
    
    with c3:
        st.markdown("""
        <div class="kpi-card">
            <h3>🧹 Auto-Clean</h3>
            <ul>
                <li>✅ Fix #N/A, #REF! errors</li>
                <li>✅ Remove duplicates</li>
                <li>✅ Fix column names</li>
                <li>✅ Auto type detection</li>
                <li>✅ Fill missing values</li>
                <li>✅ Handle any encoding</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
