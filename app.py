Here is the **complete `app.py`** — copy everything and replace your entire file:

```python
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import re
from datetime import datetime
import numpy as np
import time
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import warnings

# PDF library (safe import)
try:
    import pdfplumber
except ImportError:
    pdfplumber = None

warnings.filterwarnings('ignore')

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
if 'theme_color' not in st.session_state:
    st.session_state.theme_color = "#667eea"

# ==================== HELPER FUNCTIONS ====================

def make_columns_unique(df):
    """Ensure all column names are unique - fixes pyarrow errors"""
    if df is None or df.empty:
        return df
    cols = pd.Series([str(c) for c in df.columns])
    if cols.duplicated().any():
        for dup in cols[cols.duplicated()].unique():
            count = 0
            for i, col in enumerate(cols):
                if col == dup:
                    if count > 0:
                        cols.iloc[i] = f"{col}_{count}"
                    count += 1
    df.columns = cols
    return df

@st.cache_data
def load_csv(file):
    """Load CSV file safely"""
    try:
        df = pd.read_csv(file, on_bad_lines='skip')
        if df.empty:
            st.error("❌ File is empty!")
            return None
        df = make_columns_unique(df)
        return df
    except Exception as e:
        st.error(f"❌ Error reading CSV: {str(e)}")
        return None

@st.cache_data
def load_excel(file):
    """Load Excel file safely"""
    try:
        df = pd.read_excel(file)
        if df.empty:
            st.error("❌ File is empty!")
            return None
        df = make_columns_unique(df)
        return df
    except Exception as e:
        st.error(f"❌ Error reading Excel: {str(e)}")
        return None

@st.cache_data
def load_url_data(url):
    """Load data from URL (CSV or Google Sheets)"""
    try:
        if 'docs.google.com/spreadsheets' in url:
            if '/edit' in url:
                url = url.replace('/edit#gid=', '/export?format=csv&gid=')
                url = url.replace('/edit?usp=sharing', '/export?format=csv')
                url = url.replace('/edit', '/export?format=csv')

        df = pd.read_csv(url)
        if df.empty:
            st.error("❌ URL data is empty!")
            return None
        df = make_columns_unique(df)
        return df
    except Exception as e:
        st.error(f"❌ Error loading URL: {str(e)}")
        return None

def extract_tables_from_pdf(pdf_bytes):
    """Extract all tables from PDF"""
    tables_data = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                page_tables = page.extract_tables()
                if page_tables:
                    for table_num, table in enumerate(page_tables, start=1):
                        if table and len(table) > 1:
                            cleaned = []
                            for row in table:
                                if row is None:
                                    continue
                                new_row = [str(c).strip() if c is not None else "" for c in row]
                                if any(cell != "" for cell in new_row):
                                    cleaned.append(new_row)

                            if len(cleaned) < 2:
                                continue

                            max_cols = max(len(r) for r in cleaned)
                            cleaned = [r + [""] * (max_cols - len(r)) for r in cleaned]

                            header = cleaned[0]
                            if not any(str(h).strip() for h in header):
                                header = [f"COLUMN_{i+1}" for i in range(max_cols)]

                            temp_df = pd.DataFrame(cleaned[1:], columns=header)
                            temp_df.insert(0, 'PDF_PAGE', page_num)
                            temp_df.insert(1, 'PDF_TABLE', table_num)
                            temp_df = make_columns_unique(temp_df)
                            tables_data.append(temp_df)
        return tables_data
    except Exception as e:
        st.error(f"❌ Error extracting tables: {str(e)}")
        return []

def extract_text_from_pdf(pdf_bytes):
    """Extract text line by line from PDF"""
    text_rows = []
    try:
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page_num, page in enumerate(pdf.pages, start=1):
                text = page.extract_text()
                if text:
                    for line_no, line in enumerate(text.splitlines(), start=1):
                        line = line.strip()
                        if line:
                            text_rows.append({
                                'PDF_PAGE': page_num,
                                'LINE_NO': line_no,
                                'TEXT': line
                            })
        return text_rows
    except Exception as e:
        st.error(f"❌ Error extracting text: {str(e)}")
        return []

def load_pdf(pdf_bytes, mode="Auto"):
    """Load PDF and return DataFrame"""
    if pdfplumber is None:
        st.error("❌ pdfplumber is not installed. Add 'pdfplumber' to requirements.txt")
        return None

    df = None

    if mode in ("Tables", "Auto"):
        tables = extract_tables_from_pdf(pdf_bytes)
        if tables:
            df = pd.concat(tables, ignore_index=True, sort=False)

    if df is None and mode in ("Text", "Auto"):
        text_rows = extract_text_from_pdf(pdf_bytes)
        if text_rows:
            df = pd.DataFrame(text_rows)

    if mode == "Text" and df is None:
        text_rows = extract_text_from_pdf(pdf_bytes)
        if text_rows:
            df = pd.DataFrame(text_rows)

    if df is None or df.empty:
        st.error("❌ Could not extract any data from this PDF")
        return None

    df = make_columns_unique(df)
    return df

def generate_sample_data():
    """Generate sample inventory data"""
    np.random.seed(42)
    sample_data = {
        'ITEM': [f'Product_{i}' for i in range(1, 51)],
        'CODE': [f'PRD-{i:03d}' for i in range(1, 51)],
        'CATEGORY': np.random.choice(['Raw Material', 'Finished Goods', 'Semi-Finished', 'Work in Progress'], 50),
        'SUPPLIER': np.random.choice(['Supplier A', 'Supplier B', 'Supplier C', 'Supplier D'], 50),
        'SAFETY_STOCK': np.random.randint(100, 1000, 50),
        'OPENING_STOCK': np.random.randint(500, 5000, 50),
        'RECEIVED': np.random.randint(0, 2000, 50),
        'ISSUED': np.random.randint(0, 3000, 50),
        'CLOSING_STOCK': np.random.randint(0, 5000, 50),
        'PRICE': np.random.uniform(10, 500, 50).round(2),
        'STATUS': np.random.choice(['Active', 'Inactive', 'Discontinued'], 50)
    }
    return pd.DataFrame(sample_data)

def validate_data(df):
    """Validate data quality"""
    issues = []

    if df is None or df.empty:
        issues.append("⚠️ Dataset is empty")
        return issues
    if len(df) < 2:
        issues.append("⚠️ Less than 2 rows of data")
    if df.isnull().all().any():
        null_cols = df.columns[df.isnull().all()].tolist()
        issues.append(f"⚠️ Columns entirely empty: {', '.join(null_cols)}")

    return issues

def clean_data(df, options):
    """Enhanced data cleaning with error handling"""
    if df is None or df.empty:
        return df, {'original_rows': 0, 'final_rows': 0, 'duplicates_removed': 0, 'missing_filled': 0}

    original_rows = len(df)
    duplicates_removed = 0
    missing_filled = 0

    try:
        # Step 1: Standardize columns FIRST
        if options.get('standardize', True):
            df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')

        # Make sure columns are unique after standardization
        df = make_columns_unique(df)

        # Step 2: Get column types AFTER standardization
        object_cols = df.select_dtypes(include=['object']).columns.tolist()
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()

        # Step 3: Trim spaces safely
        if options.get('trim', True) and len(object_cols) > 0:
            for col in object_cols:
                try:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.strip()
                except Exception:
                    pass

        # Step 4: Remove duplicates
        if options.get('duplicates', True):
            before = len(df)
            df = df.drop_duplicates()
            duplicates_removed = before - len(df)

        # Step 5: Fill missing values
        if options.get('missing', True):
            missing_filled = int(df.isnull().sum().sum())

            for col in numeric_cols:
                if col in df.columns:
                    df[col] = df[col].fillna(0)

            for col in object_cols:
                if col in df.columns:
                    df[col] = df[col].fillna('N/A')

        # Step 6: Remove special characters
        if options.get('special', False):
            for col in object_cols:
                try:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.replace(r'[^\w\s]', '', regex=True)
                except Exception:
                    pass

        # Step 7: Auto convert to numeric where possible
        for col in df.columns:
            try:
                if df[col].dtype == 'object':
                    converted = pd.to_numeric(df[col], errors='coerce')
                    if len(df) > 0 and converted.notna().sum() / len(df) > 0.8:
                        df[col] = converted.fillna(0)
            except Exception:
                pass

    except Exception as e:
        st.error(f"❌ Error during cleaning: {str(e)}")

    return df, {
        'original_rows': original_rows,
        'final_rows': len(df),
        'duplicates_removed': duplicates_removed,
        'missing_filled': missing_filled
    }

def create_chart_safely(chart_type, df, x_axis, y_axis, color_by, numeric_cols, text_cols, top_n=10):
    """Create charts with error handling"""
    try:
        if df is None or df.empty:
            return None
        if x_axis not in df.columns or y_axis not in df.columns:
            st.warning("⚠️ Selected columns not found in data")
            return None

        if chart_type == "Bar Chart":
            data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(top_n)
            fig = px.bar(data, x=x_axis, y=y_axis, color=color_by or y_axis,
                        color_continuous_scale='Viridis', title=f"{y_axis} by {x_axis}")

        elif chart_type == "Horizontal Bar":
            data = df.nlargest(top_n, y_axis)
            fig = px.bar(data, x=y_axis, y=x_axis, orientation='h', color=color_by or y_axis,
                        color_continuous_scale='Viridis', title=f"Top {top_n} - {y_axis}")

        elif chart_type == "Pie Chart":
            data = df.nlargest(top_n, y_axis)
            fig = px.pie(data, names=x_axis, values=y_axis, title=f"Distribution of {y_axis}")

        elif chart_type == "Donut Chart":
            data = df.nlargest(top_n, y_axis)
            fig = px.pie(data, names=x_axis, values=y_axis, hole=0.5,
                        title=f"Distribution of {y_axis}")

        elif chart_type == "Line Chart":
            fig = px.line(df.head(50), x=x_axis, y=y_axis, color=color_by,
                         markers=True, title=f"{y_axis} Trend")

        elif chart_type == "Area Chart":
            fig = px.area(df.head(50), x=x_axis, y=y_axis, color=color_by,
                         title=f"{y_axis} Area Chart")

        elif chart_type == "Scatter Plot":
            fig = px.scatter(df, x=x_axis, y=y_axis, color=color_by, size=y_axis,
                           title=f"{x_axis} vs {y_axis}")

        elif chart_type == "Bubble Chart":
            fig = px.scatter(df, x=x_axis, y=y_axis, size=y_axis, color=color_by,
                           title=f"Bubble Chart: {x_axis} vs {y_axis}")

        elif chart_type == "Histogram":
            fig = px.histogram(df, x=y_axis, nbins=30, color=color_by,
                             title=f"Distribution of {y_axis}")

        elif chart_type == "Box Plot":
            fig = px.box(df, y=y_axis, color=color_by,
                        title=f"Box Plot of {y_axis}")

        elif chart_type == "Violin Plot":
            fig = px.violin(df, y=y_axis, color=color_by, box=True,
                           title=f"Violin Plot of {y_axis}")

        elif chart_type == "Heatmap":
            if len(numeric_cols) > 1:
                corr = df[numeric_cols].corr()
                fig = px.imshow(corr, text_auto=True, aspect='auto',
                              color_continuous_scale='RdBu_r', title="Correlation Heatmap")
            else:
                st.warning("Need at least 2 numeric columns for heatmap")
                return None

        elif chart_type == "Treemap":
            data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(top_n)
            fig = px.treemap(data, path=[x_axis], values=y_axis,
                           title=f"Treemap of {y_axis} by {x_axis}")

        elif chart_type == "Sunburst":
            if len(text_cols) >= 2:
                fig = px.sunburst(df.head(30), path=text_cols[:2], values=y_axis,
                                 title="Sunburst Chart")
            else:
                fig = px.sunburst(df.head(30), path=[x_axis], values=y_axis,
                                 title="Sunburst Chart")

        elif chart_type == "Funnel Chart":
            data = df.nlargest(top_n, y_axis)
            fig = px.funnel(data, x=y_axis, y=x_axis,
                          title=f"Funnel Chart: {y_axis}")

        elif chart_type == "Waterfall":
            data = df.nlargest(10, y_axis)
            fig = px.bar(data, x=x_axis, y=y_axis,
                        title=f"Waterfall: {y_axis}")

        else:
            st.warning(f"Chart type '{chart_type}' not implemented")
            return None

        fig.update_layout(
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            height=500,
            showlegend=True,
            font=dict(size=11)
        )
        return fig

    except Exception as e:
        st.error(f"❌ Chart error: {str(e)}")
        return None

def create_excel_with_formatting(df, filename):
    """Create formatted Excel file"""
    try:
        if df is None or df.empty:
            return None
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Data', index=False)

            workbook = writer.book
            worksheet = writer.sheets['Data']

            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width

        output.seek(0)
        return output.getvalue()
    except Exception as e:
        st.error(f"❌ Error creating Excel: {str(e)}")
        return None

# ==================== SIDEBAR ====================
with st.sidebar:
    st.title("⚙️ Control Panel")
    st.markdown("---")

    # THEME SELECTION
    st.subheader("🎨 Theme & Colors")
    theme = st.selectbox("Choose Theme", [
        "Professional Blue",
        "Dark Mode",
        "Ocean",
        "Sunset",
        "Forest",
        "Purple",
        "Mint",
        "Custom"
    ])

    color_themes = {
        "Professional Blue": ("#667eea", "#764ba2", "#00cc00"),
        "Dark Mode": ("#1e1e1e", "#333333", "#4CAF50"),
        "Ocean": ("#0077be", "#00a8cc", "#7fdbff"),
        "Sunset": ("#ff6b6b", "#feca57", "#ff9ff3"),
        "Forest": ("#2d5016", "#73a942", "#aad576"),
        "Purple": ("#9b59b6", "#8e44ad", "#3498db"),
        "Mint": ("#1abc9c", "#16a085", "#27ae60"),
    }

    if theme == "Custom":
        primary_color = st.color_picker("Primary Color", "#667eea")
        secondary_color = st.color_picker("Secondary Color", "#764ba2")
        accent_color = st.color_picker("Accent Color", "#00cc00")
    else:
        primary_color, secondary_color, accent_color = color_themes[theme]

    st.session_state.theme_color = primary_color

    st.markdown("---")

    # AUTO REFRESH
    st.subheader("🔄 Auto Refresh")
    auto_refresh = st.checkbox("Enable Auto-Refresh", value=False)
    refresh_seconds = st.slider("Refresh every (seconds)", 5, 300, 30)

    st.markdown("---")

    # DATA SOURCE
    st.subheader("📁 Data Source")
    data_source = st.radio("Choose Source", ["Upload File", "Live URL", "Sample Data", "PDF Document"])

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
    show_stats = st.checkbox("Show Statistics", value=True)
    show_raw = st.checkbox("Show Raw Data", value=True)

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
        font-weight: bold;
    }}
    .alert-warning {{
        background: linear-gradient(135deg, #feca57 0%, #ff9ff3 100%);
        color: white;
        padding: 15px;
        border-radius: 10px;
        margin: 10px 0;
        font-weight: bold;
    }}
    .alert-success {{
        background: linear-gradient(135deg, #1abc9c 0%, #16a085 100%);
        color: white;
        padding: 15px;
        border-radius: 10px;
        margin: 10px 0;
        font-weight: bold;
    }}
    .pdf-info {{
        background: #e8f4f8;
        padding: 15px;
        border-radius: 10px;
        border-left: 4px solid {primary_color};
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
    .metric-card {{
        background: white;
        padding: 15px;
        border-radius: 10px;
        text-align: center;
        box-shadow: 0 2px 10px rgba(0,0,0,0.1);
    }}
    </style>
""", unsafe_allow_html=True)

# ==================== HEADER ====================
st.markdown("""
<div class="main-header">
    <h1>📊 Ultimate Professional Dashboard</h1>
    <p>Auto-Clean | Live Data | Advanced Analytics | Real-time Updates | PDF Support</p>
</div>
""", unsafe_allow_html=True)

# ==================== DATA LOADING ====================
df = None

if data_source == "Upload File":
    uploaded_file = st.file_uploader(
        "📁 Upload Excel/CSV File",
        type=['csv', 'xlsx', 'xlsm', 'xls'],
        help="Supports Excel and CSV formats"
    )

    if uploaded_file:
        if uploaded_file.name.lower().endswith('.csv'):
            df = load_csv(uploaded_file)
        else:
            df = load_excel(uploaded_file)

        if df is not None:
            st.session_state.data = df
            st.success(f"✅ Loaded {len(df)} rows × {len(df.columns)} columns")

elif data_source == "Live URL":
    url = st.text_input(
        "📎 Paste Google Sheets or CSV URL",
        help="Make sheet public: File → Share → Anyone with link → Viewer"
    )
    st.info("💡 For Google Sheets: Replace '/edit' with '/export?format=csv'")

    if url and st.button("🔄 Load URL Data"):
        with st.spinner("Loading data..."):
            df = load_url_data(url)
            if df is not None:
                st.session_state.data = df
                st.success(f"✅ Loaded {len(df)} rows × {len(df.columns)} columns")

elif data_source == "Sample Data":
    if st.button("📊 Generate Sample Data"):
        df = generate_sample_data()
        st.session_state.data = df
        st.success(f"✅ Sample data loaded! {len(df)} rows × {len(df.columns)} columns")

elif data_source == "PDF Document":
    st.markdown("### 📄 Upload PDF Document")

    uploaded_pdf = st.file_uploader(
        "📄 Upload PDF File",
        type=['pdf'],
        help="Upload a PDF with tables or text data"
    )

    if uploaded_pdf:
        st.markdown(f"""
        <div class="pdf-info">
            <h4>📄 PDF: {uploaded_pdf.name}</h4>
            <p><strong>Size:</strong> {uploaded_pdf.size / 1024:.2f} KB</p>
        </div>
        """, unsafe_allow_html=True)

        extraction_mode = st.radio(
            "🔍 Extraction Mode",
            ["Auto", "Tables", "Text"],
            horizontal=True,
            help="Auto = try tables first, then text"
        )

        if st.button("🔄 Extract Data from PDF", type="primary"):
            with st.spinner("📄 Extracting data from PDF..."):
                pdf_bytes = uploaded_pdf.getvalue()
                df = load_pdf(pdf_bytes, mode=extraction_mode)

                if df is not None and not df.empty:
                    st.session_state.data = df
                    st.success(f"✅ Extracted {len(df)} rows × {len(df.columns)} columns from PDF!")

                    st.markdown("### 📋 Data Preview")
                    st.dataframe(df.head(10), use_container_width=True)

# ==================== MAIN DASHBOARD ====================
if st.session_state.data is not None:
    df_raw = st.session_state.data.copy()

    # Validate data
    validation_issues = validate_data(df_raw)
    if validation_issues:
        st.warning("⚠️ Data Issues Detected:")
        for issue in validation_issues:
            st.warning(issue)

    # AUTO CLEAN
    clean_options = {
        'duplicates': remove_duplicates,
        'missing': fill_missing,
        'trim': trim_spaces,
        'standardize': standardize_case,
        'special': remove_special
    }

    df, clean_stats = clean_data(df_raw.copy(), clean_options)
    df = make_columns_unique(df)
    st.session_state.clean_data = df

    # Show cleaning stats
    with st.expander("🧹 Data Cleaning Report", expanded=False):
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Original Rows", f"{clean_stats['original_rows']:,}")
        with col2:
            st.metric("Cleaned Rows", f"{clean_stats['final_rows']:,}")
        with col3:
            st.metric("Duplicates Removed", clean_stats['duplicates_removed'])
        with col4:
            st.metric("Missing Filled", clean_stats['missing_filled'])

        st.markdown("---")
        st.markdown("### 📊 Data Info")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Null Values", int(df.isnull().sum().sum()))
        with col2:
            st.metric("Duplicate Rows", int(df.duplicated().sum()))
        with col3:
            st.metric("Memory Usage", f"{df.memory_usage(deep=True).sum() / 1024:.2f} KB")

    # Get column types
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    text_cols = df.select_dtypes(include=['object']).columns.tolist()

    # ==================== TABS ====================
    tabs = st.tabs([
        "📈 Live KPIs",
        "🚨 Alerts",
        "📊 Charts",
        "🔄 Pivot",
        "📉 Statistics",
        "🔍 Filter",
        "📋 Raw Data",
        "📥 Export"
    ])

    # ========== TAB 1: KPIs ==========
    with tabs[0]:
        if show_kpis:
            st.subheader("📈 Real-Time Key Performance Indicators")

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

            # Numeric columns analysis
            if numeric_cols:
                st.markdown("### 🔢 Numeric Analysis")

                num_cols_display = numeric_cols[:8]
                num_tabs = st.tabs([f"📊 {str(col)[:20]}" for col in num_cols_display])

                for i, col in enumerate(num_tabs):
                    if i < len(num_cols_display):
                        col_name = num_cols_display[i]
                        with col:
                            c1, c2, c3, c4, c5, c6 = st.columns(6)

                            with c1:
                                st.metric("Sum", f"{df[col_name].sum():,.0f}")
                            with c2:
                                st.metric("Avg", f"{df[col_name].mean():,.2f}")
                            with c3:
                                st.metric("Median", f"{df[col_name].median():,.2f}")
                            with c4:
                                st.metric("Max", f"{df[col_name].max():,.0f}")
                            with c5:
                                st.metric("Min", f"{df[col_name].min():,.0f}")
                            with c6:
                                st.metric("Std Dev", f"{df[col_name].std():,.2f}")

                            fig = px.histogram(df, x=col_name, nbins=20, title=f"Distribution of {col_name}")
                            st.plotly_chart(fig, use_container_width=True, height=300)

            # Text columns analysis
            if text_cols:
                st.markdown("---")
                st.markdown("### 📝 Text Analysis")

                for col in text_cols[:3]:
                    with st.expander(f"📊 {col} - Value Counts"):
                        value_counts = df[col].value_counts()
                        col1, col2 = st.columns([1, 2])

                        with col1:
                            st.dataframe(value_counts.head(10), use_container_width=True)

                        with col2:
                            fig = px.bar(x=value_counts.head(10).index.astype(str),
                                       y=value_counts.head(10).values,
                                       labels={'x': col, 'y': 'Count'},
                                       title=f"Top 10 - {col}")
                            st.plotly_chart(fig, use_container_width=True, height=300)
        else:
            st.info("KPIs display is disabled. Enable in sidebar settings.")

    # ========== TAB 2: ALERTS ==========
    with tabs[1]:
        if show_alerts:
            st.subheader("🚨 Alerts & Warnings")

            alerts_found = False

            for col in numeric_cols:
                zero_count = (df[col] == 0).sum()
                if zero_count > 0 and zero_count > len(df) * 0.3:
                    alerts_found = True
                    st.markdown(f"""
                    <div class="alert-warning">
                        <h3>⚠️ WARNING: {zero_count} ({zero_count/len(df)*100:.1f}%) values are ZERO in {col}</h3>
                    </div>
                    """, unsafe_allow_html=True)

            if 'CLOSING_STOCK' in df.columns and 'SAFETY_STOCK' in df.columns:
                try:
                    critical = df[df['CLOSING_STOCK'] < df['SAFETY_STOCK']]
                    if len(critical) > 0:
                        alerts_found = True
                        st.markdown(f"""
                        <div class="alert-critical">
                            <h3>🔴 CRITICAL: {len(critical)} items below safety stock!</h3>
                        </div>
                        """, unsafe_allow_html=True)
                        st.dataframe(critical, use_container_width=True)
                except:
                    pass

            if 'ISSUED' in df.columns and 'RECEIVED' in df.columns:
                try:
                    dead = df[(df['ISSUED'] == 0) & (df['RECEIVED'] == 0)]
                    if len(dead) > 0:
                        alerts_found = True
                        st.markdown(f"""
                        <div class="alert-warning">
                            <h3>⚠️ WARNING: {len(dead)} items with NO MOVEMENT (Dead Stock)</h3>
                        </div>
                        """, unsafe_allow_html=True)
                        st.dataframe(dead, use_container_width=True)
                except:
                    pass

            if 'CLOSING_STOCK' in df.columns:
                try:
                    out_stock = df[df['CLOSING_STOCK'] == 0]
                    if len(out_stock) > 0:
                        alerts_found = True
                        st.markdown(f"""
                        <div class="alert-critical">
                            <h3>🔴 CRITICAL: {len(out_stock)} items are OUT OF STOCK!</h3>
                        </div>
                        """, unsafe_allow_html=True)
                        st.dataframe(out_stock, use_container_width=True)
                except:
                    pass

            if not alerts_found:
                st.markdown("""
                <div class="alert-success">
                    <h3>✅ No Alerts! Everything looks good!</h3>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.info("Alerts display is disabled. Enable in sidebar settings.")

    # ========== TAB 3: CHARTS ==========
    with tabs[2]:
        if show_charts:
            st.subheader("📊 Advanced Visualizations")

            col1, col2 = st.columns([1, 3])

            with col1:
                chart_type = st.selectbox("Chart Type", [
                    "Bar Chart", "Horizontal Bar", "Pie Chart", "Donut Chart",
                    "Line Chart", "Area Chart", "Scatter Plot", "Bubble Chart",
                    "Histogram", "Box Plot", "Violin Plot", "Heatmap",
                    "Treemap", "Sunburst", "Funnel Chart", "Waterfall"
                ])

                if numeric_cols and text_cols:
                    x_axis = st.selectbox("X-Axis", [None] + df.columns.tolist(), key="chart_x")
                    y_axis = st.selectbox("Y-Axis", [None] + numeric_cols, key="chart_y")
                    color_by = st.selectbox("Color By", [None] + df.columns.tolist(), key="chart_color")
                    top_n = st.slider("Top N", 5, 50, 15)
                else:
                    x_axis = text_cols[0] if text_cols else None
                    y_axis = numeric_cols[0] if numeric_cols else None
                    color_by = None
                    top_n = 15

            with col2:
                if x_axis and y_axis:
                    fig = create_chart_safely(chart_type, df, x_axis, y_axis, color_by, numeric_cols, text_cols, top_n)
                    if fig:
                        st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("⚠️ Need to select both X and Y axes")

            st.markdown("---")

            st.markdown("### 📊 Quick Auto-Charts")

            if numeric_cols and text_cols:
                chart_col1, chart_col2, chart_col3 = st.columns(3)

                with chart_col1:
                    fig1 = create_chart_safely("Bar Chart", df, text_cols[0], numeric_cols[0],
                                              None, numeric_cols, text_cols, 10)
                    if fig1:
                        st.plotly_chart(fig1, use_container_width=True, height=300)

                with chart_col2:
                    fig2 = create_chart_safely("Pie Chart", df, text_cols[0], numeric_cols[0],
                                              None, numeric_cols, text_cols, 10)
                    if fig2:
                        st.plotly_chart(fig2, use_container_width=True, height=300)

                with chart_col3:
                    if len(numeric_cols) > 1:
                        fig3 = create_chart_safely("Scatter Plot", df, numeric_cols[0], numeric_cols[1],
                                                  None, numeric_cols, text_cols)
                        if fig3:
                            st.plotly_chart(fig3, use_container_width=True, height=300)

                if len(numeric_cols) >= 2:
                    chart_col1, chart_col2 = st.columns(2)

                    with chart_col1:
                        fig4 = create_chart_safely("Histogram", df, numeric_cols[0], numeric_cols[0],
                                                  None, numeric_cols, text_cols)
                        if fig4:
                            st.plotly_chart(fig4, use_container_width=True, height=300)

                    with chart_col2:
                        fig5 = px.box(df, y=numeric_cols[0], title=f"Box Plot - {numeric_cols[0]}")
                        st.plotly_chart(fig5, use_container_width=True, height=300)

                    if len(numeric_cols) > 2:
                        corr = df[numeric_cols].corr()
                        fig6 = px.imshow(corr, text_auto=True, aspect='auto',
                                       color_continuous_scale='RdBu_r', title="Correlation Heatmap")
                        st.plotly_chart(fig6, use_container_width=True)
        else:
            st.info("Charts display is disabled. Enable in sidebar settings.")

    # ========== TAB 4: PIVOT ==========
    with tabs[3]:
        if show_pivot:
            st.subheader("🔄 Advanced Pivot Table Analysis")

            col1, col2, col3, col4 = st.columns(4)

            with col1:
                pivot_rows = st.multiselect("📊 Rows", text_cols,
                                           default=text_cols[:1] if text_cols else [],
                                           key="pivot_rows")

            with col2:
                pivot_cols = st.multiselect("📋 Columns", text_cols, key="pivot_cols")

            with col3:
                pivot_values = st.multiselect("🔢 Values", numeric_cols,
                                            default=numeric_cols[:1] if numeric_cols else [],
                                            key="pivot_values")

            with col4:
                pivot_agg = st.selectbox("⚡ Aggregation",
                                        ["sum", "mean", "count", "max", "min", "median", "std"])

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
                        margins_name='TOTAL'
                    )

                    pivot_display = pivot_table.reset_index()

                    # Flatten MultiIndex columns if needed
                    if isinstance(pivot_display.columns, pd.MultiIndex):
                        pivot_display.columns = [
                            "_".join([str(x) for x in c if str(x) != ""])
                            for c in pivot_display.columns.to_flat_index()
                        ]

                    pivot_display = make_columns_unique(pivot_display)

                    st.markdown("### 📊 Pivot Result")
                    st.dataframe(pivot_display, use_container_width=True, height=400)

                    pivot_csv = pivot_display.to_csv(index=False)
                    st.download_button("📥 Download Pivot (CSV)", pivot_csv,
                                     f"pivot_{datetime.now():%Y%m%d_%H%M%S}.csv", "text/csv")

                except Exception as e:
                    st.error(f"❌ Pivot error: {str(e)}")
            else:
                st.info("👆 Select Rows and Values to create pivot table")
        else:
            st.info("Pivot display is disabled. Enable in sidebar settings.")

    # ========== TAB 5: STATISTICS ==========
    with tabs[4]:
        if show_stats:
            st.subheader("📉 Statistical Analysis")

            col1, col2 = st.columns(2)

            with col1:
                st.markdown("### 📊 Descriptive Statistics")
                try:
                    desc_stats = df.describe(include='all').T.reset_index()
                    desc_stats = desc_stats.rename(columns={'index': 'COLUMN'})
                    desc_stats = make_columns_unique(desc_stats)
                    st.dataframe(desc_stats, use_container_width=True, hide_index=True)
                except Exception as e:
                    st.warning(f"⚠️ Could not generate statistics: {str(e)}")

            with col2:
                st.markdown("### 📈 Data Type Summary")
                try:
                    dtype_counts = df.dtypes.astype(str).value_counts()
                    dtype_summary = pd.DataFrame({
                        'Data Type': dtype_counts.index,
                        'Count': dtype_counts.values
                    })
                    st.dataframe(dtype_summary, use_container_width=True, hide_index=True)
                except Exception as e:
                    st.warning(f"⚠️ Could not generate type summary: {str(e)}")

            st.markdown("---")

            col1, col2 = st.columns(2)

            selected_col = None

            with col1:
                st.markdown("### 🔢 Column Statistics")
                if numeric_cols:
                    selected_col = st.selectbox("Select Column", numeric_cols, key="stats_col")

                    mode_val = df[selected_col].mode()
                    mode_display = f"{mode_val.iloc[0]:,.2f}" if not mode_val.empty else "N/A"

                    stats_data = {
                        'Metric': [
                            'Count', 'Sum', 'Mean', 'Median', 'Mode',
                            'Std Dev', 'Variance', 'Min', 'Q1', 'Q3',
                            'Max', 'Range', 'IQR', 'Skewness', 'Kurtosis'
                        ],
                        'Value': [
                            str(df[selected_col].count()),
                            f"{df[selected_col].sum():,.2f}",
                            f"{df[selected_col].mean():,.2f}",
                            f"{df[selected_col].median():,.2f}",
                            mode_display,
                            f"{df[selected_col].std():,.2f}",
                            f"{df[selected_col].var():,.2f}",
                            f"{df[selected_col].min():,.2f}",
                            f"{df[selected_col].quantile(0.25):,.2f}",
                            f"{df[selected_col].quantile(0.75):,.2f}",
                            f"{df[selected_col].max():,.2f}",
                            f"{df[selected_col].max() - df[selected_col].min():,.2f}",
                            f"{df[selected_col].quantile(0.75) - df[selected_col].quantile(0.25):,.2f}",
                            f"{df[selected_col].skew():,.2f}",
                            f"{df[selected_col].kurtosis():,.2f}"
                        ]
                    }
                    stats_df = pd.DataFrame(stats_data)
                    st.dataframe(stats_df, use_container_width=True, hide_index=True)

            with col2:
                st.markdown("### 📈 Distribution")
                if numeric_cols and selected_col:
                    fig = px.histogram(df, x=selected_col, nbins=30, marginal="box",
                                     title=f"Distribution of {selected_col}")
                    st.plotly_chart(fig, use_container_width=True)

            st.markdown("---")
            st.markdown("### 🔗 Correlation Analysis")
            if len(numeric_cols) > 1:
                corr = df[numeric_cols].corr()
                fig = px.imshow(corr, text_auto=True, aspect="auto", color_continuous_scale='RdBu_r',
                              title="Correlation Matrix", labels=dict(color="Correlation"))
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Statistics display is disabled. Enable in sidebar settings.")

    # ========== TAB 6: FILTER & QUERY ==========
    with tabs[5]:
        st.subheader("🔍 Advanced Filter & Query")

        search = st.text_input("🔍 Global Search (searches all columns)", placeholder="Type to search...")

        st.markdown("### 🎯 Column Filters")

        num_filters = st.slider("Number of filters", 1, 5, 1, key="filter_count")

        filters = {}
        for i in range(num_filters):
            col1, col2 = st.columns(2)
            with col1:
                filter_col = st.selectbox(f"Column {i+1}", df.columns.tolist(), key=f"filter_col_{i}")
            with col2:
                if filter_col in numeric_cols:
                    min_v = float(df[filter_col].min())
                    max_v = float(df[filter_col].max())
                    if min_v == max_v:
                        max_v = min_v + 1
                    filters[filter_col] = st.slider(f"Range for {filter_col}",
                                                    min_v, max_v, (min_v, max_v), key=f"filter_range_{i}")
                else:
                    unique_vals = sorted(df[filter_col].astype(str).unique().tolist())
                    default_vals = unique_vals[:5] if len(unique_vals) > 5 else unique_vals
                    filters[filter_col] = st.multiselect(f"Select values", unique_vals,
                                                         default=default_vals, key=f"filter_multi_{i}")

        # Apply filters
        filtered_df = df.copy()

        for col, val in filters.items():
            if col in numeric_cols:
                filtered_df = filtered_df[(filtered_df[col] >= val[0]) & (filtered_df[col] <= val[1])]
            else:
                if val:
                    filtered_df = filtered_df[filtered_df[col].astype(str).isin(val)]

        if search:
            search_filter = filtered_df.astype(str).apply(
                lambda x: x.str.contains(search, case=False, na=False)
            ).any(axis=1)
            filtered_df = filtered_df[search_filter]

        pct = (len(filtered_df) / len(df) * 100) if len(df) > 0 else 0
        st.markdown(f"### 📋 Results: {len(filtered_df):,} rows ({pct:.1f}%)")

        col1, col2 = st.columns([3, 1])
        with col2:
            st.info(f"Original: {len(df):,} rows")

        st.dataframe(filtered_df, use_container_width=True, height=500)

        # Download
        col1, col2, col3 = st.columns(3)
        with col1:
            csv_filtered = filtered_df.to_csv(index=False)
            st.download_button("📥 CSV", csv_filtered, f"filtered_{datetime.now():%Y%m%d_%H%M%S}.csv")
        with col2:
            json_filtered = filtered_df.to_json(orient='records', indent=2)
            st.download_button("📋 JSON", json_filtered, f"filtered_{datetime.now():%Y%m%d_%H%M%S}.json")
        with col3:
            excel_filtered = create_excel_with_formatting(filtered_df, "filtered")
            if excel_filtered:
                st.download_button("📊 Excel", excel_filtered,
                                 f"filtered_{datetime.now():%Y%m%d_%H%M%S}.xlsx")

    # ========== TAB 7: RAW DATA ==========
    with tabs[6]:
        if show_raw:
            st.subheader("📋 Complete Raw Data")

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Rows", f"{len(df):,}")
            with col2:
                st.metric("Columns", len(df.columns))
            with col3:
                st.metric("Missing Values", int(df.isnull().sum().sum()))
            with col4:
                st.metric("Duplicate Rows", int(df.duplicated().sum()))

            st.markdown("---")
            search_data = st.text_input("🔍 Search in raw data", key="raw_search")
            if search_data:
                search_filter = df.astype(str).apply(
                    lambda x: x.str.contains(search_data, case=False, na=False)
                ).any(axis=1)
                display_df = df[search_filter]
            else:
                display_df = df

            st.dataframe(display_df, use_container_width=True, height=600)

            st.markdown("---")
            st.markdown("### 📊 Column Information")

            try:
                info_data = {
                    'Column': df.columns.tolist(),
                    'Type': df.dtypes.astype(str).tolist(),
                    'Non-Null': df.count().tolist(),
                    'Null': df.isnull().sum().tolist(),
                    'Null %': (df.isnull().sum() / len(df) * 100).round(2).tolist() if len(df) > 0 else [0] * len(df.columns),
                    'Unique': [df.iloc[:, i].nunique() for i in range(len(df.columns))],
                }
                info_df = pd.DataFrame(info_data)
                st.dataframe(info_df, use_container_width=True, hide_index=True)
            except Exception as e:
                st.warning(f"⚠️ Could not generate column info: {str(e)}")
        else:
            st.info("Raw data display is disabled. Enable in sidebar settings.")

    # ========== TAB 8: EXPORT ==========
    with tabs[7]:
        st.subheader("📥 Export Options")

        st.markdown("### 📊 Export Full Dataset")
        col1, col2, col3, col4 = st.columns(4)

        with col1:
            csv = df.to_csv(index=False)
            st.download_button(
                "📄 CSV",
                csv,
                f"data_{datetime.now():%Y%m%d_%H%M%S}.csv",
                "text/csv",
                use_container_width=True
            )

        with col2:
            json_data = df.to_json(orient='records', indent=2)
            st.download_button(
                "📋 JSON",
                json_data,
                f"data_{datetime.now():%Y%m%d_%H%M%S}.json",
                "application/json",
                use_container_width=True
            )

        with col3:
            excel_data = create_excel_with_formatting(df, "data")
            if excel_data:
                st.download_button(
                    "📊 Excel",
                    excel_data,
                    f"data_{datetime.now():%Y%m%d_%H%M%S}.xlsx",
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

        with col4:
            try:
                export_df = make_columns_unique(df.copy())
                export_df.columns = [str(c) for c in export_df.columns]
                parquet_data = export_df.to_parquet(index=False)
                st.download_button(
                    "⚡ Parquet",
                    parquet_data,
                    f"data_{datetime.now():%Y%m%d_%H%M%S}.parquet",
                    "application/octet-stream",
                    use_container_width=True
                )
            except Exception as e:
                st.warning(f"⚠️ Parquet export not available: {str(e)}")

        st.markdown("---")
        st.markdown("### 📋 Export Summary")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Dataset Summary**")
            summary_text = f"""
            - **Rows:** {len(df):,}
            - **Columns:** {len(df.columns)}
            - **File Size:** {df.memory_usage(deep=True).sum() / 1024:.2f} KB
            - **Export Time:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            """
            st.markdown(summary_text)

        with col2:
            st.markdown("**Column Types**")
            type_counts = df.dtypes.astype(str).value_counts()
            for dtype, count in type_counts.items():
                st.markdown(f"- **{dtype}:** {count} columns")

    # ==================== AUTO REFRESH ====================
    if auto_refresh:
        st.session_state.refresh_count += 1
        col1, col2 = st.columns([3, 1])
        with col2:
            st.info(f"🔄 Auto-refresh in {refresh_seconds}s...")
        time.sleep(refresh_seconds)
        st.rerun()

else:
    # Landing page
    st.markdown("""
    <div class="main-header" style="text-align: center;">
        <h2>👋 Welcome to Ultimate Professional Dashboard</h2>
        <p>Start by uploading a file, providing a URL, using a PDF, or using sample data</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("""
        <div class="kpi-card">
            <h3>🚀 Features</h3>
            <ul>
                <li>✅ Auto-Clean Data</li>
                <li>✅ Live Updates</li>
                <li>✅ 15+ Chart Types</li>
                <li>✅ Excel-like Pivot</li>
                <li>✅ Advanced Filters</li>
                <li>✅ PDF Document Support</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="kpi-card">
            <h3>🎨 Customization</h3>
            <ul>
                <li>✅ 8 Color Themes</li>
                <li>✅ Custom Colors</li>
                <li>✅ Auto-Refresh Control</li>
                <li>✅ Toggle Features</li>
                <li>✅ Responsive Design</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    with col3:
        st.markdown("""
        <div class="kpi-card">
            <h3>📊 Analytics</h3>
            <ul>
                <li>✅ Real-time KPIs</li>
                <li>✅ Statistics</li>
                <li>✅ Correlation Analysis</li>
                <li>✅ Alert System</li>
                <li>✅ Multi-Format Export</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.success("✅ Supports: CSV, Excel, PDF, Google Sheets")
    with col2:
        st.info("ℹ️ Auto-cleans data on upload")
    with col3:
        st.warning("⚠️ No data sent to external servers")
```

---

## Also update your `requirements.txt`

Replace your `requirements.txt` with this:

```txt
streamlit
pandas
numpy
plotly
openpyxl
pdfplumber
pyarrow
```

---

## What's included in this version

| Feature | Status |
|---|---|
| ✅ PDF Document option in sidebar | Added |
| ✅ PDF table extraction | Added |
| ✅ PDF text extraction (fallback) | Added |
| ✅ Auto / Tables / Text extraction modes | Added |
| ✅ Duplicate column name fix (pyarrow error) | Fixed |
| ✅ Column Info "arrays same length" error | Fixed |
| ✅ `describe()` crash fix | Fixed |
| ✅ Pivot MultiIndex column flattening | Fixed |
| ✅ Safe Parquet export | Fixed |
| ✅ URL loader `timeout` bug removed | Fixed |

After pushing to GitHub, your sidebar will show **PDF Document** as the 4th option. Select it, upload a PDF, choose extraction mode, and click **Extract Data from PDF**.
