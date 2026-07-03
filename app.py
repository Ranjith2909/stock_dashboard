Here is the **complete, ready-to-copy code** with PDF upload functionality included:

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

# PDF libraries
import pdfplumber
from PyPDF2 import PdfReader

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

@st.cache_data
def load_csv(file):
    try:
        df = pd.read_csv(file, on_bad_lines='skip')
        if df.empty:
            st.error("❌ File is empty!")
            return None
        return df
    except Exception as e:
        st.error(f"❌ Error reading CSV: {str(e)}")
        return None

@st.cache_data
def load_excel(file):
    try:
        df = pd.read_excel(file)
        if df.empty:
            st.error("❌ File is empty!")
            return None
        return df
    except Exception as e:
        st.error(f"❌ Error reading Excel: {str(e)}")
        return None

@st.cache_data
def load_url_data(url):
    try:
        if 'docs.google.com/spreadsheets' in url:
            if '/edit' in url:
                url = url.replace('/edit#gid=', '/export?format=csv&gid=')
                url = url.replace('/edit?usp=sharing', '/export?format=csv')
                url = url.replace('/edit', '/export?format=csv')
        
        df = pd.read_csv(url, timeout=10)
        if df.empty:
            st.error("❌ URL data is empty!")
            return None
        return df
    except Exception as e:
        st.error(f"❌ Error loading URL: {str(e)}")
        return None

def extract_tables_from_pdf(pdf_file):
    tables_data = []
    try:
        with pdfplumber.open(pdf_file) as pdf:
            for i, page in enumerate(pdf.pages):
                page_tables = page.extract_tables()
                if page_tables:
                    for j, table in enumerate(page_tables):
                        if table:
                            df = pd.DataFrame(table[1:], columns=table[0])
                            df['_source_page'] = i + 1
                            df['_source_table'] = j + 1
                            tables_data.append(df)
        return tables_data
    except Exception as e:
        st.error(f"❌ Error extracting tables from PDF: {str(e)}")
        return []

def extract_text_from_pdf(pdf_file):
    try:
        reader = PdfReader(pdf_file)
        text_data = []
        for i, page in enumerate(reader.pages):
            text = page.extract_text()
            if text:
                text_data.append({
                    'page': i + 1,
                    'text': text
                })
        return text_data
    except Exception as e:
        st.error(f"❌ Error extracting text from PDF: {str(e)}")
        return []

def parse_pdf_to_dataframe(pdf_file, extraction_mode='tables'):
    try:
        if extraction_mode == 'tables':
            tables = extract_tables_from_pdf(pdf_file)
            if tables:
                return pd.concat(tables, ignore_index=True)
            else:
                st.warning("⚠️ No tables found in PDF. Trying text extraction...")
                return None
        elif extraction_mode == 'text':
            text_data = extract_text_from_pdf(pdf_file)
            if text_data:
                return pd.DataFrame(text_data)
            return None
        else:
            tables = extract_tables_from_pdf(pdf_file)
            if tables:
                return pd.concat(tables, ignore_index=True)
            return None
    except Exception as e:
        st.error(f"❌ Error parsing PDF: {str(e)}")
        return None

def convert_pdf_text_to_structured_data(text_df):
    try:
        all_data = []
        for _, row in text_df.iterrows():
            text = row['text']
            lines = text.split('\n')
            
            for line in lines:
                if ':' in line or '=' in line:
                    parts = re.split(r'[:=]', line, 1)
                    if len(parts) == 2:
                        key = parts[0].strip()
                        value = parts[1].strip()
                        if key and value:
                            all_data.append({
                                'Field': key,
                                'Value': value,
                                'Page': row['page']
                            })
        
        if all_data:
            return pd.DataFrame(all_data)
        return None
    except Exception as e:
        st.error(f"❌ Error converting text: {str(e)}")
        return None

def generate_sample_data():
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

def make_columns_unique(df):
    if df is None or df.empty:
        return df
    if df.columns.duplicated().any():
        cols = pd.Series(df.columns)
        for dup in cols[cols.duplicated()].unique():
            count = 0
            for i, col in enumerate(cols):
                if col == dup:
                    if count > 0:
                        cols.iloc[i] = f"{col}_{count}"
                    count += 1
        df.columns = cols
    return df

def clean_data(df, options):
    if df is None or df.empty:
        return df, {'original_rows': 0, 'final_rows': 0, 'duplicates_removed': 0, 'missing_filled': 0}
    
    original_rows = len(df)
    duplicates_removed = 0
    missing_filled = 0
    
    try:
        if options.get('standardize', True):
            df.columns = df.columns.str.strip().str.upper().str.replace(' ', '_')
        
        object_cols = df.select_dtypes(include=['object']).columns.tolist()
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        if options.get('trim', True) and len(object_cols) > 0:
            for col in object_cols:
                try:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.strip()
                except:
                    pass
        
        if options.get('duplicates', True):
            before = len(df)
            df = df.drop_duplicates()
            duplicates_removed = before - len(df)
        
        if options.get('missing', True):
            missing_filled = df.isnull().sum().sum()
            for col in numeric_cols:
                if col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            for col in object_cols:
                if col in df.columns:
                    df[col] = df[col].fillna('N/A')
        
        if options.get('special', False):
            for col in object_cols:
                try:
                    if col in df.columns:
                        df[col] = df[col].astype(str).str.replace(r'[^\w\s]', '', regex=True)
                except:
                    pass
        
        for col in df.columns:
            try:
                if df[col].dtype == 'object':
                    converted = pd.to_numeric(df[col], errors='coerce')
                    if converted.notna().sum() / len(df) > 0.8:
                        df[col] = converted
            except:
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
            fig = px.pie(data, names=x_axis, values=y_axis, hole=0.5, title=f"Distribution of {y_axis}")
        elif chart_type == "Line Chart":
            fig = px.line(df.head(50), x=x_axis, y=y_axis, color=color_by, markers=True, title=f"{y_axis} Trend")
        elif chart_type == "Area Chart":
            fig = px.area(df.head(50), x=x_axis, y=y_axis, color=color_by, title=f"{y_axis} Area Chart")
        elif chart_type == "Scatter Plot":
            fig = px.scatter(df, x=x_axis, y=y_axis, color=color_by, size=y_axis, title=f"{x_axis} vs {y_axis}")
        elif chart_type == "Bubble Chart":
            fig = px.scatter(df, x=x_axis, y=y_axis, size=y_axis, color=color_by, title=f"Bubble Chart: {x_axis} vs {y_axis}")
        elif chart_type == "Histogram":
            fig = px.histogram(df, x=y_axis, nbins=30, color=color_by, title=f"Distribution of {y_axis}")
        elif chart_type == "Box Plot":
            fig = px.box(df, y=y_axis, color=color_by, title=f"Box Plot of {y_axis}")
        elif chart_type == "Violin Plot":
            fig = px.violin(df, y=y_axis, color=color_by, box=True, title=f"Violin Plot of {y_axis}")
        elif chart_type == "Heatmap":
            if len(numeric_cols) > 1:
                corr = df[numeric_cols].corr()
                fig = px.imshow(corr, text_auto=True, aspect='auto', color_continuous_scale='RdBu_r', title="Correlation Heatmap")
            else:
                st.warning("Need at least 2 numeric columns for heatmap")
                return None
        elif chart_type == "Treemap":
            data = df.nlargest(top_n, y_axis) if y_axis in numeric_cols else df.head(top_n)
            fig = px.treemap(data, path=[x_axis], values=y_axis, title=f"Treemap of {y_axis} by {x_axis}")
        elif chart_type == "Sunburst":
            if len(text_cols) >= 2:
                fig = px.sunburst(df.head(30), path=text_cols[:2], values=y_axis, title="Sunburst Chart")
            else:
                fig = px.sunburst(df.head(30), path=[x_axis], values=y_axis, title="Sunburst Chart")
        elif chart_type == "Funnel Chart":
            data = df.nlargest(top_n, y_axis)
            fig = px.funnel(data, x=y_axis, y=x_axis, title=f"Funnel Chart: {y_axis}")
        elif chart_type == "Waterfall":
            data = df.nlargest(10, y_axis)
            fig = px.bar(data, x=x_axis, y=y_axis, title=f"Waterfall: {y_axis}")
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
    
    st.subheader("🎨 Theme & Colors")
    theme = st.selectbox("Choose Theme", [
        "Professional Blue", "Dark Mode", "Ocean", "Sunset", "Forest", "Purple", "Mint", "Custom"
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
    st.subheader("🔄 Auto Refresh")
    auto_refresh = st.checkbox("Enable Auto-Refresh", value=False)
    refresh_seconds = st.slider("Refresh every (seconds)", 5, 300, 30)
    
    st.markdown("---")
    st.subheader("📁 Data Source")
    data_source = st.radio("Choose Source", ["Upload File", "Live URL", "Sample Data", "PDF Document"])
    
    st.markdown("---")
    st.subheader("🧹 Auto-Clean Options")
    remove_duplicates = st.checkbox("Remove Duplicates", value=True)
    fill_missing = st.checkbox("Fill Missing Values", value=True)
    trim_spaces = st.checkbox("Trim Spaces", value=True)
    standardize_case = st.checkbox("Standardize Columns", value=True)
    remove_special = st.checkbox("Remove Special Characters", value=False)
    
    st.markdown("---")
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
    .pdf-info {{
        background: #e8f4f8;
        padding: 15px;
        border-radius: 10px;
        border-left: 4px solid {primary_color};
        margin: 10px 0;
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
        if uploaded_file.name.endswith('.csv'):
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
    st.markdown("### 📄 PDF Document Upload")
    uploaded_pdf = st.file_uploader(
        "📄 Upload PDF File",
        type=['pdf'],
        help="Upload a PDF document to extract tables or text"
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
            ["Tables (Recommended)", "Text", "Both"],
            horizontal=True
        )
        
        if st.button("🔄 Extract Data from PDF", type="primary"):
            with st.spinner("📄 Extracting data from PDF..."):
                if extraction_mode == "Both":
                    df = parse_pdf_to_dataframe(uploaded_pdf, 'tables')
                    if df is None or df.empty:
                        st.warning("⚠️ No tables found. Trying text extraction...")
                        text_df = extract_text_from_pdf(uploaded_pdf)
                        if text_df:
                            structured_df = convert_pdf_text_to_structured_data(pd.DataFrame(text_df))
                            if structured_df is not None:
                                df = structured_df
                        else:
                            st.error("❌ No data could be extracted from PDF")
                            df = None
                elif extraction_mode == "Text":
                    df = parse_pdf_to_dataframe(uploaded_pdf, 'text')
                    if df is not None:
                        structured_df = convert_pdf_text_to_structured_data(df)
                        if structured_df is not None and not structured_df.empty:
                            df = structured_df
                else:
                    df = parse_pdf_to_dataframe(uploaded_pdf, 'tables')
                
                if df is not None and not df.empty:
                    st.session_state.data = df
                    st.success(f"✅ Extracted {len(df)} rows × {len(df.columns)} columns from PDF!")
                    st.markdown("### 📋 Data Preview")
                    st.dataframe(df.head(10), use_container_width=True)
                else:
                    st.error("❌ Could not extract any data from the PDF")

# ==================== MAIN DASHBOARD ====================
if st.session_state.data is not None:
    df_raw = st.session_state.data.copy()
    
    validation_issues = validate_data(df_raw)
    if validation_issues:
        st.warning("⚠️ Data Issues Detected:")
        for issue in validation_issues:
            st.warning(issue)
    
    clean_options = {
        'duplicates': remove_duplicates,
        'missing': fill_missing,
        'trim': trim_spaces,
        'standardize': standardize_case,
        'special': remove_special
    }
    
    df, clean_stats = clean_data(df_raw.copy(), clean_options)
    st.session_state.clean_data = df
    
    if df is not None:
        df = make_columns_unique(df)
    
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
            st.metric("Null Values", df.isnull().sum().sum() if df is not None else 0)
        with col2:
            st.metric("Duplicate Rows", df.duplicated().sum() if df is not None else 0)
        with col3:
            st.metric("Memory Usage", f"{df.memory_usage(deep=True).sum() / 1024:.2f} KB" if df is not None else "0 KB")
    
    if df is not None:
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        text_cols = df.select_dtypes(include=['object']).columns.tolist()
    else:
        numeric_cols = []
        text_cols = []
    
    tabs = st.tabs([
        "📈 Live KPIs", "🚨 Alerts", "📊 Charts", "🔄 Pivot", 
        "📉 Statistics", "🔍 Filter", "📋 Raw Data", "📥 Export"
    ])
    
    # TAB 1: KPIs
    with tabs[0]:
        if show_kpis:
            st.subheader("📈 Real-Time Key Performance Indicators")
            st.markdown("### 📊 Overview Metrics")
            cols = st.columns(4)
            with cols[0]:
                st.metric("📦 Total Records", f"{len(df):,}" if df is not None else "0")
            with cols[1]:
                st.metric("📊 Columns", len(df.columns) if df is not None else 0)
            with cols[2]:
                st.metric("🔢 Numeric Fields", len(numeric_cols))
            with cols[3]:
                st.metric("📝 Text Fields", len(text_cols))
            
            st.markdown("---")
            
            if numeric_cols:
                st.markdown("### 🔢 Numeric Analysis")
                num_cols_display = numeric_cols[:8]
                if num_cols_display:
                    num_tabs = st.tabs([f"📊 {col[:20]}" for col in num_cols_display])
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
                            fig = px.bar(x=value_counts.head(10).index, y=value_counts.head(10).values,
                                       labels={'x': col, 'y': 'Count'}, title=f"Top 10 - {col}")
                            st.plotly_chart(fig, use_container_width=True, height=300)
        else:
            st.info("KPIs display is disabled.")
    
    # TAB 2: ALERTS
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
            st.info("Alerts display is disabled.")
    
    # TAB 3: CHARTS
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
                    fig1 = create_chart_safely("Bar Chart", df, text_cols[0], numeric_cols[0], None, numeric_cols, text_cols, 10)
                    if fig1:
                        st.plotly_chart(fig1, use_container_width=True, height=300)
                with chart_col2:
                    fig2 = create_chart_safely("Pie Chart", df, text_cols[0], numeric_cols[0], None, numeric_cols, text_cols, 10)
                    if fig2:
                        st.plotly_chart(fig2, use_container_width=True, height=300)
                with chart_col3:
                    if len(numeric_cols) > 1:
                        fig3 = create_chart_safely("Scatter Plot", df, numeric_cols[0], numeric_cols[1], None, numeric_cols, text_cols)
                        if fig3:
                            st.plotly_chart(fig3, use_container_width=True, height=300)
                
                if len(numeric_cols) >= 2:
                    chart_col1, chart_col2 = st.columns(2)
                    with chart_col1:
                        fig4 = create_chart_safely("Histogram", df, numeric_cols[0], numeric_cols[0], None, numeric_cols, text_cols)
                        if fig4:
                            st.plotly_chart(fig4, use_container_width=True, height=300)
                    with chart_col2:
                        fig5 = px.box(df, y=numeric_cols[0], title=f"Box Plot - {numeric_cols[0]}")
                        st.plotly_chart(fig5, use_container_width=True, height=300)
                    if len(numeric_cols) > 2:
                        corr = df[numeric_cols].corr()
                        fig6 = px.imshow(corr, text_auto=True, aspect='auto', color_continuous_scale='RdBu_r', title="Correlation Heatmap")
                        st.plotly_chart(fig6, use_container_width=True)
        else:
            st.info("Charts display is disabled.")
    
    # TAB 4: PIVOT
    with tabs[3]:
        if show_pivot:
            st.subheader("🔄 Advanced Pivot Table Analysis")
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                pivot_rows = st.multiselect("📊 Rows", text_cols, default=text_cols[:1] if text_cols else [], key="pivot_rows")
            with col2:
                pivot_cols = st.multiselect("📋 Columns", text_cols, key="pivot_cols")
            with col3:
                pivot_values = st.multiselect("🔢 Values", numeric_cols, default=numeric_cols[:1] if numeric_cols else [], key="pivot_values")
            with col4:
                pivot_agg = st.selectbox("⚡ Aggregation", ["sum", "mean", "count", "max", "min", "median", "std"])
            
            if pivot_rows and pivot_values:
                try:
                    pivot_table = pd.pivot_table(df, index=pivot_rows, columns=pivot_cols if pivot_cols else None,
                                               values=pivot_values, aggfunc=pivot_agg, fill_value=0, margins=True, margins_name='TOTAL')
                    pivot_table = make_columns_unique(pivot_table.reset_index())
                    st.markdown("### 📊 Pivot Result")
                    st.dataframe(pivot_table, use_container_width=True, height=400)
                    pivot_csv = pivot_table.to_csv()
                    st.download_button("📥 Download Pivot (CSV)", pivot_csv, f"pivot_{datetime.now():%Y%m%d_%H%M%S}.csv", "text/csv")
                except Exception as e:
                    st.error(f"❌ Pivot error: {str(e)}")
            else:
                st.info("👆 Select Rows and Values to create pivot table")
        else:
            st.info("Pivot display is disabled.")
    
    # TAB 5: STATISTICS
    with tabs[4]:
        if show_stats:
            st.subheader("📉 Statistical Analysis")
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### 📊 Descriptive Statistics")
                if df is not None and not df.empty:
                    desc_stats = df.describe(include='all').T
                    st.dataframe(desc_stats, use_container_width=True)
            
            with col2:
                st.markdown("### 📈 Data Type Summary")
                if df is not None:
                    dtype_summary = pd.DataFrame({
                        'Data Type': df.dtypes.value_counts().index.astype(str),
                        'Count': df.dtypes.value_counts().values
                    })
                    st.dataframe(dtype_summary, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### 🔢 Column Statistics")
                if numeric_cols:
                    selected_col = st.selectbox("Select Column", numeric_cols, key="stats_col")
                    mode_val = df[selected_col].mode()
                    mode_display = f"{mode_val.iloc[0]:,.2f}" if not mode_val.empty else "N/A"
                    
                    stats_data = {
                        'Metric': ['Count', 'Sum', 'Mean', 'Median', 'Mode', 'Std Dev', 'Variance', 'Min', 'Q1', 'Q3', 'Max', 'Range', 'IQR', 'Skewness', 'Kurtosis'],
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
                    fig = px.histogram(df, x=selected_col, nbins=30, marginal="box", title=f"Distribution of {selected_col}")
                    st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("---")
            st.markdown("### 🔗 Correlation Analysis")
            if len(numeric_cols) > 1:
                corr = df[numeric_cols].corr()
                fig = px.imshow(corr, text_auto=True, aspect="auto", color_continuous_scale='RdBu_r',
                              title="Correlation Matrix", labels=dict(color="Correlation"))
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Statistics display is disabled.")
    
    # TAB 6: FILTER
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
                    filters[filter_col] = st.slider(f"Range for {filter_col}", min_v, max_v, (min_v, max_v), key=f"filter_range_{i}")
                else:
                    unique_vals = sorted(df[filter_col].astype(str).unique().tolist())
                    default_vals = unique_vals[:5] if len(unique_vals) > 5 else unique_vals
                    filters[filter_col] = st.multiselect(f"Select values", unique_vals, default=default_vals, key=f"filter_multi_{i}")
        
        filtered_df = df.copy()
        for col, val in filters.items():
            if col in numeric_cols:
                filtered_df = filtered_df[(filtered_df[col] >= val[0]) & (filtered_df[col] <= val[1])]
            else:
                if val:
                    filtered_df = filtered_df[filtered_df[col].astype(str).isin(val)]
        
        if search:
            search_filter = filtered_df.astype(str).apply(lambda x: x.str.contains(search, case=False, na=False)).any(axis=1)
            filtered_df = filtered_df[search_filter]
        
        st.markdown(f"### 📋 Results: {len(filtered_df):,} rows ({len(filtered_df)/len(df)*100:.1f}%)")
        col1, col2 = st.columns([3, 1])
        with col2:
            st.info(f"Original: {len(df):,} rows")
        
        st.dataframe(filtered_df, use_container_width=True, height=500)
        
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
                st.download_button("📊 Excel", excel_filtered, f"filtered_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
    
    # TAB 7: RAW DATA
    with tabs[6]:
        if show_raw:
            st.subheader("📋 Complete Raw Data")
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Rows", f"{len(df):,}")
            with col2:
                st.metric("Columns", len(df.columns))
            with col3:
                st.metric("Missing Values", df.isnull().sum().sum())
            with col4:
                st.metric("Duplicate Rows", df.duplicated().sum())
            
            st.markdown("---")
            search_data = st.text_input("🔍 Search in raw data", key="raw_search")
            if search_data:
                search_filter = df.astype(str).apply(lambda x: x.str.contains(search_data, case=False, na=False)).any(axis=1)
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
                    'Null %': (df.isnull().sum() / len(df) * 100).round(2).tolist(),
                    'Unique': df.nunique().tolist(),
                }
                info_df = pd.DataFrame(info_data)
                st.dataframe(info_df, use_container_width=True, hide_index=True)
            except Exception as e:
                st.warning(f"⚠️ Could not generate column info: {str(e)}")
        else:
            st.info("Raw data display is disabled.")
    
    # TAB 8: EXPORT
    with tabs[7]:
        st.subheader("📥 Export Options")
        st.markdown("### 📊 Export Full Dataset")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            csv = df.to_csv(index=False)
            st.download_button("📄 CSV", csv, f"data_{datetime.now():%Y%m%d_%H%M%S}.csv", "text/csv", use_container_width=True)
        
        with col2:
            json_data = df.to_json(orient='records', indent=2)
            st.download_button("📋 JSON", json_data, f"data_{datetime.now():%Y%m%d_%H%M%S}.json", "application/json", use_container_width=True)
        
        with col3:
            excel_data = create_excel_with_formatting(df, "data")
            if excel_data:
                st.download_button("📊 Excel", excel_data, f"data_{datetime.now():%Y%m%d_%H%M%S}.xlsx", 
                                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
        
        with col4:
            parquet_data = df.to_parquet(index=False)
            st.download_button("⚡ Parquet", parquet_data, f"data_{datetime.now():%Y%m%d_%H%M%S}.parquet", 
                             "application/octet-stream", use_container_width=True)
        
        st.markdown("---")
        st.markdown("### 📋 Export Summary")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Dataset Summary**")
            summary_text = f"- **Rows:** {len(df):,}\n- **Columns:** {len(df.columns)}\n- **File Size:** {df.memory_usage(deep=True).sum() / 1024:.2f} KB\n- **Export Time:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
            st.markdown(summary_text)
        
        with col2:
            st.markdown("**Column Types**")
            type_counts = df.dtypes.value_counts()
            for dtype, count in type_counts.items():
                st.markdown(f"- **{dtype}:** {count} columns")
    
    # AUTO REFRESH
    if auto_refresh:
        st.session_state.refresh_count += 1
        col1, col2 = st.columns([3, 1])
        with col2:
            st.info(f"🔄 Auto-refresh in {refresh_seconds}s...")
        time.sleep(refresh_seconds)
        st.rerun()

else:
    st.markdown("""
    <div class="main-header" style="text-align: center;">
        <h2>👋 Welcome to Ultimate Professional Dashboard</h2>
        <p>Start by uploading a file, providing a URL, using PDF, or using sample data</p>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown(f"""
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
        st.markdown(f"""
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
        st.markdown(f"""
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

### Required Packages

Make sure you have these installed:

```bash
pip install streamlit pandas plotly numpy openpyxl pdfplumber PyPDF2
```

Just copy and paste the code above into your `app.py` file.
