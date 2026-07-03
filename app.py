import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="Stock Dashboard", page_icon="📊", layout="wide")

st.title("📊 Live Inventory Dashboard")
st.markdown("---")

with st.sidebar:
    st.header("⚙️ Settings")
    refresh = st.slider("Refresh (sec)", 10, 300, 30)

uploaded_file = st.file_uploader("📁 Upload Excel/CSV", type=['csv', 'xlsx', 'xls'])

if uploaded_file:
    try:
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        df.columns = df.columns.str.strip().str.upper()
        
        st.subheader("📈 KPIs")
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Total Items", len(df))
        with col2:
            stock = df['CLOSING STOCK'].sum() if 'CLOSING STOCK' in df.columns else 0
            st.metric("Total Stock", f"{stock:,.0f}")
        with col3:
            if 'CLOSING STOCK' in df.columns and 'SAFETY STOCK' in df.columns:
                low = len(df[df['CLOSING STOCK'] < df['SAFETY STOCK']])
                st.metric("🚨 Low Stock", low)
        with col4:
            st.metric("Status", "✅ Active")
        
        st.markdown("---")
        
        if 'CLOSING STOCK' in df.columns and 'SAFETY STOCK' in df.columns:
            critical = df[df['CLOSING STOCK'] < df['SAFETY STOCK']]
            if len(critical) > 0:
                st.error(f"🔴 {len(critical)} items below safety stock!")
                for _, row in critical.iterrows():
                    item = row.get('ITEM', row.get('ITEM NAME', 'Unknown'))
                    st.warning(f"{item}: Current={row['CLOSING STOCK']}, Safety={row['SAFETY STOCK']}")
            else:
                st.success("✅ All items OK!")
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if 'ITEM' in df.columns and 'CLOSING STOCK' in df.columns:
                top = df.nlargest(10, 'CLOSING STOCK')
                fig = px.bar(top, x='ITEM', y='CLOSING STOCK', title="Top 10 Items")
                st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            if 'CLOSING STOCK' in df.columns and 'SAFETY STOCK' in df.columns:
                in_stock = len(df[df['CLOSING STOCK'] >= df['SAFETY STOCK']])
                low = len(df[(df['CLOSING STOCK'] < df['SAFETY STOCK']) & (df['CLOSING STOCK'] > 0)])
                out = len(df[df['CLOSING STOCK'] == 0])
                fig = px.pie(values=[in_stock, low, out], names=['In Stock', 'Low', 'Out'], title="Status")
                st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        st.subheader("📋 Data Table")
        search = st.text_input("🔍 Search")
        if search:
            df_show = df[df.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]
        else:
            df_show = df
        st.dataframe(df_show, use_container_width=True)
        
        csv = df.to_csv(index=False)
        st.download_button("📥 Download CSV", csv, f"data_{datetime.now():%Y%m%d}.csv")
        
    except Exception as e:
        st.error(f"Error: {e}")
else:
    st.info("👆 Upload Excel or CSV to start")
