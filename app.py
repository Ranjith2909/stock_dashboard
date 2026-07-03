import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import numpy as np

# Page config
st.set_page_config(
    page_title="Professional Dashboard",
    page_icon="📊",
    layout="wide"
)

# Title
st.title("📊 Professional Data Analytics Dashboard")
st.markdown("### Complete Analysis with KPIs, Pivot Tables & Charts")
st.markdown("---")

# Sidebar
with st.sidebar:
    st.header("⚙️ Settings")
    st.markdown("---")
    st.info("📁 Upload your file to see complete analysis")
    st.markdown("---")
    st.markdown("### Features:")
    st.markdown("✅ KPIs")
    st.markdown("✅ Pivot Tables")
    st.markdown("✅ Multiple Charts")
    st.markdown("✅ Statistics")
    st.markdown("✅ Filters")
    st.markdown("✅ Export Data")

# File Upload
uploaded_file = st.file_uploader(
    "📁 Upload Excel/CSV File", 
    type=['csv', 'xlsx', 'xlsm', 'xls']
)

if uploaded_file:
    try:
        # Read file
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        
        # Clean column names
        df.columns = df.columns.str.strip().str.upper()
        
        # Get numeric and text columns
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        text_cols = df.select_dtypes(include=['object']).columns.tolist()
        
        # Fill missing values
        for col in numeric_cols:
            df[col] = df[col].fillna(0)
        
        # =============== TABS ===============
        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "📈 KPIs", 
            "📊 Charts", 
            "🔄 Pivot Table",
            "📉 Statistics",
            "🔍 Filter & Search",
            "📋 Raw Data",
            "📥 Export"
        ])
        
        # =============== TAB 1: KPIs ===============
        with tab1:
            st.subheader("📈 Key Performance Indicators")
            
            # Row 1: Basic KPIs
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("📦 Total Rows", f"{len(df):,}")
            
            with col2:
                st.metric("📊 Total Columns", len(df.columns))
            
            with col3:
                if numeric_cols:
                    total_sum = df[numeric_cols[0]].sum()
                    st.metric(f"Sum of {numeric_cols[0]}", f"{total_sum:,.0f}")
            
            with col4:
                if numeric_cols:
                    avg_val = df[numeric_cols[0]].mean()
                    st.metric(f"Avg of {numeric_cols[0]}", f"{avg_val:,.2f}")
            
            st.markdown("---")
            
            # Row 2: Stock-specific KPIs (if columns exist)
            st.subheader("💼 Business KPIs")
            col1, col2, col3, col4, col5 = st.columns(5)
            
            with col1:
                if 'CLOSING STOCK' in df.columns:
                    total_stock = df['CLOSING STOCK'].sum()
                    st.metric("Total Stock", f"{total_stock:,.0f}")
            
            with col2:
                if 'OPENING STOCK' in df.columns:
                    opening = df['OPENING STOCK'].sum()
                    st.metric("Opening Stock", f"{opening:,.0f}")
            
            with col3:
                if 'ISSUED' in df.columns:
                    issued = df['ISSUED'].sum()
                    st.metric("Total Issued", f"{issued:,.0f}")
            
            with col4:
                if 'RECEIVED' in df.columns:
                    received = df['RECEIVED'].sum()
                    st.metric("Total Received", f"{received:,.0f}")
            
            with col5:
                if 'CLOSING STOCK' in df.columns and 'SAFETY STOCK' in df.columns:
                    low_stock = len(df[df['CLOSING STOCK'] < df['SAFETY STOCK']])
                    st.metric("🚨 Low Stock Items", low_stock)
            
            st.markdown("---")
            
            # Alerts
            st.subheader("🚨 Alerts & Warnings")
            
            if 'CLOSING STOCK' in df.columns and 'SAFETY STOCK' in df.columns:
                critical = df[df['CLOSING STOCK'] < df['SAFETY STOCK']]
                if len(critical) > 0:
                    st.error(f"🔴 CRITICAL: {len(critical)} items below safety stock!")
                    with st.expander(f"View {len(critical)} critical items"):
                        st.dataframe(critical, use_container_width=True)
                else:
                    st.success("✅ All items are above safety stock!")
            
            # Dead stock
            if 'ISSUED' in df.columns and 'RECEIVED' in df.columns:
                dead = df[(df['ISSUED'] == 0) & (df['RECEIVED'] == 0)]
                if len(dead) > 0:
                    st.warning(f"⚠️ {len(dead)} items have NO movement (Dead Stock)")
                    with st.expander(f"View {len(dead)} dead stock items"):
                        st.dataframe(dead, use_container_width=True)
        
        # =============== TAB 2: CHARTS ===============
        with tab2:
            st.subheader("📊 Multiple Chart Types")
            
            chart_type = st.selectbox(
                "Choose Chart Type",
                ["Bar Chart", "Pie Chart", "Line Chart", "Scatter Plot", "Area Chart", "Histogram", "Box Plot", "Heatmap"]
            )
            
            col1, col2 = st.columns(2)
            with col1:
                x_col = st.selectbox("X-Axis", df.columns.tolist())
            with col2:
                y_col = st.selectbox("Y-Axis", numeric_cols if numeric_cols else df.columns.tolist())
            
            st.markdown("---")
            
            try:
                if chart_type == "Bar Chart":
                    top_n = st.slider("Show top N items", 5, 50, 10)
                    if x_col in text_cols:
                        chart_data = df.nlargest(top_n, y_col) if y_col in numeric_cols else df.head(top_n)
                        fig = px.bar(chart_data, x=x_col, y=y_col, color=y_col, 
                                    color_continuous_scale='Viridis',
                                    title=f"Bar Chart: {y_col} by {x_col}")
                    else:
                        fig = px.bar(df.head(top_n), x=x_col, y=y_col, title=f"Bar Chart")
                    st.plotly_chart(fig, use_container_width=True)
                
                elif chart_type == "Pie Chart":
                    top_n = st.slider("Show top N slices", 3, 20, 10)
                    if x_col in text_cols and y_col in numeric_cols:
                        chart_data = df.nlargest(top_n, y_col)
                        fig = px.pie(chart_data, names=x_col, values=y_col,
                                    title=f"Pie Chart: Distribution of {y_col}")
                        st.plotly_chart(fig, use_container_width=True)
                
                elif chart_type == "Line Chart":
                    fig = px.line(df.head(50), x=x_col, y=y_col, 
                                markers=True, title=f"Line Chart")
                    st.plotly_chart(fig, use_container_width=True)
                
                elif chart_type == "Scatter Plot":
                    color_col = st.selectbox("Color by", [None] + df.columns.tolist())
                    fig = px.scatter(df, x=x_col, y=y_col, color=color_col,
                                    title=f"Scatter Plot")
                    st.plotly_chart(fig, use_container_width=True)
                
                elif chart_type == "Area Chart":
                    fig = px.area(df.head(30), x=x_col, y=y_col, title=f"Area Chart")
                    st.plotly_chart(fig, use_container_width=True)
                
                elif chart_type == "Histogram":
                    fig = px.histogram(df, x=y_col, nbins=30, title=f"Histogram of {y_col}")
                    st.plotly_chart(fig, use_container_width=True)
                
                elif chart_type == "Box Plot":
                    fig = px.box(df, y=y_col, title=f"Box Plot of {y_col}")
                    st.plotly_chart(fig, use_container_width=True)
                
                elif chart_type == "Heatmap":
                    if len(numeric_cols) > 1:
                        corr = df[numeric_cols].corr()
                        fig = px.imshow(corr, text_auto=True, aspect="auto",
                                       color_continuous_scale='RdBu_r',
                                       title="Correlation Heatmap")
                        st.plotly_chart(fig, use_container_width=True)
                    else:
                        st.warning("Need more numeric columns for heatmap")
            
            except Exception as e:
                st.error(f"Error creating chart: {e}")
            
            st.markdown("---")
            
            # Multiple charts in grid
            st.subheader("📊 Auto-Generated Charts")
            
            if len(numeric_cols) >= 2 and len(text_cols) >= 1:
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write(f"**Top 10 by {numeric_cols[0]}**")
                    top10 = df.nlargest(10, numeric_cols[0])
                    fig = px.bar(top10, x=text_cols[0], y=numeric_cols[0], color=numeric_cols[0])
                    st.plotly_chart(fig, use_container_width=True)
                
                with col2:
                    st.write(f"**Distribution of {numeric_cols[0]}**")
                    fig = px.histogram(df, x=numeric_cols[0], nbins=20)
                    st.plotly_chart(fig, use_container_width=True)
        
        # =============== TAB 3: PIVOT TABLE ===============
        with tab3:
            st.subheader("🔄 Pivot Table Analysis")
            st.info("Create pivot tables like Excel!")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                rows = st.selectbox("Rows (Group by)", text_cols if text_cols else df.columns.tolist())
            
            with col2:
                columns = st.selectbox("Columns (Optional)", [None] + text_cols)
            
            with col3:
                values = st.selectbox("Values", numeric_cols if numeric_cols else df.columns.tolist())
            
            agg_func = st.selectbox("Aggregation", ["sum", "mean", "count", "max", "min", "median"])
            
            try:
                if columns:
                    pivot = pd.pivot_table(df, index=rows, columns=columns, values=values, aggfunc=agg_func, fill_value=0)
                else:
                    pivot = pd.pivot_table(df, index=rows, values=values, aggfunc=agg_func, fill_value=0)
                
                st.markdown("### 📊 Pivot Table Result:")
                st.dataframe(pivot, use_container_width=True)
                
                # Download pivot
                pivot_csv = pivot.to_csv()
                st.download_button("📥 Download Pivot Table", pivot_csv, "pivot_table.csv")
                
                # Visualize pivot
                st.markdown("### 📈 Pivot Chart:")
                if columns:
                    fig = px.bar(pivot, barmode='group', title=f"Pivot: {values} by {rows} and {columns}")
                else:
                    fig = px.bar(x=pivot.index, y=pivot.values.flatten(), title=f"Pivot: {values} by {rows}")
                st.plotly_chart(fig, use_container_width=True)
            
            except Exception as e:
                st.error(f"Error creating pivot: {e}")
        
        # =============== TAB 4: STATISTICS ===============
        with tab4:
            st.subheader("📉 Statistical Analysis")
            
            st.markdown("### 📊 Descriptive Statistics")
            st.dataframe(df.describe(), use_container_width=True)
            
            st.markdown("---")
            st.markdown("### 🔢 Column-wise Analysis")
            
            selected_col = st.selectbox("Select Column for Analysis", numeric_cols if numeric_cols else df.columns.tolist())
            
            if selected_col in numeric_cols:
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    st.metric("Sum", f"{df[selected_col].sum():,.2f}")
                    st.metric("Mean", f"{df[selected_col].mean():,.2f}")
                
                with col2:
                    st.metric("Median", f"{df[selected_col].median():,.2f}")
                    st.metric("Mode", f"{df[selected_col].mode()[0] if not df[selected_col].mode().empty else 0:,.2f}")
                
                with col3:
                    st.metric("Max", f"{df[selected_col].max():,.2f}")
                    st.metric("Min", f"{df[selected_col].min():,.2f}")
                
                with col4:
                    st.metric("Std Dev", f"{df[selected_col].std():,.2f}")
                    st.metric("Variance", f"{df[selected_col].var():,.2f}")
                
                # Distribution chart
                fig = px.histogram(df, x=selected_col, nbins=30, title=f"Distribution of {selected_col}", marginal="box")
                st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("---")
            st.markdown("### 🔗 Correlation Matrix")
            
            if len(numeric_cols) > 1:
                corr = df[numeric_cols].corr()
                fig = px.imshow(corr, text_auto=True, aspect="auto", color_continuous_scale='RdBu_r')
                st.plotly_chart(fig, use_container_width=True)
        
        # =============== TAB 5: FILTER & SEARCH ===============
        with tab5:
            st.subheader("🔍 Filter & Search Data")
            
            # Search
            search = st.text_input("🔍 Global Search (search across all columns)")
            
            # Column filters
            st.markdown("### 🎯 Column Filters")
            filter_col = st.selectbox("Filter by Column", df.columns.tolist())
            
            if filter_col in numeric_cols:
                min_val = float(df[filter_col].min())
                max_val = float(df[filter_col].max())
                range_val = st.slider(f"Range for {filter_col}", min_val, max_val, (min_val, max_val))
                filtered_df = df[(df[filter_col] >= range_val[0]) & (df[filter_col] <= range_val[1])]
            else:
                unique_values = df[filter_col].unique().tolist()
                selected = st.multiselect(f"Select {filter_col}", unique_values, default=unique_values[:10] if len(unique_values) > 10 else unique_values)
                filtered_df = df[df[filter_col].isin(selected)]
            
            # Apply search
            if search:
                filtered_df = filtered_df[filtered_df.astype(str).apply(lambda x: x.str.contains(search, case=False)).any(axis=1)]
            
            st.markdown(f"### 📋 Filtered Data ({len(filtered_df)} rows)")
            st.dataframe(filtered_df, use_container_width=True, height=500)
            
            # Download filtered
            csv_filtered = filtered_df.to_csv(index=False)
            st.download_button("📥 Download Filtered Data", csv_filtered, "filtered_data.csv")
        
        # =============== TAB 6: RAW DATA ===============
        with tab6:
            st.subheader("📋 Complete Raw Data")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", len(df))
            with col2:
                st.metric("Total Columns", len(df.columns))
            with col3:
                st.metric("Missing Values", df.isnull().sum().sum())
            
            st.markdown("---")
            st.dataframe(df, use_container_width=True, height=600)
            
            st.markdown("---")
            st.markdown("### 📊 Column Info")
            info_df = pd.DataFrame({
                'Column': df.columns,
                'Type': df.dtypes.astype(str),
                'Non-Null Count': df.count().values,
                'Null Count': df.isnull().sum().values,
                'Unique Values': [df[col].nunique() for col in df.columns]
            })
            st.dataframe(info_df, use_container_width=True)
        
        # =============== TAB 7: EXPORT ===============
        with tab7:
            st.subheader("📥 Export Data")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                csv = df.to_csv(index=False)
                st.download_button(
                    "📄 Download as CSV",
                    csv,
                    f"data_{datetime.now():%Y%m%d_%H%M%S}.csv",
                    "text/csv"
                )
            
            with col2:
                json_data = df.to_json(orient='records', indent=2)
                st.download_button(
                    "📋 Download as JSON",
                    json_data,
                    f"data_{datetime.now():%Y%m%d_%H%M%S}.json",
                    "application/json"
                )
            
            with col3:
                st.info(f"🕐 Last Updated: {datetime.now():%Y-%m-%d %H:%M:%S}")
            
            st.markdown("---")
            st.markdown("### 📊 Summary Report")
            
            summary = f"""
            # Data Summary Report
            
            **Generated:** {datetime.now():%Y-%m-%d %H:%M:%S}
            
            **Total Records:** {len(df):,}
            **Total Columns:** {len(df.columns)}
            
            ## Column Statistics
            {df.describe().to_string()}
            
            ## Column Names
            {', '.join(df.columns.tolist())}
            """
            
            st.download_button(
                "📄 Download Summary Report",
                summary,
                f"report_{datetime.now():%Y%m%d}.txt",
                "text/plain"
            )
    
    except Exception as e:
        st.error(f"❌ Error processing file: {e}")
        st.info("Please check your file format and try again")

else:
    # Landing page
    st.info("👆 Upload an Excel or CSV file to see the complete dashboard")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 📈 KPIs")
        st.markdown("- Total counts")
        st.markdown("- Sum, Average")
        st.markdown("- Business metrics")
        st.markdown("- Alerts & warnings")
    
    with col2:
        st.markdown("### 📊 Charts")
        st.markdown("- Bar, Pie, Line")
        st.markdown("- Scatter, Area")
        st.markdown("- Histogram, Box")
        st.markdown("- Heatmap")
    
    with col3:
        st.markdown("### 🔄 Analysis")
        st.markdown("- Pivot Tables")
        st.markdown("- Statistics")
        st.markdown("- Filters & Search")
        st.markdown("- Export options")
    
    st.markdown("---")
    st.markdown("### 📋 Supported Formats")
    st.markdown("- **Excel:** .xlsx, .xlsm, .xls")
    st.markdown("- **CSV:** .csv")
