"""Streamlit entrypoint for AI-powered business file analyzer."""

from __future__ import annotations

import pandas as pd
import streamlit as st

from config import ALL_SUPPORTED_EXTENSIONS, APP_CONFIG, THEMES
from cleaner import clean_dataframe
from analytics import run_analytics
from dashboard import render_dashboard, render_data_quality, render_statistics
from export import (
    export_csv,
    export_excel,
    export_html_report,
    export_json,
    export_payload,
    export_pdf_report,
    export_powerpoint_report,
    export_word_report,
)
from loader import load_documents
from reports import build_executive_summary, generate_ai_insights


st.set_page_config(
    page_title=APP_CONFIG.app_name,
    page_icon=APP_CONFIG.app_icon,
    layout=APP_CONFIG.page_layout,
)


@st.cache_data(show_spinner=False)
def _concat_documents_tables(documents) -> pd.DataFrame:
    chunks = []
    for doc in documents:
        for table_name, frame in doc.tables.items():
            if frame is not None and not frame.empty:
                tagged = frame.copy()
                tagged["_source_file"] = doc.source_name
                tagged["_source_table"] = table_name
                chunks.append(tagged)

    if not chunks:
        return pd.DataFrame()
    return pd.concat(chunks, ignore_index=True, sort=False)


def apply_theme(theme_name: str) -> None:
    theme = THEMES[theme_name]
    css = f"""
    <style>
      .stApp {{ background: {theme['background']}; color: {theme['text']}; }}
      .block-container {{ padding-top: 1rem; padding-bottom: 1rem; }}
      .stMetric {{ background: {theme['surface']}; border-radius: 10px; padding: 10px; }}
      div[data-testid='stDataFrame'] {{ background: {theme['surface']}; border-radius: 10px; }}
    </style>
    """
    st.markdown(css, unsafe_allow_html=True)


def render_exports(clean_df: pd.DataFrame, summary: str, kpis) -> None:
    st.subheader("Export Center")
    frames = {"clean_data": clean_df}

    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("Export CSV", data=export_csv(clean_df), file_name="analysis.csv", mime="text/csv")
        st.download_button(
            "Export Excel",
            data=export_excel(frames),
            file_name="analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.download_button("Export JSON", data=export_json(clean_df), file_name="analysis.json", mime="application/json")

    with c2:
        st.download_button(
            "Export HTML Report",
            data=export_html_report("AI Business Report", summary, kpis, clean_df),
            file_name="report.html",
            mime="text/html",
        )
        pdf_bytes = export_pdf_report("AI Business Report", summary, kpis)
        st.download_button("Export PDF", data=pdf_bytes, file_name="report.pdf", mime="application/pdf")

    with c3:
        docx_bytes = export_word_report("AI Business Report", summary, kpis)
        pptx_bytes = export_powerpoint_report("AI Business Report", summary, kpis)
        st.download_button(
            "Export Word",
            data=docx_bytes,
            file_name="report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        st.download_button(
            "Export PowerPoint",
            data=pptx_bytes,
            file_name="report.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
        st.download_button(
            "Export API Payload",
            data=export_payload(clean_df, kpis, summary),
            file_name="report_payload.json",
            mime="application/json",
        )


def main() -> None:
    st.title("📊 AI-Powered Business File Analyzer")
    st.caption("Upload business files and get automatic dashboards, KPIs, pivots, AI insights, and exports.")

    with st.sidebar:
        st.header("Input Sources")
        theme = st.selectbox("Theme", list(THEMES.keys()), index=0)
        apply_theme(theme)

        uploaded_files = st.file_uploader(
            "Upload Files",
            type=[ext.strip(".") for ext in ALL_SUPPORTED_EXTENSIONS],
            accept_multiple_files=True,
            help="Supports Excel, CSV, PDF, Word, PowerPoint, JSON, XML, TXT, and Images.",
        )

        remote_url = st.text_input("File URL", placeholder="https://example.com/report.xlsx")

        st.markdown("**Connected Source Placeholders**")
        st.caption("OneDrive / Google Drive / SharePoint / Dropbox can be accessed via public or signed URLs.")

    if not uploaded_files and not remote_url:
        st.info("Upload files or provide a URL to begin automated analysis.")
        return

    with st.spinner("Loading and parsing files..."):
        documents = load_documents(uploaded_files, remote_url)

    if not documents:
        st.error("No valid documents were parsed.")
        return

    raw_df = _concat_documents_tables(documents)
    if raw_df.empty:
        st.warning("No tabular data extracted from uploaded documents. Text extraction may still be available.")
        text_content = []
        for doc in documents:
            text_content.extend(doc.text_blocks)

        if text_content:
            st.text_area("Extracted Text", value="\n\n".join(text_content)[:50000], height=400)
        return

    with st.spinner("Cleaning and profiling data..."):
        clean_df, cleaning_report = clean_dataframe(raw_df)

    analytics = run_analytics(clean_df)
    insights = generate_ai_insights(clean_df, analytics.domain, analytics.kpis)
    exec_summary = build_executive_summary(insights)

    st.success("Analysis complete.")

    render_dashboard(clean_df, analytics.kpis, insights, exec_summary)

    quality_df = pd.DataFrame(
        {
            "Metric": [
                "Rows Before",
                "Rows After",
                "Columns Before",
                "Columns After",
                "Duplicates Removed",
                "Blank Rows Removed",
                "Blank Cols Removed",
            ],
            "Value": [
                cleaning_report.rows_before,
                cleaning_report.rows_after,
                cleaning_report.columns_before,
                cleaning_report.columns_after,
                cleaning_report.duplicates_removed,
                cleaning_report.blank_rows_removed,
                cleaning_report.blank_cols_removed,
            ],
        }
    )

    render_data_quality(cleaning_report.notes, quality_df)
    render_statistics(analytics.descriptive_stats, analytics.correlation, analytics.anomalies)
    render_exports(clean_df, exec_summary, analytics.kpis)


if __name__ == "__main__":
    main()
