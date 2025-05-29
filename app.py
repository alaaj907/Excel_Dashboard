import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

st.set_page_config(page_title="Enhanced Audience Dashboard", layout="wide")

st.title("📊 Enhanced Audience Dashboard")

uploaded_file = st.file_uploader("Upload Excel File", type=[".xlsx"])
if uploaded_file:
    # Read Excel file and allow user to pick a sheet
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Select a sheet", xls.sheet_names)
    df = pd.read_excel(xls, sheet_name=sheet_name)

    # Display raw data
    with st.expander("🔍 Preview Raw Data"):
        st.dataframe(df.head(20))

    # Normalize column names
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_").str.replace("%", "percent")

    # Audience percentage binning
    if 'audience' in df.columns:
        df['audience_bin'] = pd.cut(df['audience'], bins=[0, 10, 25, 50, 75, 90, 99, 100],
                                    labels=["0–10%", "10–25%", "25–50%", "50–75%", "75–90%", "90–99%", "99–100%"])

    # Descriptive statistics
    with st.expander("📈 Descriptive Statistics"):
        desc_stats = df.describe(include='all')
        st.dataframe(desc_stats)
        if 'index' in df.columns and 'relative_lift' in df.columns:
            stats_df = pd.DataFrame({
                'Metric': ['Index', 'Relative Lift'],
                'Mean': [df['index'].mean(), df['relative_lift'].mean()],
                'Std': [df['index'].std(), df['relative_lift'].std()]
            })
            fig, ax = plt.subplots()
            stats_df.set_index("Metric")[['Mean', 'Std']].plot(kind='bar', ax=ax)
            st.pyplot(fig)

    # Statistical Highlighting
    st.subheader("✨ Highlighted Data Table")
    highlight_cols = ['index', 'relative_lift']
    means = df[highlight_cols].mean()
    stds = df[highlight_cols].std()

    def highlight_outliers(val, col):
        if col in highlight_cols:
            if abs(val - means[col]) > 2 * stds[col]:
                return 'background-color: yellow'
        return ''

    styled_df = df.style.applymap(lambda val: highlight_outliers(val, 'index'), subset=['index']) \
                        .applymap(lambda val: highlight_outliers(val, 'relative_lift'), subset=['relative_lift'])
    st.dataframe(styled_df, use_container_width=True)

    # Primary Pie Chart (e.g., attribute group proportions)
    if 'attribute_group' in df.columns:
        st.subheader("🧩 Attribute Group Distribution")
        group_counts = df['attribute_group'].value_counts()
        fig1, ax1 = plt.subplots()
        ax1.pie(group_counts, labels=group_counts.index, autopct='%1.1f%%', startangle=90)
        ax1.axis('equal')
        st.pyplot(fig1)

    # Secondary Pie Chart: Unique ID counts by audience bin
    if 'audience_bin' in df.columns:
        st.subheader("🥧 Audience Bin Distribution (by count)")
        bin_counts = df['audience_bin'].value_counts().sort_index()
        fig2, ax2 = plt.subplots()
        ax2.pie(bin_counts, labels=bin_counts.index, autopct='%1.1f%%', startangle=90)
        ax2.axis('equal')
        st.pyplot(fig2)

    # Volume Analysis by Segment
    if 'attribute_group' in df.columns and 'audience_bin' in df.columns:
        st.subheader("🔬 Volume Analysis by Segment and Audience Bin")
        volume_df = df.groupby(['attribute_group', 'audience_bin']).size().unstack(fill_value=0)
        st.dataframe(volume_df)

        st.subheader("📊 Heatmap of ID Volume by Segment")
        fig3, ax3 = plt.subplots(figsize=(10, 6))
        sns.heatmap(volume_df, cmap="YlGnBu", annot=True, fmt="d")
        st.pyplot(fig3)
else:
    st.info("👈 Please upload an Excel file to get started.")
