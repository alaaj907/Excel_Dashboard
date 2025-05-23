import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Audience Analytics Dashboard", layout="wide")
st.title("Audience Analytics Dashboard")
st.markdown("""
<style>
    .css-1d391kg { padding-top: 1rem; }
    .css-1v0mbdj { padding-top: 0rem; }
    .main .block-container { padding-top: 2rem; }
    .reportview-container .markdown-text-container p {
        font-size: 1.1rem;
    }
</style>
""", unsafe_allow_html=True)

uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx"])

def clean_dataframe(df):
    df.columns = df.columns.str.strip().str.lower().str.replace(' ', '_').str.replace('&', 'and')
    return df

def summarize_data(df):
    group_summary = df.groupby('attribute_group').agg(
        count=('attribute_name', 'count'),
        avg_index=('index', 'mean'),
        avg_lift=('relative_lift', 'mean'),
        total_size=('attribute_size', 'sum')
    ).reset_index().sort_values(by='avg_index', ascending=False)
    return group_summary

def download_excel(df):
    return df.to_csv(index=False).encode('utf-8')

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    sheet = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet)
    df = clean_dataframe(df)

    st.markdown("### Key Metrics")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    kpi1.metric("Total Attributes", len(df))
    kpi2.metric("Avg Index Score", f"{df['index'].mean():.1f}")
    kpi3.metric("High Performers (Index > 120)", (df['index'] > 120).sum())
    kpi4.metric("Attribute Groups", df['attribute_group'].nunique())

    st.markdown("### Filter by Group")
    df['attribute_group'] = df['attribute_group'].fillna('Unknown').astype(str)
    group_options = ['All'] + sorted(df['attribute_group'].unique())
    group_filter = st.selectbox("Select Attribute Group", options=group_options)
    if group_filter != 'All':
        df = df[df['attribute_group'] == group_filter]

    st.markdown("### Top Performing Attributes")
    top_performers = df.sort_values(by='index', ascending=False).head(10)
    fig_bar = px.bar(top_performers, x='attribute_name', y='index', color='attribute_group',
                     labels={'index': 'Index Score'}, height=400)
    st.plotly_chart(fig_bar, use_container_width=True)

    st.markdown("### Group Performance Summary")
    group_summary = summarize_data(df)
    fig_pie = px.pie(group_summary.head(8), values='count', names='attribute_group',
                     title='Top Attribute Groups by Count')
    st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("### Index vs Relative Lift")
    fig_scatter = px.scatter(df, x='index', y='relative_lift', color='attribute_group',
                             hover_data=['attribute_name'], height=400)
    st.plotly_chart(fig_scatter, use_container_width=True)

    st.markdown("### Full Data Table")
    st.dataframe(df)

    csv_data = download_excel(df)
    st.download_button("Download Filtered Data as CSV", data=csv_data, file_name="filtered_data.csv", mime="text/csv")

else:
    st.info("Please upload an Excel file to begin.")
