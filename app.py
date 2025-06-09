import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import tempfile
import os
import numpy as np

# Page configuration
st.set_page_config(page_title="Advanced Audience Analytics Dashboard", layout="wide", initial_sidebar_state="expanded")

# Custom CSS for better styling
st.markdown("""
<style>
    .metric-card {
        background-color: #f8f9fa;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #dc3545;
    }
    .insight-box {
        background-color: #e3f2fd;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #2196f3;
        margin: 1rem 0;
    }
    .warning-box {
        background-color: #fff3e0;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #ff9800;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

st.title("🎯 Advanced Audience Analytics Dashboard")
st.markdown("*Comprehensive analysis with performance segmentation, geographic insights, and strategic recommendations*")

def create_performance_buckets(df):
    """Create performance segmentation"""
    buckets = {
        'very_high': df[df['index'] > 150],
        'high': df[(df['index'] > 120) & (df['index'] <= 150)],
        'medium': df[(df['index'] >= 80) & (df['index'] <= 120)],
        'low': df[df['index'] < 80]
    }
    return buckets

def create_comprehensive_ppt(df, charts_data, insights):
    """Create an enhanced professional PowerPoint presentation"""
    prs = Presentation()
    slide_layout = prs.slide_layouts[6]  # Blank layout
    slide1 = prs.slides.add_slide(slide_layout)
    
    # Set background
    background = slide1.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(248, 249, 250)
    
    # Title
    title_box = slide1.shapes.add_textbox(Inches(0.5), Inches(0.2), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "Growing Households - Executive Dashboard"
    title_para = title_frame.paragraphs[0]
    title_para.alignment = PP_ALIGN.CENTER
    title_para.font.size = Pt(32)
    title_para.font.bold = True
    title_para.font.color.rgb = RGBColor(220, 53, 69)
    
    # Key metrics section
    metrics = [
        ("Total Attributes", f"{len(df):,}"),
        ("Avg Index Score", f"{df['index'].mean():.1f}"),
        ("High Performers", f"{len(df[df['index'] > 120]):,}"),
        ("Coverage", f"{insights.get('total_audience', 'N/A')}")
    ]
    
    for i, (label, value) in enumerate(metrics):
        x_pos = 0.5 + (i * 2.25)
        box = slide1.shapes.add_textbox(Inches(x_pos), Inches(1.2), Inches(2), Inches(1))
        box_frame = box.text_frame
        box_frame.text = label
        box_frame.paragraphs[0].font.size = Pt(11)
        box_frame.paragraphs[0].font.bold = True
        
        value_para = box_frame.add_paragraph()
        value_para.text = value
        value_para.font.size = Pt(16)
        value_para.font.bold = True
        value_para.font.color.rgb = RGBColor(220, 53, 69)
    
    # Key insights box
    insights_box = slide1.shapes.add_textbox(Inches(0.5), Inches(5.8), Inches(9), Inches(1.2))
    insights_frame = insights_box.text_frame
    insights_frame.text = "KEY INSIGHTS"
    insights_frame.paragraphs[0].font.size = Pt(14)
    insights_frame.paragraphs[0].font.bold = True
    
    key_insights = [
        f"• Growing households show strong affinity (Index: {insights.get('top_index', 'N/A')})",
        "• Lower-middle income segments over-index significantly",
        "• Young families with children 0-2 are primary targets",
        f"• {insights.get('high_performer_count', 0)} attributes exceed 120 index threshold"
    ]
    
    for insight in key_insights:
        para = insights_frame.add_paragraph()
        para.text = insight
        para.font.size = Pt(10)
    
    # Save presentation
    temp_ppt = tempfile.NamedTemporaryFile(delete=False, suffix=".pptx")
    prs.save(temp_ppt.name)
    return temp_ppt.name

def analyze_audience_insights(df):
    """Generate comprehensive audience insights"""
    insights = {}
    insights['total_attributes'] = len(df)
    insights['avg_index'] = df['index'].mean()
    insights['high_performer_count'] = len(df[df['index'] > 120])
    
    # Top performer
    top_performer = df.loc[df['index'].idxmax()]
    insights['top_index'] = top_performer['index']
    insights['top_attribute'] = top_performer['attribute_name']
    
    # Total audience estimation
    if 'attribute_size' in df.columns:
        max_audience = df['attribute_size'].max()
        insights['total_audience'] = f"{max_audience:,.0f}" if pd.notna(max_audience) else "N/A"
    
    return insights

# Sidebar for navigation
st.sidebar.title("📊 Dashboard Navigation")
analysis_sections = [
    "🏠 Overview",
    "📈 Performance Analysis", 
    "💰 Income Profiling",
    "👨‍👩‍👧‍👦 Family Lifecycle",
    "🗺️ Geographic Analysis",
    "🎯 Audience Overlap",
    "📋 Category Performance",
    "⚠️ Exclusion Opportunities",
    "🧮 Sizing Intelligence",
    "📑 Strategic Recommendations"
]

selected_section = st.sidebar.selectbox("Select Analysis Section", analysis_sections)

# File upload
uploaded_file = st.file_uploader("Upload Excel File with Index Report", type=["xlsx"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        if "Index Report" in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name="Index Report")
            
            # Clean column names
            df.columns = df.columns.str.strip().str.lower().str.replace(" ", "_").str.replace("&", "and")
            
            # Create performance buckets
            performance_buckets = create_performance_buckets(df)
            
            # Generate insights
            insights = analyze_audience_insights(df)
            
            # Navigation logic
            if selected_section == "🏠 Overview":
                st.header("📊 Executive Overview")
                
                # Key metrics row
                col1, col2, col3, col4, col5 = st.columns(5)
                
                with col1:
                    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                    st.metric("Total Attributes", f"{len(df):,}")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                with col2:
                    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                    st.metric("Avg Index Score", f"{df['index'].mean():.1f}")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                with col3:
                    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                    st.metric("High Performers (>120)", f"{len(performance_buckets['high']) + len(performance_buckets['very_high']):,}")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                with col4:
                    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                    st.metric("Performance Rate", f"{((len(performance_buckets['high']) + len(performance_buckets['very_high'])) / len(df) * 100):.1f}%")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                with col5:
                    st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                    st.metric("Top Index Score", f"{df['index'].max():.1f}")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                # Performance distribution
                st.subheader("🎯 Performance Distribution")
                col1, col2 = st.columns([1, 1])
                
                with col1:
                    bucket_counts = {
                        'Very High (>150)': len(performance_buckets['very_high']),
                        'High (120-150)': len(performance_buckets['high']),
                        'Medium (80-120)': len(performance_buckets['medium']),
                        'Low (<80)': len(performance_buckets['low'])
                    }
                    
                    fig_pie = px.pie(
                        values=list(bucket_counts.values()),
                        names=list(bucket_counts.keys()),
                        title="Performance Segmentation",
                        color_discrete_sequence=['#dc3545', '#fd7e14', '#ffc107', '#6c757d']
                    )
                    st.plotly_chart(fig_pie, use_container_width=True)
                
                with col2:
                    performance_df = pd.DataFrame([
                        {'Segment': 'Very High (>150)', 'Count': len(performance_buckets['very_high']), 'Percentage': f"{len(performance_buckets['very_high'])/len(df)*100:.1f}%"},
                        {'Segment': 'High (120-150)', 'Count': len(performance_buckets['high']), 'Percentage': f"{len(performance_buckets['high'])/len(df)*100:.1f}%"},
                        {'Segment': 'Medium (80-120)', 'Count': len(performance_buckets['medium']), 'Percentage': f"{len(performance_buckets['medium'])/len(df)*100:.1f}%"},
                        {'Segment': 'Low (<80)', 'Count': len(performance_buckets['low']), 'Percentage': f"{len(performance_buckets['low'])/len(df)*100:.1f}%"}
                    ])
                    st.dataframe(performance_df, use_container_width=True)
                
                # Key insights
                st.markdown('<div class="insight-box">', unsafe_allow_html=True)
                st.markdown("**🔍 Key Insights:**")
                st.markdown(f"• **Growing Households** is your top performer with an index of {insights['top_index']:.1f}")
                st.markdown(f"• **{insights['high_performer_count']:,} attributes** exceed the 120 index threshold")
                st.markdown(f"• **{len(performance_buckets['low']):,} attributes** under-perform and may need exclusion")
                st.markdown(f"• Average index of **{insights['avg_index']:.1f}** indicates moderate audience alignment")
                st.markdown('</div>', unsafe_allow_html=True)
            
            elif selected_section == "📈 Performance Analysis":
                st.header("📈 Performance Analysis Deep Dive")
                
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    fig_hist = px.histogram(
                        df, x='index', nbins=50,
                        title='Index Score Distribution',
                        labels={'index': 'Index Score', 'count': 'Number of Attributes'}
                    )
                    fig_hist.add_vline(x=120, line_dash="dash", line_color="red", annotation_text="High Performance Threshold")
                    fig_hist.add_vline(x=80, line_dash="dash", line_color="orange", annotation_text="Low Performance Threshold")
                    st.plotly_chart(fig_hist, use_container_width=True)
                
                with col2:
                    st.subheader("Performance Statistics")
                    stats_df = df['index'].describe().round(2)
                    st.dataframe(stats_df.to_frame('Index Statistics'))
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("🏆 Top 10 Performers")
                    top_performers = df.nlargest(10, 'index')[['attribute_name', 'attribute_group', 'index', 'relative_lift']]
                    top_performers['index'] = top_performers['index'].round(1)
                    top_performers['relative_lift'] = top_performers['relative_lift'].round(1)
                    st.dataframe(top_performers, use_container_width=True)
                
                with col2:
                    st.subheader("📉 Bottom 10 Performers")
                    bottom_performers = df.nsmallest(10, 'index')[['attribute_name', 'attribute_group', 'index']]
                    bottom_performers['index'] = bottom_performers['index'].round(1)
                    st.dataframe(bottom_performers, use_container_width=True)
            
            elif selected_section == "💰 Income Profiling":
                st.header("💰 Income Profile Analysis")
                
                if 'attribute_group' in df.columns:
                    income_data = df[df['attribute_group'].str.contains('Income', na=False)]
                    
                    if not income_data.empty:
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            income_sorted = income_data.sort_values('index', ascending=False)
                            
                            fig_income = go.Figure()
                            fig_income.add_trace(go.Bar(
                                x=income_sorted['attribute_name'],
                                y=income_sorted['index'],
                                marker_color=['#dc3545' if x > 120 else '#ffc107' for x in income_sorted['index']],
                                text=income_sorted['index'].round(1),
                                textposition='outside'
                            ))
                            
                            fig_income.update_layout(
                                title="Income Segments - Index Performance",
                                xaxis_title="Income Brackets",
                                yaxis_title="Index Score",
                                xaxis_tickangle=-45,
                                height=500
                            )
                            fig_income.add_hline(y=120, line_dash="dash", line_color="red", annotation_text="High Performance")
                            st.plotly_chart(fig_income, use_container_width=True)
                        
                        with col2:
                            st.subheader("Income Insights")
                            avg_income_index = income_data['index'].mean()
                            high_income_performers = len(income_data[income_data['index'] > 120])
                            
                            st.metric("Avg Income Index", f"{avg_income_index:.1f}")
                            st.metric("High Performing Segments", high_income_performers)
                            
                            income_summary = income_data[['attribute_name', 'index', 'relative_lift']].sort_values('index', ascending=False)
                            income_summary['index'] = income_summary['index'].round(1)
                            income_summary['relative_lift'] = income_summary['relative_lift'].round(1)
                            st.dataframe(income_summary, use_container_width=True)
                        
                        st.markdown('<div class="insight-box">', unsafe_allow_html=True)
                        st.markdown("**💡 Income Profile Insights:**")
                        st.markdown("• **Lower-middle income households** show highest affinity")
                        st.markdown("• **Inverse relationship** between income level and index performance")
                        st.markdown(f"• **{high_income_performers} income segments** exceed 120 index threshold")
                        st.markdown("• Target households in **$15K-$75K** range for maximum efficiency")
                        st.markdown('</div>', unsafe_allow_html=True)
                    else:
                        st.warning("No income data found in the dataset.")
                else:
                    st.warning("Attribute group column not found.")
            
            elif selected_section == "👨‍👩‍👧‍👦 Family Lifecycle":
                st.header("👨‍👩‍👧‍👦 Family Lifecycle Analysis")
                
                family_keywords = ['child', 'children', 'family', 'age', 'parent', 'household']
                family_data = df[df['attribute_group'].str.contains('|'.join(family_keywords), case=False, na=False)]
                
                if not family_data.empty:
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        family_grouped = family_data.groupby('attribute_group').agg({
                            'index': ['mean', 'count']
                        }).round(1)
                        
                        family_grouped.columns = ['avg_index', 'count']
                        family_grouped = family_grouped.reset_index()
                        
                        fig_family = px.bar(
                            family_grouped, 
                            x='attribute_group', 
                            y='avg_index',
                            title='Family Lifecycle Stage Performance',
                            color='avg_index',
                            color_continuous_scale='RdYlBu_r'
                        )
                        fig_family.update_xaxes(tickangle=-45)
                        st.plotly_chart(fig_family, use_container_width=True)
                    
                    with col2:
                        st.subheader("Family Insights")
                        st.metric("Family-Related Attributes", len(family_data))
                        st.metric("Avg Family Index", f"{family_data['index'].mean():.1f}")
                        
                        top_family = family_data.nlargest(5, 'index')[['attribute_name', 'index']]
                        top_family['index'] = top_family['index'].round(1)
                        st.dataframe(top_family, use_container_width=True)
                else:
                    st.warning("No family-related data found in the dataset.")
            
            elif selected_section == "🗺️ Geographic Analysis":
                st.header("🗺️ Geographic Performance Analysis")
                
                geo_keywords = ['state', 'census', 'geographic', 'region', 'zip']
                geo_data = df[df['attribute_group'].str.contains('|'.join(geo_keywords), case=False, na=False)]
                
                if not geo_data.empty:
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        top_geo = geo_data.nlargest(20, 'index')
                        
                        fig_geo = px.bar(
                            top_geo,
                            x='index',
                            y='attribute_name',
                            orientation='h',
                            title='Top 20 Geographic Areas by Index',
                            color='index',
                            color_continuous_scale='RdYlBu_r'
                        )
                        fig_geo.update_layout(height=600, yaxis={'categoryorder': 'total ascending'})
                        st.plotly_chart(fig_geo, use_container_width=True)
                    
                    with col2:
                        st.subheader("Geographic Insights")
                        st.metric("Geographic Attributes", len(geo_data))
                        st.metric("Avg Geographic Index", f"{geo_data['index'].mean():.1f}")
                        st.metric("High Performing Areas", len(geo_data[geo_data['index'] > 120]))
                        
                        geo_performance = {
                            'High (>120)': len(geo_data[geo_data['index'] > 120]),
                            'Medium (80-120)': len(geo_data[(geo_data['index'] >= 80) & (geo_data['index'] <= 120)]),
                            'Low (<80)': len(geo_data[geo_data['index'] < 80])
                        }
                        
                        geo_df = pd.DataFrame(list(geo_performance.items()), columns=['Performance', 'Count'])
                        st.dataframe(geo_df, use_container_width=True)
                    
                    st.markdown('<div class="insight-box">', unsafe_allow_html=True)
                    st.markdown("**🗺️ Geographic Targeting Insights:**")
                    top_state = geo_data.loc[geo_data['index'].idxmax()]
                    st.markdown(f"• **{top_state['attribute_name']}** is the top performing area (Index: {top_state['index']:.1f})")
                    high_geo_count = len(geo_data[geo_data['index'] > 120])
                    st.markdown(f"• **{high_geo_count} geographic areas** exceed 120 index threshold")
                    st.markdown("• Focus media spend and targeting on high-indexing regions")
                    st.markdown("• Consider geo-fencing and location-based campaigns")
                    st.markdown('</div>', unsafe_allow_html=True)
                else:
                    st.warning("No geographic data found in the dataset.")
            
            elif selected_section == "🎯 Audience Overlap":
                st.header("🎯 Audience Overlap & Penetration Analysis")
                
                if 'audience_attribute_proportion' in df.columns and 'attribute_size' in df.columns:
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        sample_df = df.sample(min(1000, len(df)))
                        fig_scatter = px.scatter(
                            sample_df,
                            x='audience_attribute_proportion',
                            y='index',
                            size='attribute_size',
                            color='attribute_group',
                            title='Penetration vs Performance Analysis',
                            labels={
                                'audience_attribute_proportion': 'Audience Penetration Rate',
                                'index': 'Index Score',
                                'attribute_size': 'Audience Size'
                            },
                            hover_data=['attribute_name']
                        )
                        fig_scatter.add_hline(y=120, line_dash="dash", line_color="red")
                        st.plotly_chart(fig_scatter, use_container_width=True)
                    
                    with col2:
                        st.subheader("Overlap Insights")
                        
                        high_pen_high_perf = df[
                            (df['audience_attribute_proportion'] > df['audience_attribute_proportion'].quantile(0.75)) &
                            (df['index'] > 120)
                        ]
                        
                        st.metric("High Penetration + High Performance", len(high_pen_high_perf))
                        st.metric("Avg Penetration Rate", f"{df['audience_attribute_proportion'].mean():.3f}")
                        
                        if len(high_pen_high_perf) > 0:
                            st.subheader("Best Overlap Opportunities")
                            best_overlap = high_pen_high_perf.nlargest(5, 'index')[['attribute_name', 'index', 'audience_attribute_proportion']]
                            best_overlap['index'] = best_overlap['index'].round(1)
                            best_overlap['audience_attribute_proportion'] = best_overlap['audience_attribute_proportion'].round(4)
                            st.dataframe(best_overlap, use_container_width=True)
                else:
                    st.warning("Required columns for overlap analysis not found.")
            
            elif selected_section == "📋 Category Performance":
                st.header("📋 Category Performance Leaderboard")
                
                if 'attribute_group' in df.columns:
                    category_stats = df.groupby('attribute_group').agg({
                        'index': ['mean', 'count', 'std', 'min', 'max']
                    }).round(1)
                    
                    category_stats.columns = ['avg_index', 'count', 'std_dev', 'min_index', 'max_index']
                    category_stats = category_stats[category_stats['count'] >= 2]
                    category_stats['high_performers'] = df.groupby('attribute_group').apply(lambda x: len(x[x['index'] > 120]))
                    category_stats = category_stats.sort_values('avg_index', ascending=False)
                    
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        top_categories = category_stats.head(20).reset_index()
                        
                        fig_cat = go.Figure()
                        fig_cat.add_trace(go.Bar(
                            x=top_categories['attribute_group'],
                            y=top_categories['avg_index'],
                            marker_color=['#dc3545' if x > 120 else '#ffc107' if x > 100 else '#6c757d' for x in top_categories['avg_index']],
                            text=top_categories['avg_index'],
                            textposition='outside',
                            name='Average Index'
                        ))
                        
                        fig_cat.update_layout(
                            title="Top 20 Category Performance Rankings",
                            xaxis_title="Categories",
                            yaxis_title="Average Index Score",
                            height=600,
                            xaxis_tickangle=-45,
                            margin=dict(b=200)
                        )
                        fig_cat.add_hline(y=120, line_dash="dash", line_color="red", annotation_text="High Performance")
                        st.plotly_chart(fig_cat, use_container_width=True)
                    
                    with col2:
                        st.subheader("Category Insights")
                        st.metric("Total Categories", len(category_stats))
                        st.metric("High Performing Categories", len(category_stats[category_stats['avg_index'] > 120]))
                        st.metric("Categories with Consistency", len(category_stats[category_stats['std_dev'] < 20]))
                        
                        st.subheader("Top 10 Categories")
                        top_10_display = category_stats.head(10)[['avg_index', 'count', 'high_performers']].reset_index()
                        top_10_display.columns = ['Category', 'Avg Index', 'Total Attrs', 'High Performers']
                        st.dataframe(top_10_display, use_container_width=True)
                else:
                    st.warning("Attribute group column not found.")
            
            elif selected_section == "⚠️ Exclusion Opportunities":
                st.header("⚠️ Exclusion Opportunities & Optimization")
                
                low_performers = performance_buckets['low']
                
                col1, col2 = st.columns([2, 1])
                
                with col1:
                    if not low_performers.empty:
                        low_perf_by_category = low_performers.groupby('attribute_group').agg({
                            'index': ['mean', 'count']
                        }).round(1)
                        
                        low_perf_by_category.columns = ['avg_index', 'count']
                        low_perf_by_category = low_perf_by_category.sort_values('count', ascending=False).head(15)
                        low_perf_by_category = low_perf_by_category.reset_index()
                        
                        fig_exclusion = px.bar(
                            low_perf_by_category,
                            x='attribute_group',
                            y='count',
                            title='Categories with Most Low Performers (Index < 80)',
                            color='avg_index',
                            color_continuous_scale='RdYlBu'
                        )
                        fig_exclusion.update_xaxes(tickangle=-45)
                        st.plotly_chart(fig_exclusion, use_container_width=True)
                    else:
                        st.info("No low performers found in the dataset.")
                
                with col2:
                    st.subheader("Exclusion Insights")
                    st.metric("Total Low Performers", len(low_performers))
                    if not low_performers.empty:
                        st.metric("Categories Affected", low_performers['attribute_group'].nunique())
                        st.metric("Potential Budget Savings", f"{(len(low_performers)/len(df)*100):.1f}%")
                        
                        st.subheader("Worst 10 Performers")
                        worst_performers = low_performers.nsmallest(10, 'index')[['attribute_name', 'attribute_group', 'index']]
                        worst_performers['index'] = worst_performers['index'].round(1)
                        st.dataframe(worst_performers, use_container_width=True)
                
                if not low_performers.empty:
                    st.markdown('<div class="warning-box">', unsafe_allow_html=True)
                    st.markdown("**⚠️ Exclusion Recommendations:**")
                    worst_category = low_performers.groupby('attribute_group').size().idxmax()
                    worst_count = low_performers.groupby('attribute_group').size().max()
                    st.markdown(f"• **{worst_category}** has the most low performers ({worst_count} attributes)")
                    st.markdown(f"• Consider excluding **{len(low_performers):,} attributes** with index < 80")
                    st.markdown(f"• This could improve overall efficiency by **{(len(low_performers)/len(df)*100):.1f}%**")
                    st.markdown("• Review budget allocation from low performers to high performers")
                    st.markdown('</div>', unsafe_allow_html=True)
            
            elif selected_section == "🧮 Sizing Intelligence":
                st.header("🧮 Audience Sizing Intelligence")
                
                if 'attribute_size' in df.columns:
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        sample_df = df.sample(min(1000, len(df)))
                        
                        fig_reach = px.scatter(
                            sample_df,
                            x='attribute_size',
                            y='index',
                            color='attribute_group',
                            size='index',
                            title='Reach vs Relevance Trade-off Analysis',
                            labels={
                                'attribute_size': 'Audience Size (Reach)',
                                'index': 'Index Score (Relevance)'
                            },
                            hover_data=['attribute_name']
                        )
                        fig_reach.add_hline(y=120, line_dash="dash", line_color="red", annotation_text="High Relevance")
                        fig_reach.update_xaxes(type="log")
                        st.plotly_chart(fig_reach, use_container_width=True)
                    
                    with col2:
                        st.subheader("Sizing Insights")
                        
                        high_reach_high_rel = df[(df['attribute_size'] > df['attribute_size'].quantile(0.75)) & (df['index'] > 120)]
                        low_reach_high_rel = df[(df['attribute_size'] < df['attribute_size'].quantile(0.25)) & (df['index'] > 120)]
                        
                        st.metric("High Reach + High Relevance", len(high_reach_high_rel))
                        st.metric("Low Reach + High Relevance", len(low_reach_high_rel))
                        st.metric("Sweet Spot Opportunities", len(df[(df['attribute_size'] > 100000) & (df['index'] > 150)]))
                        
                        st.subheader("Optimal Audience Sizes")
                        size_buckets = pd.cut(df['attribute_size'], bins=5, labels=['XS', 'S', 'M', 'L', 'XL'])
                        size_performance = df.groupby(size_buckets)['index'].mean().round(1)
                        st.dataframe(size_performance.to_frame('Avg Index'), use_container_width=True)
                else:
                    st.warning("Attribute size column not found.")
            
            elif selected_section == "📑 Strategic Recommendations":
                st.header("📑 Strategic Recommendations & Action Plan")
                
                recommendations = []
                
                if len(performance_buckets['very_high']) > 0:
                    top_performer = performance_buckets['very_high'].iloc[0]
                    recommendations.append({
                        'Priority': 'High',
                        'Category': 'Performance Optimization',
                        'Recommendation': f"Prioritize '{top_performer['attribute_name']}' segment (Index: {top_performer['index']:.1f})",
                        'Impact': 'High ROI',
                        'Implementation': 'Immediate'
                    })
                
                if 'attribute_group' in df.columns:
                    income_data = df[df['attribute_group'].str.contains('Income', na=False)]
                    if not income_data.empty:
                        top_income = income_data.loc[income_data['index'].idxmax()]
                        recommendations.append({
                            'Priority': 'High',
                            'Category': 'Income Targeting',
                            'Recommendation': f"Focus on {top_income['attribute_name']} segment",
                            'Impact': 'Improved targeting efficiency',
                            'Implementation': '2-4 weeks'
                        })
                
                if len(performance_buckets['low']) > 0:
                    recommendations.append({
                        'Priority': 'Medium',
                        'Category': 'Budget Optimization',
                        'Recommendation': f"Exclude {len(performance_buckets['low'])} low-performing attributes (Index < 80)",
                        'Impact': f"{(len(performance_buckets['low'])/len(df)*100):.1f}% budget reallocation",
                        'Implementation': '1-2 weeks'
                    })
                
                geo_data = df[df['attribute_group'].str.contains('State|Geographic|Census', case=False, na=False)]
                if not geo_data.empty:
                    top_geo = geo_data.loc[geo_data['index'].idxmax()]
                    recommendations.append({
                        'Priority': 'Medium',
                        'Category': 'Geographic Focus',
                        'Recommendation': f"Increase investment in {top_geo['attribute_name']} market",
                        'Impact': 'Geographic optimization',
                        'Implementation': '2-3 weeks'
                    })
                
                family_data = df[df['attribute_group'].str.contains('Children|Family', case=False, na=False)]
                if not family_data.empty and family_data['index'].mean() > 110:
                    recommendations.append({
                        'Priority': 'High',
                        'Category': 'Demographic Targeting',
                        'Recommendation': 'Develop family-focused creative and messaging',
                        'Impact': 'Improved engagement',
                        'Implementation': '4-6 weeks'
                    })
                
                if recommendations:
                    rec_df = pd.DataFrame(recommendations)
                    
                    for priority in ['High', 'Medium', 'Low']:
                        priority_recs = rec_df[rec_df['Priority'] == priority]
                        if not priority_recs.empty:
                            st.subheader(f"🎯 {priority} Priority Actions")
                            st.dataframe(priority_recs.drop('Priority', axis=1), use_container_width=True)
                
                st.subheader("🚀 Campaign Optimization Framework")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.markdown("**📈 Short-term Actions (1-4 weeks)**")
                    st.markdown("• Exclude low-performing attributes (Index < 80)")
                    st.markdown("• Increase budget allocation to top 10 performers")
                    st.markdown("• Implement negative targeting for under-performers")
                    st.markdown("• A/B test high vs medium performance segments")
                
                with col2:
                    st.markdown("**🎯 Long-term Strategy (1-6 months)**")
                    st.markdown("• Develop lookalike audiences from top performers")
                    st.markdown("• Create family-focused creative strategies")
                    st.markdown("• Implement sequential targeting by performance tiers")
                    st.markdown("• Build predictive models for audience expansion")
                
                st.subheader("📊 Success Metrics & KPIs")
                
                kpi_data = {
                    'KPI': [
                        'Overall Index Improvement',
                        'High Performer Concentration',
                        'Budget Efficiency Gain',
                        'Audience Quality Score',
                        'Geographic Performance Ratio'
                    ],
                    'Current State': [
                        f"{df['index'].mean():.1f}",
                        f"{((len(performance_buckets['high']) + len(performance_buckets['very_high']))/len(df)*100):.1f}%",
                        "Baseline",
                        f"{len(df[df['index'] > 100])/len(df)*100:.1f}%",
                        f"{geo_data['index'].mean():.1f}" if not geo_data.empty else "N/A"
                    ],
                    'Target': [
                        f"{df['index'].mean() * 1.15:.1f}",
                        f"{((len(performance_buckets['high']) + len(performance_buckets['very_high']))/len(df)*100) * 1.3:.1f}%",
                        "+25%",
                        f"{(len(df[df['index'] > 100])/len(df)*100) * 1.2:.1f}%",
                        f"{geo_data['index'].mean() * 1.1:.1f}" if not geo_data.empty else "N/A"
                    ]
                }
                
                kpi_df = pd.DataFrame(kpi_data)
                st.dataframe(kpi_df, use_container_width=True)
                
                st.markdown('<div class="insight-box">', unsafe_allow_html=True)
                st.markdown("**🎯 Executive Summary & Next Steps:**")
                st.markdown(f"• **Primary Opportunity**: {insights['top_attribute']} segment shows exceptional performance")
                st.markdown(f"• **Quick Wins**: Exclude {len(performance_buckets['low'])} low performers for immediate efficiency gains")
                st.markdown("• **Strategic Focus**: Target growing households in lower-middle income brackets")
                st.markdown("• **Long-term Vision**: Build comprehensive family lifecycle targeting strategy")
                st.markdown('</div>', unsafe_allow_html=True)
            
            # PowerPoint Generation Section
            st.sidebar.markdown("---")
            st.sidebar.subheader("📥 Export Dashboard")
            
            if st.sidebar.button("🎯 Generate Executive Presentation", type="primary"):
                with st.spinner("Creating comprehensive presentation..."):
                    try:
                        charts_data = {}
                        
                        comprehensive_insights = analyze_audience_insights(df)
                        
                        geo_data = df[df['attribute_group'].str.contains('State|Geographic|Census', case=False, na=False)]
                        if not geo_data.empty:
                            top_geo_insights = geo_data.nlargest(5, 'index')
                            comprehensive_insights['geographic_insights'] = [
                                f"{row['attribute_name']}: {row['index']:.1f}" 
                                for _, row in top_geo_insights.iterrows()
                            ]
                        
                        ppt_path = create_comprehensive_ppt(df, charts_data, comprehensive_insights)
                        
                        st.success("✅ Executive presentation generated successfully!")
                        
                        with open(ppt_path, "rb") as file:
                            st.sidebar.download_button(
                                label="📥 Download Executive Presentation",
                                data=file,
                                file_name="Executive_Audience_Analytics_Dashboard.pptx",
                                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                            )
                        
                        try:
                            os.unlink(ppt_path)
                        except:
                            pass
                            
                    except Exception as e:
                        st.error(f"❌ Error generating presentation: {str(e)}")
                        st.write("Please check your data format and try again.")
        
        else:
            st.error("❌ Sheet 'Index Report' not found in uploaded file.")
            st.info("Available sheets: " + ", ".join(xls.sheet_names))
    
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.write("Please ensure your Excel file is properly formatted.")

else:
    st.markdown("## 👋 Welcome to Advanced Audience Analytics")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        This comprehensive dashboard provides deep insights into your audience data with:
        
        **🎯 Performance Analysis**
        - Performance segmentation and distribution analysis
        - Top and bottom performer identification
        - Index score optimization opportunities
        
        **💰 Income Profiling** 
        - Detailed income bracket performance analysis
        - Inverse relationship insights
        - Targeting optimization recommendations
        
        **👨‍👩‍👧‍👦 Family Lifecycle Analysis**
        - Children age group performance
        - Family-focused targeting insights
        - Lifecycle stage optimization
        
        **🗺️ Geographic Intelligence**
        - State and region performance mapping
        - Geographic concentration analysis
        - Location-based targeting recommendations
        
        **🎯 Advanced Analytics**
        - Audience overlap and penetration analysis
        - Reach vs relevance trade-offs
        - Budget optimization insights
        - Strategic recommendations and action plans
        """)
    
    with col2:
        st.markdown('<div class="insight-box">', unsafe_allow_html=True)
        st.markdown("**📋 Data Requirements:**")
        st.markdown("Upload an Excel file with an **'Index Report'** sheet containing:")
        st.markdown("• `Attribute Name` - Name of each attribute")
        st.markdown("• `Attribute Group` - Category classification")
        st.markdown("• `Index` - Performance index score")
        st.markdown("• `Relative_Lift` - Relative lift percentage")
        st.markdown("• `Attribute Size` - Audience size (optional)")
        st.markdown("• Additional demographic columns")
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.info("👆 Upload your Excel file to begin comprehensive audience analysis")
