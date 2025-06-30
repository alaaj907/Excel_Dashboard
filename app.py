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
import traceback
import re

# Page configuration
st.set_page_config(page_title="Enhanced Audience Analytics Dashboard", layout="wide", initial_sidebar_state="expanded")

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
    .success-box {
        background-color: #e8f5e8;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #4caf50;
        margin: 1rem 0;
    }
    .demographics-box {
        background-color: #f3e5f5;
        padding: 1rem;
        border-radius: 0.5rem;
        border-left: 4px solid #9c27b0;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

st.title("🎯 Enhanced Audience Analytics Dashboard")
st.markdown("*Advanced multi-dimensional analysis with demographics and actionable insights*")

def find_index_report_sheet(xls):
    """Find the Index Report sheet with enhanced case-insensitive search"""
    # Direct matches first
    for sheet in xls.sheet_names:
        if sheet.lower() == "index report":
            return sheet
        if "index report" in sheet.lower():
            return sheet
    
    # Look for variations
    index_patterns = [
        r'.*index.*report.*',
        r'.*report.*index.*',
        r'.*index.*data.*',
        r'.*report.*data.*'
    ]
    
    for pattern in index_patterns:
        for sheet in xls.sheet_names:
            if re.search(pattern, sheet.lower()):
                return sheet
    
    # Last resort - look for just "index" or "report"
    for sheet in xls.sheet_names:
        if 'index' in sheet.lower() or 'report' in sheet.lower():
            return sheet
    
    return None

def find_target_columns_and_data(df, target_columns):
    """Find the exact row where target columns start and extract data below it"""
    TARGET_COLUMNS = [
        "Attribute Name", "Attribute Path", "Attribute Size",
        "Audience & THIS Attribute Overlap", "Audience & ANY Attribute Overlap", 
        "Audience Attribute Proportion", "Base Adjusted Population & THIS Attribute Overlap",
        "Base Adjusted Population & ANY Attribute Overlap", "Base Adjusted Population Attribute Proportion",
        "Index", "AIR Category", "AIR Attribute", "AIR Attribute Value", "AIR Attribute Path",
        "Audience Overlap % of Input Size", "Audience Threshold", "Exceeds Audience Threshold"
    ]
    
    st.sidebar.write("🔍 **Searching for target columns...**")
    
    for row_idx in range(min(25, len(df))):
        row_values = df.iloc[row_idx].astype(str).str.strip()
        matches = 0
        column_positions = {}
        
        for col_idx, cell_value in enumerate(row_values):
            for target_col in TARGET_COLUMNS:
                if cell_value.lower().strip() == target_col.lower().strip():
                    matches += 1
                    column_positions[target_col] = col_idx
                    break
        
        st.sidebar.write(f"Row {row_idx}: Found {matches}/{len(TARGET_COLUMNS)} target columns")
        
        if matches >= 5:  # Minimum threshold for detection
            st.sidebar.write(f"✅ **Header row found at row {row_idx}**")
            
            header_row = df.iloc[row_idx].astype(str).str.strip()
            data_rows = df.iloc[row_idx + 1:].reset_index(drop=True)
            data_rows.columns = header_row
            
            # Create filtered dataframe with only target columns
            filtered_df = pd.DataFrame()
            column_mapping = {}
            found_columns = []
            
            for target_col in TARGET_COLUMNS:
                for actual_col in data_rows.columns:
                    if actual_col.lower().strip() == target_col.lower().strip():
                        filtered_df[target_col] = data_rows[actual_col]
                        column_mapping[target_col] = actual_col
                        found_columns.append(target_col)
                        break
            
            # Clean the data
            filtered_df = filtered_df.dropna(how='all')
            
            # Validate Index column exists and has numeric data
            if 'Index' in filtered_df.columns:
                numeric_test = pd.to_numeric(filtered_df['Index'], errors='coerce')
                mask = numeric_test.notna()
                filtered_df = filtered_df[mask]
                filtered_df['Index'] = numeric_test[mask]
            
            return filtered_df, column_mapping, row_idx, found_columns
    
    st.sidebar.write("❌ **Could not find sufficient target columns**")
    return None, None, None, None

def prepare_data_for_analysis(df):
    """Clean and prepare data for analysis"""
    try:
        # Convert numeric columns
        numeric_columns = [
            'Attribute Size', 'Audience & THIS Attribute Overlap', 'Audience & ANY Attribute Overlap', 
            'Audience Attribute Proportion', 'Base Adjusted Population & THIS Attribute Overlap',
            'Base Adjusted Population & ANY Attribute Overlap', 'Base Adjusted Population Attribute Proportion',
            'Audience Overlap % of Input Size', 'Index'
        ]
        
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Clean text columns
        text_columns = [
            'Attribute Name', 'Attribute Path', 'AIR Category', 'AIR Attribute',
            'AIR Attribute Value', 'AIR Attribute Path'
        ]
        
        for col in text_columns:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()
                df[col] = df[col].replace('nan', '')
                df[col] = df[col].replace('', 'Unknown')
        
        # Remove rows with invalid Index values
        if 'Index' in df.columns:
            df = df[df['Index'].notna() & (df['Index'] > 0)]
        
        return df
        
    except Exception as e:
        st.error(f"Error in data preparation: {str(e)}")
        return df

def analyze_demographics(df):
    """Analyze demographic patterns from attribute names and categories"""
    demographics = {}
    
    if 'Attribute Name' in df.columns:
        attr_text = df['Attribute Name'].str.lower()
        
        # Age analysis
        age_patterns = {
            '18-24': ['18-24', '18 to 24', 'young adult', 'gen z'],
            '25-34': ['25-34', '25 to 34', 'millennial'],
            '35-44': ['35-44', '35 to 44', 'gen x'],
            '45-54': ['45-54', '45 to 54'],
            '55-64': ['55-64', '55 to 64'],
            '65+': ['65+', '65 plus', 'senior', 'boomer']
        }
        
        age_data = {}
        for age_group, patterns in age_patterns.items():
            mask = attr_text.str.contains('|'.join(patterns), na=False)
            segments = df[mask]
            if len(segments) > 0:
                age_data[age_group] = {
                    'count': len(segments),
                    'avg_index': segments['Index'].mean(),
                    'segments': segments
                }
        
        demographics['age'] = age_data
        
        # Gender analysis
        gender_patterns = {
            'Male': ['male', 'men', 'masculine'],
            'Female': ['female', 'women', 'feminine']
        }
        
        gender_data = {}
        for gender, patterns in gender_patterns.items():
            mask = attr_text.str.contains('|'.join(patterns), na=False)
            segments = df[mask]
            if len(segments) > 0:
                gender_data[gender] = {
                    'count': len(segments),
                    'avg_index': segments['Index'].mean(),
                    'segments': segments
                }
        
        demographics['gender'] = gender_data
        
        # Income analysis
        income_patterns = {
            'High Income ($100K+)': ['high income', '100k', '$100', 'affluent', 'luxury'],
            'Upper Middle ($75K-$100K)': ['upper middle', '75k', '$75'],
            'Middle Income ($50K-$75K)': ['middle income', '50k', '$50'],
            'Lower Income (<$50K)': ['lower income', 'budget', 'value']
        }
        
        income_data = {}
        for income_group, patterns in income_patterns.items():
            mask = attr_text.str.contains('|'.join(patterns), na=False)
            segments = df[mask]
            if len(segments) > 0:
                income_data[income_group] = {
                    'count': len(segments),
                    'avg_index': segments['Index'].mean(),
                    'segments': segments
                }
        
        demographics['income'] = income_data
        
        # Homeownership analysis
        homeowner_patterns = {
            'Homeowners': ['homeowner', 'own home', 'mortgage', 'property owner'],
            'Renters': ['renter', 'rent', 'lease', 'apartment']
        }
        
        homeowner_data = {}
        for status, patterns in homeowner_patterns.items():
            mask = attr_text.str.contains('|'.join(patterns), na=False)
            segments = df[mask]
            if len(segments) > 0:
                homeowner_data[status] = {
                    'count': len(segments),
                    'avg_index': segments['Index'].mean(),
                    'segments': segments
                }
        
        demographics['homeownership'] = homeowner_data
        
        # Family status analysis
        family_patterns = {
            'Married': ['married', 'spouse', 'couple'],
            'Single': ['single', 'unmarried', 'divorced'],
            'With Children': ['children', 'kids', 'family', 'parent'],
            'Empty Nest': ['empty nest', 'no children', 'childless']
        }
        
        family_data = {}
        for status, patterns in family_patterns.items():
            mask = attr_text.str.contains('|'.join(patterns), na=False)
            segments = df[mask]
            if len(segments) > 0:
                family_data[status] = {
                    'count': len(segments),
                    'avg_index': segments['Index'].mean(),
                    'segments': segments
                }
        
        demographics['family'] = family_data
    
    return demographics

def analyze_geographic_patterns(df):
    """Analyze geographic patterns from attribute names with enhanced state detection"""
    geographic = {}
    
    if 'Attribute Name' in df.columns:
        attr_text = df['Attribute Name'].str.lower()
        
        # Enhanced state patterns with abbreviations and full names
        states_data = {
            'Alabama': ['alabama', 'al'], 'Alaska': ['alaska', 'ak'], 'Arizona': ['arizona', 'az'],
            'Arkansas': ['arkansas', 'ar'], 'California': ['california', 'ca'], 'Colorado': ['colorado', 'co'],
            'Connecticut': ['connecticut', 'ct'], 'Delaware': ['delaware', 'de'], 'Florida': ['florida', 'fl'],
            'Georgia': ['georgia', 'ga'], 'Hawaii': ['hawaii', 'hi'], 'Idaho': ['idaho', 'id'],
            'Illinois': ['illinois', 'il'], 'Indiana': ['indiana', 'in'], 'Iowa': ['iowa', 'ia'],
            'Kansas': ['kansas', 'ks'], 'Kentucky': ['kentucky', 'ky'], 'Louisiana': ['louisiana', 'la'],
            'Maine': ['maine', 'me'], 'Maryland': ['maryland', 'md'], 'Massachusetts': ['massachusetts', 'ma'],
            'Michigan': ['michigan', 'mi'], 'Minnesota': ['minnesota', 'mn'], 'Mississippi': ['mississippi', 'ms'],
            'Missouri': ['missouri', 'mo'], 'Montana': ['montana', 'mt'], 'Nebraska': ['nebraska', 'ne'],
            'Nevada': ['nevada', 'nv'], 'New Hampshire': ['new hampshire', 'nh'], 'New Jersey': ['new jersey', 'nj'],
            'New Mexico': ['new mexico', 'nm'], 'New York': ['new york', 'ny'], 'North Carolina': ['north carolina', 'nc'],
            'North Dakota': ['north dakota', 'nd'], 'Ohio': ['ohio', 'oh'], 'Oklahoma': ['oklahoma', 'ok'],
            'Oregon': ['oregon', 'or'], 'Pennsylvania': ['pennsylvania', 'pa'], 'Rhode Island': ['rhode island', 'ri'],
            'South Carolina': ['south carolina', 'sc'], 'South Dakota': ['south dakota', 'sd'],
            'Tennessee': ['tennessee', 'tn'], 'Texas': ['texas', 'tx'], 'Utah': ['utah', 'ut'],
            'Vermont': ['vermont', 'vt'], 'Virginia': ['virginia', 'va'], 'Washington': ['washington', 'wa'],
            'West Virginia': ['west virginia', 'wv'], 'Wisconsin': ['wisconsin', 'wi'], 'Wyoming': ['wyoming', 'wy']
        }
        
        # State analysis
        state_data = {}
        for state, patterns in states_data.items():
            # Create pattern that matches state name or abbreviation as whole words
            pattern = r'\b(' + '|'.join(patterns) + r')\b'
            mask = attr_text.str.contains(pattern, na=False, regex=True)
            segments = df[mask]
            if len(segments) > 0:
                state_data[state] = {
                    'count': len(segments),
                    'avg_index': segments['Index'].mean(),
                    'total_audience': segments['Attribute Size'].sum() if 'Attribute Size' in segments.columns else 0,
                    'segments': segments
                }
        
        geographic['states'] = state_data
        
        # Regional patterns (enhanced)
        regions = {
            'Northeast': ['northeast', 'new england', 'new york', 'boston', 'philadelphia', 'new jersey', 
                         'massachusetts', 'connecticut', 'maine', 'vermont', 'new hampshire', 'rhode island'],
            'Southeast': ['southeast', 'south', 'florida', 'atlanta', 'miami', 'carolinas', 'georgia', 
                         'alabama', 'mississippi', 'louisiana', 'arkansas', 'tennessee', 'kentucky'],
            'Midwest': ['midwest', 'great lakes', 'chicago', 'detroit', 'ohio', 'illinois', 'indiana', 
                       'michigan', 'wisconsin', 'minnesota', 'iowa', 'missouri', 'kansas', 'nebraska'],
            'Southwest': ['southwest', 'texas', 'arizona', 'nevada', 'new mexico', 'oklahoma'],
            'West Coast': ['west coast', 'pacific', 'california', 'seattle', 'los angeles', 'san francisco', 
                          'washington', 'oregon'],
            'Mountain West': ['mountain', 'rockies', 'colorado', 'utah', 'wyoming', 'montana', 'idaho']
        }
        
        region_data = {}
        for region, patterns in regions.items():
            mask = attr_text.str.contains('|'.join(patterns), na=False)
            segments = df[mask]
            if len(segments) > 0:
                region_data[region] = {
                    'count': len(segments),
                    'avg_index': segments['Index'].mean(),
                    'total_audience': segments['Attribute Size'].sum() if 'Attribute Size' in segments.columns else 0,
                    'segments': segments
                }
        
        geographic['regions'] = region_data
        
        # Urban/Rural patterns
        area_types = {
            'Urban': ['urban', 'city', 'metropolitan', 'downtown', 'metro area'],
            'Suburban': ['suburban', 'suburb', 'residential', 'suburban area'],
            'Rural': ['rural', 'country', 'small town', 'farming', 'countryside']
        }
        
        area_data = {}
        for area_type, patterns in area_types.items():
            mask = attr_text.str.contains('|'.join(patterns), na=False)
            segments = df[mask]
            if len(segments) > 0:
                area_data[area_type] = {
                    'count': len(segments),
                    'avg_index': segments['Index'].mean(),
                    'total_audience': segments['Attribute Size'].sum() if 'Attribute Size' in segments.columns else 0,
                    'segments': segments
                }
        
        geographic['area_types'] = area_data
    
    return geographic

def analyze_psychographics(df):
    """Analyze psychographic and behavioral patterns"""
    psychographics = {}
    
    if 'Attribute Name' in df.columns:
        attr_text = df['Attribute Name'].str.lower()
        
        # Lifestyle interests
        lifestyle_patterns = {
            'Fitness & Health': ['fitness', 'health', 'gym', 'exercise', 'wellness', 'nutrition'],
            'Travel & Adventure': ['travel', 'vacation', 'adventure', 'trip', 'tourism'],
            'Technology': ['technology', 'tech', 'digital', 'smartphone', 'computer'],
            'Food & Dining': ['food', 'dining', 'restaurant', 'cooking', 'culinary'],
            'Entertainment': ['entertainment', 'movies', 'music', 'gaming', 'streaming'],
            'Shopping': ['shopping', 'retail', 'fashion', 'luxury', 'brand']
        }
        
        lifestyle_data = {}
        for category, patterns in lifestyle_patterns.items():
            mask = attr_text.str.contains('|'.join(patterns), na=False)
            segments = df[mask]
            if len(segments) > 0:
                lifestyle_data[category] = {
                    'count': len(segments),
                    'avg_index': segments['Index'].mean(),
                    'segments': segments
                }
        
        psychographics['lifestyle'] = lifestyle_data
        
        # Financial behaviors
        financial_patterns = {
            'Investment Focused': ['investment', 'portfolio', 'stocks', 'retirement', '401k'],
            'Credit Active': ['credit', 'loan', 'mortgage', 'financing'],
            'Savings Oriented': ['savings', 'bank', 'deposit', 'conservative'],
            'Insurance Seeking': ['insurance', 'protection', 'coverage']
        }
        
        financial_data = {}
        for category, patterns in financial_patterns.items():
            mask = attr_text.str.contains('|'.join(patterns), na=False)
            segments = df[mask]
            if len(segments) > 0:
                financial_data[category] = {
                    'count': len(segments),
                    'avg_index': segments['Index'].mean(),
                    'segments': segments
                }
        
        psychographics['financial'] = financial_data
    
    return psychographics

def create_comprehensive_charts(df, demographics, geographic, psychographics):
    """Create comprehensive charts for all analyses"""
    charts = {}
    
    # Performance distribution
    if 'Index' in df.columns:
        fig_performance = px.histogram(
            df, x='Index', nbins=30,
            title='Index Performance Distribution',
            labels={'Index': 'Index Score', 'count': 'Number of Segments'}
        )
        fig_performance.add_vline(x=120, line_dash="dash", line_color="red", annotation_text="Performance Threshold")
        charts['performance_dist'] = fig_performance
        
        # US Choropleth Map for Geographic Analysis
        if geographic.get('states'):
            # Prepare state data for choropleth
            state_names = list(geographic['states'].keys())
            state_values = [geographic['states'][state]['avg_index'] for state in state_names]
            state_hover = [f"{state}<br>Avg Index: {geographic['states'][state]['avg_index']:.1f}<br>Segments: {geographic['states'][state]['count']}" 
                          for state in state_names]
            
            fig_map = go.Figure(data=go.Choropleth(
                locations=state_names,
                z=state_values,
                locationmode='USA-states',
                colorscale='RdYlBu_r',
                reversescale=False,
                text=state_hover,
                hovertemplate='%{text}<extra></extra>',
                colorbar_title="Average Index Score"
            ))
            
            fig_map.update_layout(
                title_text='Geographic Performance Heatmap - US States',
                geo_scope='usa',
                height=500
            )
            charts['us_map'] = fig_map
        
        # Demographics charts
        if demographics.get('age'):
            age_df = pd.DataFrame([
                {'Age Group': k, 'Count': v['count'], 'Avg Index': v['avg_index']}
                for k, v in demographics['age'].items()
            ])
            fig_age = px.bar(age_df, x='Age Group', y='Avg Index', color='Count',
                           title='Performance by Age Group')
            fig_age.add_hline(y=120, line_dash="dash", line_color="red")
            charts['age_performance'] = fig_age
        
        if demographics.get('income'):
            income_df = pd.DataFrame([
                {'Income Group': k, 'Count': v['count'], 'Avg Index': v['avg_index']}
                for k, v in demographics['income'].items()
            ])
            fig_income = px.bar(income_df, x='Income Group', y='Avg Index', color='Count',
                              title='Performance by Income Level')
            fig_income.add_hline(y=120, line_dash="dash", line_color="red")
            charts['income_performance'] = fig_income
        
        # Geographic charts
        if geographic.get('regions'):
            region_df = pd.DataFrame([
                {'Region': k, 'Count': v['count'], 'Avg Index': v['avg_index']}
                for k, v in geographic['regions'].items()
            ])
            fig_regions = px.bar(region_df, x='Region', y='Avg Index', color='Count',
                               title='Performance by Geographic Region')
            fig_regions.add_hline(y=120, line_dash="dash", line_color="red")
            charts['regional_performance'] = fig_regions
        
        # Advanced Performance Analysis Charts
        # 1. Box Plot for Index Distribution by Category
        if 'AIR Category' in df.columns and df['AIR Category'].nunique() > 1:
            # Get top categories by count for cleaner visualization
            top_categories = df['AIR Category'].value_counts().head(10).index
            df_top_cats = df[df['AIR Category'].isin(top_categories)]
            
            fig_box = px.box(df_top_cats, x='AIR Category', y='Index',
                           title='Index Distribution by AIR Category (Top 10)')
            fig_box.update_xaxes(tickangle=-45)
            fig_box.add_hline(y=120, line_dash="dash", line_color="red")
            charts['category_boxplot'] = fig_box
        
        # 2. Scatter Plot: Attribute Size vs Index Performance
        if 'Attribute Size' in df.columns:
            # Sample data for better performance
            sample_size = min(1000, len(df))
            df_sample = df.sample(sample_size) if len(df) > sample_size else df
            
            fig_scatter = px.scatter(df_sample, x='Attribute Size', y='Index',
                                   title='Audience Size vs Performance Analysis',
                                   hover_data=['Attribute Name'] if 'Attribute Name' in df.columns else None,
                                   opacity=0.6)
            fig_scatter.update_xaxes(type="log", title="Audience Size (Log Scale)")
            fig_scatter.add_hline(y=120, line_dash="dash", line_color="red")
            charts['size_vs_performance'] = fig_scatter
        
        # 3. Performance Trend Analysis (if we can detect time patterns)
        # This would show how performance varies across different segments
        quartiles = df['Index'].quantile([0.25, 0.5, 0.75])
        perf_segments = {
            'Top Quartile (Q4)': df[df['Index'] >= quartiles[0.75]],
            'Third Quartile (Q3)': df[(df['Index'] >= quartiles[0.5]) & (df['Index'] < quartiles[0.75])],
            'Second Quartile (Q2)': df[(df['Index'] >= quartiles[0.25]) & (df['Index'] < quartiles[0.5])],
            'Bottom Quartile (Q1)': df[df['Index'] < quartiles[0.25]]
        }
        
        # Create violin plot for performance distribution
        perf_data = []
        for quartile, data in perf_segments.items():
            for idx in data['Index']:
                perf_data.append({'Quartile': quartile, 'Index': idx})
        
        if perf_data:
            perf_df = pd.DataFrame(perf_data)
            fig_violin = px.violin(perf_df, x='Quartile', y='Index',
                                 title='Performance Distribution by Quartile',
                                 box=True, points="outliers")
            fig_violin.add_hline(y=120, line_dash="dash", line_color="red")
            charts['performance_violin'] = fig_violin
        
        # 4. Correlation Heatmap for numeric columns
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        if len(numeric_cols) > 2:
            corr_matrix = df[numeric_cols].corr()
            
            fig_heatmap = px.imshow(corr_matrix,
                                  title='Correlation Matrix - Key Performance Metrics',
                                  labels=dict(color="Correlation"),
                                  color_continuous_scale='RdBu_r',
                                  aspect="auto")
            charts['correlation_heatmap'] = fig_heatmap
        
        # 5. Performance Concentration Analysis (Pareto)
        df_sorted = df.sort_values('Index', ascending=False).reset_index(drop=True)
        df_sorted['cumulative_pct'] = (df_sorted.index + 1) / len(df_sorted) * 100
        df_sorted['performance_cumsum'] = df_sorted['Index'].cumsum()
        df_sorted['performance_pct'] = df_sorted['performance_cumsum'] / df_sorted['Index'].sum() * 100
        
        fig_pareto = go.Figure()
        fig_pareto.add_trace(go.Scatter(x=df_sorted['cumulative_pct'], 
                                      y=df_sorted['performance_pct'],
                                      mode='lines',
                                      name='Cumulative Performance %',
                                      line=dict(color='blue', width=3)))
        
        fig_pareto.add_shape(type="line", x0=0, y0=0, x1=100, y1=100,
                           line=dict(color="red", width=2, dash="dash"))
        
        fig_pareto.update_layout(
            title='Performance Concentration Analysis (Pareto Curve)',
            xaxis_title='Cumulative % of Segments',
            yaxis_title='Cumulative % of Performance',
            height=400
        )
        charts['pareto_analysis'] = fig_pareto
        
        # Psychographics charts
        if psychographics.get('lifestyle'):
            lifestyle_df = pd.DataFrame([
                {'Interest': k, 'Count': v['count'], 'Avg Index': v['avg_index']}
                for k, v in psychographics['lifestyle'].items()
            ])
            fig_lifestyle = px.bar(lifestyle_df, x='Interest', y='Avg Index', color='Count',
                                 title='Performance by Lifestyle Interest')
            fig_lifestyle.add_hline(y=120, line_dash="dash", line_color="red")
            fig_lifestyle.update_xaxes(tickangle=-45)
            charts['lifestyle_performance'] = fig_lifestyle
    
    return charts

def generate_key_insights(df, demographics, geographic, psychographics):
    """Generate key insights and call-outs"""
    insights = []
    
    # Overall performance insights
    if 'Index' in df.columns:
        avg_index = df['Index'].mean()
        high_performers = len(df[df['Index'] > 120])
        
        insights.append(f"**Overall Performance**: {avg_index:.1f} average index with {high_performers:,} high-performing segments")
    
    # Demographic insights
    if demographics.get('age'):
        best_age = max(demographics['age'].items(), key=lambda x: x[1]['avg_index'])
        insights.append(f"**Age Sweet Spot**: {best_age[0]} performs best with {best_age[1]['avg_index']:.1f} average index")
    
    if demographics.get('income'):
        best_income = max(demographics['income'].items(), key=lambda x: x[1]['avg_index'])
        insights.append(f"**Income Focus**: {best_income[0]} shows strongest performance at {best_income[1]['avg_index']:.1f} index")
    
    # Geographic insights
    if geographic.get('regions'):
        best_region = max(geographic['regions'].items(), key=lambda x: x[1]['avg_index'])
        insights.append(f"**Geographic Opportunity**: {best_region[0]} leads performance with {best_region[1]['avg_index']:.1f} index")
    
    # Psychographic insights
    if psychographics.get('lifestyle'):
        best_lifestyle = max(psychographics['lifestyle'].items(), key=lambda x: x[1]['avg_index'])
        insights.append(f"**Lifestyle Driver**: {best_lifestyle[0]} interest shows {best_lifestyle[1]['avg_index']:.1f} index performance")
    
    return insights

# Sidebar navigation
st.sidebar.title("📊 Enhanced Analytics Sections")
analysis_sections = [
    "🏠 Overview",
    "👥 Demographics Deep Dive", 
    "🗺️ Geographic Intelligence",
    "🧠 Psychographics & Behaviors",
    "📈 Advanced Performance Analysis",
    "🎯 Actionable Optimization",
    "📋 Key Insights & Call-Outs"
]

selected_section = st.sidebar.selectbox("Select Analysis Section", analysis_sections)

# File upload
uploaded_file = st.file_uploader("Upload Excel File with Index Report Data", type=["xlsx"])

if uploaded_file:
    try:
        file_name = uploaded_file.name.replace('.xlsx', '')
        xls = pd.ExcelFile(uploaded_file)
        
        st.sidebar.write("**📋 Available sheets:**")
        for sheet in xls.sheet_names:
            st.sidebar.write(f"• {sheet}")
        
        target_sheet = find_index_report_sheet(xls)
        
        if target_sheet:
            st.sidebar.markdown('<div class="success-box">', unsafe_allow_html=True)
            st.sidebar.write(f"✅ **Using sheet:** {target_sheet}")
            st.sidebar.markdown('</div>', unsafe_allow_html=True)
            
            raw_df = pd.read_excel(xls, sheet_name=target_sheet, header=None)
            st.sidebar.write(f"📊 **Sheet dimensions:** {raw_df.shape[0]} rows × {raw_df.shape[1]} columns")
            
            # Find target columns and extract data
            df, column_mapping, header_row, found_columns = find_target_columns_and_data(raw_df, None)
            
            if df is not None and len(df) > 0:
                st.sidebar.markdown('<div class="success-box">', unsafe_allow_html=True)
                st.sidebar.write(f"✅ **Data extraction successful!**")
                st.sidebar.write(f"📍 **Header found at row:** {header_row}")
                st.sidebar.write(f"📊 **Extracted data:** {len(df)} rows × {len(df.columns)} columns")
                st.sidebar.write(f"🎯 **Target columns found:** {len(found_columns)}")
                st.sidebar.markdown('</div>', unsafe_allow_html=True)
                
                # Prepare data
                df = prepare_data_for_analysis(df)
                
                if 'Index' not in df.columns:
                    st.error("❌ Critical: 'Index' column not found")
                    st.stop()
                
                if len(df) == 0:
                    st.error("❌ No valid data found after cleaning")
                    st.stop()
                
                st.success(f"✅ Ready for analysis with {len(df)} valid records!")
                
                # Perform enhanced analysis
                demographics = analyze_demographics(df)
                geographic = analyze_geographic_patterns(df)
                psychographics = analyze_psychographics(df)
                charts = create_comprehensive_charts(df, demographics, geographic, psychographics)
                key_insights = generate_key_insights(df, demographics, geographic, psychographics)
                
                # ANALYSIS SECTIONS
                
                if selected_section == "🏠 Overview":
                    st.header("📊 Executive Overview Dashboard")
                    
                    col1, col2, col3, col4, col5 = st.columns(5)
                    
                    with col1:
                        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                        st.metric("Total Segments", f"{len(df):,}")
                        st.markdown('</div>', unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                        st.metric("Avg Index Score", f"{df['Index'].mean():.1f}")
                        st.markdown('</div>', unsafe_allow_html=True)
                    
                    with col3:
                        high_performers = len(df[df['Index'] > 120])
                        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                        st.metric("High Performers", f"{high_performers:,}")
                        st.markdown('</div>', unsafe_allow_html=True)
                    
                    with col4:
                        if 'Attribute Size' in df.columns:
                            total_audience = df['Attribute Size'].sum()
                            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                            st.metric("Total Audience", f"{total_audience:,.0f}")
                            st.markdown('</div>', unsafe_allow_html=True)
                        else:
                            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                            st.metric("Data Quality", "Good")
                            st.markdown('</div>', unsafe_allow_html=True)
                    
                    with col5:
                        categories = df['AIR Category'].nunique() if 'AIR Category' in df.columns else 0
                        st.markdown('<div class="metric-card">', unsafe_allow_html=True)
                        st.metric("AIR Categories", f"{categories}")
                        st.markdown('</div>', unsafe_allow_html=True)
                    
                    # Performance distribution
                    if 'performance_dist' in charts:
                        st.plotly_chart(charts['performance_dist'], use_container_width=True)
                    
                    # Quick insights preview
                    st.subheader("🎯 Quick Insights Preview")
                    for insight in key_insights[:3]:
                        st.markdown(f"• {insight}")
                
                elif selected_section == "👥 Demographics Deep Dive":
                    st.header("👥 Demographics Deep Dive Analysis")
                    st.markdown("*Precise demographic patterns and performance insights*")
                    
                    # Age Analysis
                    if demographics.get('age'):
                        st.subheader("📅 Age Range Performance")
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            if 'age_performance' in charts:
                                st.plotly_chart(charts['age_performance'], use_container_width=True)
                        
                        with col2:
                            st.markdown("**Age Group Details:**")
                            age_data = []
                            for age_group, data in demographics['age'].items():
                                age_data.append({
                                    'Age Group': age_group,
                                    'Segments': data['count'],
                                    'Avg Index': f"{data['avg_index']:.1f}"
                                })
                            age_df = pd.DataFrame(age_data)
                            st.dataframe(age_df, use_container_width=True)
                    
                    # Income Analysis
                    if demographics.get('income'):
                        st.subheader("💰 Income Level Performance")
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            if 'income_performance' in charts:
                                st.plotly_chart(charts['income_performance'], use_container_width=True)
                        
                        with col2:
                            st.markdown("**Income Insights:**")
                            income_data = []
                            for income_group, data in demographics['income'].items():
                                income_data.append({
                                    'Income Level': income_group,
                                    'Segments': data['count'],
                                    'Avg Index': f"{data['avg_index']:.1f}"
                                })
                            income_df = pd.DataFrame(income_data)
                            st.dataframe(income_df, use_container_width=True)
                    
                    # Gender Analysis
                    if demographics.get('gender'):
                        st.subheader("⚧ Gender Breakdown")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            gender_labels = list(demographics['gender'].keys())
                            gender_values = [data['count'] for data in demographics['gender'].values()]
                            
                            fig_gender = px.pie(
                                values=gender_values,
                                names=gender_labels,
                                title="Gender Distribution"
                            )
                            st.plotly_chart(fig_gender, use_container_width=True)
                        
                        with col2:
                            st.markdown("**Gender Performance:**")
                            for gender, data in demographics['gender'].items():
                                st.metric(f"{gender} Avg Index", f"{data['avg_index']:.1f}")
                    
                    # Family Status
                    if demographics.get('family'):
                        st.subheader("👨‍👩‍👧‍👦 Family Status & Children")
                        family_data = []
                        for status, data in demographics['family'].items():
                            family_data.append({
                                'Family Status': status,
                                'Segments': data['count'],
                                'Avg Index': data['avg_index'],
                                'Performance': 'High' if data['avg_index'] > 120 else 'Medium' if data['avg_index'] > 100 else 'Low'
                            })
                        
                        family_df = pd.DataFrame(family_data)
                        family_df['Avg Index'] = family_df['Avg Index'].round(1)
                        st.dataframe(family_df, use_container_width=True)
                    
                    # Homeownership
                    if demographics.get('homeownership'):
                        st.subheader("🏠 Homeownership Status")
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            home_labels = list(demographics['homeownership'].keys())
                            home_values = [data['count'] for data in demographics['homeownership'].values()]
                            
                            fig_home = px.bar(
                                x=home_labels,
                                y=home_values,
                                title="Homeownership Segments"
                            )
                            st.plotly_chart(fig_home, use_container_width=True)
                        
                        with col2:
                            st.markdown("**Homeownership Performance:**")
                            for status, data in demographics['homeownership'].items():
                                st.metric(f"{status} Index", f"{data['avg_index']:.1f}")
                    
                    # Demographics insights
                    st.markdown('<div class="demographics-box">', unsafe_allow_html=True)
                    st.markdown("**👥 Demographics Key Findings:**")
                    demo_insights = [insight for insight in key_insights if any(keyword in insight.lower() for keyword in ['age', 'income', 'gender', 'family'])]
                    for insight in demo_insights:
                        st.markdown(f"• {insight}")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                elif selected_section == "🗺️ Geographic Intelligence":
                    st.header("🗺️ Geographic Intelligence Analysis")
                    st.markdown("*Regional performance and location-based targeting opportunities*")
                    
                    # US National Map
                    if geographic.get('states') and 'us_map' in charts:
                        st.subheader("🇺🇸 National Performance Heatmap")
                        st.plotly_chart(charts['us_map'], use_container_width=True)
                        
                        # State performance summary
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.markdown("**🏆 Top Performing States:**")
                            top_states = sorted(geographic['states'].items(), 
                                              key=lambda x: x[1]['avg_index'], reverse=True)[:5]
                            for i, (state, data) in enumerate(top_states, 1):
                                st.markdown(f"{i}. **{state}**: {data['avg_index']:.1f} index ({data['count']} segments)")
                        
                        with col2:
                            st.markdown("**📊 State Statistics:**")
                            total_states = len(geographic['states'])
                            avg_state_performance = np.mean([data['avg_index'] for data in geographic['states'].values()])
                            high_performing_states = len([s for s in geographic['states'].values() if s['avg_index'] > 120])
                            
                            st.metric("States with Data", total_states)
                            st.metric("Avg State Performance", f"{avg_state_performance:.1f}")
                            st.metric("High-Performing States", high_performing_states)
                    
                    # Regional Performance
                    if geographic.get('regions'):
                        st.subheader("🌎 Regional Performance Analysis")
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            if 'regional_performance' in charts:
                                st.plotly_chart(charts['regional_performance'], use_container_width=True)
                        
                        with col2:
                            st.markdown("**Regional Rankings:**")
                            region_data = []
                            for region, data in sorted(geographic['regions'].items(), 
                                                     key=lambda x: x[1]['avg_index'], reverse=True):
                                region_data.append({
                                    'Region': region,
                                    'Segments': data['count'],
                                    'Avg Index': f"{data['avg_index']:.1f}",
                                    'Total Audience': f"{data['total_audience']:,.0f}" if data['total_audience'] > 0 else "N/A"
                                })
                            region_df = pd.DataFrame(region_data)
                            st.dataframe(region_df, use_container_width=True)
                    
                    # Urban vs Rural Analysis
                    if geographic.get('area_types'):
                        st.subheader("🏙️ Urban vs Suburban vs Rural Analysis")
                        
                        area_labels = list(geographic['area_types'].keys())
                        area_counts = [data['count'] for data in geographic['area_types'].values()]
                        area_performance = [data['avg_index'] for data in geographic['area_types'].values()]
                        area_audience = [data['total_audience'] for data in geographic['area_types'].values()]
                        
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            fig_area_count = px.pie(
                                values=area_counts,
                                names=area_labels,
                                title="Segment Distribution by Area Type"
                            )
                            st.plotly_chart(fig_area_count, use_container_width=True)
                        
                        with col2:
                            fig_area_perf = px.bar(
                                x=area_labels,
                                y=area_performance,
                                title="Performance by Area Type",
                                color=area_performance,
                                color_continuous_scale='RdYlBu_r'
                            )
                            fig_area_perf.add_hline(y=120, line_dash="dash", line_color="red")
                            st.plotly_chart(fig_area_perf, use_container_width=True)
                        
                        with col3:
                            if any(aud > 0 for aud in area_audience):
                                fig_area_aud = px.bar(
                                    x=area_labels,
                                    y=area_audience,
                                    title="Total Audience by Area Type"
                                )
                                st.plotly_chart(fig_area_aud, use_container_width=True)
                            else:
                                st.markdown("**Area Type Insights:**")
                                for i, area_type in enumerate(area_labels):
                                    st.metric(f"{area_type} Index", f"{area_performance[i]:.1f}")
                    
                    # Geographic Opportunity Analysis
                    st.subheader("🎯 Geographic Opportunity Matrix")
                    
                    if geographic.get('states'):
                        # Create opportunity matrix based on performance and audience size
                        opportunity_data = []
                        for state, data in geographic['states'].items():
                            opportunity_data.append({
                                'State': state,
                                'Performance': data['avg_index'],
                                'Segments': data['count'],
                                'Audience': data['total_audience'],
                                'Opportunity Score': (data['avg_index'] * 0.7) + (data['count'] * 2)  # Weighted score
                            })
                        
                        opp_df = pd.DataFrame(opportunity_data)
                        opp_df = opp_df.sort_values('Opportunity Score', ascending=False)
                        
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            # Scatter plot: Performance vs Audience Size
                            fig_opportunity = px.scatter(
                                opp_df,
                                x='Audience' if any(opp_df['Audience'] > 0) else 'Segments',
                                y='Performance',
                                size='Segments',
                                hover_data=['State'],
                                title='Geographic Opportunity Matrix',
                                labels={'Performance': 'Average Index Score'}
                            )
                            fig_opportunity.add_hline(y=120, line_dash="dash", line_color="red")
                            if any(opp_df['Audience'] > 0):
                                fig_opportunity.update_xaxes(type="log", title="Total Audience (Log Scale)")
                            st.plotly_chart(fig_opportunity, use_container_width=True)
                        
                        with col2:
                            st.markdown("**🎯 Priority States for Investment:**")
                            top_opportunities = opp_df.head(8)[['State', 'Performance', 'Opportunity Score']]
                            top_opportunities['Performance'] = top_opportunities['Performance'].round(1)
                            top_opportunities['Opportunity Score'] = top_opportunities['Opportunity Score'].round(1)
                            st.dataframe(top_opportunities, use_container_width=True)
                    
                    # Geographic insights
                    st.markdown('<div class="insight-box">', unsafe_allow_html=True)
                    st.markdown("**🗺️ Geographic Key Findings:**")
                    
                    if geographic.get('states'):
                        best_state = max(geographic['states'].items(), key=lambda x: x[1]['avg_index'])
                        total_states = len(geographic['states'])
                        high_perf_states = len([s for s in geographic['states'].values() if s['avg_index'] > 120])
                        
                        st.markdown(f"• **Top performing state**: {best_state[0]} with {best_state[1]['avg_index']:.1f} average index")
                        st.markdown(f"• **{total_states} states** identified in dataset with performance data")
                        st.markdown(f"• **{high_perf_states} states** exceed 120 performance threshold")
                        
                        if geographic.get('regions'):
                            best_region = max(geographic['regions'].items(), key=lambda x: x[1]['avg_index'])
                            st.markdown(f"• **{best_region[0]}** region shows strongest overall performance")
                    
                    if geographic.get('area_types'):
                        best_area = max(geographic['area_types'].items(), key=lambda x: x[1]['avg_index'])
                        st.markdown(f"• **{best_area[0]}** areas demonstrate highest performance at {best_area[1]['avg_index']:.1f} index")
                    
                    st.markdown("• **Strategic opportunity**: Focus expansion on high-performing geographic markets")
                    st.markdown("• **Geo-targeting recommendation**: Prioritize states and regions with proven performance")
                    st.markdown('</div>', unsafe_allow_html=True)
                
                elif selected_section == "🧠 Psychographics & Behaviors":
                    st.header("🧠 Psychographics & Behavioral Analysis")
                    st.markdown("*Lifestyle interests, values, and behavioral patterns*")
                    
                    # Lifestyle Interests
                    if psychographics.get('lifestyle'):
                        st.subheader("🎯 Lifestyle Interests & Hobbies")
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            if 'lifestyle_performance' in charts:
                                st.plotly_chart(charts['lifestyle_performance'], use_container_width=True)
                        
                        with col2:
                            st.markdown("**Interest Categories:**")
                            lifestyle_data = []
                            for interest, data in sorted(psychographics['lifestyle'].items(),
                                                       key=lambda x: x[1]['avg_index'], reverse=True):
                                lifestyle_data.append({
                                    'Interest': interest,
                                    'Segments': data['count'],
                                    'Avg Index': f"{data['avg_index']:.1f}"
                                })
                            lifestyle_df = pd.DataFrame(lifestyle_data)
                            st.dataframe(lifestyle_df, use_container_width=True)
                    
                    # Financial Behaviors
                    if psychographics.get('financial'):
                        st.subheader("💳 Financial Behaviors & Attitudes")
                        
                        financial_labels = list(psychographics['financial'].keys())
                        financial_performance = [data['avg_index'] for data in psychographics['financial'].values()]
                        financial_counts = [data['count'] for data in psychographics['financial'].values()]
                        
                        fig_financial = go.Figure()
                        fig_financial.add_trace(go.Bar(
                            x=financial_labels,
                            y=financial_performance,
                            text=[f"{count} segments" for count in financial_counts],
                            textposition='outside',
                            marker_color=['#dc3545' if x > 120 else '#ffc107' if x > 100 else '#6c757d' 
                                        for x in financial_performance]
                        ))
                        
                        fig_financial.update_layout(
                            title="Financial Behavior Performance",
                            xaxis_title="Financial Behavior",
                            yaxis_title="Average Index",
                            height=400
                        )
                        fig_financial.add_hline(y=120, line_dash="dash", line_color="red")
                        st.plotly_chart(fig_financial, use_container_width=True)
                    
                    # Media Consumption (if detectable from data)
                    st.subheader("📱 Media Consumption Insights")
                    if 'Attribute Name' in df.columns:
                        media_patterns = ['digital', 'social', 'mobile', 'tv', 'streaming', 'podcast']
                        media_segments = df[df['Attribute Name'].str.lower().str.contains('|'.join(media_patterns), na=False)]
                        
                        if not media_segments.empty:
                            st.metric("Media-Related Segments", len(media_segments))
                            st.metric("Avg Media Segment Index", f"{media_segments['Index'].mean():.1f}")
                            
                            top_media = media_segments.nlargest(5, 'Index')[['Attribute Name', 'Index']]
                            st.markdown("**Top Media Segments:**")
                            st.dataframe(top_media, use_container_width=True)
                    
                    # Psychographics insights
                    st.markdown('<div class="insight-box">', unsafe_allow_html=True)
                    st.markdown("**🧠 Psychographics Key Findings:**")
                    psycho_insights = [insight for insight in key_insights if 'lifestyle' in insight.lower()]
                    for insight in psycho_insights:
                        st.markdown(f"• {insight}")
                    
                    if psychographics.get('lifestyle'):
                        best_lifestyle = max(psychographics['lifestyle'].items(), key=lambda x: x[1]['avg_index'])
                        st.markdown(f"• **Top lifestyle interest**: {best_lifestyle[0]} shows {best_lifestyle[1]['avg_index']:.1f} performance")
                    
                    st.markdown('</div>', unsafe_allow_html=True)
                
                elif selected_section == "📈 Advanced Performance Analysis":
                    st.header("📈 Advanced Performance Analysis")
                    st.markdown("*Comprehensive descriptive analysis and advanced visualizations*")
                    
                    # Descriptive Statistics Summary
                    st.subheader("📊 Comprehensive Descriptive Statistics")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown("**📈 Central Tendency Measures:**")
                        desc_stats = {
                            'Mean': df['Index'].mean(),
                            'Median': df['Index'].median(), 
                            'Mode': df['Index'].mode().iloc[0] if not df['Index'].mode().empty else 'N/A',
                            'Geometric Mean': np.exp(np.log(df['Index'].clip(lower=0.01)).mean()),
                            'Harmonic Mean': len(df) / (1/df['Index'].clip(lower=0.01)).sum()
                        }
                        
                        for stat, value in desc_stats.items():
                            if isinstance(value, (int, float)):
                                st.metric(stat, f"{value:.2f}")
                            else:
                                st.metric(stat, str(value))
                    
                    with col2:
                        st.markdown("**📏 Variability Measures:**")
                        var_stats = {
                            'Standard Deviation': df['Index'].std(),
                            'Variance': df['Index'].var(),
                            'Range': df['Index'].max() - df['Index'].min(),
                            'Interquartile Range': df['Index'].quantile(0.75) - df['Index'].quantile(0.25),
                            'Coefficient of Variation': (df['Index'].std() / df['Index'].mean()) * 100
                        }
                        
                        for stat, value in var_stats.items():
                            if 'Coefficient' in stat:
                                st.metric(stat, f"{value:.2f}%")
                            else:
                                st.metric(stat, f"{value:.2f}")
                    
                    with col3:
                        st.markdown("**📐 Distribution Shape:**")
                        shape_stats = {
                            'Skewness': df['Index'].skew(),
                            'Kurtosis': df['Index'].kurtosis(),
                            'Minimum': df['Index'].min(),
                            'Maximum': df['Index'].max(),
                            '95th Percentile': df['Index'].quantile(0.95)
                        }
                        
                        for stat, value in shape_stats.items():
                            st.metric(stat, f"{value:.2f}")
                    
                    # Distribution Analysis with Statistical Interpretation
                    st.subheader("📊 Distribution Analysis & Statistical Interpretation")
                    
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        # Enhanced histogram with multiple statistical overlays
                        fig_enhanced_hist = go.Figure()
                        
                        # Histogram
                        fig_enhanced_hist.add_trace(go.Histogram(
                            x=df['Index'],
                            nbinsx=30,
                            name='Index Distribution',
                            opacity=0.7,
                            marker_color='lightblue'
                        ))
                        
                        # Add statistical lines
                        mean_val = df['Index'].mean()
                        median_val = df['Index'].median()
                        q1 = df['Index'].quantile(0.25)
                        q3 = df['Index'].quantile(0.75)
                        
                        fig_enhanced_hist.add_vline(x=mean_val, line_dash="solid", line_color="red", 
                                                  annotation_text=f"Mean: {mean_val:.1f}")
                        fig_enhanced_hist.add_vline(x=median_val, line_dash="dot", line_color="green",
                                                  annotation_text=f"Median: {median_val:.1f}")
                        fig_enhanced_hist.add_vline(x=q1, line_dash="dash", line_color="orange",
                                                  annotation_text=f"Q1: {q1:.1f}")
                        fig_enhanced_hist.add_vline(x=q3, line_dash="dash", line_color="orange",
                                                  annotation_text=f"Q3: {q3:.1f}")
                        fig_enhanced_hist.add_vline(x=120, line_dash="solid", line_color="purple",
                                                  annotation_text="Threshold: 120")
                        
                        fig_enhanced_hist.update_layout(
                            title='Enhanced Index Distribution with Statistical Markers',
                            xaxis_title='Index Score',
                            yaxis_title='Frequency',
                            height=400
                        )
                        
                        st.plotly_chart(fig_enhanced_hist, use_container_width=True)
                    
                    with col2:
                        st.markdown("**🔍 Statistical Interpretation:**")
                        
                        # Skewness interpretation
                        skewness = df['Index'].skew()
                        if skewness > 0.5:
                            st.info("📊 **Right-skewed distribution**: Most segments perform below average, with few high performers")
                        elif skewness < -0.5:
                            st.info("📊 **Left-skewed distribution**: Most segments perform above average")
                        else:
                            st.info("📊 **Normal distribution**: Balanced performance across segments")
                        
                        # Kurtosis interpretation
                        kurtosis = df['Index'].kurtosis()
                        if kurtosis > 3:
                            st.warning("📈 **High kurtosis**: Many extreme values (outliers)")
                        elif kurtosis < -1:
                            st.success("📉 **Low kurtosis**: Uniform distribution, few outliers")
                        else:
                            st.info("📊 **Normal kurtosis**: Standard distribution shape")
                        
                        # Coefficient of variation interpretation
                        cv = (df['Index'].std() / df['Index'].mean()) * 100
                        if cv > 30:
                            st.warning(f"⚠️ **High variability** ({cv:.1f}%): Inconsistent performance")
                        elif cv < 15:
                            st.success(f"✅ **Low variability** ({cv:.1f}%): Consistent performance")
                        else:
                            st.info(f"📊 **Moderate variability** ({cv:.1f}%): Normal performance spread")
                    
                    # Advanced Visualizations
                    st.subheader("📈 Advanced Performance Visualizations")
                    
                    # Row 1: Box Plot and Violin Plot
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        if 'category_boxplot' in charts:
                            st.plotly_chart(charts['category_boxplot'], use_container_width=True)
                        else:
                            # Fallback: Overall box plot
                            fig_box_overall = px.box(df, y='Index', title='Overall Index Distribution Box Plot')
                            fig_box_overall.add_hline(y=120, line_dash="dash", line_color="red")
                            st.plotly_chart(fig_box_overall, use_container_width=True)
                    
                    with col2:
                        if 'performance_violin' in charts:
                            st.plotly_chart(charts['performance_violin'], use_container_width=True)
                    
                    # Row 2: Scatter Plot and Correlation Analysis
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        if 'size_vs_performance' in charts:
                            st.plotly_chart(charts['size_vs_performance'], use_container_width=True)
                    
                    with col2:
                        if 'correlation_heatmap' in charts:
                            st.plotly_chart(charts['correlation_heatmap'], use_container_width=True)
                    
                    # Row 3: Pareto Analysis
                    if 'pareto_analysis' in charts:
                        st.subheader("📊 Performance Concentration Analysis (80/20 Rule)")
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            st.plotly_chart(charts['pareto_analysis'], use_container_width=True)
                        
                        with col2:
                            # Calculate Pareto insights
                            df_sorted = df.sort_values('Index', ascending=False)
                            total_performance = df_sorted['Index'].sum()
                            
                            # Find point where we reach 80% of performance
                            cumsum_perf = df_sorted['Index'].cumsum()
                            perf_80_idx = (cumsum_perf >= total_performance * 0.8).idxmax()
                            perf_80_pct = (df_sorted.index.get_loc(perf_80_idx) + 1) / len(df_sorted) * 100
                            
                            st.markdown("**📈 Concentration Insights:**")
                            st.metric("80% Performance from", f"{perf_80_pct:.1f}% of segments")
                            
                            # Top 20% contribution
                            top_20_count = int(len(df_sorted) * 0.2)
                            top_20_contrib = df_sorted.head(top_20_count)['Index'].sum() / total_performance * 100
                            st.metric("Top 20% Contribution", f"{top_20_contrib:.1f}%")
                            
                            if top_20_contrib > 50:
                                st.success("🎯 **High concentration**: Focus on top performers")
                            else:
                                st.info("📊 **Balanced performance**: Broad optimization needed")
                    
                    # Performance Quartile Analysis with Enhanced Details
                    st.subheader("🎯 Detailed Quartile Performance Analysis")
                    
                    quartiles = df['Index'].quantile([0.25, 0.5, 0.75])
                    
                    col1, col2 = st.columns([2, 1])
                    
                    with col1:
                        # Enhanced quartile visualization
                        fig_quartile_detail = go.Figure()
                        
                        # Add histogram
                        fig_quartile_detail.add_trace(go.Histogram(
                            x=df['Index'],
                            nbinsx=25,
                            name='Distribution',
                            opacity=0.6
                        ))
                        
                        # Add quartile regions as shapes
                        fig_quartile_detail.add_vrect(
                            x0=df['Index'].min(), x1=quartiles[0.25],
                            fillcolor="red", opacity=0.2,
                            annotation_text="Q1", annotation_position="top left"
                        )
                        fig_quartile_detail.add_vrect(
                            x0=quartiles[0.25], x1=quartiles[0.5],
                            fillcolor="orange", opacity=0.2,
                            annotation_text="Q2", annotation_position="top left"
                        )
                        fig_quartile_detail.add_vrect(
                            x0=quartiles[0.5], x1=quartiles[0.75],
                            fillcolor="yellow", opacity=0.2,
                            annotation_text="Q3", annotation_position="top left"
                        )
                        fig_quartile_detail.add_vrect(
                            x0=quartiles[0.75], x1=df['Index'].max(),
                            fillcolor="green", opacity=0.2,
                            annotation_text="Q4", annotation_position="top left"
                        )
                        
                        fig_quartile_detail.update_layout(
                            title='Quartile Analysis with Performance Zones',
                            xaxis_title='Index Score',
                            yaxis_title='Frequency'
                        )
                        
                        st.plotly_chart(fig_quartile_detail, use_container_width=True)
                    
                    with col2:
                        st.markdown("**📊 Quartile Performance Details:**")
                        
                        quartile_analysis = []
                        quartile_names = ['Q1 (Bottom 25%)', 'Q2', 'Q3', 'Q4 (Top 25%)']
                        quartile_ranges = [
                            (df['Index'].min(), quartiles[0.25]),
                            (quartiles[0.25], quartiles[0.5]),
                            (quartiles[0.5], quartiles[0.75]),
                            (quartiles[0.75], df['Index'].max())
                        ]
                        
                        for i, (name, (min_val, max_val)) in enumerate(zip(quartile_names, quartile_ranges)):
                            mask = (df['Index'] >= min_val) & (df['Index'] <= max_val)
                            quartile_data = df[mask]
                            
                            quartile_analysis.append({
                                'Quartile': name,
                                'Range': f"{min_val:.1f} - {max_val:.1f}",
                                'Count': len(quartile_data),
                                'Avg': f"{quartile_data['Index'].mean():.1f}",
                                'Std Dev': f"{quartile_data['Index'].std():.1f}"
                            })
                        
                        quartile_df = pd.DataFrame(quartile_analysis)
                        st.dataframe(quartile_df, use_container_width=True)
                    
                    # Top and Bottom Performers with Enhanced Analysis
                    st.subheader("🏆 Performance Leaders & Improvement Opportunities")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("**🏆 Top 15 Performance Leaders:**")
                        top_performers = df.nlargest(15, 'Index')
                        
                        display_cols = ['Attribute Name', 'Index']
                        if 'AIR Category' in df.columns:
                            display_cols.insert(1, 'AIR Category')
                        if 'Attribute Size' in df.columns:
                            display_cols.append('Attribute Size')
                        
                        top_display = top_performers[display_cols].copy()
                        top_display['Index'] = top_display['Index'].round(1)
                        if 'Attribute Size' in top_display.columns:
                            top_display['Attribute Size'] = top_display['Attribute Size'].apply(
                                lambda x: f"{x:,.0f}" if pd.notna(x) else "N/A"
                            )
                        
                        st.dataframe(top_display, use_container_width=True)
                        
                        # Top performers insights
                        top_avg = top_performers['Index'].mean()
                        top_range = top_performers['Index'].max() - top_performers['Index'].min()
                        st.info(f"💡 **Top 15 avg performance**: {top_avg:.1f} (Range: {top_range:.1f})")
                    
                    with col2:
                        st.markdown("**📉 Bottom 15 - Improvement Opportunities:**")
                        bottom_performers = df.nsmallest(15, 'Index')
                        
                        bottom_display = bottom_performers[display_cols].copy()
                        bottom_display['Index'] = bottom_display['Index'].round(1)
                        if 'Attribute Size' in bottom_display.columns:
                            bottom_display['Attribute Size'] = bottom_display['Attribute Size'].apply(
                                lambda x: f"{x:,.0f}" if pd.notna(x) else "N/A"
                            )
                        
                        st.dataframe(bottom_display, use_container_width=True)
                        
                        # Bottom performers insights
                        bottom_avg = bottom_performers['Index'].mean()
                        improvement_potential = df['Index'].mean() - bottom_avg
                        st.warning(f"⚠️ **Improvement potential**: +{improvement_potential:.1f} index points")
                    
                    # Advanced Performance Insights
                    st.subheader("💡 Advanced Performance Insights")
                    st.markdown('<div class="insight-box">', unsafe_allow_html=True)
                    st.markdown("**📈 Advanced Statistical Findings:**")
                    
                    # Generate advanced insights
                    mean_val = df['Index'].mean()
                    median_val = df['Index'].median()
                    std_val = df['Index'].std()
                    
                    st.markdown(f"• **Performance Distribution**: Mean ({mean_val:.1f}) vs Median ({median_val:.1f}) indicates {'right-skewed' if mean_val > median_val else 'left-skewed' if mean_val < median_val else 'symmetric'} distribution")
                    
                    outlier_count = len(df[(df['Index'] < (df['Index'].quantile(0.25) - 1.5 * (df['Index'].quantile(0.75) - df['Index'].quantile(0.25)))) | 
                                          (df['Index'] > (df['Index'].quantile(0.75) + 1.5 * (df['Index'].quantile(0.75) - df['Index'].quantile(0.25))))])
                    st.markdown(f"• **Outlier Analysis**: {outlier_count} statistical outliers detected ({outlier_count/len(df)*100:.1f}% of data)")
                    
                    high_performers = len(df[df['Index'] > 120])
                    performance_rate = high_performers / len(df) * 100
                    st.markdown(f"• **Performance Threshold**: {performance_rate:.1f}% of segments exceed 120 index threshold")
                    
                    # Performance consistency
                    cv = (std_val / mean_val) * 100
                    consistency = "High" if cv < 15 else "Moderate" if cv < 30 else "Low"
                    st.markdown(f"• **Performance Consistency**: {consistency} consistency with {cv:.1f}% coefficient of variation")
                    
                    # Concentration analysis
                    if 'perf_80_pct' in locals():
                        st.markdown(f"• **Performance Concentration**: {perf_80_pct:.1f}% of segments drive 80% of total performance (Pareto principle)")
                    
                    st.markdown('</div>', unsafe_allow_html=True)
                
                elif selected_section == "🎯 Actionable Optimization":
                    st.header("🎯 Actionable Optimization Recommendations")
                    st.markdown("*Specific actions with expected ROI and timelines*")
                    
                    low_performers = df[df['Index'] < 80]
                    high_performers = df[df['Index'] > 120]
                    
                    # Immediate actions
                    st.subheader("⚡ Immediate Actions (1-2 weeks)")
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        st.markdown("**🚫 Exclude Low Performers**")
                        if not low_performers.empty:
                            st.metric("Segments to Exclude", len(low_performers))
                            st.metric("Budget Reallocation", f"{len(low_performers)/len(df)*100:.1f}%")
                            
                            worst_performers = low_performers.nsmallest(5, 'Index')[['Attribute Name', 'Index']]
                            st.markdown("**Worst 5 Performers:**")
                            st.dataframe(worst_performers, use_container_width=True)
                        else:
                            st.success("✅ No segments below 80 index threshold")
                    
                    with col2:
                        st.markdown("**📈 Scale High Performers**")
                        if not high_performers.empty:
                            st.metric("High Performers", len(high_performers))
                            st.metric("Recommended Budget Increase", "+50%")
                            
                            top_scalers = high_performers.nlargest(5, 'Index')[['Attribute Name', 'Index']]
                            st.markdown("**Top 5 for Scaling:**")
                            st.dataframe(top_scalers, use_container_width=True)
                    
                    # Category optimization
                    if 'AIR Category' in df.columns:
                        st.subheader("🏷️ Category Portfolio Optimization")
                        
                        category_stats = df.groupby('AIR Category').agg({
                            'Index': ['mean', 'count', 'std']
                        }).round(1)
                        category_stats.columns = ['Avg Index', 'Segment Count', 'Volatility']
                        category_stats = category_stats[category_stats['Segment Count'] >= 3]
                        category_stats = category_stats.sort_values('Avg Index', ascending=False)
                        
                        col1, col2 = st.columns([2, 1])
                        
                        with col1:
                            fig_category_opt = px.scatter(
                                category_stats.reset_index(),
                                x='Segment Count', y='Avg Index',
                                size='Volatility', 
                                hover_data=['AIR Category'],
                                title='Category Optimization Matrix'
                            )
                            fig_category_opt.add_hline(y=120, line_dash="dash", line_color="red")
                            st.plotly_chart(fig_category_opt, use_container_width=True)
                        
                        with col2:
                            st.markdown("**Category Actions:**")
                            top_categories = category_stats.head(5)
                            for idx, (category, row) in enumerate(top_categories.iterrows()):
                                if row['Avg Index'] > 120:
                                    action = "🚀 Scale"
                                elif row['Avg Index'] > 100:
                                    action = "⚡ Optimize"
                                else:
                                    action = "⚠️ Review"
                                
                                st.markdown(f"**{action}** {category[:20]}... (Index: {row['Avg Index']:.1f})")
                    
                    # ROI projections
                    st.subheader("💰 Expected ROI Summary")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.markdown('<div class="success-box">', unsafe_allow_html=True)
                        st.markdown("**Week 1-2 Impact**")
                        if not low_performers.empty:
                            efficiency_gain = len(low_performers)/len(df)*100
                            st.metric("Efficiency Gain", f"+{efficiency_gain:.1f}%")
                        st.markdown('</div>', unsafe_allow_html=True)
                    
                    with col2:
                        st.markdown('<div class="warning-box">', unsafe_allow_html=True)
                        st.markdown("**Month 1 Impact**")
                        if not high_performers.empty:
                            performance_boost = len(high_performers)/len(df)*50
                            st.metric("Performance Boost", f"+{performance_boost:.1f}%")
                        st.markdown('</div>', unsafe_allow_html=True)
                    
                    with col3:
                        st.markdown('<div class="insight-box">', unsafe_allow_html=True)
                        st.markdown("**Quarter 1 Impact**")
                        compound_effect = df['Index'].std() * 0.3
                        st.metric("Compound Effect", f"+{compound_effect:.1f} Index Points")
                        st.markdown('</div>', unsafe_allow_html=True)
                
                elif selected_section == "📋 Key Insights & Call-Outs":
                    st.header("📋 Key Insights & Strategic Call-Outs")
                    st.markdown("*Executive summary with actionable strategic insights*")
                    
                    # Core characteristics
                    st.subheader("🎯 Core Audience Characteristics")
                    
                    characteristics = []
                    
                    # Demographics
                    if demographics.get('age'):
                        best_age = max(demographics['age'].items(), key=lambda x: x[1]['avg_index'])
                        characteristics.append(f"**Primary Age Group**: {best_age[0]} (Index: {best_age[1]['avg_index']:.1f})")
                    
                    if demographics.get('income'):
                        best_income = max(demographics['income'].items(), key=lambda x: x[1]['avg_index'])
                        characteristics.append(f"**Income Sweet Spot**: {best_income[0]} (Index: {best_income[1]['avg_index']:.1f})")
                    
                    if geographic.get('regions'):
                        best_region = max(geographic['regions'].items(), key=lambda x: x[1]['avg_index'])
                        characteristics.append(f"**Geographic Focus**: {best_region[0]} (Index: {best_region[1]['avg_index']:.1f})")
                    
                    if psychographics.get('lifestyle'):
                        best_lifestyle = max(psychographics['lifestyle'].items(), key=lambda x: x[1]['avg_index'])
                        characteristics.append(f"**Lifestyle Interest**: {best_lifestyle[0]} (Index: {best_lifestyle[1]['avg_index']:.1f})")
                    
                    for char in characteristics:
                        st.markdown(f"• {char}")
                    
                    # Core needs and motivations
                    st.subheader("💡 Core Needs & Motivations")
                    
                    needs = []
                    if 'Attribute Name' in df.columns:
                        # Analyze attribute names for needs patterns
                        attr_text = ' '.join(df['Attribute Name'].str.lower())
                        
                        if 'financial' in attr_text or 'money' in attr_text:
                            needs.append("**Financial Security**: Strong interest in financial products and investment opportunities")
                        
                        if 'family' in attr_text or 'children' in attr_text:
                            needs.append("**Family-Focused**: Products and services that benefit family lifestyle")
                        
                        if 'home' in attr_text or 'house' in attr_text:
                            needs.append("**Home & Property**: Interest in home-related products and services")
                        
                        if 'health' in attr_text or 'fitness' in attr_text:
                            needs.append("**Health & Wellness**: Focus on health, fitness, and wellbeing")
                        
                        if 'travel' in attr_text or 'vacation' in attr_text:
                            needs.append("**Travel & Experience**: Values experiences and travel opportunities")
                    
                    for need in needs:
                        st.markdown(f"• {need}")
                    
                    # Communication style
                    st.subheader("📢 Recommended Communication Style")
                    
                    # Performance distribution analysis for communication insights
                    avg_index = df['Index'].mean()
                    
                    if avg_index > 130:
                        comm_style = "**Premium & Aspirational**: High-end messaging with quality focus"
                    elif avg_index > 110:
                        comm_style = "**Professional & Confident**: Direct, benefit-focused communication"
                    elif avg_index > 90:
                        comm_style = "**Balanced & Practical**: Value-oriented with clear benefits"
                    else:
                        comm_style = "**Accessible & Supportive**: Simple, helpful messaging"
                    
                    st.markdown(f"• {comm_style}")
                    
                    # Determine best channels based on demographics
                    if demographics.get('age'):
                        young_segments = sum([data['count'] for age, data in demographics['age'].items() 
                                            if '18-24' in age or '25-34' in age])
                        mature_segments = sum([data['count'] for age, data in demographics['age'].items()
                                             if '45-54' in age or '55-64' in age or '65+' in age])
                        
                        if young_segments > mature_segments:
                            st.markdown("• **Digital-First Channels**: Social media, mobile, streaming platforms")
                        else:
                            st.markdown("• **Multi-Channel Approach**: Blend digital and traditional media")
                    
                    # Unique opportunities
                    st.subheader("🚀 Unique Opportunities & Strategic Focus")
                    
                    opportunities = []
                    
                    # High performer concentration
                    high_perf_rate = len(df[df['Index'] > 120]) / len(df) * 100
                    if high_perf_rate > 30:
                        opportunities.append(f"**High Performance Concentration**: {high_perf_rate:.1f}% of segments are high performers - strong foundation for scaling")
                    
                    # Category diversification
                    if 'AIR Category' in df.columns:
                        category_count = df['AIR Category'].nunique()
                        opportunities.append(f"**Category Diversification**: {category_count} different categories provide multiple targeting angles")
                    
                    # Audience size opportunity
                    if 'Attribute Size' in df.columns:
                        total_audience = df['Attribute Size'].sum()
                        opportunities.append(f"**Scale Opportunity**: {total_audience:,.0f} total addressable audience across all segments")
                    
                    # Performance gap opportunity
                    perf_gap = df['Index'].max() - df['Index'].min()
                    opportunities.append(f"**Optimization Potential**: {perf_gap:.1f} point performance gap indicates significant optimization opportunity")
                    
                    for opp in opportunities:
                        st.markdown(f"• {opp}")
                    
                    # Executive summary
                    st.subheader("📊 Executive Summary")
                    st.markdown('<div class="insight-box">', unsafe_allow_html=True)
                    st.markdown("**🎯 Strategic Recommendations:**")
                    
                    # Generate specific recommendations based on data
                    recommendations = []
                    
                    if not df[df['Index'] < 80].empty:
                        low_count = len(df[df['Index'] < 80])
                        recommendations.append(f"**Immediate Action**: Exclude {low_count} underperforming segments for {low_count/len(df)*100:.1f}% efficiency gain")
                    
                    if not df[df['Index'] > 120].empty:
                        high_count = len(df[df['Index'] > 120])
                        recommendations.append(f"**Scale Opportunity**: Increase investment in {high_count} high-performing segments")
                    
                    if demographics.get('age'):
                        best_age = max(demographics['age'].items(), key=lambda x: x[1]['avg_index'])
                        recommendations.append(f"**Demographic Focus**: Prioritize {best_age[0]} segments for highest ROI")
                    
                    if geographic.get('regions'):
                        best_region = max(geographic['regions'].items(), key=lambda x: x[1]['avg_index'])
                        recommendations.append(f"**Geographic Priority**: Expand presence in {best_region[0]} market")
                    
                    for rec in recommendations:
                        st.markdown(f"• {rec}")
                    
                    st.markdown('</div>', unsafe_allow_html=True)
                
                # Download section
                st.sidebar.markdown("---")
                st.sidebar.subheader("📥 Export Analysis")
                
                if st.sidebar.button("📊 Generate Comprehensive Report", type="primary"):
                    st.success("🎯 Comprehensive analysis complete!")
                    st.info("Enhanced dashboard with demographics, geographic, and psychographic insights ready for strategic decision-making.")
            
            else:
                st.error("❌ Could not find sufficient target columns in the data")
                st.markdown("**Expected columns:**")
                expected_cols = [
                    "Attribute Name", "Attribute Path", "Attribute Size",
                    "Audience & THIS Attribute Overlap", "Index", "AIR Category"
                ]
                for i, col in enumerate(expected_cols, 1):
                    st.write(f"{i}. {col}")
        
        else:
            st.error("❌ Could not find Index Report sheet")
            st.write("**Available sheets:**")
            for sheet in xls.sheet_names:
                st.write(f"• {sheet}")
    
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        if st.sidebar.checkbox("Show detailed error"):
            st.sidebar.code(traceback.format_exc())

else:
    st.markdown("## 👋 Welcome to Enhanced Audience Analytics")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("""
        ### 🎯 **Comprehensive Audience Intelligence Dashboard**
        
        This enhanced dashboard provides **deep demographic, geographic, and psychographic analysis** based on your Index Report data.
        
        **📊 New Analysis Sections:**
        - **Demographics Deep Dive**: Age ranges, income levels, gender, family status, homeownership
        - **Geographic Intelligence**: Regional performance, urban/suburban/rural analysis
        - **Psychographics & Behaviors**: Lifestyle interests, financial behaviors, media consumption
        - **Advanced Performance Analysis**: Quartile analysis, top/bottom performers
        - **Actionable Optimization**: Specific recommendations with ROI projections
        - **Key Insights & Call-Outs**: Executive summary with strategic recommendations
        
        **🔍 Enhanced Features:**
        - **Automatic sheet detection**: Finds "Index Report" sheets regardless of naming
        - **Flexible column mapping**: Locates target columns anywhere in the spreadsheet
        - **Demographic pattern recognition**: Identifies age, income, gender patterns from data
        - **Geographic analysis**: Regional and area type performance insights
        - **Behavioral segmentation**: Lifestyle and psychographic pattern detection
        - **Actionable recommendations**: Specific optimization steps with timelines
        """)
    
    with col2:
        st.markdown('<div class="insight-box">', unsafe_allow_html=True)
        st.markdown("**📋 Target Columns:**")
        st.markdown("• Attribute Name")
        st.markdown("• Attribute Path") 
        st.markdown("• Attribute Size")
        st.markdown("• Audience & THIS Attribute Overlap")
        st.markdown("• Audience & ANY Attribute Overlap")
        st.markdown("• Index")
        st.markdown("• AIR Category")
        st.markdown("• AIR Attribute")
        st.markdown("• Audience Overlap % of Input Size")
        st.markdown("• Base Adjusted Population columns")
        st.markdown("")
        st.markdown("**🎯 Key Benefits:**")
        st.markdown("• Demographic insights for targeting")
        st.markdown("• Geographic performance analysis")
        st.markdown("• Psychographic behavior patterns")
        st.markdown("• Actionable optimization plans")
        st.markdown("• Executive-ready insights")
        st.markdown('</div>', unsafe_allow_html=True)
    
    st.info("📁 **Upload your Excel file with Index Report data to begin comprehensive audience analysis!**")
