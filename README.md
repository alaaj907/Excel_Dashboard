# Excel_Dashboard

#  Audience Analytics Dashboard

This Streamlit-based web app lets users upload an Excel file and instantly generate an interactive analytics dashboard — featuring dynamic filtering, performance metrics, and rich visualizations.

##  Features

*  Upload any Excel file with "Index Report" sheet- the excel file must have an "Index Report" sheet.
*  Dashboard KPIs: Total attributes, average index score, high performers
*  Filter by attribute group
* Top performing attributes (bar chart)
*  Attribute group distribution (pie chart)
*  Index vs. Relative Lift (scatterplot)
*  Full data table with search and scroll
*  Export filtered data as CSV

##  Ideal Use Cases

The dashboard now includes all the requested features with proper indentation:
🎯 10 Comprehensive Analysis Sections:

🏠 Overview - Executive metrics and performance distribution
📈 Performance Analysis - Detailed performance breakdown with charts
💰 Income Profiling - Income bracket analysis with inverse relationship
👨‍👩‍👧‍👦 Family Lifecycle - Family and children targeting insights
🗺️ Geographic Analysis - State-level performance mapping
🎯 Audience Overlap - Penetration vs performance analysis
📋 Category Performance - Category rankings and leaderboards
⚠️ Exclusion Opportunities - Low performer identification
🧮 Sizing Intelligence - Reach vs relevance optimization
📑 Strategic Recommendations - Action plans and KPIs

🚀 Key Features:

Sidebar Navigation for easy section switching
Performance Buckets automatically categorize attributes
Interactive Charts with Plotly visualizations
Professional Styling with custom CSS
Error Handling for missing columns
PowerPoint Export with executive insights
Data-Driven Recommendations based on your analysis

💡 Smart Analytics:

Automated Performance Segmentation
Income Inverse Relationship Analysis
Geographic Hot Spot Identification
Family Targeting Optimization
Budget Reallocation Recommendations
Strategic Action Plans with Timelines

##  How to Use

### 1. **Upload Your File**

Upload any `.xlsx` Excel file. The app will automatically read the **first sheet**.

### 2. **Interact with the Dashboard**

* View KPIs and charts instantly
* Use the dropdown to filter by `attribute_group`
* Explore detailed data in the table

### 3. **Download Your Results**

* Use the download button to export filtered data

##  Requirements (for local use)

```bash
pip install streamlit pandas plotly openpyxl
```

##  Run the App Locally

```bash
streamlit run app.py
```

##  Deploy on Streamlit Cloud

1. Push your code to a GitHub repository
2. Go to [streamlit.io/cloud](https://streamlit.io/cloud)
3. Click "New app" and select your repo and `app.py`

## File Structure

```
streamlit-dashboard/
├── app.py               # Main application file
├── requirements.txt     # Python dependencies
└── README.md            # You’re reading it 
```

---
