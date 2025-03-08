abc
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import numpy as np
from io import BytesIO
import os
import datetime
import json
from datetime import timedelta
from PIL import Image

# Page configuration
st.set_page_config(
    page_title="Sales Analytics Dashboard", 
    layout="wide",
    initial_sidebar_state="expanded"
)

# Apply custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: 600;
        color: #1E3A8A;
        margin-bottom: 1rem;
    }
    .metric-card {
        background-color: #f8f9fa;
        border-radius: 10px;
        padding: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        text-align: center;
        transition: transform 0.3s;
    }
    .metric-card:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 8px rgba(0, 0, 0, 0.15);
    }
    .metric-value {
        font-size: 2rem;
        font-weight: 700;
        color: #2563EB;
    }
    .metric-label {
        font-size: 1rem;
        color: #4B5563;
        margin-bottom: 0.5rem;
    }
    .section-header {
        font-size: 1.5rem;
        font-weight: 600;
        color: #1F2937;
        margin-top: 2rem;
        margin-bottom: 1rem;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #E5E7EB;
    }
    .tab-content {
        padding: 20px 0;
    }
    .filters-container {
        background-color: #f1f5f9;
        padding: 15px;
        border-radius: 8px;
        margin-bottom: 20px;
    }
    .stDataFrame {
        border-radius: 8px !important;
        overflow: hidden !important;
    }
    div[data-testid="stSidebarContent"] {
        background-color: #f8fafc;
    }
    .stButton>button {
        background-color: #2563EB;
        color: white;
        border-radius: 4px;
        padding: 0.5rem 1rem;
        border: none;
    }
    .stButton>button:hover {
        background-color: #1D4ED8;
    }
    .tooltip {
        position: relative;
        display: inline-block;
        border-bottom: 1px dotted black;
    }
    .tooltip .tooltiptext {
        visibility: hidden;
        width: 200px;
        background-color: #555;
        color: #fff;
        text-align: center;
        border-radius: 6px;
        padding: 5px;
        position: absolute;
        z-index: 1;
        bottom: 125%;
        left: 50%;
        margin-left: -100px;
        opacity: 0;
        transition: opacity 0.3s;
    }
    .tooltip:hover .tooltiptext {
        visibility: visible;
        opacity: 1;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state for data storage
if 'geographic_data' not in st.session_state:
    st.session_state.geographic_data = None
if 'state_data' not in st.session_state:
    st.session_state.state_data = None
if 'zip_data' not in st.session_state:
    st.session_state.zip_data = None
if 'last_upload_date' not in st.session_state:
    st.session_state.last_upload_date = None
if 'time_period' not in st.session_state:
    st.session_state.time_period = "Year"
if 'selected_year' not in st.session_state:
    st.session_state.selected_year = datetime.datetime.now().year
if 'selected_quarter' not in st.session_state:
    st.session_state.selected_quarter = (datetime.datetime.now().month - 1) // 3 + 1
if 'selected_month' not in st.session_state:
    st.session_state.selected_month = datetime.datetime.now().month

# Function to clean state names
def clean_state_names(df):
    state_mapping = {
        'NY': 'New York',
        'NewYork': 'New York',
        'New York, NY': 'New York',
        'CA': 'California',
        'FL': 'Florida',
        'TX': 'Texas',
        'PA': 'Pennsylvania',
        'IL': 'Illinois',
        'OH': 'Ohio',
        'GA': 'Georgia',
        'NC': 'North Carolina',
        'MI': 'Michigan',
        'NJ': 'New Jersey',
        'VA': 'Virginia',
        'WA': 'Washington',
        'AZ': 'Arizona',
        'MA': 'Massachusetts',
        'TN': 'Tennessee',
        'IN': 'Indiana',
        'MO': 'Missouri',
        'MD': 'Maryland',
        'WI': 'Wisconsin',
        'MN': 'Minnesota',
        'CO': 'Colorado',
        'AL': 'Alabama',
        'SC': 'South Carolina',
        'LA': 'Louisiana',
        'KY': 'Kentucky',
        'OR': 'Oregon',
        'OK': 'Oklahoma',
        'CT': 'Connecticut',
        'IA': 'Iowa',
        'MS': 'Mississippi',
        'AR': 'Arkansas',
        'KS': 'Kansas',
        'UT': 'Utah',
        'NV': 'Nevada',
        'NM': 'New Mexico',
        'WV': 'West Virginia',
        'NE': 'Nebraska',
        'ID': 'Idaho',
        'HI': 'Hawaii',
        'ME': 'Maine',
        'NH': 'New Hampshire',
        'RI': 'Rhode Island',
        'MT': 'Montana',
        'DE': 'Delaware',
        'SD': 'South Dakota',
        'AK': 'Alaska',
        'ND': 'North Dakota',
        'VT': 'Vermont',
        'WY': 'Wyoming',
        'DC': 'District of Columbia'
    }
    if 'State' in df.columns:
        df['State'] = df['State'].replace(state_mapping)
    return df

# Function to load Excel data
def load_excel_data(file, sheet_name=None):
    try:
        if sheet_name:
            df = pd.read_excel(file, sheet_name=sheet_name)
        else:
            df = pd.read_excel(file)
        
        # Add date fields for filtering - use current date for all records
        # This is needed for the time period filtering to work with the actual data
        if 'Date' not in df.columns:
            df['Date'] = datetime.datetime.now()
            
        # Extract year, quarter, month
        df['Year'] = datetime.datetime.now().year
        df['Quarter'] = (datetime.datetime.now().month - 1) // 3 + 1
        df['Month'] = datetime.datetime.now().month
        df['Month_Name'] = datetime.datetime.now().strftime('%B')
        
        return df
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return None

# Function to download data as Excel
def download_excel(df, sheet_name='Data'):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
    return output.getvalue()

# Function to merge geographic data with sales data
def merge_data(geo_df, sales_df, key='ASIN'):
    if geo_df is None or sales_df is None:
        return None
    
    # Merge on ASIN
    merged_df = pd.merge(sales_df, geo_df, on=key, how='left')
    return merged_df

# Function to filter data based on time period
def filter_data_by_time(df, time_period, year, quarter=None, month=None):
    if df is None:
        return None
    
    filtered_df = df.copy()
    
    if 'Year' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['Year'] == year]
        
        if time_period == 'Quarter' and 'Quarter' in filtered_df.columns and quarter is not None:
            filtered_df = filtered_df[filtered_df['Quarter'] == quarter]
        
        elif time_period == 'Month' and 'Month' in filtered_df.columns and month is not None:
            filtered_df = filtered_df[filtered_df['Month'] == month]
    
    return filtered_df

# Function to create improved bar chart
def create_bar_chart(data, x_column, y_column, title, color_column=None):
    if data is None or data.empty:
        return go.Figure()
        
    # Sort by value for better visualization
    sorted_data = data.sort_values(y_column, ascending=False)
    
    # Create figure
    if color_column:
        fig = px.bar(
            sorted_data,
            x=x_column,
            y=y_column,
            title=title,
            color=color_column,
            text_auto='.2s',
            color_continuous_scale='Viridis'
        )
    else:
        fig = px.bar(
            sorted_data,
            x=x_column,
            y=y_column,
            title=title,
            color=y_column,
            text_auto='.2s',
            color_continuous_scale='Blues'
        )
    
    # Improve layout
    fig.update_layout(
        title=dict(
            text=title,
            font=dict(size=20, color="#1E3A8A"),
            x=0.5,
            xanchor='center'
        ),
        xaxis_title=None,
        yaxis_title=y_column.replace('_', ' '),
        legend_title_text=color_column.replace('_', ' ') if color_column else None,
        height=400,
        margin=dict(l=20, r=20, t=50, b=30),
        hoverlabel=dict(
            bgcolor="white",
            font_size=12,
            font_family="Arial"
        ),
        plot_bgcolor='rgba(240,242,246,0.3)'
    )
    
    # Improve bar appearance
    fig.update_traces(
        marker_line_color='rgb(255,255,255)',
        marker_line_width=1,
        opacity=0.9
    )
    
    return fig

# Function to create enhanced US state map
# Function to create enhanced US state map
# Function to create enhanced US state map with premium styling
def create_state_map(df, metric='Shipped Revenue'):
    if df is None or df.empty:
        return go.Figure()
        
    # Group by state
    state_summary = df.groupby('State').agg({
        'Shipped Revenue': 'sum',
        'Shipped Units': 'sum',
        'Shipped COGS': 'sum'
    }).reset_index()
    
    # Calculate profit and margin
    state_summary['Profit'] = state_summary['Shipped Revenue'] - state_summary['Shipped COGS']
    state_summary['Margin %'] = (state_summary['Profit'] / state_summary['Shipped Revenue'] * 100).round(2)
    
    # Create enhanced choropleth map with multiple metrics
    fig = go.Figure()
    
    # Premium color scales with gradient effects
    if metric == 'Shipped Revenue':
        colorscale = [[0, "#e6f2ff"], [0.2, "#99ccff"], [0.4, "#3399ff"], [0.6, "#0066cc"], [0.8, "#004080"], [1, "#002147"]]
        colorbar_title = "Revenue ($)"
        z_values = state_summary['Shipped Revenue']
        bubble_color = "rgba(0, 102, 204, 0.8)"  # Blue themed
    elif metric == 'Profit':
        colorscale = [[0, "#e6ffe6"], [0.2, "#99ff99"], [0.4, "#33cc33"], [0.6, "#009900"], [0.8, "#006600"], [1, "#003300"]]
        colorbar_title = "Profit ($)"
        z_values = state_summary['Profit']
        bubble_color = "rgba(0, 153, 0, 0.8)"  # Green themed
    elif metric == 'Margin %':
        colorscale = [[0, "#fff5e6"], [0.2, "#ffcc80"], [0.4, "#ffaa33"], [0.6, "#ff8800"], [0.8, "#cc5500"], [1, "#802b00"]]
        colorbar_title = "Margin (%)"
        z_values = state_summary['Margin %']
        bubble_color = "rgba(255, 136, 0, 0.8)"  # Orange themed
    else:  # Default to units
        colorscale = [[0, "#f5e6ff"], [0.2, "#d699ff"], [0.4, "#ac33ff"], [0.6, "#8800ff"], [0.8, "#5500cc"], [1, "#330080"]]
        colorbar_title = "Units Sold"
        z_values = state_summary['Shipped Units']
        bubble_color = "rgba(136, 0, 255, 0.8)"  # Purple themed
    
    # Add choropleth layer with premium styling
    fig.add_trace(go.Choropleth(
        locations=state_summary['State'],
        z=z_values,
        locationmode='USA-states',
        colorscale=colorscale,
        colorbar_title=dict(
            text=colorbar_title,
            font=dict(size=14, family="Arial", color="#333333")
        ),
        colorbar=dict(
            thickness=15,
            len=0.7,
            outlinewidth=0,
            bgcolor='rgba(255,255,255,0.0)',
            tickfont=dict(size=12, family="Arial")
        ),
        marker_line_color='white',
        marker_line_width=0.7,
        name=metric,
        hovertemplate='<b>%{location}</b><br>' +
                     f'{metric}: ' + 
                     ('%{z:,.2f}%' if metric == 'Margin %' else '$%{z:,.2f}' if metric in ['Shipped Revenue', 'Profit'] else '%{z:,}') +
                     '<extra></extra>'
    ))
    
    # Add custom hover data with additional metrics
    hover_text = []
    for index, row in state_summary.iterrows():
        hover_text.append(
            f"<b>{row['State']}</b><br>" +
            f"Revenue: ${row['Shipped Revenue']:,.2f}<br>" +
            f"Units: {row['Shipped Units']:,}<br>" +
            f"Profit: ${row['Profit']:,.2f}<br>" +
            f"Margin: {row['Margin %']:.1f}%"
        )
    
    # Calculate state centroids (approximate)
    states_centroids = {
        'Alabama': [32.806671, -86.791130], 'Alaska': [61.370716, -152.404419], 'Arizona': [33.729759, -111.431221],
        'Arkansas': [34.969704, -92.373123], 'California': [36.116203, -119.681564], 'Colorado': [39.059811, -105.311104],
        'Connecticut': [41.597782, -72.755371], 'Delaware': [39.318523, -75.507141], 'Florida': [27.766279, -81.686783],
        'Georgia': [33.040619, -83.643074], 'Hawaii': [21.094318, -157.498337], 'Idaho': [44.240459, -114.478828],
        'Illinois': [40.349457, -88.986137], 'Indiana': [39.849426, -86.258278], 'Iowa': [42.011539, -93.210526],
        'Kansas': [38.526600, -96.726486], 'Kentucky': [37.668140, -84.670067], 'Louisiana': [31.169546, -91.867805],
        'Maine': [44.693947, -69.381927], 'Maryland': [39.063946, -76.802101], 'Massachusetts': [42.230171, -71.530106],
        'Michigan': [43.326618, -84.536095], 'Minnesota': [45.694454, -93.900192], 'Mississippi': [32.741646, -89.678696],
        'Missouri': [38.456085, -92.288368], 'Montana': [46.921925, -110.454353], 'Nebraska': [41.125370, -98.268082],
        'Nevada': [38.313515, -117.055374], 'New Hampshire': [43.452492, -71.563896], 'New Jersey': [40.298904, -74.521011],
        'New Mexico': [34.840515, -106.248482], 'New York': [42.165726, -74.948051], 'North Carolina': [35.630066, -79.806419],
        'North Dakota': [47.528912, -99.784012], 'Ohio': [40.388783, -82.764915], 'Oklahoma': [35.565342, -96.928917],
        'Oregon': [44.572021, -122.070938], 'Pennsylvania': [40.590752, -77.209755], 'Rhode Island': [41.680893, -71.511780],
        'South Carolina': [33.856892, -80.945007], 'South Dakota': [44.299782, -99.438828], 'Tennessee': [35.747845, -86.692345],
        'Texas': [31.054487, -97.563461], 'Utah': [40.150032, -111.862434], 'Vermont': [44.045876, -72.710686],
        'Virginia': [37.769337, -78.169968], 'Washington': [47.400902, -121.490494], 'West Virginia': [38.491226, -80.954453],
        'Wisconsin': [44.268543, -89.616508], 'Wyoming': [42.755966, -107.302490], 'District of Columbia': [38.9072, -77.0369]
    }
    
    # Add markers for units sold with premium styling
    marker_sizes = []
    marker_lats = []
    marker_lons = []
    marker_texts = []
    
    for index, row in state_summary.iterrows():
        if row['State'] in states_centroids:
            marker_lats.append(states_centroids[row['State']][0])
            marker_lons.append(states_centroids[row['State']][1])
            # Better size scaling for bubbles
            marker_sizes.append(np.sqrt(row['Shipped Units']) / 4 + 12)
            marker_texts.append(hover_text[index])
    
    # Add bubbles with 3D effect
    fig.add_trace(go.Scattergeo(
        lon=marker_lons,
        lat=marker_lats,
        text=marker_texts,
        mode='markers',
        marker=dict(
            size=marker_sizes,
            color=bubble_color,
            opacity=0.8,
            line=dict(
                width=1.5,
                color='rgba(255, 255, 255, 0.8)'
            ),
            gradient=dict(
                type='radial',
                color='rgba(255, 255, 255, 0.5)'
            ),
            sizemode='diameter',
            sizeref=0.9
        ),
        hoverinfo='text',
        name='Units Sold'
    ))
    
    # Set premium map layout
    fig.update_layout(
        title=dict(
            text="Sales Analytics by State",
            font=dict(size=26, family="Arial", color="#1E3A8A"),
            x=0.5,
            xanchor='center',
            y=0.97
        ),
        legend_title=dict(
            text="Metrics",
            font=dict(size=14, family="Arial")
        ),
        geo=dict(
            scope='usa',
            projection_type='albers usa',
            showlakes=True,
            lakecolor='rgba(211, 233, 250, 0.8)',
            showcoastlines=True,
            coastlinecolor='rgba(220, 220, 220, 0.8)',
            showland=True,
            landcolor='rgba(250, 250, 250, 0.95)',
            countrycolor='rgba(220, 220, 220, 0.8)',
            showsubunits=True,
            subunitcolor='rgba(220, 220, 220, 0.8)',
            showcountries=True,
            resolution=50,
            lonaxis=dict(range=[-125, -66]),
            lataxis=dict(range=[24, 50]),
            bgcolor='rgba(255, 255, 255, 0)'
        ),
        paper_bgcolor='rgba(255, 255, 255, 0)',
        plot_bgcolor='rgba(255, 255, 255, 0)',
        height=650,  # Slightly larger for better visibility
        margin=dict(l=0, r=0, t=60, b=0),
        hoverlabel=dict(
            bgcolor="white",
            font_size=13,
            font_family="Arial",
            bordercolor="rgba(0, 0, 0, 0.3)",
            align="left"
        ),
        annotations=[
            dict(
                x=0.5,
                y=0.02,
                xref="paper",
                yref="paper",
                text=f"<i>Displaying {metric} data across U.S. states. Bubble size indicates unit sales volume.</i>",
                showarrow=False,
                font=dict(size=12, color="#666666", family="Arial")
            )
        ]
    )
    
    return fig

# Function to create zip code map visualization with premium styling
# Function to create zip code map visualization with premium styling
def create_zip_map(df):
    if df is None or df.empty:
        return go.Figure()
        
    # Group by postal code
    zip_summary = df.groupby('Postal Code').agg({
        'Shipped Revenue': 'sum',
        'Shipped Units': 'sum',
        'Shipped COGS': 'sum'
    }).reset_index()
    
    # Add profit calculations
    zip_summary['Profit'] = zip_summary['Shipped Revenue'] - zip_summary['Shipped COGS']
    zip_summary['Margin'] = (zip_summary['Profit'] / zip_summary['Shipped Revenue'] * 100).round(1)
    
    # Create a map visualization with premium styling
    fig = go.Figure()
    
    # Add premium USA base map
    fig.add_trace(go.Choropleth(
        locations=['Alabama', 'Alaska', 'Arizona', 'Arkansas', 'California', 'Colorado', 
                  'Connecticut', 'Delaware', 'Florida', 'Georgia', 'Hawaii', 'Idaho', 
                  'Illinois', 'Indiana', 'Iowa', 'Kansas', 'Kentucky', 'Louisiana', 
                  'Maine', 'Maryland', 'Massachusetts', 'Michigan', 'Minnesota', 'Mississippi', 
                  'Missouri', 'Montana', 'Nebraska', 'Nevada', 'New Hampshire', 'New Jersey', 
                  'New Mexico', 'New York', 'North Carolina', 'North Dakota', 'Ohio', 'Oklahoma', 
                  'Oregon', 'Pennsylvania', 'Rhode Island', 'South Carolina', 'South Dakota', 
                  'Tennessee', 'Texas', 'Utah', 'Vermont', 'Virginia', 'Washington', 'West Virginia', 
                  'Wisconsin', 'Wyoming'],
        z=[1] * 50,  # Just to show state outlines
        locationmode='USA-states',
        colorscale=[[0, 'rgba(240, 240, 240, 0.8)'], [1, 'rgba(240, 240, 240, 0.8)']],
        showscale=False,
        marker_line_color='rgba(255, 255, 255, 0.9)',
        marker_line_width=0.7,
        hoverinfo='skip'
    ))
    
    # Function to generate realistic lat/lon from zip code
    def zip_to_latlon(zip_code):
        # Create a numeric seed from the zip code (works with alphanumeric codes)
        seed = sum(ord(c) for c in str(zip_code)[:3])
        np.random.seed(seed)
        
        # Generate a position within the continental US (approximately)
        lat = np.random.uniform(25, 49)
        lon = np.random.uniform(-124, -66)
        
        # Add some clustering to make it more realistic
        # Use first character to determine region
        first_char = str(zip_code)[0]
        
        # Northeast
        if first_char in ['0', '1', '2', 'A', 'B', 'C', 'E', 'G', 'H', 'J']:
            lat = np.random.uniform(37, 47)
            lon = np.random.uniform(-80, -67)
        # South
        elif first_char in ['3', '4', 'K', 'L', 'M', 'N']:
            lat = np.random.uniform(25, 36)
            lon = np.random.uniform(-98, -75)
        # Midwest
        elif first_char in ['5', '6', 'P', 'R', 'S']:
            lat = np.random.uniform(36, 49)
            lon = np.random.uniform(-97, -80)
        # West
        else:
            lat = np.random.uniform(32, 49)
            lon = np.random.uniform(-124, -100)
            
        return lat, lon
    
    # Generate positions for each zip code
    lats = []
    lons = []
    for zip_code in zip_summary['Postal Code']:
        lat, lon = zip_to_latlon(zip_code)
        lats.append(lat)
        lons.append(lon)
    
    # Sort zip data by revenue for better visualization (higher revenue on top)
    zip_summary['Latitude'] = lats
    zip_summary['Longitude'] = lons
    zip_summary = zip_summary.sort_values('Shipped Revenue', ascending=True)
    
    # Create hover text with premium formatting
    hover_text = []
    for index, row in zip_summary.iterrows():
        hover_text.append(
            f"<b>Zip Code: {row['Postal Code']}</b><br>" +
            f"<b>Revenue:</b> ${row['Shipped Revenue']:,.2f}<br>" +
            f"<b>Units:</b> {row['Shipped Units']:,}<br>" +
            f"<b>Profit:</b> ${row['Profit']:,.2f}<br>" +
            f"<b>Margin:</b> {row['Margin']:.1f}%"
        )
    
    # Normalize the values for sizing and coloring
    max_revenue = zip_summary['Shipped Revenue'].max()
    max_units = zip_summary['Shipped Units'].max()
    
    # Create 1-mile radius circles with premium styling
    # Calculate sizes for 1-mile radius circles, scaled by revenue for visibility
    # Use better logarithmic scaling for visualization
    sizes = np.log1p(zip_summary['Shipped Revenue']) / np.log1p(max_revenue) * 18 + 7
    
    # Premium color scale
    custom_colorscale = [
        [0, "#e6f2ff"], [0.2, "#99ccff"], 
        [0.4, "#3399ff"], [0.6, "#0066cc"], 
        [0.8, "#004080"], [1, "#002147"]
    ]
    
    # Add the circles with premium styling
    fig.add_trace(go.Scattergeo(
        lon=zip_summary['Longitude'],
        lat=zip_summary['Latitude'],
        text=hover_text,
        mode='markers',
        marker=dict(
            size=sizes,
            color=zip_summary['Shipped Revenue'],
            colorscale=custom_colorscale,
            cmin=0,
            cmax=max_revenue,
            colorbar_title="Revenue ($)",
            colorbar=dict(
                title=dict(
                    text="Revenue ($)",
                    font=dict(size=14, family="Arial", color="#333333")
                ),
                tickformat="$,.0f",
                xanchor="left",
                len=0.7,
                thickness=15,
                outlinewidth=0,
                bgcolor='rgba(255,255,255,0.0)',
                tickfont=dict(size=12, family="Arial")
            ),
            opacity=0.85,
            line=dict(width=1, color='rgba(255,255,255,0.8)')
        ),
        name='Zip Code Sales (1-mile radius)',
        hoverinfo='text'
    ))
    
    # Premium map layout
    fig.update_layout(
        title=dict(
            text="Sales by Zip Code (1-mile radius visualization)",
            font=dict(size=26, family="Arial", color="#1E3A8A"),
            x=0.5,
            xanchor='center',
            y=0.97
        ),
        geo=dict(
            scope='usa',
            projection_type='albers usa',
            showland=True,
            landcolor='rgba(250, 250, 250, 0.95)',
            countrycolor='rgba(220, 220, 220, 0.8)',
            showlakes=True,
            lakecolor='rgba(211, 233, 250, 0.8)',
            showsubunits=True,
            subunitcolor='rgba(220, 220, 220, 0.8)',
            showcoastlines=True,
            coastlinecolor='rgba(220, 220, 220, 0.8)',
            showcountries=True,
            resolution=50,
            lonaxis=dict(range=[-125, -66]),
            lataxis=dict(range=[24, 50]),
            bgcolor='rgba(255, 255, 255, 0)'
        ),
        paper_bgcolor='rgba(255, 255, 255, 0)',
        plot_bgcolor='rgba(255, 255, 255, 0)',
        height=650,  # Slightly larger
        margin=dict(l=0, r=0, t=60, b=30),
        hoverlabel=dict(
            bgcolor="white",
            font_size=13,
            font_family="Arial",
            bordercolor="rgba(0, 0, 0, 0.3)",
            align="left"
        ),
        annotations=[
            dict(
                x=0.5,
                y=0.02,
                xref="paper",
                yref="paper",
                text="<i>Circle size and color intensity represent revenue scale. Each circle represents a 1-mile radius around the zip code center.</i>",
                showarrow=False,
                font=dict(size=12, color="#666666", family="Arial")
            )
        ]
    )
    
    return fig

# Function to create enhanced pivot tables with better styling
def create_enhanced_pivot(merged_data, pivot_type, value_type='Shipped Revenue'):
    if merged_data is None or merged_data.empty:
        return pd.DataFrame()
        
    # Determine index and columns based on pivot type
    if pivot_type == "State by Product Family":
        index = 'State'
        columns = 'Family'
    elif pivot_type == "Product Family by State":
        index = 'Family'
        columns = 'State'
    elif pivot_type == "Size by State":
        index = 'Size'
        columns = 'State'
    elif pivot_type == "Family by Month":
        index = 'Family'
        columns = 'Month_Name'
        # Ensure month names are in correct order
        month_order = ['January', 'February', 'March', 'April', 'May', 'June', 
                       'July', 'August', 'September', 'October', 'November', 'December']
        merged_data['Month_Name'] = pd.Categorical(merged_data['Month_Name'], categories=month_order, ordered=True)
    elif pivot_type == "SKU by State":
        index = 'ASIN'
        columns = 'State'
    else:  # Default
        index = 'State'
        columns = 'Family'
    
    # Create pivot table
    pivot = pd.pivot_table(
        merged_data,
        values=value_type,
        index=index,
        columns=columns,
        aggfunc='sum',
        fill_value=0
    )
    
    # Add totals
    pivot['Total'] = pivot.sum(axis=1)
    pivot.loc['Total'] = pivot.sum()
    
    return pivot

# Function to create metric cards with improved styling
def create_metrics(df, column='Shipped Revenue'):
    if df is None or df.empty:
        return 0, 0, 0, 0
        
    total = df[column].sum()
    avg = df[column].mean()
    max_val = df[column].max()
    
    # Add period-over-period comparison if date data available
    if 'Date' in df.columns and not df.empty:
        # Get current and previous period data based on time_period in session state
        if st.session_state.time_period == 'Year':
            current_year = st.session_state.selected_year
            prev_year = current_year - 1
            current_data = df[df['Year'] == current_year]
            prev_data = df[df['Year'] == prev_year]
        elif st.session_state.time_period == 'Quarter':
            current_year = st.session_state.selected_year
            current_quarter = st.session_state.selected_quarter
            
            # Calculate previous quarter and year
            if current_quarter == 1:
                prev_quarter = 4
                prev_year = current_year - 1
            else:
                prev_quarter = current_quarter - 1
                prev_year = current_year
                
            current_data = df[(df['Year'] == current_year) & (df['Quarter'] == current_quarter)]
            prev_data = df[(df['Year'] == prev_year) & (df['Quarter'] == prev_quarter)]
        else:  # Month
            current_year = st.session_state.selected_year
            current_month = st.session_state.selected_month
            
            # Calculate previous month and year
            if current_month == 1:
                prev_month = 12
                prev_year = current_year - 1
            else:
                prev_month = current_month - 1
                prev_year = current_year
                
            current_data = df[(df['Year'] == current_year) & (df['Month'] == current_month)]
            prev_data = df[(df['Year'] == prev_year) & (df['Month'] == prev_month)]
        
        # Calculate period over period change
        current_total = current_data[column].sum() if not current_data.empty else 0
        prev_total = prev_data[column].sum() if not prev_data.empty else 0
        
        if prev_total > 0:
            change_pct = ((current_total - prev_total) / prev_total) * 100
        else:
            change_pct = 100 if current_total > 0 else 0
            
        return total, avg, max_val, change_pct
    
    return total, avg, max_val, 0

# Sidebar for uploading files and filters
with st.sidebar:
    st.markdown("<h3 style='color:#1E3A8A; padding-bottom:10px;'>Sales Analytics Dashboard</h3>", unsafe_allow_html=True)
    
    # Add information about the dashboard
    with st.expander("ℹ️ Dashboard Information", expanded=False):
        st.markdown("""
        This enhanced sales analytics dashboard allows you to:
        
        - Visualize sales by state and zip code
        - Analyze performance by SKU Model & SKU Parent
        - Filter by time period (month/quarter/year)
        - See 1-mile radius visualization for zip code data
        - Create interactive pivot tables with multiple views
        - Download reports for further analysis
        
        **To get started, upload your data files below.**
        """)
    
    st.markdown("<h4 style='color:#4B5563; margin-top:15px;'>Upload Data Files</h4>", unsafe_allow_html=True)
    
    # File uploaders
    geo_file = st.file_uploader("Upload Geographic Reference Data", type=['xlsx'])
    sales_file = st.file_uploader("Upload Sales Data by State and Zip", type=['xlsx'])
    
    if geo_file and sales_file:
        # Load geographic data
        geo_data = load_excel_data(geo_file)
        
        # Load state and zip code data
        state_data = load_excel_data(sales_file, sheet_name='By State')
        zip_data = load_excel_data(sales_file, sheet_name='By Zip Code')
        
        if geo_data is not None and state_data is not None and zip_data is not None:
            # Clean data
            state_data = clean_state_names(state_data)
            
            # Store in session state
            st.session_state.geographic_data = geo_data
            st.session_state.state_data = state_data
            st.session_state.zip_data = zip_data
            st.session_state.last_upload_date = datetime.datetime.now()
            
            st.success("Data loaded successfully!")
    
    # Display data info
    if st.session_state.state_data is not None:
        st.write(f"State data: {len(st.session_state.state_data)} records")
        st.write(f"Zip code data: {len(st.session_state.zip_data)} records")
        if st.session_state.last_upload_date:
            st.write(f"Last updated: {st.session_state.last_upload_date.strftime('%Y-%m-%d %H:%M')}")
    
    # Filter options header
    st.markdown("<h4 style='color:#4B5563; margin-top:20px;'>Filter Options</h4>", unsafe_allow_html=True)
    
    # SKU filters (from geographic data)
    if st.session_state.geographic_data is not None:
        # Family filter
        all_families = ["All"] + list(st.session_state.geographic_data["Family"].unique())
        selected_family = st.selectbox("Select Family", all_families)
        
        # Flavor filter
        all_flavors = ["All"] + list(st.session_state.geographic_data["Flavor "].unique())
        selected_flavor = st.selectbox("Select Flavor", all_flavors)
        
        # Size filter
        all_sizes = ["All"] + list(st.session_state.geographic_data["Size"].unique())
        selected_size = st.selectbox("Select Size", all_sizes)
        
        # ASIN filter based on selections
        filtered_geo = st.session_state.geographic_data.copy()
        
        if selected_family != "All":
            filtered_geo = filtered_geo[filtered_geo["Family"] == selected_family]
        
        if selected_flavor != "All":
            filtered_geo = filtered_geo[filtered_geo["Flavor "] == selected_flavor]
        
        if selected_size != "All":
            filtered_geo = filtered_geo[filtered_geo["Size"] == selected_size]
        
        selected_asins = filtered_geo["ASIN"].tolist()
    else:
        selected_asins = []

# Main content
st.markdown("<h1 class='main-header'>📊 Sales Analytics Dashboard</h1>", unsafe_allow_html=True)

# Check if data is loaded
if (st.session_state.geographic_data is None or 
    st.session_state.state_data is None or 
    st.session_state.zip_data is None):
    
    st.markdown("<div style='text-align: center; padding: 50px 0;'>", unsafe_allow_html=True)
    st.markdown("<h2 style='color:#4B5563;'>Welcome to the Enhanced Sales Analytics Dashboard</h2>", unsafe_allow_html=True)
    st.markdown("<p style='font-size: 1.2rem; color:#6B7280; margin-bottom:30px;'>Please upload your data files using the sidebar to get started.</p>", unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        st.markdown("""
        <div style='background-color: #f8fafc; padding: 20px; border-radius: 10px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);'>
            <h3 style='color:#1E3A8A; margin-bottom:15px;'>Expected Data Files:</h3>
            <div style='margin-bottom:15px;'>
                <strong style='color:#1E40AF;'>1. Geographic Reference Data</strong>
                <p>Excel file with ASIN, Family, Flavor, and Size columns</p>
            </div>
            <div>
                <strong style='color:#1E40AF;'>2. Sales Data</strong>
                <p>Excel file with two sheets:
                <ul>
                    <li>'By State' sheet - with State, ASIN, Shipped Revenue, Shipped COGS, and Shipped Units</li>
                    <li>'By Zip Code' sheet - with Postal Code, ASIN, Shipped Revenue, Shipped COGS, and Shipped Units</li>
                </ul>
                </p>
            </div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)
    
else:
    # Filter data based on selections
    if selected_asins:
        filtered_state_data = st.session_state.state_data[st.session_state.state_data['ASIN'].isin(selected_asins)]
        filtered_zip_data = st.session_state.zip_data[st.session_state.zip_data['ASIN'].isin(selected_asins)]
    else:
        filtered_state_data = st.session_state.state_data.copy()
        filtered_zip_data = st.session_state.zip_data.copy()
    
    # Create tabs
    tab1, tab2, tab3, tab4 = st.tabs(["Summary", "State Analysis", "Zip Code Analysis", "Pivot Tables"])
    
    with tab1:
        st.markdown("<h2 class='section-header'>Sales Summary Dashboard</h2>", unsafe_allow_html=True)
        
        # Time period selector
        st.markdown("<div class='filters-container'>", unsafe_allow_html=True)
        period_col1, period_col2, period_col3 = st.columns([1,1,1])
        
        with period_col1:
            time_period = st.selectbox(
                "Select Time Period",
                ["Year", "Quarter", "Month"]
            )
            st.session_state.time_period = time_period
            
        with period_col2:
            selected_year = st.selectbox(
                "Select Year",
                [datetime.datetime.now().year, datetime.datetime.now().year-1]
            )
            st.session_state.selected_year = selected_year
            
        with period_col3:
            if time_period == "Quarter":
                selected_quarter = st.selectbox(
                    "Select Quarter",
                    [1, 2, 3, 4]
                )
                st.session_state.selected_quarter = selected_quarter
            elif time_period == "Month":
                selected_month = st.selectbox(
                    "Select Month",
                    list(range(1, 13)),
                    format_func=lambda x: datetime.date(2022, x, 1).strftime('%B')
                )
                st.session_state.selected_month = selected_month
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Create summary metrics
        state_total_revenue, state_avg_revenue, state_max_revenue, revenue_change = create_metrics(filtered_state_data, 'Shipped Revenue')
        state_total_units, state_avg_units, state_max_units, units_change = create_metrics(filtered_state_data, 'Shipped Units')
        profit = state_total_revenue - filtered_state_data['Shipped COGS'].sum()
        margin = (profit / state_total_revenue * 100) if state_total_revenue > 0 else 0
        revenue_per_unit = state_total_revenue / state_total_units if state_total_units > 0 else 0
        
        # Create enhanced metrics row with cards
        st.markdown("<div style='padding: 10px 0 30px 0;'>", unsafe_allow_html=True)
        metric_col1, metric_col2, metric_col3, metric_col4 = st.columns(4)
        
        with metric_col1:
            st.markdown(
                f"""
                <div class='metric-card'>
                    <div class='metric-label'>Total Revenue</div>
                    <div class='metric-value'>${state_total_revenue:,.0f}</div>
                    <div>vs prev period: {"+" if revenue_change >= 0 else ""}{revenue_change:.1f}%</div>
                </div>
                """, 
                unsafe_allow_html=True
            )
            
        with metric_col2:
            st.markdown(
                f"""
                <div class='metric-card'>
                    <div class='metric-label'>Total Units Sold</div>
                    <div class='metric-value'>{int(state_total_units):,}</div>
                    <div>vs prev period: {"+" if units_change >= 0 else ""}{units_change:.1f}%</div>
                </div>
                """, 
                unsafe_allow_html=True
            )
            
        with metric_col3:
            st.markdown(
                f"""
                <div class='metric-card'>
                    <div class='metric-label'>Profit Margin</div>
                    <div class='metric-value'>{margin:.1f}%</div>
                    <div>Total Profit: ${profit:,.0f}</div>
                </div>
                """, 
                unsafe_allow_html=True
            )
            
        with metric_col4:
            st.markdown(
                f"""
                <div class='metric-card'>
                    <div class='metric-label'>Revenue per Unit</div>
                    <div class='metric-value'>${revenue_per_unit:.2f}</div>
                    <div>Avg Revenue per State: ${state_avg_revenue:,.0f}</div>
                </div>
                """, 
                unsafe_allow_html=True
            )
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Create sales by family chart - only show bar chart as requested
        if st.session_state.geographic_data is not None:
            # Merge state data with geographic data
            merged_data = merge_data(st.session_state.geographic_data, filtered_state_data)
            
            if merged_data is not None:
                # Create sales by family chart
                family_sales = merged_data.groupby('Family')[['Shipped Revenue', 'Shipped Units']].sum().reset_index()
                
                # Create improved bar chart
                fig = create_bar_chart(
                    family_sales,
                    'Family',
                    'Shipped Revenue',
                    "Revenue by Product Family",
                    'Family'
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # Top states chart with improved styling
                top_states = filtered_state_data.groupby('State')[['Shipped Revenue', 'Shipped Units']].sum().reset_index()
                top_states = top_states.sort_values('Shipped Revenue', ascending=False).head(10)
                
                fig = create_bar_chart(
                    top_states,
                    'State',
                    'Shipped Revenue',
                    "Top 10 States by Revenue"
                )
                st.plotly_chart(fig, use_container_width=True)
                
                # Add SKU performance section
                st.markdown("<h3 class='section-header'>SKU Performance Analysis</h3>", unsafe_allow_html=True)
                
                sku_sales = merged_data.groupby(['ASIN', 'Family', 'Flavor ', 'Size'])[['Shipped Revenue', 'Shipped Units', 'Shipped COGS']].sum().reset_index()
                sku_sales['Profit'] = sku_sales['Shipped Revenue'] - sku_sales['Shipped COGS']
                sku_sales['Margin %'] = (sku_sales['Profit'] / sku_sales['Shipped Revenue'] * 100).round(2)
                sku_sales = sku_sales.sort_values('Shipped Revenue', ascending=False)
                
                st.dataframe(
                    sku_sales.head(20).style
                    .format({
                        'Shipped Revenue': '${:,.2f}',
                        'Shipped COGS': '${:,.2f}',
                        'Profit': '${:,.2f}',
                        'Margin %': '{:.2f}%'
                    })
                    .background_gradient(cmap='Blues', subset=['Shipped Revenue'])
                    .background_gradient(cmap='Greens', subset=['Profit'])
                    .background_gradient(cmap='RdYlGn', subset=['Margin %']),
                    use_container_width=True,
                    height=400
                )
        
    with tab2:
        st.markdown("<h2 class='section-header'>Sales by State</h2>", unsafe_allow_html=True)
        
        # Add metric selector for map
        metric_options = ["Shipped Revenue", "Profit", "Margin %", "Shipped Units"]
        selected_metric = st.selectbox("Select Map Metric", metric_options)
        
        # Create choropleth map with selected metric
        state_map = create_state_map(filtered_state_data, selected_metric)
        st.plotly_chart(state_map, use_container_width=True)
        
        # Add filters for SKU analysis
        st.markdown("<div class='filters-container'>", unsafe_allow_html=True)
        filter_col1, filter_col2, filter_col3 = st.columns([1,1,1])
        
        with filter_col1:
            # SKU Parent filter (Family)
            if st.session_state.geographic_data is not None:
                all_families = ["All"] + sorted(list(st.session_state.geographic_data["Family"].unique()))
                selected_family_state = st.selectbox("Select Product Family", all_families, key='state_family')
            else:
                selected_family_state = "All"
                
        with filter_col2:
            # Flavor filter
            if st.session_state.geographic_data is not None and selected_family_state != "All":
                filtered_flavors = st.session_state.geographic_data[st.session_state.geographic_data["Family"] == selected_family_state]["Flavor "].unique()
                all_flavors = ["All"] + sorted(list(filtered_flavors))
                selected_flavor_state = st.selectbox("Select Flavor", all_flavors, key='state_flavor')
            else:
                all_flavors = ["All"]
                if st.session_state.geographic_data is not None:
                    all_flavors += sorted(list(st.session_state.geographic_data["Flavor "].unique()))
                selected_flavor_state = st.selectbox("Select Flavor", all_flavors, key='state_flavor_all')
                
        with filter_col3:
            # Size filter
            if st.session_state.geographic_data is not None:
                all_sizes = ["All"] + sorted([str(x) for x in st.session_state.geographic_data["Size"].unique()])
                selected_size_state = st.selectbox("Select Size", all_sizes, key='state_size')
            else:
                selected_size_state = "All"
        st.markdown("</div>", unsafe_allow_html=True)
        
        # State analytics
        st.markdown("<h3 class='section-header'>State-by-State Performance Analysis</h3>", unsafe_allow_html=True)
        st.markdown("Comprehensive breakdown of sales performance across all states")
        
        # Apply SKU filters to state data if geographic data is available
        if (st.session_state.geographic_data is not None and 
            (selected_family_state != "All" or selected_flavor_state != "All" or selected_size_state != "All")):
            
            # Filter geographic data based on selections
            filtered_geo = st.session_state.geographic_data.copy()
            
            if selected_family_state != "All":
                filtered_geo = filtered_geo[filtered_geo["Family"] == selected_family_state]
            
            if selected_flavor_state != "All":
                filtered_geo = filtered_geo[filtered_geo["Flavor "] == selected_flavor_state]
            
            if selected_size_state != "All":
                filtered_geo = filtered_geo[filtered_geo["Size"] == selected_size_state]
            
            # Get filtered ASINs and apply to state data
            filtered_asins = filtered_geo["ASIN"].unique()
            filtered_state_data_sku = filtered_state_data[filtered_state_data['ASIN'].isin(filtered_asins)]
        else:
            filtered_state_data_sku = filtered_state_data
        
        # Group by state
        state_summary = filtered_state_data_sku.groupby('State').agg({
            'Shipped Revenue': 'sum',
            'Shipped COGS': 'sum',
            'Shipped Units': 'sum'
        }).reset_index()
        
        # Add profit and margin columns
        state_summary['Profit'] = state_summary['Shipped Revenue'] - state_summary['Shipped COGS']
        state_summary['Margin %'] = (state_summary['Profit'] / state_summary['Shipped Revenue'] * 100).round(2)
        
        # Sort by revenue
        state_summary = state_summary.sort_values('Shipped Revenue', ascending=False)
        
        # Set up the columns
        col1, col2 = st.columns([3, 1])
        
        with col1:
            # Enhanced table with better styling and metrics - using Blues instead of viridis
            st.dataframe(
                state_summary.style
                .format({
                    'Shipped Revenue': '${:,.2f}',
                    'Shipped COGS': '${:,.2f}',
                    'Profit': '${:,.2f}',
                    'Margin %': '{:.2f}%'
                })
                .background_gradient(cmap='Blues', subset=['Shipped Revenue'])
                .background_gradient(cmap='Greens', subset=['Profit'])
                .background_gradient(cmap='RdYlGn', subset=['Margin %'])
                .bar(subset=['Shipped Units'], color='#4b6cb7')
                .set_properties(**{'font-size': '12pt', 'text-align': 'center'})
                .set_caption("Complete State Performance Metrics")
                .highlight_max(subset=['Shipped Revenue', 'Profit', 'Margin %'], color='#dbeafe')
                .highlight_min(subset=['Shipped Revenue', 'Profit', 'Margin %'], color='#fee2e2'),
                use_container_width=True,
                height=500
            )
        
        with col2:
            # Add key insights about state performance with improved styling
            st.markdown("<h4 style='color:#1E3A8A;'>Key Insights</h4>", unsafe_allow_html=True)
            
            # Calculate insights
            if not state_summary.empty:
                top_state = state_summary.iloc[0]['State']
                top_revenue = state_summary.iloc[0]['Shipped Revenue']
                top_margin_index = state_summary['Margin %'].idxmax() if len(state_summary) > 0 else 0
                top_margin_state = state_summary.loc[top_margin_index]['State'] if len(state_summary) > 0 else "N/A"
                top_margin = state_summary['Margin %'].max() if len(state_summary) > 0 else 0
                avg_margin = state_summary['Margin %'].mean() if len(state_summary) > 0 else 0
                
                # Display insights with improved styling
                st.markdown(
                    f"""
                    <div class='metric-card' style='background-color: #dbeafe; margin-bottom: 15px;'>
                        <div style='font-weight: 600; color: #1E3A8A;'>Top Performing State</div>
                        <div style='font-size: 1.5rem; font-weight: 700; color: #1E40AF;'>{top_state}</div>
                        <div style='font-size: 1.2rem; color: #1E40AF;'>${top_revenue:,.0f}</div>
                        <div style='font-size: 0.9rem;'>in total revenue</div>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )
                
                st.markdown(
                    f"""
                    <div class='metric-card' style='background-color: #dcfce7; margin-bottom: 15px;'>
                        <div style='font-weight: 600; color: #166534;'>Highest Margin State</div>
                        <div style='font-size: 1.5rem; font-weight: 700; color: #166534;'>{top_margin_state}</div>
                        <div style='font-size: 1.2rem; color: #166534;'>{top_margin:.2f}%</div>
                        <div style='font-size: 0.9rem;'>profit margin</div>
                    </div>
                    """, 
                    unsafe_allow_html=True
                )
                
                # Add geographical insight
                regions = {
                    'Northeast': ['Connecticut', 'Maine', 'Massachusetts', 'New Hampshire', 'Rhode Island', 
                                'Vermont', 'New Jersey', 'New York', 'Pennsylvania'],
                    'Midwest': ['Illinois', 'Indiana', 'Michigan', 'Ohio', 'Wisconsin', 'Iowa', 'Kansas', 
                               'Minnesota', 'Missouri', 'Nebraska', 'North Dakota', 'South Dakota'],
                    'South': ['Delaware', 'Florida', 'Georgia', 'Maryland', 'North Carolina', 'South Carolina', 
                             'Virginia', 'West Virginia', 'Alabama', 'Kentucky', 'Mississippi', 'Tennessee', 
                             'Arkansas', 'Louisiana', 'Oklahoma', 'Texas'],
                    'West': ['Arizona', 'Colorado', 'Idaho', 'Montana', 'Nevada', 'New Mexico', 'Utah', 
                     'Wyoming', 'Alaska', 'California', 'Hawaii', 'Oregon', 'Washington']
                }
                
                # Create region summaries
                region_data = []
                for region, states in regions.items():
                    region_states = state_summary[state_summary['State'].isin(states)]
                    if not region_states.empty:
                        region_revenue = region_states['Shipped Revenue'].sum()
                        region_data.append({
                            'Region': region,
                            'Revenue': region_revenue,
                            'Percentage': region_revenue / state_summary['Shipped Revenue'].sum() * 100
                        })
                
                region_df = pd.DataFrame(region_data)
                if not region_df.empty:
                    top_region = region_df.loc[region_df['Revenue'].idxmax()]['Region']
                    top_region_pct = region_df.loc[region_df['Revenue'].idxmax()]['Percentage']
                    
                    st.markdown(
                        f"""
                        <div class='metric-card' style='background-color: #f1f5f9;'>
                            <div style='font-weight: 600; color: #334155;'>Top Performing Region</div>
                            <div style='font-size: 1.5rem; font-weight: 700; color: #334155;'>{top_region}</div>
                            <div style='font-size: 1.2rem; color: #334155;'>{top_region_pct:.1f}%</div>
                            <div style='font-size: 0.9rem;'>of total revenue</div>
                        </div>
                        """, 
                        unsafe_allow_html=True
                    )
        
        # Add download button for state data with improved styling
        state_excel = download_excel(state_summary, 'State_Summary')
        st.download_button(
            label="📥 Download State Summary",
            data=state_excel,
            file_name=f"state_sales_summary_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with tab3:
        st.markdown("<h2 class='section-header'>Sales by Zip Code</h2>", unsafe_allow_html=True)
        
        # Add filters for zip code analysis
        st.markdown("<div class='filters-container'>", unsafe_allow_html=True)
        zip_filter_col1, zip_filter_col2, zip_filter_col3 = st.columns([1,1,1])
        
        with zip_filter_col1:
            # SKU Parent filter (Family)
            if st.session_state.geographic_data is not None:
                all_families = ["All"] + sorted(list(st.session_state.geographic_data["Family"].unique()))
                selected_family_zip = st.selectbox("Select Product Family", all_families, key='zip_family')
            else:
                selected_family_zip = "All"
                
        with zip_filter_col2:
            # Add revenue threshold filter
            if not filtered_zip_data.empty:
                min_rev = int(filtered_zip_data['Shipped Revenue'].min())
                max_rev = int(filtered_zip_data['Shipped Revenue'].max())
                rev_step = max(1, int((max_rev - min_rev) / 100))
                revenue_threshold = st.slider(
                    "Minimum Revenue ($)", 
                    min_value=min_rev,
                    max_value=max_rev,
                    value=min_rev,
                    step=rev_step
                )
            else:
                revenue_threshold = 0
                
        with zip_filter_col3:
            # Add option to show top N zip codes
            if not filtered_zip_data.empty:
                top_n = st.number_input("Show Top N Zip Codes", min_value=10, max_value=1000, value=100, step=10)
            else:
                top_n = 100
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Create zip code visualization with 1-mile radius
        st.markdown("<h3 class='section-header'>1-Mile Radius Zip Code Visualization</h3>", unsafe_allow_html=True)
        st.markdown("""
        This map shows sales by zip code, with each circle representing a 1-mile radius around the zip code center.
        The size and color of circles represent the revenue volume.
        """)
        
        # Apply filters to zip data
        filtered_zip_for_map = filtered_zip_data.copy()
        
        # Filter by revenue threshold
        if revenue_threshold > 0:
            filtered_zip_for_map = filtered_zip_for_map[filtered_zip_for_map['Shipped Revenue'] >= revenue_threshold]
        
        # Filter by family if selected
        if selected_family_zip != "All" and st.session_state.geographic_data is not None:
            # Get ASINs for selected family
            family_asins = st.session_state.geographic_data[
                st.session_state.geographic_data["Family"] == selected_family_zip
            ]["ASIN"].unique()
            
            filtered_zip_for_map = filtered_zip_for_map[filtered_zip_for_map['ASIN'].isin(family_asins)]
        
        # Create enhanced zip code map
        zip_map = create_zip_map(filtered_zip_for_map)
        st.plotly_chart(zip_map, use_container_width=True)
        
        # Zip code analytics with improved UI
        st.markdown("<h3 class='section-header'>Zip Code Performance Analytics</h3>", unsafe_allow_html=True)
        
        # Group by zip code
        zip_summary = filtered_zip_data.groupby('Postal Code').agg({
            'Shipped Revenue': 'sum',
            'Shipped COGS': 'sum',
            'Shipped Units': 'sum'
        }).reset_index()
        
        # Add profit and margin columns
        zip_summary['Profit'] = zip_summary['Shipped Revenue'] - zip_summary['Shipped COGS']
        zip_summary['Margin %'] = (zip_summary['Profit'] / zip_summary['Shipped Revenue'] * 100).round(2)
        
        # Apply revenue threshold filter
        if revenue_threshold > 0:
            zip_summary = zip_summary[zip_summary['Shipped Revenue'] >= revenue_threshold]
        
        # Sort by revenue
        zip_summary = zip_summary.sort_values('Shipped Revenue', ascending=False)
        
        # Setup columns for summary stats and table
        zip_col1, zip_col2 = st.columns([1, 3])
        
        with zip_col1:
            # Show summary statistics
            total_zip_count = len(zip_summary)
            total_zip_revenue = zip_summary['Shipped Revenue'].sum()
            avg_zip_revenue = zip_summary['Shipped Revenue'].mean()
            med_zip_revenue = zip_summary['Shipped Revenue'].median()
            avg_zip_margin = zip_summary['Margin %'].mean()
    
            st.markdown(
                f"""
                <div class='metric-card' style='margin-bottom: 15px;'>
                    <div class='metric-label'>Total Zip Codes</div>
                    <div class='metric-value'>{total_zip_count:,}</div>
                </div>
                
                <div class='metric-card' style='margin-bottom: 15px;'>
                    <div class='metric-label'>Total Revenue</div>
                    <div class='metric-value'>${total_zip_revenue:,.0f}</div>
                </div>
                
                <div class='metric-card' style='margin-bottom: 15px;'>
                    <div class='metric-label'>Average Revenue per Zip</div>
                    <div class='metric-value'>${avg_zip_revenue:,.0f}</div>
                </div>
                
                <div class='metric-card'>
                    <div class='metric-label'>Average Margin</div>
                    <div class='metric-value'>{avg_zip_margin:.2f}%</div>
                </div>
                """,
                unsafe_allow_html=True
            )

        with zip_col2:
            # Show top N zip codes with enhanced styling
            st.dataframe(
                zip_summary.head(top_n).style
                .format({
                    'Shipped Revenue': '${:,.2f}',
                    'Shipped COGS': '${:,.2f}',
                    'Profit': '${:,.2f}',
                    'Margin %': '{:.2f}%'
                })
                .background_gradient(cmap='Blues', subset=['Shipped Revenue'])
                .background_gradient(cmap='Greens', subset=['Profit'])
                .background_gradient(cmap='RdYlGn', subset=['Margin %'])
                .bar(subset=['Shipped Units'], color='#4b6cb7')
                .set_properties(**{'font-size': '11pt', 'text-align': 'center'})
                .set_caption(f"Top {top_n} Zip Codes by Revenue"),
                use_container_width=True,
                height=400
            )

        # Add download button for zip data with improved styling
        zip_excel = download_excel(zip_summary, 'Zip_Summary')
        st.download_button(
            label="📥 Download Zip Code Summary",
            data=zip_excel,
            file_name=f"zip_sales_summary_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        with tab4:
            st.markdown("<h2 class='section-header'>Interactive Pivot Tables</h2>", unsafe_allow_html=True)
            
            # Enhanced pivot table options with more dimensions
            st.markdown("<div class='filters-container'>", unsafe_allow_html=True)
            pivot_col1, pivot_col2, pivot_col3 = st.columns([1,1,1])
            
            with pivot_col1:
                # Select pivot type with more options
                pivot_type = st.selectbox(
                    "Select Pivot Table View",
                    ["State by Product Family", 
                        "Product Family by State", 
                        "Size by State", 
                        "Family by Month",
                        "SKU by State"]
                )
            
            with pivot_col2:
                # Select value type
                value_type = st.selectbox(
                    "Select Value to Analyze",
                    ["Shipped Revenue", "Shipped Units", "Profit", "Margin %"]
                )
                
            with pivot_col3:
                # Add color scheme selector
                color_scheme = st.selectbox(
                    "Select Color Scheme",
                    ["Blues", "Greens", "Purples", "Oranges", "RdYlGn", "viridis"]
                )
            st.markdown("</div>", unsafe_allow_html=True)
        
            # Create pivot table based on merged data
            if st.session_state.geographic_data is not None:
                # Merge state data with geographic data
                merged_data = merge_data(st.session_state.geographic_data, filtered_state_data)
                
                if merged_data is not None:
                    # Add calculated columns for pivot tables
                    if value_type == "Profit":
                        merged_data['Profit'] = merged_data['Shipped Revenue'] - merged_data['Shipped COGS']
                    
                    if value_type == "Margin %":
                        merged_data['Profit'] = merged_data['Shipped Revenue'] - merged_data['Shipped COGS']
                        merged_data['Margin %'] = (merged_data['Profit'] / merged_data['Shipped Revenue'] * 100).round(2)
                    
                    # Create enhanced pivot tables
                    pivot = create_enhanced_pivot(merged_data, pivot_type, value_type)
                    
                    if not pivot.empty:
                        # Display enhanced pivot table with better styling
                        st.markdown("<h3 class='section-header'>Interactive Sales Pivot Table</h3>", unsafe_allow_html=True)
                        st.markdown("Analyze sales patterns across multiple dimensions with enhanced visualization")
                        
                        # Format based on value type
                        if value_type in ["Shipped Revenue", "Profit"]:
                            format_str = '${:,.2f}'
                        elif value_type == "Margin %":
                            format_str = '{:.2f}%'
                        else:  # Units
                            format_str = '{:,.0f}'
                        
                        # Apply enhanced styling with selected color scheme
                        styled_pivot = pivot.style.format(format_str)\
                            .background_gradient(cmap=color_scheme, axis=None)\
                            .highlight_max(axis=0, color='#dbeafe')\
                            .highlight_min(axis=0, color='#fee2e2')\
                            .set_properties(**{'font-size': '11pt', 'text-align': 'center'})\
                            .set_caption(f"Sales Analysis: {pivot_type} - {value_type}")
                        
                        # Display the styled pivot table with increased height
                        st.dataframe(styled_pivot, use_container_width=True, height=500)
                        
                        # Add chart visualization of pivot data
                        st.markdown("<h3 class='section-header'>Pivot Data Visualization</h3>", unsafe_allow_html=True)
                        
                        # Create visualization based on pivot type
                        pivot_for_chart = pivot.reset_index().melt(id_vars=pivot.index.name)
                        
                        if pivot_type in ["State by Product Family", "SKU by State"]:
                            # For these types, show a grouped bar chart of top 10 states/SKUs
                            top_indices = pivot['Total'].nlargest(10).index
                            pivot_top10 = pivot.loc[top_indices].drop(columns=['Total'])
                            pivot_chart_data = pivot_top10.reset_index().melt(id_vars=pivot.index.name)
                            
                            # First, check what column names you actually have
                            pivot_for_chart = pivot.reset_index().melt(id_vars=pivot.index.name)
                            print(pivot_for_chart.columns)  # You can use st.write() instead for Streamlit

                            # Then use the correct column name for the color parameter
                            fig = px.bar(
                                pivot_chart_data,
                                x=pivot.index.name,
                                y='value',
                                color=pivot_chart_data.columns[1],  # Use the actual column name from melting
                                title=f"Top 10 {pivot.index.name}s by {value_type}",
                                labels={
                                    'value': value_type,
                                    pivot_chart_data.columns[1]: 'Category'  # Update this too
                                },
                                height=500
                            )
                            
                            fig.update_layout(
                                xaxis_title=pivot.index.name,
                                yaxis_title=value_type,
                                legend_title="Category",
                                font=dict(family="Arial", size=12),
                                plot_bgcolor='rgba(240,242,246,0.3)'
                            )
                            
                        else:
                            # For other types, show a heatmap
                            fig = px.imshow(
                                pivot.drop(columns=['Total']).drop('Total') if 'Total' in pivot.columns else pivot,
                                text_auto='.2s',
                                aspect="auto",
                                color_continuous_scale=color_scheme,
                                labels=dict(x="Category", y="Product", color=value_type),
                                height=500
                            )
                            
                            fig.update_layout(
                                title=f"Heatmap: {pivot_type} - {value_type}",
                                font=dict(family="Arial", size=12)
                            )
                        
                        st.plotly_chart(fig, use_container_width=True)
                        
                        # Download pivot table with enhanced button
                        pivot_excel = BytesIO()
                        with pd.ExcelWriter(pivot_excel, engine='xlsxwriter') as writer:
                            pivot.to_excel(writer, sheet_name='Pivot')
                        
                        st.download_button(
                            label="📥 Download Pivot Table",
                            data=pivot_excel.getvalue(),
                            file_name=f"sales_pivot_{pivot_type.replace(' ', '_')}_{value_type.replace(' ', '_')}_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.warning("No data available for the selected pivot configuration. Try changing your filters.")
                else:
                    st.warning("Unable to create pivot table. Please ensure both geographic and sales data are loaded.")
            else:
                st.info("Please upload geographic reference data to use pivot tables.")

        # Footer
        st.markdown("---")
        st.markdown("""
        <div style='display: flex; justify-content: space-between; align-items: center; padding: 10px 0;'>
        <div style='color: #6B7280; font-size: 0.9rem;'>Sales Analytics Dashboard | Created with Streamlit</div>
        <div style='color: #6B7280; font-size: 0.9rem;'>Last updated: {}</div>
        </div>
        """.format(datetime.datetime.now().strftime("%Y-%m-%d %H:%M")), unsafe_allow_html=True)
