from dotenv import load_dotenv
import os
import streamlit as st
import pandas as pd
import tempfile
import shutil
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from io import BytesIO
from streamlit_option_menu import option_menu
import json
from datetime import datetime
import time
import base64
import numpy as np
import requests
from pathlib import Path
import httpx
import asyncio

# Load environment variables first
load_dotenv()

# FastAPI backend URL
BACKEND_URL = os.getenv("BACKEND_URL", "http://localhost:8000")

st.set_page_config(page_title="Excel → PPT Generator", page_icon="📊", layout="wide")

# Initialize session state with proper defaults
def initialize_session_state():
    """Initialize all session state variables with proper defaults"""
    defaults = {
        "theme": "dark",
        "uploaded_csv_file": None,
        "uploaded_template_file": None,
        "tmp_dir": None,
        "csv_path": None,
        "preview": None,
        "df": None,
        "template_path": None,
        "chosen_columns": [],
        "slides_count": 3,
        "slide_layout": "Title and Content",
        "chart_types": {},
        "custom_texts": {},
        "presentation_title": "Data Analysis Presentation",
        "company_logo": None,
        "generation_history": [],
        "ai_suggestions": {},
        "custom_styles": {
            "primary_color": "#f26f21",
            "secondary_color": "#ffa800",
            "chart_style": "plotly_white",
            "template_style": "professional",
            "text_color": "#000000"  # Default to black for sample templates
        },
        "logo_path": None,
        "analysis_complete": False,
        "sample_templates": [
            {"name": "Professional", "value": "professional", "description": "Clean and corporate style", "icon": "💼", "text_color": "#000000"},
            {"name": "Creative", "value": "creative", "description": "Modern and colorful design", "icon": "🎨", "text_color": "#000000"},
            {"name": "Minimal", "value": "minimal", "description": "Simple and elegant layout", "icon": "⚪", "text_color": "#000000"},
            {"name": "Technical", "value": "technical", "description": "Data-focused with charts emphasis", "icon": "📈", "text_color": "#000000"}
        ],
        "current_template": "none",
        "outlier_data": None,
        "correlation_data": None,
        "use_ai_content": True,
        "template_content": None,
        "detailed_analyses": {},
        "chart_htmls": {},
        "dashboard_config": {"enabled": False, "charts": []},
        "insight_config": {"enabled": False, "charts": [], "content": {}, "selected_columns": []},
        "comparison_config": {"enabled": False, "comparisons": []},
        "custom_query": {"enabled": False, "queries": [], "download_as_excel": False},
        "slide_content_preview": {},
        "selected_features": {
            "dashboard": False,
            "charts_insights": False,
            "comparison": False,
            "query": False,
            "insights": False
        },
        "backend_file_id": None,
        "processed_data": None,
        "template_file_id": None,
        "logo_file_id": None,
        "upload_status": {},
        "presentation_config": {},
        "config_saved": False,
        "show_chart_preview": False,
        "edit_mode": False,
        "feature_selection_confirmed": False,
        "charts_preview_ready": False,
        "presentation_generated": False,
        "generated_ppt_id": None,
        "slide_preview_data": [],
        "current_slide": 0,
        "slide_images": [],
        "slide_images_loaded": False,
        "slide_edit_mode": False,
        "current_editing_slide": None,
        "slide_content_edit": {},
        "column_types": {},
        "debug_info": {},
        "generation_errors": [],
        "backend_logs": [],
        "show_success_popup": False,
        "template_text_color": "#000000"  # NEW: For template-based text color
    }
   
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

# Initialize session state
initialize_session_state()

# Enhanced error handling decorator
def handle_backend_errors(func):
    """Decorator to handle backend errors and provide informative messages"""
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except httpx.ConnectError as e:
            error_msg = f"🔌 Connection Error: Cannot connect to backend server at {BACKEND_URL}"
            st.error(error_msg)
            st.error(f"💡 Debug Info: Please ensure the backend server is running and accessible")
            st.session_state.debug_info["last_error"] = f"ConnectionError: {str(e)}"
            st.session_state.generation_errors.append(error_msg)
            return None
        except httpx.TimeoutException as e:
            error_msg = f"⏰ Timeout Error: Backend request took too long"
            st.error(error_msg)
            st.error(f"💡 Debug Info: The server might be overloaded or the file is too large")
            st.session_state.debug_info["last_error"] = f"TimeoutError: {str(e)}"
            st.session_state.generation_errors.append(error_msg)
            return None
        except httpx.HTTPStatusError as e:
            error_msg = f"🚨 HTTP Error {e.response.status_code}: {e.response.text}"
            st.error(error_msg)
            st.error(f"💡 Debug Info: Backend returned an error status")
            st.session_state.debug_info["last_error"] = f"HTTPError {e.response.status_code}: {e.response.text}"
            st.session_state.generation_errors.append(error_msg)
            return None
        except Exception as e:
            error_msg = f"❌ Unexpected Error: {str(e)}"
            st.error(error_msg)
            st.error(f"💡 Debug Info: Check backend logs for more details")
            st.session_state.debug_info["last_error"] = f"UnexpectedError: {str(e)}"
            st.session_state.generation_errors.append(error_msg)
            return None
    return wrapper

# Backend API functions with enhanced error handling
@handle_backend_errors
async def upload_file_to_backend_async(file, file_type: str = "csv") -> str:
    """Upload file to backend and return file ID"""
    try:
        if hasattr(file, 'getvalue'):
            file_bytes = file.getvalue()
            filename = file.name
        else:
            with open(file, 'rb') as f:
                file_bytes = f.read()
            filename = os.path.basename(file)
       
        files = {"file": (filename, file_bytes, "application/octet-stream")}
        data = {"file_type": file_type}
       
        async with httpx.AsyncClient(timeout=600.0) as client:
            response = await client.post(
                f"{BACKEND_URL}/upload/",
                files=files,
                data=data
            )
            response.raise_for_status()
            result = response.json()
            st.session_state.upload_status[file_type] = "success"
            return result["file_id"]
    except Exception as e:
        st.error(f"📤 Upload Error: Failed to upload {file_type} file")
        st.session_state.upload_status[file_type] = "failed"
        raise e

def upload_file_to_backend(file, file_type: str = "csv") -> str:
    return asyncio.run(upload_file_to_backend_async(file, file_type))

@handle_backend_errors
async def analyze_data_backend_async(file_id: str):
    """Analyze data using backend"""
    try:
        async with httpx.AsyncClient(timeout=300.0) as client:
            response = await client.get(f"{BACKEND_URL}/analyze/{file_id}")
            response.raise_for_status()
            return response.json()
    except Exception as e:
        st.error(f"🔍 Analysis Error: Failed to analyze data")
        raise e

def analyze_data_backend(file_id: str):
    return asyncio.run(analyze_data_backend_async(file_id))

@handle_backend_errors
async def generate_ppt_backend_async(config: dict):
    """Generate PPT using backend with detailed error reporting"""
    try:
        async with httpx.AsyncClient(timeout=1800.0) as client:
            response = await client.post(
                f"{BACKEND_URL}/generate-ppt/",
                json=config,
                timeout=1200.0
            )
            response.raise_for_status()
            result = response.json()
           
            if result.get("warnings"):
                st.session_state.generation_errors.extend(result["warnings"])
            if result.get("errors"):
                st.session_state.generation_errors.extend(result["errors"])
               
            return result
    except Exception as e:
        st.error(f"🎯 Generation Error: Failed to generate presentation")
        raise e

def generate_ppt_backend(config: dict):
    return asyncio.run(generate_ppt_backend_async(config))

@handle_backend_errors
async def download_ppt_backend_async(ppt_id: str):
    """Download generated PPT from backend"""
    try:
        async with httpx.AsyncClient(timeout=120.0) as client:
            response = await client.get(f"{BACKEND_URL}/download/{ppt_id}")
            response.raise_for_status()
            return response.content
    except Exception as e:
        st.error(f"📥 Download Error: Failed to download presentation")
        raise e

def download_ppt_backend(ppt_id: str):
    return asyncio.run(download_ppt_backend_async(ppt_id))

@handle_backend_errors
async def save_config_backend_async(config: dict, csv_file_id: str):
    """Save configuration JSON to backend"""
    try:
        async with httpx.AsyncClient(timeout=60.0) as client:
            response = await client.post(
                f"{BACKEND_URL}/save-config/{csv_file_id}",
                json=config
            )
            response.raise_for_status()
            return response.json()
    except Exception as e:
        st.error(f"💾 Save Error: Failed to save configuration")
        raise e

def save_config_backend(config: dict, csv_file_id: str):
    return asyncio.run(save_config_backend_async(config, csv_file_id))

@handle_backend_errors
async def get_slide_images_backend_async(ppt_id: str):
    """Get actual slide images from backend - ENHANCED with better error handling"""
    try:
        async with httpx.AsyncClient(timeout=120.0) as client:
            response = await client.get(f"{BACKEND_URL}/slide-images/{ppt_id}")
            response.raise_for_status()
            result = response.json()
           
            if result.get("success") and "slide_images" in result:
                slide_images = result["slide_images"]
                valid_slides = []
               
                for i, slide in enumerate(slide_images):
                    if slide.get('image_data') or slide.get('content'):
                        valid_slides.append(slide)
                    else:
                        st.warning(f"⚠️ Slide {i+1} has no image data or content")
                       
                if len(valid_slides) < len(slide_images):
                    st.warning(f"⚠️ Only {len(valid_slides)} out of {len(slide_images)} slides have valid content")
               
                result["slide_images"] = valid_slides
               
            return result
    except Exception as e:
        st.error(f"🖼️ Slide Image Error: Failed to load slide images")
        raise e

def get_slide_images_backend(ppt_id: str):
    return asyncio.run(get_slide_images_backend_async(ppt_id))

def analyze_column_types(df):
    """Enhanced column type analysis for feature engineering"""
    column_types = {}
   
    for col in df.columns:
        col_data = df[col]
       
        # Check if numeric
        if pd.api.types.is_numeric_dtype(col_data):
            column_types[col] = {
                'type': 'numerical',
                'subtype': 'continuous' if col_data.nunique() > 10 else 'discrete',
                'chart_types': ['Bar', 'Line', 'Histogram', 'Scatter', 'Box', 'Area']
            }
       
        # Check if datetime
        elif pd.api.types.is_datetime64_any_dtype(col_data):
            column_types[col] = {
                'type': 'datetime',
                'chart_types': ['Line', 'Area', 'Scatter']
            }
       
        # Check if categorical (limited unique values)
        elif col_data.nunique() <= 20 and col_data.nunique() > 1:
            column_types[col] = {
                'type': 'categorical',
                'subtype': 'ordinal' if col_data.dtype.name == 'category' else 'nominal',
                'chart_types': ['Bar', 'Pie', 'Box']
            }
       
        # Check if text (many unique values)
        elif col_data.nunique() > 20 and pd.api.types.is_string_dtype(col_data):
            column_types[col] = {
                'type': 'text',
                'chart_types': ['Bar']
            }
       
        # Default fallback
        else:
            column_types[col] = {
                'type': 'unknown',
                'chart_types': ['Bar', 'Histogram']
            }
   
    return column_types

def get_valid_chart_types(column_name, second_column=None):
    """Get valid chart types based on column data types"""
    if column_name not in st.session_state.column_types:
        return ['Bar', 'Line', 'Histogram', 'Scatter', 'Box', 'Area', 'Pie']
   
    col_info = st.session_state.column_types[column_name]
   
    # If second column is provided, adjust chart types based on combination
    if second_column and second_column != "None":
        if second_column in st.session_state.column_types:
            second_col_info = st.session_state.column_types[second_column]
           
            # Numerical vs Numerical
            if col_info['type'] == 'numerical' and second_col_info['type'] == 'numerical':
                return ['Scatter', 'Line', 'Area']
           
            # Categorical vs Numerical
            elif col_info['type'] == 'categorical' and second_col_info['type'] == 'numerical':
                return ['Bar', 'Box', 'Line']
           
            # Numerical vs Categorical
            elif col_info['type'] == 'numerical' and second_col_info['type'] == 'categorical':
                return ['Bar', 'Box', 'Line']
   
    # Single column chart types
    if col_info['type'] == 'numerical':
        return ['Histogram', 'Box', 'Bar', 'Line', 'Area']
    elif col_info['type'] == 'categorical':
        return ['Bar', 'Pie']
    elif col_info['type'] == 'datetime':
        return ['Line', 'Area', 'Bar']
    elif col_info['type'] == 'text':
        return ['Bar']
    else:
        return ['Bar', 'Histogram']

def create_advanced_plotly_chart(df, column, chart_type, style="plotly_white", second_column=None):
    """Create ADVANCED Plotly chart with maximum animations and interactivity"""
    try:
        # Define MULTI-COLOR schemes for different chart types
        color_scales = {
            'Bar': ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F'],
            'Line': ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F'],
            'Histogram': ['#6A89CC', '#4A69BD', '#1E3799', '#0C2461', '#82CCDD', '#60A3BC', '#3C6382', '#0A3D62'],
            'Scatter': ['#E55039', '#EB2F06', '#B71540', '#F6B93B', '#E58E26', '#78E08F', '#38ADA9', '#079992'],
            'Box': ['#78E08F', '#38ADA9', '#079992', '#82CCDD', '#60A3BC', '#3C6382', '#0A3D62', '#6A89CC'],
            'Area': ['#D6A2E8', '#BDC581', '#F8EFBA', '#FD7272', '#9AECDB', '#81ECEC', '#78E08F', '#38ADA9'],
            'Pie': ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4', '#FFEAA7', '#DDA0DD', '#98D8C8', '#F7DC6F']
        }
       
        colors = color_scales.get(chart_type, ['#FF6B6B', '#4ECDC4', '#45B7D1', '#96CEB4'])
       
        if chart_type == "Bar":
            if df[column].dtype == 'object':
                value_counts = df[column].value_counts().head(10)  # Limit to top 10 for better visualization
                fig = px.bar(
                    x=value_counts.index,
                    y=value_counts.values,
                    title=f"<b>📊 {column}</b> - Bar Chart",
                    color=value_counts.values,
                    color_continuous_scale=colors,
                    template="plotly_white"
                )
                fig.update_traces(
                    marker_line_color='black',
                    marker_line_width=1.5,
                    opacity=0.8,
                    hovertemplate="<b>%{x}</b><br>Count: %{y}<extra></extra>"
                )
            else:
                fig = px.histogram(
                    df, x=column,
                    title=f"<b>📊 {column}</b> - Distribution",
                    color_discrete_sequence=[colors[0]],
                    template="plotly_white",
                    nbins=20
                )
                fig.update_traces(
                    marker_line_color='black',
                    marker_line_width=1.5,
                    opacity=0.7
                )
               
        elif chart_type == "Line":
            if second_column and second_column != "None":
                # Create sample data for line chart if needed
                sample_df = df.head(50)  # Limit to 50 points for better performance
                fig = px.line(
                    sample_df, x=column, y=second_column,
                    title=f"<b>📈 {column} vs {second_column}</b> - Line Chart",
                    color_discrete_sequence=[colors[0]],
                    template="plotly_white"
                )
                fig.update_traces(
                    line=dict(width=4, shape='spline'),
                    marker=dict(size=8, symbol='circle'),
                    hovertemplate="<b>%{x}</b><br>Value: %{y}<extra></extra>"
                )
            else:
                # Create trend line with sample data
                sample_df = df.head(100)
                fig = px.line(
                    sample_df, x=sample_df.index, y=column,
                    title=f"<b>📈 {column}</b> - Trend Analysis",
                    color_discrete_sequence=[colors[0]],
                    template="plotly_white"
                )
                fig.update_traces(
                    line=dict(width=4, shape='spline'),
                    marker=dict(size=6),
                    fill='tozeroy',
                    fillcolor=f'rgba{tuple(int(colors[0].lstrip("#")[i:i+2], 16) for i in (0, 2, 4)) + (0.2,)}'
                )
               
        elif chart_type == "Histogram":
            fig = px.histogram(
                df, x=column,
                title=f"<b>📊 {column}</b> - Distribution Analysis",
                color_discrete_sequence=colors,
                template="plotly_white",
                nbins=30,
                marginal="box"
            )
            fig.update_traces(
                marker_line_color='black',
                marker_line_width=1.5,
                opacity=0.7
            )
           
        elif chart_type == "Scatter":
            if second_column and second_column != "None":
                sample_df = df.head(100)  # Limit for performance
                fig = px.scatter(
                    sample_df, x=column, y=second_column,
                    title=f"<b>🔍 {column} vs {second_column}</b> - Scatter Plot",
                    color_discrete_sequence=[colors[0]],
                    template="plotly_white",
                    trendline="lowess"
                )
                fig.update_traces(
                    marker=dict(size=10, opacity=0.6, line=dict(width=1, color='black')),
                    hovertemplate=f"<b>{column}: %{{x}}</b><br>{second_column}: %{{y}}<extra></extra>"
                )
            elif len(df.columns) > 1:
                second_col = [c for c in df.columns if c != column][0]
                sample_df = df.head(100)
                fig = px.scatter(
                    sample_df, x=column, y=second_col,
                    title=f"<b>🔍 {column} vs {second_col}</b> - Scatter Analysis",
                    color_discrete_sequence=[colors[0]],
                    template="plotly_white"
                )
                fig.update_traces(
                    marker=dict(size=8, opacity=0.7, line=dict(width=1, color='black'))
                )
            else:
                sample_df = df.head(100)
                fig = px.scatter(
                    sample_df, x=sample_df.index, y=column,
                    title=f"<b>🔍 {column}</b> - Scatter Distribution",
                    color_discrete_sequence=[colors[0]],
                    template="plotly_white"
                )
                fig.update_traces(marker=dict(size=6, opacity=0.8))
               
        elif chart_type == "Box":
            fig = px.box(
                df, y=column,
                title=f"<b>📦 {column}</b> - Statistical Distribution",
                color_discrete_sequence=[colors[0]],
                template="plotly_white"
            )
            fig.update_traces(
                boxpoints='all',
                jitter=0.3,
                pointpos=-1.8,
                marker=dict(size=4, opacity=0.6)
            )
           
        elif chart_type == "Area":
            if second_column and second_column != "None":
                sample_df = df.head(50)
                fig = px.area(
                    sample_df, x=column, y=second_column,
                    title=f"<b>📈 {column} vs {second_column}</b> - Area Chart",
                    color_discrete_sequence=[colors[0]],
                    template="plotly_white"
                )
            else:
                sample_df = df.head(50)
                fig = px.area(
                    sample_df, x=sample_df.index, y=column,
                    title=f"<b>📈 {column}</b> - Area Analysis",
                    color_discrete_sequence=[colors[0]],
                    template="plotly_white"
                )
            fig.update_traces(
                line=dict(width=3),
                hovertemplate="<b>Value: %{y}</b><extra></extra>"
            )
           
        elif chart_type == "Pie":
            value_counts = df[column].value_counts().head(8)  # Limit to top 8 for better visualization
            fig = px.pie(
                values=value_counts.values,
                names=value_counts.index,
                title=f"<b>🥧 {column}</b> - Composition Analysis",
                color_discrete_sequence=colors,
                template="plotly_white"
            )
            fig.update_traces(
                textposition='inside',
                textinfo='percent+label',
                pull=[0.1 if i == 0 else 0 for i in range(len(value_counts))],
                hovertemplate="<b>%{label}</b><br>Count: %{value}<br>Percentage: %{percent}<extra></extra>"
            )
        else:
            # Default fallback
            fig = px.histogram(df, x=column, title=f"<b>📊 {column}</b> - Distribution")
       
        # Enhanced layout with maximum animations and interactivity
        fig.update_layout(
            template=style,
            height=500,  # Increased height for better visibility
            showlegend=True,
            font=dict(
                color='#1f2937',  # Dark color for better visibility in both modes
                size=14,
                family="Arial, sans-serif"
            ),
            title_font=dict(
                color='#1f2937',  # Dark color for titles
                size=20,
                family="Arial Black, sans-serif"
            ),
            paper_bgcolor='rgba(255,255,255,0.9)',
            plot_bgcolor='rgba(255,255,255,0.9)',
            xaxis=dict(
                title_font=dict(color='#1f2937', size=14, weight='bold'),
                tickfont=dict(color='#4b5563', size=12),
                gridcolor='rgba(0,0,0,0.1)',
                linecolor='rgba(0,0,0,0.2)',
                linewidth=2
            ),
            yaxis=dict(
                title_font=dict(color='#1f2937', size=14, weight='bold'),
                tickfont=dict(color='#4b5563', size=12),
                gridcolor='rgba(0,0,0,0.1)',
                linecolor='rgba(0,0,0,0.2)',
                linewidth=2
            ),
            legend=dict(
                font=dict(color='#1f2937', size=12),
                bgcolor='rgba(255,255,255,0.8)',
                bordercolor='rgba(0,0,0,0.2)',
                borderwidth=1
            ),
            hoverlabel=dict(
                bgcolor='white',
                bordercolor='black',
                font=dict(color='black', size=12)
            ),
            # Animation configurations
            transition=dict(
                duration=500,
                easing='cubic-in-out'
            )
        )

        return fig
       
    except Exception as e:
        st.error(f"📊 Chart Error: Failed to create advanced chart for {column}: {str(e)}")
        # Fallback to basic chart
        return create_plotly_chart(df, column, chart_type, style, second_column)

def create_plotly_chart(df, column, chart_type, style="plotly_white", second_column=None):
    """Create basic Plotly chart as fallback"""
    try:
        if chart_type == "Bar":
            if df[column].dtype == 'object':
                value_counts = df[column].value_counts()
                fig = px.bar(x=value_counts.index, y=value_counts.values,
                           title=f"{column} - Bar Chart")
            else:
                fig = px.histogram(df, x=column, title=f"{column} - Distribution")
        elif chart_type == "Line":
            if second_column and second_column != "None":
                fig = px.line(df, x=column, y=second_column, title=f"{column} vs {second_column} - Line Chart")
            else:
                fig = px.line(df, x=df.index, y=column, title=f"{column} - Line Chart")
        elif chart_type == "Histogram":
            fig = px.histogram(df, x=column, title=f"{column} - Histogram")
        elif chart_type == "Scatter":
            if second_column and second_column != "None":
                fig = px.scatter(df, x=column, y=second_column, title=f"{column} vs {second_column} - Scatter Plot")
            elif len(df.columns) > 1:
                second_col = [c for c in df.columns if c != column][0]
                fig = px.scatter(df, x=column, y=second_col, title=f"{column} vs {second_col} - Scatter Plot")
            else:
                fig = px.scatter(df, x=df.index, y=column, title=f"{column} - Scatter Plot")
        elif chart_type == "Box":
            fig = px.box(df, y=column, title=f"{column} - Box Plot")
        elif chart_type == "Area":
            if second_column and second_column != "None":
                fig = px.area(df, x=column, y=second_column, title=f"{column} vs {second_column} - Area Chart")
            else:
                fig = px.area(df, x=df.index, y=column, title=f"{column} - Area Chart")
        elif chart_type == "Pie":
            value_counts = df[column].value_counts()
            fig = px.pie(values=value_counts.values, names=value_counts.index,
                       title=f"{column} - Pie Chart")
        else:
            fig = px.histogram(df, x=column, title=f"{column} - Distribution")
       
        fig.update_layout(
            template=style,
            height=400,
            showlegend=True,
            font=dict(
                color='#1f2937',
                size=14,
                family="Arial, sans-serif"
            ),
            title_font=dict(
                color='#1f2937',
                size=18,
                family="Arial, sans-serif"
            ),
            paper_bgcolor='white',
            plot_bgcolor='white'
        )
                   
        return fig
    except Exception as e:
        st.error(f"📊 Chart Error: Failed to create chart for {column}: {str(e)}")
        fig = go.Figure()
        fig.add_annotation(text=f"Error creating chart: {str(e)}", x=0.5, y=0.5,
                         xref="paper", yref="paper", showarrow=False,
                         font=dict(color='#1f2937', size=14))
        return fig

def clean_dataframe(df):
    """Clean dataframe to fix Arrow serialization issues"""
    try:
        df_clean = df.copy()
       
        for col in df_clean.columns:
            if df_clean[col].dtype == 'object':
                df_clean[col] = pd.to_numeric(df_clean[col], errors='ignore')
               
                if df_clean[col].dtype == 'object':
                    df_clean[col] = df_clean[col].astype(str)
       
        return df_clean
    except Exception as e:
        st.error(f"🧹 Data Cleaning Error: {str(e)}")
        return df
   
def clear_temp():
    """Clear temporary files and reset session state"""
    d = st.session_state.get('tmp_dir')
    try:
        if d and os.path.isdir(d):
            shutil.rmtree(d)
    except Exception as e:
        st.error(f"🗑️ Cleanup Error: {str(e)}")
   
    initialize_session_state()
    st.success("✅ All data cleared! You can upload new files.")

def generate_presentation_config():
    """Generate the complete presentation configuration JSON in required format"""
   
    if st.session_state.current_template == "none" and not st.session_state.uploaded_template_file:
        st.error("❌ Please select a template or upload a custom template!")
        return None
       
    if st.session_state.uploaded_template_file:
        template_name = st.session_state.uploaded_template_file.name
    else:
        template_name = st.session_state.current_template
   
    config = {
        "presentation_title": st.session_state.presentation_title,
        "template_name": template_name,
        "timestamp": datetime.now().isoformat(),
        "text_color": st.session_state.template_text_color,  # Use template-based text color
        "dashboards": [],
        "insights": [],
        "insights_charts": {},
        "comparison": {},
        "queries": [],
        "styles": st.session_state.custom_styles
    }
   
    if st.session_state.selected_features["dashboard"]:
        for chart in st.session_state.dashboard_config.get("charts", []):
            column1 = chart["column"]
            column2 = chart.get("second_column", "None")
            chart_type = chart["type"].lower()
           
            config["dashboards"].append([column1, column2, chart_type])
   
    if st.session_state.selected_features["charts_insights"]:
        for chart in st.session_state.insight_config.get("charts", []):
            column_name = chart["column"]
            chart_type = chart["type"].lower()
            config["insights_charts"][column_name] = chart_type
   
    if st.session_state.selected_features["insights"]:
        config["insights"] = st.session_state.insight_config.get("selected_columns", [])
   
    if st.session_state.selected_features["comparison"]:
        for comp in st.session_state.comparison_config.get("comparisons", []):
            key = f"{comp['compare']}_vs_{comp['group_by']}"
            config["comparison"][key] = comp["chart_type"].lower()
   
    if st.session_state.selected_features["query"]:
        config["queries"] = st.session_state.custom_query.get("queries", [])
   
    return config

def display_advanced_charts_grid(charts_data, feature_type):
    """Display charts in ADVANCED grid layout with maximum animations - SIMPLIFIED"""
    if not charts_data:
        st.info("📊 No charts configured yet. Add charts to see live previews!")
        return
   
    st.subheader("🎯 Live Advanced Charts Preview")
   
    # SIMPLIFIED: Single view with advanced animations
    if len(charts_data) <= 2:
        cols = st.columns(len(charts_data))
    else:
        cols = st.columns(2)
   
    for i, chart_data in enumerate(charts_data):
        col_idx = i % len(cols)
        with cols[col_idx]:
            with st.container():
                # Enhanced card with ULTRA animations
                if feature_type == "dashboard":
                    chart_title = chart_data.get('column', 'Chart')
                    if "second_column" in chart_data and chart_data["second_column"] != "None":
                        chart_title = f"{chart_data['column']} vs {chart_data['second_column']}"
                elif feature_type == "charts_insights":
                    chart_title = chart_data.get('column', 'Chart')
                    if "second_column" in chart_data and chart_data["second_column"] != "None":
                        chart_title = f"{chart_data['column']} vs {chart_data['second_column']}"
                else:
                    chart_title = f"{chart_data.get('compare', 'Chart')} vs {chart_data.get('group_by', 'Group')}"
               
                st.markdown(f"""
                    <div class='chart-card advanced-chart-card slide-reveal'>
                        <h4>🚀 {chart_title}</h4>
                        <p class='chart-type-badge'>{chart_data.get('type', 'Chart')}</p>
                    </div>
                """, unsafe_allow_html=True)
               
                # Create advanced chart with MULTI-COLOR scheme
                if feature_type == "dashboard":
                    fig = create_advanced_plotly_chart(
                        st.session_state.df,
                        chart_data["column"],
                        chart_data["type"],
                        second_column=chart_data.get("second_column", None)
                    )
                elif feature_type == "charts_insights":
                    fig = create_advanced_plotly_chart(
                        st.session_state.df,
                        chart_data["column"],
                        chart_data["type"],
                        second_column=chart_data.get("second_column", None)
                    )
                elif feature_type == "comparison":
                    fig = create_advanced_plotly_chart(
                        st.session_state.df,
                        chart_data["compare"],
                        chart_data["chart_type"],
                        second_column=chart_data["group_by"]
                    )
               
                st.plotly_chart(fig, use_container_width=True, config={
                    'displayModeBar': True,
                    'scrollZoom': True,
                    'doubleClick': 'reset'
                })

def create_slide_preview_data():
    """Create slide preview data based on ACTUAL generated presentation"""
    if not st.session_state.presentation_config:
        return []
   
    if st.session_state.slide_images_loaded and st.session_state.slide_images:
        return st.session_state.slide_images
   
    config = st.session_state.presentation_config
   
    slide_data = []
   
    slide_data.append({
        "title": "📋 Title Slide",
        "content": [
            f"**Title:** {config.get('presentation_title', 'Data Analysis Presentation')}",
            f"**Template:** {config.get('template_name', 'Professional').title()}",
            f"**Text Color:** {config.get('text_color', '#000000')}",
            f"**Generated:** {datetime.now().strftime('%Y-%m-%d %H:%M')}"
        ],
        "slide_number": 1
    })
   
    if st.session_state.df is not None:
        slide_data.append({
            "title": "📊 Data Overview",
            "content": [
                f"**Total Rows:** {len(st.session_state.df):,}",
                f"**Total Columns:** {len(st.session_state.df.columns)}",
                f"**Numeric Columns:** {len(st.session_state.df.select_dtypes(include=['number']).columns)}",
                f"**Categorical Columns:** {len(st.session_state.df.select_dtypes(include=['object']).columns)}"
            ],
            "slide_number": 2
    })
   
    if config.get("dashboards"):
        slide_number = len(slide_data) + 1
        dashboard_content = ["**Dashboard Charts:**"]
        for chart_config in config["dashboards"]:
            if isinstance(chart_config, list) and len(chart_config) == 3:
                column1, column2, chart_type = chart_config
                if column2 != "None":
                    dashboard_content.append(f"• {column1} vs {column2}: {chart_type} chart")
                else:
                    dashboard_content.append(f"• {column1}: {chart_type} chart")
   
        slide_data.append({
            "title": "📈 Dashboard Overview",
            "content": dashboard_content,
            "slide_number": slide_number
        })
   
    if config.get("insights"):
        slide_number = len(slide_data) + 1
        insights_content = ["**AI Insights for:**"]
        for col in config["insights"]:
            insights_content.append(f"• {col}")
       
        slide_data.append({
            "title": "💡 AI Insights",
            "content": insights_content,
            "slide_number": slide_number
        })
   
    if config.get("insights_charts"):
        slide_number = len(slide_data) + 1
        charts_content = ["**Charts with Insights:**"]
        for col, chart_type in config["insights_charts"].items():
            charts_content.append(f"• {col}: {chart_type} chart")
       
        slide_data.append({
            "title": "📊 Charts with Insights",
            "content": charts_content,
            "slide_number": slide_number
        })
   
    if config.get("comparison"):
        slide_number = len(slide_data) + 1
        comparison_content = ["**Data Comparisons:**"]
        for comp, chart_type in config["comparison"].items():
            comparison_content.append(f"• {comp}: {chart_type} chart")
       
        slide_data.append({
            "title": "⚖️ Data Comparisons",
            "content": comparison_content,
            "slide_number": slide_number
        })
   
    if config.get("queries"):
        slide_number = len(slide_data) + 1
        query_content = ["**Custom Queries:**"]
       
        for i, query in enumerate(config["queries"]):
            query_data = query if isinstance(query, dict) else {"text": str(query)}
            query_text = f"• {query_data.get('text', 'Unknown query')}"
           
            if not query_data.get('include_in_slides', True):
                query_text += " [EXCLUDED]"
           
            query_content.append(query_text)
       
        slide_data.append({
            "title": "🔍 Custom Queries",
            "content": query_content,
            "slide_number": slide_number
        })
   
    slide_data.append({
        "title": "📝 Summary",
        "content": [
            "**Presentation Summary:**",
            f"• Total Slides: {len(slide_data) + 1}",
            f"• Data Source: {st.session_state.uploaded_csv_file.name if st.session_state.uploaded_csv_file else 'Unknown'}",
            f"• Template: {config.get('template_name', 'Professional').title()}",
            f"• Text Color: {config.get('text_color', '#000000')}",
            "• Analysis Complete: Yes",
            f"• Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        ],
        "slide_number": len(slide_data) + 1
    })
   
    return slide_data

def display_generation_errors():
    """Display generation errors and warnings to user"""
    if st.session_state.generation_errors:
        st.markdown("---")
        st.subheader("⚠️ Generation Warnings & Errors")
       
        for i, error in enumerate(st.session_state.generation_errors):
            if "warning" in error.lower() or "warn" in error.lower():
                st.warning(f"**Warning {i+1}:** {error}")
            elif "error" in error.lower() or "failed" in error.lower():
                st.error(f"**Error {i+1}:** {error}")
            else:
                st.info(f"**Info {i+1}:** {error}")

def show_success_popup():
    """Show success popup when presentation is generated"""
    if st.session_state.show_success_popup:
        # Create a clean, professional success notification
        st.markdown("""
            <div style='
                position: fixed;
                top: 20px;
                right: 20px;
                background: linear-gradient(135deg, #48bb78 0%, #38a169 100%);
                color: white;
                padding: 20px 30px;
                border-radius: 15px;
                box-shadow: 0 10px 30px rgba(72, 187, 120, 0.4);
                z-index: 1000;
                text-align: center;
                animation: slideInRight 0.6s cubic-bezier(0.25, 0.46, 0.45, 0.94);
                border: 2px solid rgba(255,255,255,0.3);
                backdrop-filter: blur(10px);
                min-width: 300px;
            '>
                <div style='font-size: 48px; margin-bottom: 10px;'>✅</div>
                <h3 style='margin: 0 0 8px 0; color: white; font-weight: 700; font-size: 18px;'>Success!</h3>
                <p style='font-size: 14px; margin: 0; font-weight: 600; opacity: 0.9;'>Presentation Generated Successfully</p>
            </div>
            <style>
                @keyframes slideInRight {
                    from { transform: translateX(100%); opacity: 0; }
                    to { transform: translateX(0); opacity: 1; }
                }
            </style>
        """, unsafe_allow_html=True)
       
        # Reset the popup state after showing
        st.session_state.show_success_popup = False

def load_css():
    """Load custom CSS styles with DUAL THEME support (Dark/Light)"""
    st.markdown(f"""
        <style>
            /* DUAL THEME VARIABLES */
            [data-theme="light"] {{
                --primary-gradient: linear-gradient(135deg, #f26f21 0%, #ffa800 100%);
                --success-gradient: linear-gradient(135deg, #48bb78, #38a169);
                --warning-gradient: linear-gradient(135deg, #ed8936, #dd6b20);
                --error-gradient: linear-gradient(135deg, #f56565, #e53e3e);
                --glass-bg: rgba(255, 255, 255, 0.1);
                --glass-border: rgba(0, 0, 0, 0.1);
                --text-primary: #1f2937;
                --text-secondary: #4b5563;
                --text-muted: #6b7280;
                --border-light: rgba(0, 0, 0, 0.1);
                --bg-light: #f8fafc;
                --background-color: #ffffff;
                --card-background: #ffffff;
                --dark-grey-bg: #f8fafc;
                --darker-grey-bg: #f1f5f9;
                --slide-content-color: #374151;
                --chart-bg: #ffffff;
                --chart-text: #1f2937;
            }}
           
            [data-theme="dark"] {{
                --primary-gradient: linear-gradient(135deg, #f26f21 0%, #ffa800 100%);
                --success-gradient: linear-gradient(135deg, #48bb78, #38a169);
                --warning-gradient: linear-gradient(135deg, #ed8936, #dd6b20);
                --error-gradient: linear-gradient(135deg, #f56565, #e53e3e);
                --glass-bg: #2d374899;
                --glass-border: rgba(255, 255, 255, 0.15);
                --text-primary: #ffffff;
                --text-secondary: #e2e8f0;
                --text-muted: #a0aec0;
                --border-light: rgba(255, 255, 255, 0.1);
                --bg-light: #4a5568;
                --background-color: #1a202c;
                --card-background: #2d3748;
                --dark-grey-bg: #2d3748;
                --darker-grey-bg: #1a202c;
                --slide-content-color: #e2e8f0;
                --chart-bg: #2d3748;
                --chart-text: #ffffff;
            }}

            /* Base styles that work in both themes */
            .main {{
                background: var(--background-color);
                color: var(--text-primary);
            }}
           
            .stApp {{
                background: var(--background-color);
                min-height: 100vh;
            }}

            /* ULTRA ENHANCED FILE UPLOADER STYLING */
            .stFileUploader > div {{
                background: var(--card-background) !important;
                border: 3px dashed var(--glass-border) !important;
                border-radius: 20px !important;
                padding: 40px 20px !important;
                transition: all 0.6s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
                backdrop-filter: blur(15px);
                box-shadow: 0 8px 25px rgba(0, 0, 0, 0.1) !important;
                position: relative;
                overflow: hidden;
            }}

            .stFileUploader > div::before {{
                content: '';
                position: absolute;
                top: 0;
                left: -100%;
                width: 100%;
                height: 100%;
                background: linear-gradient(90deg, transparent, rgba(255,255,255,0.3), transparent);
                transition: left 0.8s;
            }}

            .stFileUploader > div:hover::before {{
                left: 100%;
            }}

            .stFileUploader > div:hover {{
                border-color: #f26f21 !important;
                border-style: solid !important;
                transform: translateY(-8px) scale(1.02) !important;
                box-shadow: 0 20px 40px rgba(242, 111, 33, 0.3), 0 15px 35px rgba(0, 0, 0, 0.2) !important;
                background: linear-gradient(135deg, var(--card-background) 0%, var(--bg-light) 100%) !important;
                animation: uploaderHover 2s infinite !important;
            }}

            @keyframes uploaderHover {{
                0%, 100% {{ transform: translateY(-8px) scale(1.02); }}
                50% {{ transform: translateY(-10px) scale(1.03); }}
            }}

            .stFileUploader > div:active {{
                transform: translateY(-2px) scale(1.01) !important;
            }}

            /* ENHANCED BUTTON STYLING FOR QUERY SECTION */
            .query-button {{
                min-height: 50px !important;
                height: 50px !important;
                padding: 12px 24px !important;
                font-size: 14px !important;
                font-weight: 700 !important;
                border-radius: 12px !important;
                transition: all 0.4s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
                margin: 8px 0 !important;
                border: none !important;
                box-shadow: 0 6px 20px rgba(242, 111, 33, 0.4) !important;
                background: linear-gradient(135deg, #f26f21 0%, #ffa800 100%) !important;
                color: white !important;
            }}

            .query-button:hover {{
                transform: translateY(-4px) scale(1.08) !important;
                box-shadow: 0 12px 30px rgba(255, 168, 0, 0.6), 0 8px 25px rgba(0, 0, 0, 0.15) !important;
                background: linear-gradient(135deg, #ff7b2e 0%, #ffb733 100%) !important;
                animation: buttonPulse 2s infinite !important;
            }}

            @keyframes buttonPulse {{
                0%, 100% {{ transform: translateY(-4px) scale(1.08); }}
                50% {{ transform: translateY(-6px) scale(1.1); }}
            }}

            /* STATIC NAVIGATION BUTTONS STYLING */
            .nav-section-button {{
                background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
                color: white !important;
                border: none !important;
                border-radius: 12px !important;
                padding: 14px 28px !important;
                font-weight: 700 !important;
                font-size: 14px !important;
                transition: all 0.5s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
                box-shadow: 0 6px 20px rgba(102, 126, 234, 0.4) !important;
                margin: 8px !important;
                min-height: 50px !important;
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
            }}

            .nav-section-button:hover {{
                transform: translateY(-4px) scale(1.06) !important;
                box-shadow: 0 12px 30px rgba(118, 75, 162, 0.5) !important;
                background: linear-gradient(135deg, #768efa 0%, #865bb2 100%) !important;
            }}

            /* MAXIMUM ENHANCED ANIMATIONS AND TRANSITIONS */
            .feature-card {{
                background: var(--card-background);
                backdrop-filter: blur(20px);
                border: 1px solid var(--glass-border);
                border-radius: 20px;
                padding: 25px;
                margin: 15px 0;
                transition: all 0.6s cubic-bezier(0.25, 0.46, 0.45, 0.94);
                box-shadow: 0 12px 35px rgba(0, 0, 0, 0.15);
                position: relative;
                overflow: hidden;
                height: 160px !important;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                text-align: center;
                transform-style: preserve-3d;
                perspective: 1000px;
            }}
           
            .advanced-chart-card {{
                background: var(--card-background);
                border: 2px solid var(--glass-border);
                border-radius: 18px;
                padding: 22px;
                margin: 15px 0;
                transition: all 0.5s cubic-bezier(0.25, 0.46, 0.45, 0.94);
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.2);
                height: auto !important;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                text-align: center;
                transform-style: preserve-3d;
            }}
           
            .metric-card {{
                background: var(--card-background);
                border-radius: 18px;
                padding: 25px;
                margin: 12px;
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.15);
                border-left: 6px solid #f26f21;
                transition: all 0.5s cubic-bezier(0.25, 0.46, 0.45, 0.94);
                position: relative;
                overflow: hidden;
                height: 140px !important;
                display: flex;
                flex-direction: column;
                justify-content: center;
                align-items: center;
                text-align: center;
                transform-style: preserve-3d;
            }}
           
            /* PERFECTLY CENTERED CARD CONTENT */
            .card-content {{
                display: flex !important;
                flex-direction: column !important;
                justify-content: center !important;
                align-items: center !important;
                text-align: center !important;
                width: 100% !important;
                height: 100% !important;
                padding: 0 !important;
                margin: 0 !important;
                transform: translateZ(20px);
            }}
           
            .card-title {{
                font-size: 18px !important;
                font-weight: 800 !important;
                margin-bottom: 8px !important;
                color: var(--text-primary) !important;
                text-align: center !important;
                line-height: 1.3 !important;
            }}
           
            .card-value {{
                font-size: 28px !important;
                font-weight: 900 !important;
                margin: 10px 0 !important;
                color: var(--text-primary) !important;
                text-align: center !important;
                line-height: 1 !important;
            }}
           
            .card-description {{
                font-size: 14px !important;
                color: var(--text-muted) !important;
                font-weight: 600 !important;
                text-align: center !important;
                line-height: 1.3 !important;
                margin: 0 !important;
            }}
           
            /* ULTRA ENHANCED: ORANGE-YELLOW NAVIGATION BUTTONS WITH MAX ANIMATIONS */
            .stButton button {{
                background: linear-gradient(135deg, #f26f21 0%, #ffa800 100%) !important;
                color: #ffffff !important;
                border: none !important;
                border-radius: 16px !important;
                padding: 16px 32px !important;
                font-weight: 800 !important;
                font-size: 16px !important;
                transition: all 0.6s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
                box-shadow: 0 8px 25px rgba(242, 111, 33, 0.4) !important;
                position: relative !important;
                overflow: hidden !important;
                min-height: 55px !important;
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
                transform: translateY(0) scale(1);
                border: 2px solid transparent !important;
            }}
           
            .stButton button:hover {{
                transform: translateY(-6px) scale(1.08) !important;
                box-shadow: 0 20px 40px rgba(242, 111, 33, 0.6), 0 15px 35px rgba(0, 0, 0, 0.2) !important;
                color: #ffffff !important;
                background: linear-gradient(135deg, #ff7b2e 0%, #ffb733 100%) !important;
                border-color: #ffffff !important;
                animation: buttonGlow 2s infinite !important;
            }}
           
            .stButton button:active {{
                transform: translateY(-3px) scale(1.04) !important;
                box-shadow: 0 12px 30px rgba(242, 111, 33, 0.5) !important;
                transition: all 0.2s ease !important;
            }}

            /* Navigation buttons specific styling */
            .nav-button {{
                background: linear-gradient(135deg, #f26f21 0%, #ffa800 100%) !important;
                color: white !important;
                border: none !important;
                border-radius: 12px !important;
                padding: 12px 24px !important;
                font-weight: 700 !important;
                font-size: 14px !important;
                transition: all 0.4s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
                box-shadow: 0 6px 20px rgba(242, 111, 33, 0.4) !important;
                margin: 5px !important;
            }}
           
            .nav-button:hover {{
                transform: translateY(-3px) scale(1.05) !important;
                box-shadow: 0 10px 25px rgba(255, 168, 0, 0.6) !important;
                background: linear-gradient(135deg, #ff7b2e 0%, #ffb733 100%) !important;
            }}
           
            /* ULTRA ENHANCED: ORANGE-YELLOW NAVIGATION MENU */
            .css-1d391kg {{
                background: var(--card-background) !important;
                border-radius: 20px !important;
                border: 2px solid var(--glass-border) !important;
                padding: 15px !important;
                box-shadow: 0 15px 35px rgba(0, 0, 0, 0.1) !important;
                transition: all 0.5s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
                backdrop-filter: blur(15px);
            }}
           
            /* ULTRA ENHANCED: Navigation items with maximum animations */
            .nav-item {{
                transition: all 0.6s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
                border-radius: 16px !important;
                margin: 8px 0 !important;
                padding: 16px 20px !important;
                background: rgba(242, 111, 33, 0.1) !important;
                border: 2px solid transparent !important;
                position: relative !important;
                overflow: hidden !important;
                transform: translateX(0) scale(1);
                box-shadow: 0 6px 20px rgba(0, 0, 0, 0.1);
                backdrop-filter: blur(10px);
            }}
           
            .nav-item:hover {{
                background: rgba(242, 111, 33, 0.2) !important;
                transform: translateX(12px) scale(1.06) !important;
                border-color: #f26f21 !important;
                box-shadow: 0 15px 35px rgba(242, 111, 33, 0.3), 0 10px 25px rgba(0, 0, 0, 0.1) !important;
                animation: navPulse 3s infinite !important;
            }}
           
            .nav-item-selected {{
                background: linear-gradient(135deg, #f26f21 0%, #ffa800 100%) !important;
                color: #ffffff !important;
                box-shadow: 0 15px 35px rgba(242, 111, 33, 0.4), 0 10px 25px rgba(0, 0, 0, 0.1) !important;
                transform: translateX(10px) scale(1.08) !important;
                border: none !important;
                animation: selectedGlow 2s infinite !important;
            }}

            /* MAXIMUM ENHANCED ANIMATIONS */
            @keyframes float {{
                0%, 100% {{ transform: translateY(0px) rotate(0deg) scale(1); }}
                25% {{ transform: translateY(-8px) rotate(1deg) scale(1.02); }}
                50% {{ transform: translateY(-12px) rotate(-1deg) scale(1.01); }}
                75% {{ transform: translateY(-6px) rotate(0.5deg) scale(1.03); }}
            }}
           
            @keyframes glow {{
                0% {{ box-shadow: 0 0 10px #f26f21, 0 0 20px rgba(242, 111, 33, 0.3); }}
                50% {{ box-shadow: 0 0 25px #ffa800, 0 0 40px rgba(255, 168, 0, 0.4), 0 0 60px rgba(255, 168, 0, 0.2); }}
                100% {{ box-shadow: 0 0 10px #f26f21, 0 0 20px rgba(242, 111, 33, 0.3); }}
            }}
           
            @keyframes buttonGlow {{
                0%, 100% {{ box-shadow: 0 8px 25px rgba(242, 111, 33, 0.4), 0 0 10px rgba(242, 111, 33, 0.2); }}
                50% {{ box-shadow: 0 12px 35px rgba(255, 168, 0, 0.6), 0 0 20px rgba(255, 168, 0, 0.3), 0 0 30px rgba(255, 168, 0, 0.1); }}
            }}
           
            @keyframes navPulse {{
                0%, 100% {{ transform: translateX(12px) scale(1.06); }}
                50% {{ transform: translateX(12px) scale(1.08); }}
            }}
           
            @keyframes selectedGlow {{
                0%, 100% {{ box-shadow: 0 15px 35px rgba(242, 111, 33, 0.4), 0 10px 25px rgba(0, 0, 0, 0.1); }}
                50% {{ box-shadow: 0 20px 45px rgba(255, 168, 0, 0.6), 0 15px 35px rgba(0, 0, 0, 0.15), 0 0 25px rgba(255, 168, 0, 0.2); }}
            }}
           
            @keyframes pulse {{
                0%, 100% {{ transform: scale(1) rotate(0deg); }}
                25% {{ transform: scale(1.05) rotate(1deg); }}
                50% {{ transform: scale(1.08) rotate(-1deg); }}
                75% {{ transform: scale(1.03) rotate(0.5deg); }}
            }}
           
            @keyframes slideIn {{
                from {{ transform: translateX(-30px); opacity: 0; filter: blur(10px); }}
                to {{ transform: translateX(0); opacity: 1; filter: blur(0); }}
            }}
           
            @keyframes fadeInUp {{
                from {{ transform: translateY(30px) scale(0.95); opacity: 0; filter: blur(10px); }}
                to {{ transform: translateY(0) scale(1); opacity: 1; filter: blur(0); }}
            }}
           
            @keyframes shimmer {{
                0% {{ background-position: -1000px 0; }}
                100% {{ background-position: 1000px 0; }}
            }}

            @keyframes slideInRight {{
                from {{ transform: translateX(100%); opacity: 0; }}
                to {{ transform: translateX(0); opacity: 1; }}
            }}
           
            /* NEW: Advanced realistic animations */
            @keyframes slideReveal {{
                0% {{ transform: translateX(-100%) rotateY(45deg); opacity: 0; }}
                60% {{ transform: translateX(20px) rotateY(-15deg); opacity: 0.8; }}
                100% {{ transform: translateX(0) rotateY(0deg); opacity: 1; }}
            }}
           
            @keyframes imageZoom {{
                0% {{ transform: scale(0.8) rotate(-2deg); opacity: 0; }}
                100% {{ transform: scale(1) rotate(0deg); opacity: 1; }}
            }}
           
            @keyframes contentSlide {{
                0% {{ transform: translateY(50px); opacity: 0; }}
                100% {{ transform: translateY(0); opacity: 1; }}
            }}
           
            .float-animation {{
                animation: float 4s ease-in-out infinite;
                transform-style: preserve-3d;
            }}
            .glow-animation {{
                animation: glow 3s ease-in-out infinite;
            }}
            .pulse-animation {{
                animation: pulse 3s ease-in-out infinite;
            }}
            .slide-in {{
                animation: slideIn 0.8s cubic-bezier(0.25, 0.46, 0.45, 0.94) forwards;
            }}
            .fade-in-up {{
                animation: fadeInUp 0.9s cubic-bezier(0.25, 0.46, 0.45, 0.94) forwards;
            }}
            .slide-reveal {{
                animation: slideReveal 1.2s cubic-bezier(0.25, 0.46, 0.45, 0.94) forwards;
            }}
            .image-zoom {{
                animation: imageZoom 1s cubic-bezier(0.25, 0.46, 0.45, 0.94) forwards;
            }}
            .content-slide {{
                animation: contentSlide 0.8s cubic-bezier(0.25, 0.46, 0.45, 0.94) forwards;
            }}
           
            /* ULTRA ENHANCED: Card hover effects with 3D transforms */
            .feature-card::before {{
                content: '';
                position: absolute;
                top: 0;
                left: -150%;
                width: 150%;
                height: 100%;
                background: linear-gradient(90deg, transparent, rgba(255,255,255,0.15), transparent);
                transition: left 0.8s;
                transform: skewX(-25deg);
            }}
           
            .feature-card:hover::before {{
                left: 150%;
            }}
           
            .feature-card:hover {{
                transform: translateY(-10px) scale(1.05) rotateY(5deg);
                box-shadow: 0 25px 50px rgba(242, 111, 33, 0.3), 0 15px 35px rgba(0, 0, 0, 0.15);
                border-color: #f26f21;
                animation: cardHover 2s infinite !important;
            }}
           
            @keyframes cardHover {{
                0%, 100% {{ transform: translateY(-10px) scale(1.05) rotateY(5deg); }}
                50% {{ transform: translateY(-12px) scale(1.06) rotateY(3deg); }}
            }}
           
            .advanced-chart-card:hover {{
                transform: translateY(-8px) scale(1.04) rotateX(5deg);
                box-shadow: 0 20px 45px rgba(242, 111, 33, 0.25), 0 12px 30px rgba(0, 0, 0, 0.15);
                border-color: #ffa800;
            }}
           
            .metric-card::before {{
                content: '';
                position: absolute;
                top: 0;
                left: 0;
                right: 0;
                height: 4px;
                background: var(--primary-gradient);
                transform: scaleX(0);
                transition: transform 0.6s cubic-bezier(0.25, 0.46, 0.45, 0.94);
                border-radius: 18px 18px 0 0;
            }}
           
            .metric-card:hover::before {{
                transform: scaleX(1);
            }}
           
            .metric-card:hover {{
                transform: translateY(-8px) scale(1.06) rotateZ(1deg);
                box-shadow: 0 20px 45px rgba(242, 111, 33, 0.3), 0 12px 30px rgba(0, 0, 0, 0.2);
                animation: metricPulse 2s infinite;
            }}
           
            @keyframes metricPulse {{
                0%, 100% {{ transform: translateY(-8px) scale(1.06) rotateZ(1deg); }}
                50% {{ transform: translateY(-10px) scale(1.08) rotateZ(-1deg); }}
            }}
           
            /* FIXED: Download section specific styling */
            .download-section {{
                background: linear-gradient(135deg, var(--card-background) 0%, var(--bg-light) 100%) !important;
                border-radius: 20px !important;
                padding: 30px !important;
                margin: 20px 0 !important;
                border: 3px solid var(--glass-border) !important;
                box-shadow: 0 20px 50px rgba(0, 0, 0, 0.15) !important;
                transition: all 0.6s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
                backdrop-filter: blur(20px);
                position: relative;
                overflow: hidden;
            }}
           
            .download-section::before {{
                content: '';
                position: absolute;
                top: -50%;
                left: -50%;
                width: 200%;
                height: 200%;
                background: radial-gradient(circle, rgba(242, 111, 33, 0.1) 0%, transparent 70%);
                animation: shimmer 6s infinite linear;
            }}
           
            .download-section:hover {{
                transform: translateY(-8px) scale(1.02) !important;
                box-shadow: 0 30px 60px rgba(242, 111, 33, 0.2), 0 20px 40px rgba(0, 0, 0, 0.1) !important;
                border-color: #ffa800 !important;
            }}
           
            /* FIXED: Slide preview content styling - Better color for both themes */
            .slide-content-text {{
                color: var(--slide-content-color) !important;
                font-size: 16px !important;
                line-height: 1.7 !important;
                font-weight: 500 !important;
                padding: 8px 0 !important;
                border-left: 3px solid #f26f21 !important;
                padding-left: 15px !important;
                margin: 8px 0 !important;
                background: rgba(242, 111, 33, 0.05) !important;
                border-radius: 0 8px 8px 0 !important;
                transition: all 0.4s ease !important;
            }}
           
            .slide-content-text:hover {{
                background: rgba(242, 111, 33, 0.1) !important;
                transform: translateX(8px) scale(1.02) !important;
                border-left-color: #ffa800 !important;
                box-shadow: 0 5px 15px rgba(242, 111, 33, 0.1) !important;
            }}
           
            /* NEW: Chart type badge */
            .chart-type-badge {{
                background: linear-gradient(135deg, #f26f21, #ffa800);
                color: white;
                padding: 4px 12px;
                border-radius: 20px;
                font-size: 12px;
                font-weight: 700;
                margin: 5px 0;
                display: inline-block;
                box-shadow: 0 3px 10px rgba(242, 111, 33, 0.3);
            }}

            /* SLIDE IMAGE CONTAINER STYLING */
            .slide-image-container {{
                background: var(--card-background);
                border-radius: 20px;
                padding: 30px;
                margin: 20px 0;
                border: 3px solid var(--glass-border);
                box-shadow: 0 20px 50px rgba(0, 0, 0, 0.15);
                transition: all 0.6s cubic-bezier(0.25, 0.46, 0.45, 0.94);
                position: relative;
                overflow: hidden;
            }}

            .slide-image-container:hover {{
                transform: translateY(-5px) scale(1.01);
                box-shadow: 0 25px 60px rgba(242, 111, 33, 0.2), 0 15px 40px rgba(0, 0, 0, 0.1);
                border-color: #ffa800;
            }}

            .slide-image {{
                width: 100%;
                height: auto;
                border-radius: 15px;
                box-shadow: 0 15px 35px rgba(0, 0, 0, 0.2);
                transition: all 0.5s ease;
            }}

            .slide-image:hover {{
                transform: scale(1.02);
                box-shadow: 0 20px 45px rgba(0, 0, 0, 0.3);
            }}

            /* QUERY SECTION FIXES */
            .query-section {{
                background: var(--card-background);
                border-radius: 20px;
                padding: 25px;
                margin: 15px 0;
                border: 2px solid var(--glass-border);
                box-shadow: 0 10px 30px rgba(0, 0, 0, 0.1);
            }}

            .query-button-container {{
                display: flex;
                gap: 10px;
                margin: 15px 0;
                align-items: center;
            }}

            .query-button-fixed {{
                min-width: 120px !important;
                height: 45px !important;
                padding: 10px 20px !important;
                font-size: 14px !important;
                font-weight: 700 !important;
                border-radius: 12px !important;
                transition: all 0.4s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
                border: none !important;
                box-shadow: 0 6px 20px rgba(242, 111, 33, 0.4) !important;
                background: linear-gradient(135deg, #f26f21 0%, #ffa800 100%) !important;
                color: white !important;
            }}

            .query-button-fixed:hover {{
                transform: translateY(-4px) scale(1.08) !important;
                box-shadow: 0 12px 30px rgba(255, 168, 0, 0.6) !important;
                background: linear-gradient(135deg, #ff7b2e 0%, #ffb733 100%) !important;
            }}

            .query-button-secondary-fixed {{
                min-width: 120px !important;
                height: 45px !important;
                padding: 10px 20px !important;
                font-size: 14px !important;
                font-weight: 700 !important;
                border-radius: 12px !important;
                transition: all 0.4s cubic-bezier(0.25, 0.46, 0.45, 0.94) !important;
                display: flex !important;
                align-items: center !important;
                justify-content: center !important;
                border: none !important;
                box-shadow: 0 6px 20px rgba(78, 205, 196, 0.4) !important;
                background: linear-gradient(135deg, #4ECDC4 0%, #45B7D1 100%) !important;
                color: white !important;
            }}

            .query-button-secondary-fixed:hover {{
                transform: translateY(-4px) scale(1.08) !important;
                box-shadow: 0 12px 30px rgba(69, 183, 209, 0.6) !important;
                background: linear-gradient(135deg, #5EDDD6 0%, #55C7E1 100%) !important;
            }}
        </style>
    """, unsafe_allow_html=True)

# Load custom CSS
load_css()

# Sidebar Navigation
with st.sidebar:
    st.markdown('<div class="float-animation">', unsafe_allow_html=True)
    st.title("🚀 PPT Studio")
    st.markdown('</div>', unsafe_allow_html=True)
   
    # Template selection with TEXT COLOR OPTION FOR CUSTOM TEMPLATES
    with st.expander("🎭 Choose Template", expanded=True):
        template_options = [{"name": "None", "value": "none", "description": "No template selected", "icon": "⚪"}] + st.session_state.sample_templates
       
        selected_template = st.radio(
            "Select a template:",
            options=[t['value'] for t in template_options],
            format_func=lambda x: next((t['name'] for t in template_options if t['value'] == x), x),
            index=0,
            key="template_selector"
        )
       
        if selected_template != st.session_state.current_template:
            st.session_state.current_template = selected_template
            current_template = next((t for t in template_options if t['value'] == selected_template), None)
            if current_template:
                if selected_template == "none":
                    st.warning("⚠️ No template selected! Please choose a template.")
                else:
                    st.success(f"✅ {current_template['name']} template selected!")
                    # Set text color based on template type
                    if st.session_state.uploaded_template_file:
                        # Custom template - show color picker
                        st.session_state.template_text_color = st.color_picker(
                            "🎨 Choose Text Color for Custom Template",
                            value="#000000",
                            key="custom_template_color"
                        )
                    else:
                        # Sample template - use static black
                        st.session_state.template_text_color = "#000000"
       
        current_template = next((t for t in template_options if t['value'] == st.session_state.current_template), None)
        if current_template and st.session_state.current_template != "none":
            template_display_name = current_template['name']
            template_icon = current_template['icon']
            template_description = current_template['description']
           
            if st.session_state.uploaded_template_file:
                template_display_name = "Custom Template"
                template_icon = "📁"
                template_description = st.session_state.uploaded_template_file.name
           
            st.markdown(f"""
                <div class='metric-card pulse-animation'>
                    <div class='card-content'>
                        <div style='font-size: 24px; margin-bottom: 8px;'>{template_icon}</div>
                        <div class='card-title'>{template_display_name}</div>
                        <div class='card-description'>{template_description}</div>
                        <div class='card-description' style='color: {st.session_state.template_text_color}; font-weight: 800;'>
                            Text Color: {st.session_state.template_text_color}
                        </div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
   
    # Enhanced Navigation with orange-yellow theme
    selected = option_menu(
        "🧭 Navigation",
        ["📂 Upload", "🎛️ Preview & Customize", "⚡ Generate", "👁️ Slide Preview", "📥 Download"],
        icons=["cloud-upload", "sliders", "lightning", "eye", "download"],
        menu_icon="compass",
        default_index=0,
        styles={
            "container": {
                "padding": "15px!important",
                "background": "var(--card-background)",
                "border-radius": "16px",
                "border": "2px solid var(--glass-border)",
                "box-shadow": "0 8px 25px rgba(0, 0, 0, 0.1)",
                "transition": "all 0.3s ease"
            },
            "icon": {"font-size": "20px", "margin-right": "12px", "color": "var(--text-primary)"},
            "nav-link": {
                "font-size": "16px",
                "text-align": "left",
                "margin": "6px",
                "border-radius": "12px",
                "color": "var(--text-primary)",
                "transition": "all 0.4s ease",
                "padding": "14px 16px",
                "font-weight": "700",
                "background": "rgba(242, 111, 33, 0.1)",
                "border": "2px solid transparent",
                "box-shadow": "0 4px 15px rgba(0, 0, 0, 0.1)",
                "transform": "translateX(0)"
            },
            "nav-link-selected": {
                "background": "linear-gradient(135deg, #f26f21 0%, #ffa800 100%)",
                "color": "#ffffff",
                "box-shadow": "0 8px 25px rgba(242, 111, 33, 0.4), 0 6px 20px rgba(0, 0, 0, 0.1)",
                "transform": "translateX(6px) scale(1.04)",
                "border": "none"
            },
            "nav-link:hover": {
                "background": "rgba(242, 111, 33, 0.15)",
                "transform": "translateX(8px) scale(1.03)",
                "border-color": "#f26f21",
                "box-shadow": "0 8px 25px rgba(242, 111, 33, 0.3), 0 6px 20px rgba(0, 0, 0, 0.1)"
            },
        }
    )

# Main content with enhanced animations
st.markdown('<div class="float-animation">', unsafe_allow_html=True)
st.title("🚀 AI-Powered Excel → PPT Generator")
st.markdown("</div>", unsafe_allow_html=True)

# Show success popup if needed
show_success_popup()

# UPLOAD SECTION
if selected == "📂 Upload":
    st.subheader("📂 Upload your Files")
   
    try:
        health_response = requests.get(f"{BACKEND_URL}/health", timeout=5)
        if health_response.status_code != 200:
            st.error("⚠️ Backend server is not responding.")
    except:
        st.error("❌ Cannot connect to backend server.")
        if st.button("🔄 Retry Connection", use_container_width=True):
            st.rerun()

    col1, col2 = st.columns(2)
   
    with col1:
        uploaded_csv = st.file_uploader(
            "📑 Upload CSV or Excel file",
            type=['csv', 'xlsx', 'xls'],
            help="Upload your data file (CSV or Excel format) - NO SIZE LIMITS"
        )
        if uploaded_csv is not None:
            st.session_state.uploaded_csv_file = uploaded_csv
            st.success(f"✅ {uploaded_csv.name} ready for upload")
            file_size = len(uploaded_csv.getvalue()) / (1024 * 1024)
            st.info(f"📏 File size: {file_size:.2f} MB")

    with col2:
        uploaded_template = st.file_uploader(
            "🎨 Upload PPTX template (optional)",
            type=['pptx'],
            help="Upload a custom PowerPoint template (optional) - NO SIZE LIMITS"
        )
        if uploaded_template is not None:
            st.session_state.uploaded_template_file = uploaded_template
            st.success(f"✅ {uploaded_template.name} ready for upload")
            # Show color picker for custom template
            st.session_state.template_text_color = st.color_picker(
                "🎨 Choose Text Color for Custom Template",
                value="#000000",
                key="custom_template_color_upload"
            )

    if st.button("🔄 Clear / Reset", use_container_width=True):
        clear_temp()

    if st.session_state.uploaded_csv_file and st.session_state.backend_file_id is None:
        with st.spinner("📤 Uploading CSV to backend (this may take a while for large files)..."):
            file_id = upload_file_to_backend(st.session_state.uploaded_csv_file, "csv")
            if file_id:
                st.session_state.backend_file_id = file_id
                st.success("✅ CSV file uploaded successfully!")
               
                with st.spinner("🔍 Analyzing data (processing large files may take time)..."):
                    analysis_result = analyze_data_backend(file_id)
                    if analysis_result:
                        st.session_state.processed_data = analysis_result
                        if "sample_data" in analysis_result:
                            st.session_state.df = pd.DataFrame(analysis_result["sample_data"])
                        else:
                            if st.session_state.uploaded_csv_file.name.endswith(('.xlsx', '.xls')):
                                st.session_state.df = pd.read_excel(st.session_state.uploaded_csv_file)
                            else:
                                st.session_state.df = pd.read_csv(st.session_state.uploaded_csv_file)
                       
                        st.session_state.column_types = analyze_column_types(st.session_state.df)
                       
                        st.session_state.preview = {
                            "columns": analysis_result.get("columns", list(st.session_state.df.columns)),
                            "summary": analysis_result.get("summary", {})
                        }
                        st.session_state.analysis_complete = True
                        st.success("✅ Data analyzed successfully!")
                       
                        with st.expander("📊 Data Preview", expanded=True):
                            st.dataframe(st.session_state.df, use_container_width=True)
                            st.write(f"**Shape:** {st.session_state.df.shape[0]} rows × {st.session_state.df.shape[1]} columns")
                           
                            st.write("**📋 Column Types Analysis:**")
                            type_summary = {}
                            for col, info in st.session_state.column_types.items():
                                col_type = info['type']
                                type_summary[col_type] = type_summary.get(col_type, 0) + 1
                           
                            for col_type, count in type_summary.items():
                                st.write(f"• {col_type.title()}: {count} columns")

    if st.session_state.uploaded_template_file and st.session_state.template_file_id is None:
        with st.spinner("📤 Uploading template..."):
            template_file_id = upload_file_to_backend(st.session_state.uploaded_template_file, "template")
            if template_file_id:
                st.session_state.template_file_id = template_file_id
                st.success("✅ Template uploaded successfully!")

    if st.session_state.backend_file_id:
        st.markdown("---")
        st.subheader("✅ Upload Summary")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("""
                <div class='metric-card fade-in-up'>
                    <div class='card-content'>
                        <div class='card-title'>CSV File</div>
                        <div class='card-value'>Uploaded</div>
                        <div class='card-description' style='color: #48bb78;'>✓ Ready</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with col2:
            status = "Uploaded" if st.session_state.template_file_id else "Not Provided"
            color = "#48bb78" if st.session_state.template_file_id else "#a0aec0"
            st.markdown(f"""
                <div class='metric-card fade-in-up'>
                    <div class='card-content'>
                        <div class='card-title'>Template</div>
                        <div class='card-value'>{status}</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with col3:
            if st.session_state.uploaded_template_file:
                template_status = "Custom Template"
            elif st.session_state.current_template != "none":
                current_template = next((t for t in st.session_state.sample_templates if t['value'] == st.session_state.current_template), None)
                template_status = current_template['name'] if current_template else "Selected"
            else:
                template_status = "None Selected"
               
            st.markdown(f"""
                <div class='metric-card fade-in-up'>
                    <div class='card-content'>
                        <div class='card-title'>Template Type</div>
                        <div class='card-value' style='font-size: 18px;'>{template_status}</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)

# PREVIEW & CUSTOMIZE SECTION
elif selected == "🎛️ Preview & Customize":
    if st.session_state.preview is not None and st.session_state.analysis_complete:
        df = st.session_state.df

        st.subheader("🎛️ Presentation Customization Studio")
       
        st.session_state.presentation_title = st.text_input(
            "📝 Presentation Title",
            value="Data Analysis Presentation"
        )
       
        st.subheader("🚀 Select Presentation Features")
        st.info("✅ Check the features you want to include in your presentation")
       
        col1, col2, col3, col4, col5 = st.columns(5)
       
        with col1:
            dashboard_selected = st.checkbox(
                "📊 Dashboard",
                value=st.session_state.selected_features["dashboard"],
                help="Interactive dashboard with multiple charts"
            )
       
        with col2:
            charts_insights_selected = st.checkbox(
                "📈 Charts+Insights",
                value=st.session_state.selected_features["charts_insights"],
                help="Charts with AI-generated insights"
            )
       
        with col3:
            comparison_selected = st.checkbox(
                "⚖️ Comparison",
                value=st.session_state.selected_features["comparison"],
                help="Data comparisons across categories"
            )
       
        with col4:
            query_selected = st.checkbox(
                "🔍 Query",
                value=st.session_state.selected_features["query"],
                help="Custom data queries and analysis"
            )
       
        with col5:
            insights_selected = st.checkbox(
                "💡 Insights",
                value=st.session_state.selected_features["insights"],
                help="AI-generated data insights"
            )
       
        if st.button("✅ Confirm Feature Selection", type="primary", use_container_width=True):
            st.session_state.selected_features = {
                "dashboard": dashboard_selected,
                "charts_insights": charts_insights_selected,
                "comparison": comparison_selected,
                "query": query_selected,
                "insights": insights_selected
            }
            st.session_state.feature_selection_confirmed = True
            st.session_state.charts_preview_ready = False
            st.success("✅ Feature selection confirmed! Configure your charts below.")
            st.rerun()

        active_features = [f for f, active in st.session_state.selected_features.items() if active]
        if active_features:
            feature_names = {
                "dashboard": "📊 Dashboard",
                "charts_insights": "📈 Charts with Insights",
                "comparison": "⚖️ Comparisons",
                "query": "🔍 Custom Queries",
                "insights": "💡 AI Insights"
            }
            active_list = [feature_names[f] for f in active_features]
            st.markdown(f"""
                <div class='stSuccess fade-in-up'>
                    <h4 style='margin: 0;'>✅ Active Features</h4>
                    <p style='margin: 10px 0 0 0; font-weight: 600;'>{', '.join(active_list)}</p>
                </div>
            """, unsafe_allow_html=True)

        if st.session_state.feature_selection_confirmed and any(st.session_state.selected_features.values()):
            st.markdown("---")
            st.subheader("⚙️ Feature Configuration")
           
            if st.session_state.selected_features["dashboard"]:
                with st.expander("📊 Dashboard Settings", expanded=True):
                    st.info("🎯 Add up to 6 charts for your dashboard")
                    available_columns = df.columns.tolist()
                   
                    col1, col2, col3 = st.columns([2, 2, 2])
                    with col1:
                        new_column = st.selectbox("Select X-axis column", options=available_columns, key="db_col_x")
                    with col2:
                        valid_chart_types = get_valid_chart_types(new_column)
                        new_chart_type = st.selectbox("Chart type", options=valid_chart_types, key="db_type")
                    with col3:
                        y_axis_options = ["None"] + [col for col in available_columns if col != new_column]
                        second_col = st.selectbox("Y-axis column (optional)", options=y_axis_options, key="db_col_y")
                       
                        if second_col and second_col != "None":
                            valid_chart_types = get_valid_chart_types(new_column, second_col)
                            if new_chart_type not in valid_chart_types and valid_chart_types:
                                new_chart_type = valid_chart_types[0]
                   
                    if new_column in st.session_state.column_types:
                        col_info = st.session_state.column_types[new_column]
                        st.info(f"📋 **Column Type:** {col_info['type'].title()} | **Valid Charts:** {', '.join(col_info['chart_types'])}")
                   
                    col4, col5 = st.columns([1, 1])
                    with col4:
                        if st.button("➕ Add Chart", key="add_db", use_container_width=True):
                            if len(st.session_state.dashboard_config["charts"]) < 6:
                                chart_config = {"column": new_column, "type": new_chart_type}
                               
                                if second_col and second_col != "None":
                                    chart_config["second_column"] = second_col
                               
                                st.session_state.dashboard_config["charts"].append(chart_config)
                                st.success(f"✅ Added {new_chart_type} chart for {new_column}")
                            else:
                                st.warning("🚫 Maximum 6 charts allowed for dashboard")
                    with col5:
                        if st.button("🎯 Preview All", key="preview_all_db", use_container_width=True):
                            st.session_state.charts_preview_ready = True
                   
                    if st.session_state.dashboard_config["charts"]:
                        st.subheader("📋 Dashboard Charts")
                        for i, chart in enumerate(st.session_state.dashboard_config["charts"]):
                            col1, col2, col3 = st.columns([4, 1, 1])
                            with col1:
                                if "second_column" in chart and chart["second_column"] != "None":
                                    st.write(f"**{i+1}. {chart['column']} (X) + {chart['second_column']} (Y)** - *{chart['type']}*")
                                else:
                                    st.write(f"**{i+1}. {chart['column']}** - *{chart['type']}*")
                            with col2:
                                if st.button("👁️ Preview", key=f"preview_db_{i}", use_container_width=True):
                                    st.session_state.charts_preview_ready = True
                            with col3:
                                if st.button("🗑️ Remove", key=f"rm_db_{i}", use_container_width=True):
                                    st.session_state.dashboard_config["charts"].pop(i)
                                    st.rerun()

            if st.session_state.selected_features["charts_insights"]:
                with st.expander("📈 Charts + Insights Settings", expanded=True):
                    available_columns = df.columns.tolist()
                   
                    col1, col2, col3 = st.columns([2, 2, 2])
                    with col1:
                        chart_column = st.selectbox("Select X-axis column", options=available_columns, key="ci_col_x")
                    with col2:
                        valid_chart_types = get_valid_chart_types(chart_column)
                        chart_type = st.selectbox("Chart type", options=valid_chart_types, key="ci_type")
                    with col3:
                        y_axis_options = ["None"] + [col for col in available_columns if col != chart_column]
                        second_col = st.selectbox("Y-axis column (optional)", options=y_axis_options, key="ci_col_y")
                       
                        if second_col and second_col != "None":
                            valid_chart_types = get_valid_chart_types(chart_column, second_col)
                            if chart_type not in valid_chart_types and valid_chart_types:
                                chart_type = valid_chart_types[0]
                   
                    if chart_column in st.session_state.column_types:
                        col_info = st.session_state.column_types[chart_column]
                        st.info(f"📋 **Column Type:** {col_info['type'].title()} | **Valid Charts:** {', '.join(col_info['chart_types'])}")
                   
                    col4, col5 = st.columns([1, 1])
                    with col4:
                        if st.button("➕ Add Chart", key="add_ci", use_container_width=True):
                            chart_config = {"column": chart_column, "type": chart_type}
                           
                            if second_col and second_col != "None":
                                chart_config["second_column"] = second_col
                           
                            st.session_state.insight_config["charts"].append(chart_config)
                            st.success(f"✅ Added {chart_type} chart for {chart_column}")
                    with col5:
                        if st.button("🎯 Preview All", key="preview_all_ci", use_container_width=True):
                            st.session_state.charts_preview_ready = True
                   
                    if st.session_state.insight_config.get("charts"):
                        st.subheader("📋 Charts with Insights")
                        for i, chart in enumerate(st.session_state.insight_config["charts"]):
                            col1, col2, col3 = st.columns([4, 1, 1])
                            with col1:
                                if "second_column" in chart and chart["second_column"] != "None":
                                    st.write(f"**{i+1}. {chart['column']} (X) + {chart['second_column']} (Y)** - *{chart['type']}*")
                                else:
                                    st.write(f"**{i+1}. {chart['column']}** - *{chart['type']}*")
                            with col2:
                                if st.button("👁️ Preview", key=f"preview_ci_{i}", use_container_width=True):
                                    st.session_state.charts_preview_ready = True
                            with col3:
                                if st.button("🗑️ Remove", key=f"rm_ci_{i}", use_container_width=True):
                                    st.session_state.insight_config["charts"].pop(i)
                                    st.rerun()

            if st.session_state.selected_features["insights"]:
                with st.expander("💡 Insights Settings", expanded=True):
                    st.info("✅ Select columns for AI-generated insights")
                    available_columns = df.columns.tolist()
                   
                    selected_columns = st.multiselect(
                        "Select columns for insights:",
                        options=available_columns,
                        default=st.session_state.insight_config.get("selected_columns", []),
                        key="insight_columns"
                    )
                    st.session_state.insight_config["selected_columns"] = selected_columns
                   
                    if selected_columns:
                        st.success(f"✅ {len(selected_columns)} columns selected for AI insights")

            if st.session_state.selected_features["comparison"]:
                with st.expander("⚖️ Comparison Settings", expanded=True):
                    numeric_cols = [col for col in df.columns if pd.api.types.is_numeric_dtype(df[col])]
                    categorical_cols = [col for col in df.columns if not pd.api.types.is_numeric_dtype(df[col])]
                   
                    if numeric_cols and categorical_cols:
                        col1, col2, col3, col4 = st.columns([2, 2, 1, 1])
                        with col1:
                            compare_col = st.selectbox("Column to compare", options=numeric_cols, key="comp_col")
                        with col2:
                            group_col = st.selectbox("Group by", options=categorical_cols, key="group_col")
                        with col3:
                            valid_comp_charts = get_valid_chart_types(compare_col, group_col)
                            chart_type = st.selectbox("Chart type", options=valid_comp_charts, key="comp_type")
                        with col4:
                            if st.button("➕ Add", key="add_comp", use_container_width=True):
                                st.session_state.comparison_config["comparisons"].append({
                                    "compare": compare_col, "group_by": group_col, "chart_type": chart_type
                                })
                                st.success("✅ Comparison added!")
                   
                    if st.session_state.comparison_config.get("comparisons"):
                        st.subheader("📋 Comparisons")
                        for i, comp in enumerate(st.session_state.comparison_config["comparisons"]):
                            col1, col2, col3 = st.columns([4, 1, 1])
                            with col1:
                                st.write(f"**{i+1}. Compare {comp['compare']} by {comp['group_by']}**")
                            with col2:
                                if st.button("👁️ Preview", key=f"preview_comp_{i}", use_container_width=True):
                                    st.session_state.charts_preview_ready = True
                            with col3:
                                if st.button("🗑️ Remove", key=f"rm_comp_{i}", use_container_width=True):
                                    st.session_state.comparison_config["comparisons"].pop(i)
                                    st.rerun()
                       
                        if st.button("🎯 Preview All Comparisons", key="preview_all_comp", use_container_width=True):
                            st.session_state.charts_preview_ready = True

            if st.session_state.selected_features["query"]:
                with st.expander("🔍 Query Settings", expanded=True):
                    include_in_slides = st.toggle(
                        "📑 Include in Presentation",
                        value=True,
                        help="Include this query in the final presentation"
                    )
                   
                    query_text = st.text_area("Enter your query:", placeholder="e.g., Show top 10 products by sales", key="query_text_area")
                   
                    # FIXED: Query button alignment with static size
                    st.markdown('<div class="query-button-container">', unsafe_allow_html=True)
                    col3, col4 = st.columns([3, 1])
                    with col3:
                        if st.button("➕ Add Query", key="add_query", use_container_width=True, 
                                   type="primary" if query_text.strip() else "secondary") and query_text.strip():
                            if "queries" not in st.session_state.custom_query:
                                st.session_state.custom_query["queries"] = []
                            st.session_state.custom_query["queries"].append({
                                "text": query_text.strip(),
                                "include_in_slides": include_in_slides
                            })
                            st.success("✅ Query added!")
                    st.markdown('</div>', unsafe_allow_html=True)
                   
                    if st.session_state.custom_query.get("queries"):
                        st.subheader("📋 Custom Queries")
                        for i, query in enumerate(st.session_state.custom_query["queries"]):
                            col1, col2, col3 = st.columns([4, 1, 1])
                            with col1:
                                st.write(f"**{i+1}. {query['text']}**")
                                st.caption(f"Include in Slides: {'Yes' if query['include_in_slides'] else 'No'}")
                            with col2:
                                if st.button("✏️ Edit", key=f"edit_query_{i}", use_container_width=True, type="secondary"):
                                    # Edit functionality placeholder
                                    pass
                            with col3:
                                if st.button("🗑️ Remove", key=f"rm_query_{i}", use_container_width=True, type="secondary"):
                                    st.session_state.custom_query["queries"].pop(i)
                                    st.rerun()

        if st.session_state.charts_preview_ready:
            st.markdown("---")
           
            preview_data = []
           
            if st.session_state.selected_features["dashboard"] and st.session_state.dashboard_config.get("charts"):
                preview_data.extend([(chart, "dashboard") for chart in st.session_state.dashboard_config["charts"]])
           
            if st.session_state.selected_features["charts_insights"] and st.session_state.insight_config.get("charts"):
                preview_data.extend([(chart, "charts_insights") for chart in st.session_state.insight_config["charts"]])
           
            if st.session_state.selected_features["comparison"] and st.session_state.comparison_config.get("comparisons"):
                preview_data.extend([(chart, "comparison") for chart in st.session_state.comparison_config.get("comparisons")])
           
            if preview_data:
                # Use the SIMPLIFIED ADVANCED charts display function
                display_advanced_charts_grid([data[0] for data in preview_data], preview_data[0][1] if preview_data else "dashboard")
            else:
                st.info("No charts configured for preview.")

        st.markdown("---")
        st.subheader("💾 Save Configuration")
       
        if st.button("💾 Save Configuration to Backend", type="primary", use_container_width=True):
            if st.session_state.current_template == "none" and not st.session_state.uploaded_template_file:
                st.error("❌ Please select a template or upload a custom template!")
            elif st.session_state.backend_file_id:
                with st.spinner("Saving configuration..."):
                    config = generate_presentation_config()
                    if config:
                        result = save_config_backend(config, st.session_state.backend_file_id)
                       
                        if result and result.get("success"):
                            st.session_state.presentation_config = config
                            st.session_state.config_saved = True
                            st.session_state.edit_mode = False
                           
                            st.markdown("""
                                <div style="text-align: center; padding: 30px;">
                                    <div style="font-size: 64px; animation: fadeInUp 1s ease-in;">✨</div>
                                    <h3 style="color: #48bb78; margin: 20px 0;">Configuration Saved Successfully!</h3>
                                    <p style="color: var(--text-muted); font-weight: 600;">Your presentation settings have been saved and are ready for generation.</p>
                                </div>
                            """, unsafe_allow_html=True)
                        else:
                            st.error("❌ Failed to save configuration")
            else:
                st.error("❌ Please upload a CSV file first")

    else:
        st.info("📂 Please upload a CSV file first in the Upload section")

# GENERATE SECTION
elif selected == "⚡ Generate":
    st.subheader("⚡ Generate PowerPoint Presentation")
   
    if st.session_state.preview is not None and st.session_state.analysis_complete:
       
        if not st.session_state.config_saved:
            st.warning("⚠️ Please save your configuration first.")
            st.info("Go to 'Preview & Customize' to configure and save your presentation settings.")
       
        # Display any existing generation errors
        display_generation_errors()
       
        if st.session_state.config_saved:
            st.success("✅ Configuration ready for generation!")
           
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                active_count = sum(st.session_state.selected_features.values())
                st.markdown(f"""
                    <div class='metric-card glow-animation'>
                        <div class='card-content'>
                            <div class='card-title'>🚀 Active Features</div>
                            <div class='card-value'>{active_count}</div>
                            <div class='card-description' style='color: #48bb78;'>Selected</div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
            with col2:
                chart_count = len(st.session_state.dashboard_config.get("charts", [])) + \
                             len(st.session_state.insight_config.get("charts", [])) + \
                             len(st.session_state.comparison_config.get("comparisons", []))
                st.markdown(f"""
                    <div class='metric-card glow-animation'>
                        <div class='card-content'>
                            <div class='card-title'>📈 Total Charts</div>
                            <div class='card-value'>{chart_count}</div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
            with col3:
                if st.session_state.uploaded_template_file:
                    template_name = "Custom Template"
                elif st.session_state.current_template != "none":
                    current_template = next((t for t in st.session_state.sample_templates if t['value'] == st.session_state.current_template), None)
                    template_name = current_template['name'] if current_template else st.session_state.current_template
                else:
                    template_name = "No Template"
                   
                st.markdown(f"""
                    <div class='metric-card glow-animation'>
                        <div class='card-content'>
                            <div class='card-title'>🎭 Template</div>
                            <div class='card-value' style='font-size: 20px;'>{template_name}</div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
            with col4:
                st.markdown("""
                    <div class='metric-card glow-animation'>
                        <div class='card-content'>
                            <div class='card-title'>💾 Status</div>
                            <div class='card-value'>Ready</div>
                            <div class='card-description' style='color: #48bb78;'>✓</div>
                        </div>
                    </div>
                """, unsafe_allow_html=True)
       
        if st.button("🎯 Generate Presentation", use_container_width=True, type="primary",
                    disabled=not st.session_state.config_saved):
            if st.session_state.backend_file_id and st.session_state.config_saved:
                try:
                    # Clear previous errors before generation
                    st.session_state.generation_errors = []
                   
                    with st.status("🚀 Generating presentation...", expanded=True) as status:
                        st.write("📋 Preparing configuration...")
                        config = {
                            "csv_file_id": st.session_state.backend_file_id,
                            "presentation_config": st.session_state.presentation_config,
                            "options": {
                                "use_ai_content": True,
                                "template_style": st.session_state.current_template
                            }
                        }
                       
                        if st.session_state.template_file_id:
                            config["template_file_id"] = st.session_state.template_file_id
                       
                        st.write("⚡ Generating PowerPoint...")
                        generation_result = generate_ppt_backend(config)
                       
                        if generation_result and generation_result.get("success"):
                            st.write("📥 Downloading presentation...")
                            ppt_id = generation_result["ppt_id"]
                            ppt_bytes = download_ppt_backend(ppt_id)
                           
                            if ppt_bytes:
                                st.write("🖼️ Loading actual slide images...")
                                slide_images_result = get_slide_images_backend(ppt_id)
                                if slide_images_result and slide_images_result.get("success"):
                                    st.session_state.slide_images = slide_images_result["slide_images"]
                                    st.session_state.slide_images_loaded = True
                                    st.write(f"✅ Loaded {len(st.session_state.slide_images)} actual slide images")
                                else:
                                    st.warning("Could not load actual slide images, will use text preview")
                               
                                status.update(label="✅ Presentation generated successfully!", state="complete")
                               
                                history_entry = {
                                    "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                    "filename": f"presentation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx",
                                    "features": sum(st.session_state.selected_features.values()),
                                    "charts": len(st.session_state.dashboard_config.get("charts", [])) + \
                                             len(st.session_state.insight_config.get("charts", [])) + \
                                             len(st.session_state.comparison_config.get("comparisons", [])),
                                    "ppt_id": ppt_id
                                }
                                st.session_state.generation_history.append(history_entry)
                                st.session_state.presentation_generated = True
                                st.session_state.generated_ppt_id = ppt_id
                                st.session_state.slide_preview_data = create_slide_preview_data()
                               
                                # SET SUCCESS POPUP STATE
                                st.session_state.show_success_popup = True
                               
                                st.rerun()
                            else:
                                st.error("❌ Failed to download presentation")
                        else:
                            st.error("❌ Failed to generate presentation")
                           
                except Exception as e:
                    st.error(f"❌ Error generating PPT: {e}")
            else:
                st.error("❌ Configuration not saved or backend not connected")
    else:
        st.info("📂 Please upload and analyze data first.")

# SLIDE PREVIEW SECTION - ENHANCED WITH ACTUAL SLIDE IMAGES AND EDIT OPTIONS
elif selected == "👁️ Slide Preview":
    st.subheader("👁️ Slide Preview - Actual PowerPoint Slides")
   
    # Display generation errors at the top
    display_generation_errors()
   
    if st.session_state.presentation_generated and st.session_state.generated_ppt_id:
        st.success("✅ Presentation generated successfully! Review your actual slides below.")
       
        if not st.session_state.slide_images_loaded:
            with st.spinner("🖼️ Loading actual slide images..."):
                slide_images_result = get_slide_images_backend(st.session_state.generated_ppt_id)
                if slide_images_result and slide_images_result.get("success"):
                    st.session_state.slide_images = slide_images_result["slide_images"]
                    st.session_state.slide_images_loaded = True
                    st.rerun()
                else:
                    st.warning("⚠️ Could not load actual slide images. Showing text preview instead.")
       
        if st.session_state.slide_images_loaded and st.session_state.slide_images:
            # SHOW ACTUAL SLIDE IMAGES (REAL GENERATED SLIDES)
            total_slides = len(st.session_state.slide_images)
            current_slide_index = st.session_state.get('current_slide', 0)
            current_slide = st.session_state.slide_images[current_slide_index]
           
            st.markdown("### 📋 Slide Navigation")
           
            # Navigation buttons with enhanced styling
            cols = st.columns(min(6, total_slides))
            for i in range(total_slides):
                with cols[i % len(cols)]:
                    slide_num = i + 1
                    is_current = (i == current_slide_index)
                   
                    if st.button(f"Slide {slide_num}", key=f"nav_{i}", use_container_width=True):
                        st.session_state.current_slide = i
                        st.rerun()
           
            # Display actual slide image with enhanced styling
            st.markdown(f"""
                <div class='slide-image-container image-zoom'>
                    <div style='text-align: center; margin-bottom: 20px;'>
                        <h2 style='color: #ffa800; margin: 0; font-weight: 700;'>Slide {current_slide_index + 1}</h2>
                        <p style='color: var(--text-muted); font-weight: 600; font-size: 16px;'>Slide {current_slide_index + 1} of {total_slides}</p>
                    </div>
                </div>
            """, unsafe_allow_html=True)
           
            # Display the actual slide image
            col1, col2, col3 = st.columns([1, 3, 1])
            with col2:
                st.markdown('<div class="content-slide">', unsafe_allow_html=True)
               
                # Display the actual generated slide image
                if 'image_data' in current_slide:
                    try:
                        st.image(
                            current_slide['image_data'],
                            use_column_width=True,
                            caption=f"🎯 Slide {current_slide_index + 1} - Actual Generated Slide",
                            output_format="PNG"
                        )
                    except:
                        st.image(
                            current_slide['image_data'],
                            use_column_width=True,
                            caption=f"🎯 Slide {current_slide_index + 1} - Actual Generated Slide"
                        )
                else:
                    st.warning("No image data available for this slide")
               
                st.markdown('</div>', unsafe_allow_html=True)
           
            # Display slide content if available
            if current_slide.get('content'):
                with st.expander("📝 Slide Content", expanded=False):
                    st.markdown('<div class="content-slide">', unsafe_allow_html=True)
                    for content_line in current_slide['content']:
                        if content_line.strip():
                            st.markdown(f'<div class="slide-content-text">{content_line}</div>', unsafe_allow_html=True)
                    st.markdown('</div>', unsafe_allow_html=True)
           
            # Enhanced Slide navigation controls with NEXT/BACK buttons
            st.markdown("---")
            st.subheader("🧭 Navigation Controls")
           
            col_nav1, col_nav2, col_nav3, col_nav4 = st.columns([1, 2, 1, 1])
            with col_nav1:
                if current_slide_index > 0:
                    if st.button("⬅️ Previous", use_container_width=True, key="prev_slide"):
                        st.session_state.current_slide = current_slide_index - 1
                        st.rerun()
                else:
                    st.button("⬅️ Previous", disabled=True, use_container_width=True)
           
            with col_nav2:
                st.info(f"📊 Slide {current_slide_index + 1} of {total_slides} (Actual Generated Slide)")
           
            with col_nav3:
                if current_slide_index < total_slides - 1:
                    if st.button("Next ➡️", use_container_width=True, key="next_slide"):
                        st.session_state.current_slide = current_slide_index + 1
                        st.rerun()
                else:
                    st.button("Next ➡️", disabled=True, use_container_width=True)
           
            with col_nav4:
                if st.button("🔄 Reload Images", use_container_width=True):
                    st.session_state.slide_images_loaded = False
                    st.session_state.slide_images = []
                    st.rerun()
       
        else:
            # FALLBACK: Show text preview
            st.warning("⚠️ Actual slide images not available. Showing text preview.")
            if st.session_state.slide_preview_data:
                total_slides = len(st.session_state.slide_preview_data)
                current_slide_index = st.session_state.get('current_slide', 0)
                slide_data = st.session_state.slide_preview_data[current_slide_index]
               
                st.markdown("### 📋 Slide Navigation")
                cols = st.columns(min(6, total_slides))
                for i, slide_data_nav in enumerate(st.session_state.slide_preview_data):
                    with cols[i % len(cols)]:
                        is_current = (i == current_slide_index)
                        button_label = f"📍 {slide_data_nav['title']}" if is_current else slide_data_nav['title']
                       
                        if st.button(button_label, key=f"nav_{i}", use_container_width=True):
                            st.session_state.current_slide = i
                            st.rerun()
                                                                       
                st.markdown(f"""
                    <div class='slide-preview fade-in-up'>
                        <div style='text-align: center; margin-bottom: 20px;'>
                            <h2 style='color: #ffa800; margin: 0; font-weight: 700;'>{slide_data['title']}</h2>
                            <p style='color: var(--text-muted); font-weight: 600;'>Slide {slide_data['slide_number']} of {total_slides}</p>
                        </div>
                """, unsafe_allow_html=True)
               
                for content in slide_data['content']:
                    st.markdown(f'<div class="slide-content-text">{content}</div>', unsafe_allow_html=True)
               
                st.markdown("</div>", unsafe_allow_html=True)
            else:
                st.info("No preview data available.")
       
        # Regenerate presentation option
        st.markdown("---")
        st.subheader("🔄 Regenerate Presentation")
        st.info("If you made edits or want to regenerate with different settings:")
       
        if st.button("🔄 Regenerate PPT", type="secondary", use_container_width=True):
            st.session_state.presentation_generated = False
            st.session_state.current_slide = 0
            st.session_state.slide_images_loaded = False
            st.session_state.slide_images = []
            st.rerun()
   
    elif st.session_state.config_saved and st.session_state.presentation_config:
        st.info("📋 Your presentation configuration is saved. Generate the presentation to preview actual slides.")
        if st.button("⚡ Generate Presentation First", use_container_width=True):
            st.session_state.current_slide = 0
            st.rerun()
    else:
        st.info("💡 No presentation generated yet. Configure and generate your presentation first.")

# DOWNLOAD SECTION
elif selected == "📥 Download":
    st.subheader("📥 Download Your Presentation")  
    if st.session_state.presentation_generated and st.session_state.generated_ppt_id:
        st.success("✅ Your presentation is ready for download!")
       
        st.markdown("""
            <div class='download-section fade-in-up'>
                <div style='text-align: center; padding: 20px;'>
                    <h2 style='color: #48bb78; margin-bottom: 15px;'>🎉 Presentation Ready!</h2>
                    <p style='color: var(--text-secondary); font-size: 18px; font-weight: 600; margin-bottom: 25px;'>
                        Your PowerPoint has been successfully generated and is ready for download.
                    </p>
                </div>
            </div>
        """, unsafe_allow_html=True)
       
        ppt_bytes = download_ppt_backend(st.session_state.generated_ppt_id)
        if ppt_bytes:
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.download_button(
                    "📥 Download Presentation",
                    data=ppt_bytes,
                    file_name=f"presentation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True,
                    type="primary"
                )
       
        st.markdown("---")
        st.subheader("📊 Presentation Details")
       
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            total_slides = len(st.session_state.slide_images) if st.session_state.slide_images_loaded else len(st.session_state.slide_preview_data)
            st.markdown(f"""
                <div class='metric-card fade-in-up'>
                    <div class='card-content'>
                        <div class='card-title'>Total Slides</div>
                        <div class='card-value'>{total_slides}</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with col2:
            active_count = sum(st.session_state.selected_features.values())
            st.markdown(f"""
                <div class='metric-card fade-in-up'>
                    <div class='card-content'>
                        <div class='card-title'>Active Features</div>
                        <div class='card-value'>{active_count}</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with col3:
            chart_count = len(st.session_state.dashboard_config.get("charts", [])) + \
                         len(st.session_state.insight_config.get("charts", [])) + \
                         len(st.session_state.comparison_config.get("comparisons", []))
            st.markdown(f"""
                <div class='metric-card fade-in-up'>
                    <div class='card-content'>
                        <div class='card-title'>Total Charts</div>
                        <div class='card-value'>{chart_count}</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
        with col4:
            st.markdown(f"""
                <div class='metric-card fade-in-up'>
                    <div class='card-content'>
                        <div class='card-title'>Generated</div>
                        <div class='card-value' style='font-size: 20px;'>{datetime.now().strftime("%H:%M")}</div>
                    </div>
                </div>
            """, unsafe_allow_html=True)
       
        st.markdown("---")
        st.subheader("🔄 Need Changes?")
        st.info("If you need to make changes, you can go back to Preview & Customize section.")
       
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("✏️ Back to Customization", use_container_width=True, type="secondary"):
                st.session_state.current_slide = 0
                st.rerun()
   
    else:
        st.info("📋 No presentation generated yet. Please generate your presentation first.")
       
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.session_state.config_saved:
                if st.button("⚡ Generate Presentation", use_container_width=True):
                    st.session_state.current_slide = 0
                    st.rerun()
            else:
                if st.button("🎛️ Configure Presentation", use_container_width=True):
                    st.session_state.current_slide = 0
                    st.rerun()

# Footer with enhanced styling
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; padding: 30px; background: var(--card-background); border-radius: 20px; margin: 20px 0; border: 2px solid var(--glass-border); box-shadow: 0 15px 35px rgba(0, 0, 0, 0.1);'>
        <div style='font-size: 24px; color: #ffa800; margin-bottom: 10px; font-weight: 800;'>✨ AI-Powered Excel to PPT Generator</div>
        <div style='color: var(--text-muted); font-weight: 600;'>Built with Streamlit & FastAPI | Professional Presentation Automation</div>
    </div>
    """,
    unsafe_allow_html=True
)
# Re-load CSS when styles change
load_css()