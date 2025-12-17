import json
import os
import random
import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from collections import Counter
import re


class HotelBookingVisualization:
    def __init__(self, dataset_path="latest_file",output_path = "outputs_dir"):
        self.color_palette = [
            '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
            '#8c564b', "#b81f8a", '#7f7f7f', '#bcbd22', '#17becf',
            '#aec7e8', '#ffbb78', "#41a82d", '#ff9896', '#c5b0d5',
            '#c49c94', '#f7b6d2', "#c5b3b3", '#dbdb8d', '#9edae5'
        ]
        self.used_colors = set()
        self.output_dir = output_path
        self.dataset_path = dataset_path
        self.dataset = None
        
        # Chart compatibility mapping
        self.chart_compatibility = {
            'numerical': ['histogram', 'box', 'violin', 'scatter', 'line', 'bar'],
            'categorical': ['bar', 'pie', 'sunburst', 'treemap'],
            'date': ['line', 'bar', 'area', 'scatter']
        }
        
        # Preferred fallback charts by data type
        self.fallback_charts = {
            'numerical': ['histogram', 'box', 'violin'],
            'categorical': ['bar', 'pie'],
            'date': ['line', 'bar']
        }
        
        self.setup_directories()
        self.load_dataset()

    def setup_directories(self):
        """Create necessary output directories"""
        os.makedirs(os.path.join(self.output_dir, "charts"), exist_ok=True)

    def load_dataset(self):
        """Load the hotel booking dataset"""
        if os.path.exists(self.dataset_path):
            self.dataset = pd.read_csv(self.dataset_path)
            print(f"✓ Dataset loaded: {self.dataset.shape}")
        else:
            # Create sample data if file doesn't exist
            print("⚠ Dataset not found, creating sample data...")
            self.create_sample_data()

    def create_sample_data(self):
        """Create sample hotel booking data for testing"""
        np.random.seed(42)
        n_rows = 1000
        
        sample_data = {
            'hotel': np.random.choice(['City Hotel', 'Resort Hotel'], n_rows),
            'lead_time': np.random.randint(0, 365, n_rows),
            'arrival_date': pd.date_range('2023-01-01', periods=n_rows, freq='D'),
            'adr': np.random.normal(120, 50, n_rows).clip(50, 300),
            'adults': np.random.randint(1, 4, n_rows),
            'children': np.random.randint(0, 3, n_rows),
            'country': np.random.choice(['PRT', 'GBR', 'FRA', 'ESP', 'DEU', 'ITA'], n_rows),
            'customer_type': np.random.choice(['Transient', 'Contract', 'Group', 'Transient-Party'], n_rows),
            'is_canceled': np.random.choice([0, 1], n_rows, p=[0.7, 0.3]),
            'required_car_parking_spaces': np.random.randint(0, 2, n_rows),
            'total_of_special_requests': np.random.randint(0, 5, n_rows),
            'reservation_status': np.random.choice(['Check-Out', 'Canceled', 'No-Show'], n_rows),
            'stays_in_week_nights': np.random.randint(1, 8, n_rows),
            'stays_in_weekend_nights': np.random.randint(0, 3, n_rows)
        }
        
        self.dataset = pd.DataFrame(sample_data)
        print("✓ Sample data created successfully")

    def get_unique_color(self):
        """Get a unique color from the palette"""
        available_colors = [c for c in self.color_palette if c not in self.used_colors]
        if not available_colors:
            self.used_colors = set()
            available_colors = self.color_palette.copy()
        color = random.choice(available_colors)
        self.used_colors.add(color)
        return color

    def detect_data_type(self, series):
        """Detect the data type of a column with enhanced logic"""
        if series.isnull().all():
            return "unknown"
        
        # Check for datetime
        try:
            pd.to_datetime(series.dropna().head(10))
            return "date"
        except:
            pass
        
        # Check for numeric
        if self.is_numeric_column(series):
            return "numerical"
        
        # Check for categorical (limited unique values)
        unique_ratio = series.nunique() / len(series.dropna())
        if unique_ratio < 0.1 or series.nunique() <= 20:
            return "categorical"
        
        # Default to categorical for text data
        return "categorical"

    def find_best_column_match(self, metric_name):
        """Find the best matching column in dataset"""
        if self.dataset is None:
            return None
            
        metric_lower = metric_name.lower()
        
        # Direct match
        for col in self.dataset.columns:
            if col.lower() == metric_lower:
                return col
        
        # Partial match
        for col in self.dataset.columns:
            col_lower = col.lower()
            if metric_lower in col_lower or col_lower in metric_lower:
                return col
        
        # Word-based match
        words = metric_lower.split('_')
        for col in self.dataset.columns:
            col_lower = col.lower()
            if any(word in col_lower for word in words if len(word) > 2):
                return col
        
        # Return first column as fallback
        return self.dataset.columns[0] if len(self.dataset.columns) > 0 else None

    def is_numeric_column(self, series):
        """Check if column contains numeric data"""
        try:
            pd.to_numeric(series)
            return True
        except (ValueError, TypeError):
            return False

    def assess_data_distribution(self, data, column_name, data_type):
        """Assess data distribution to determine best chart type"""
        assessment = {
            "recommended_charts": [],
            "distribution_type": "unknown",
            "unique_values": 0,
            "data_points": 0,
            "reasoning": []
        }
        
        try:
            clean_data = data[column_name].dropna()
            assessment["data_points"] = len(clean_data)
            
            if data_type == "numerical":
                numeric_data = pd.to_numeric(clean_data, errors='coerce').dropna()
                assessment["unique_values"] = len(numeric_data.unique())
                
                # Assess distribution
                if len(numeric_data) > 0:
                    skewness = numeric_data.skew()
                    
                    if abs(skewness) > 1:
                        assessment["distribution_type"] = "highly_skewed"
                        assessment["reasoning"].append(f"Data is highly skewed (skewness: {skewness:.2f})")
                        assessment["recommended_charts"] = ["box", "violin", "histogram"]
                    elif abs(skewness) > 0.5:
                        assessment["distribution_type"] = "moderately_skewed"
                        assessment["reasoning"].append(f"Data is moderately skewed (skewness: {skewness:.2f})")
                        assessment["recommended_charts"] = ["histogram", "box", "violin"]
                    else:
                        assessment["distribution_type"] = "normal"
                        assessment["reasoning"].append(f"Data is approximately normal (skewness: {skewness:.2f})")
                        assessment["recommended_charts"] = ["histogram", "violin", "box"]
                    
                    # Add density plot for large datasets
                    if len(numeric_data) > 100:
                        assessment["recommended_charts"].append("density")
                        
            elif data_type == "categorical":
                assessment["unique_values"] = len(clean_data.unique())
                
                if assessment["unique_values"] <= 5:
                    assessment["distribution_type"] = "few_categories"
                    assessment["reasoning"].append(f"Only {assessment['unique_values']} categories - good for pie/sunburst")
                    assessment["recommended_charts"] = ["pie", "bar", "sunburst"]
                elif assessment["unique_values"] <= 15:
                    assessment["distribution_type"] = "moderate_categories"
                    assessment["reasoning"].append(f"Moderate number of categories ({assessment['unique_values']})")
                    assessment["recommended_charts"] = ["bar", "treemap", "pie"]
                else:
                    assessment["distribution_type"] = "many_categories"
                    assessment["reasoning"].append(f"Many categories ({assessment['unique_values']}) - bar chart recommended")
                    assessment["recommended_charts"] = ["bar", "treemap"]
                    
            elif data_type == "date":
                assessment["recommended_charts"] = ["line", "bar", "area"]
                assessment["reasoning"].append("Date data - line/bar charts work best")
                
        except Exception as e:
            assessment["reasoning"].append(f"Assessment error: {str(e)}")
            
        return assessment

    def select_appropriate_chart(self, requested_chart, data_type, distribution_assessment):
        """Select the most appropriate chart type with fallback logic"""
        # Check if requested chart is compatible
        compatible_charts = self.chart_compatibility.get(data_type, [])
        
        if requested_chart in compatible_charts:
            return {
                "selected_chart": requested_chart,
                "was_fallback": False,
                "reason": "Requested chart type is compatible with data type"
            }
        
        # Fallback logic
        fallback_options = self.fallback_charts.get(data_type, ['bar'])
        
        # Prioritize distribution-based recommendations
        recommended_charts = distribution_assessment.get("recommended_charts", [])
        
        # Find the best fallback
        for chart in recommended_charts:
            if chart in compatible_charts:
                return {
                    "selected_chart": chart,
                    "was_fallback": True,
                    "reason": f"Requested chart '{requested_chart}' not suitable. Selected '{chart}' based on data distribution: {distribution_assessment['distribution_type']}"
                }
        
        # Use general fallback
        for chart in fallback_options:
            if chart in compatible_charts:
                return {
                    "selected_chart": chart,
                    "was_fallback": True,
                    "reason": f"Requested chart '{requested_chart}' not suitable. Selected '{chart}' as general fallback for {data_type} data"
                }
        
        # Ultimate fallback
        ultimate_fallback = "bar" if data_type in ["categorical", "numerical"] else "line"
        return {
            "selected_chart": ultimate_fallback,
            "was_fallback": True,
            "reason": f"Requested chart '{requested_chart}' not suitable. Selected '{ultimate_fallback}' as ultimate fallback"
        }

    def generate_numerical_insights(self, data, column_name):
        """Generate comprehensive insights for numerical columns"""
        try:
            numeric_data = pd.to_numeric(data[column_name], errors='coerce').dropna()
            
            if len(numeric_data) == 0:
                return {"error": "No valid numerical data found"}
            
            # Basic statistics
            stats = {
                "mean": round(float(numeric_data.mean()), 2),
                "median": round(float(numeric_data.median()), 2),
                "std_dev": round(float(numeric_data.std()), 2),
                "min": round(float(numeric_data.min()), 2),
                "max": round(float(numeric_data.max()), 2),
                "range": round(float(numeric_data.max() - numeric_data.min()), 2)
            }
            
            # Quartiles
            q1 = float(numeric_data.quantile(0.25))
            q3 = float(numeric_data.quantile(0.75))
            iqr = q3 - q1
            
            # Outlier analysis
            lower_bound = q1 - 1.5 * iqr
            upper_bound = q3 + 1.5 * iqr
            outliers = numeric_data[(numeric_data < lower_bound) | (numeric_data > upper_bound)]
            
            # Value distribution
            value_ranges = {
                "low": f"< {q1:.1f}",
                "medium": f"{q1:.1f} - {q3:.1f}", 
                "high": f"> {q3:.1f}"
            }
            
            return {
                "data_type": "numerical",
                "statistical_summary": stats,
                "distribution_analysis": {
                    "skewness": round(float(numeric_data.skew()), 3),
                    "total_values": int(len(numeric_data)),
                    "zero_values": int((numeric_data == 0).sum())
                },
                "quartiles": {
                    "q1": round(q1, 2),
                    "q3": round(q3, 2),
                    "iqr": round(iqr, 2)
                },
                "outlier_analysis": {
                    "outlier_count": int(len(outliers)),
                    "outlier_percentage": round(len(outliers) / len(numeric_data) * 100, 1),
                    "bounds": {
                        "lower": round(lower_bound, 2),
                        "upper": round(upper_bound, 2)
                    }
                },
                "value_ranges": value_ranges,
                "business_insights": {
                    "average_value": f"The average {column_name} is {stats['mean']}",
                    "variability": f"Values typically range from {stats['min']} to {stats['max']}",
                    "consistency": f"Standard deviation of {stats['std_dev']} indicates moderate variability"
                }
            }
        except Exception as e:
            return {"error": f"Error processing numerical data: {str(e)}"}

    def generate_categorical_insights(self, data, column_name):
        """Generate comprehensive insights for categorical columns"""
        try:
            categorical_data = data[column_name].astype(str)
            value_counts = categorical_data.value_counts()
            total_count = len(categorical_data.dropna())
            
            if total_count == 0:
                return {"error": "No valid categorical data found"}
            
            # Basic statistics
            basic_stats = {
                "unique_categories": int(len(value_counts)),
                "total_values": int(total_count),
                "most_frequent_value": str(value_counts.index[0]),
                "most_frequent_count": int(value_counts.iloc[0]),
                "most_frequent_percentage": round(value_counts.iloc[0] / total_count * 100, 1)
            }
            
            # Top categories (up to 10)
            top_categories = {}
            for category, count in value_counts.head(10).items():
                percentage = (count / total_count) * 100
                top_categories[str(category)] = {
                    "count": int(count),
                    "percentage": round(percentage, 1)
                }
            
            # Diversity analysis
            if len(value_counts) > 0:
                dominance = (value_counts.iloc[0] / total_count) * 100
                top_5_concentration = (value_counts.head(5).sum() / total_count) * 100
                
                diversity = {
                    "dominance_percentage": round(dominance, 1),
                    "top_5_concentration": round(top_5_concentration, 1),
                    "diversity_level": "High" if len(value_counts) > 20 else "Medium" if len(value_counts) > 10 else "Low"
                }
            else:
                diversity = {}
            
            return {
                "data_type": "categorical",
                "basic_statistics": basic_stats,
                "top_categories": top_categories,
                "diversity_analysis": diversity,
                "business_insights": {
                    "dominant_category": f"'{basic_stats['most_frequent_value']}' dominates with {basic_stats['most_frequent_percentage']}% share",
                    "category_count": f"Found {basic_stats['unique_categories']} unique categories",
                    "concentration": f"Top 5 categories account for {diversity.get('top_5_concentration', 0):.1f}% of total"
                }
            }
        except Exception as e:
            return {"error": f"Error processing categorical data: {str(e)}"}

    def generate_date_insights(self, data, column_name):
        """Generate insights for date columns - FIXED: only show non-zero values"""
        try:
            date_data = pd.to_datetime(data[column_name], errors='coerce').dropna()
            
            if len(date_data) == 0:
                return {"error": "No valid date data found"}
            
            # Date range analysis
            date_range = {
                "earliest_date": date_data.min().strftime("%Y-%m-%d"),
                "latest_date": date_data.max().strftime("%Y-%m-%d"),
                "total_days": int((date_data.max() - date_data.min()).days),
                "total_records": int(len(date_data))
            }
            
            # Monthly distribution - ONLY NON-ZERO VALUES
            monthly_counts = date_data.dt.month.value_counts()
            monthly_distribution = {
                str(month): int(count) for month, count in monthly_counts.items() if count > 0
            }
            
            # Weekday analysis - ONLY NON-ZERO VALUES
            weekday_counts = date_data.dt.dayofweek.value_counts()
            weekdays = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
            weekday_distribution = {
                weekdays[day]: int(count) for day, count in weekday_counts.items() if count > 0
            }
            
            return {
                "data_type": "date",
                "date_range": date_range,
                "temporal_patterns": {
                    "monthly_distribution": monthly_distribution,
                    "weekday_distribution": weekday_distribution
                },
                "business_insights": {
                    "time_span": f"Data covers {date_range['total_days']} days from {date_range['earliest_date']} to {date_range['latest_date']}",
                    "peak_month": f"Peak activity in month {monthly_counts.idxmax()} with {monthly_counts.max()} records",
                    "weekday_pattern": f"Most active day: {max(weekday_distribution, key=weekday_distribution.get) if weekday_distribution else 'N/A'}"
                }
            }
        except Exception as e:
            return {"error": f"Error processing date data: {str(e)}"}

    def generate_insights(self, data, column_name, metric_name):
        """Generate insights based on data type detection"""
        if data is None or column_name not in data.columns:
            return {"error": f"Column '{column_name}' not found in dataset"}
        
        # Detect data type
        data_type = self.detect_data_type(data[column_name])
        
        if data_type == "date":
            return self.generate_date_insights(data, column_name)
        elif data_type == "numerical":
            return self.generate_numerical_insights(data, column_name)
        else:
            return self.generate_categorical_insights(data, column_name)

    def create_chart(self, data, column_name, metric_name, requested_chart_type):
        """Create charts with intelligent fallback mechanism"""
        color = self.get_unique_color()
        fig = None
        chart_selection_info = {}
        actual_chart_type = requested_chart_type
        
        try:
            # Detect data type and assess distribution
            data_type = self.detect_data_type(data[column_name])
            distribution_assessment = self.assess_data_distribution(data, column_name, data_type)
            
            # Select appropriate chart type
            chart_selection = self.select_appropriate_chart(
                requested_chart_type, data_type, distribution_assessment
            )
            actual_chart_type = chart_selection["selected_chart"]
            
            # Only include chart selection info if fallback was used
            if chart_selection["was_fallback"]:
                chart_selection_info = {
                    "was_fallback": True,
                    "reason": chart_selection["reason"]
                }
            
            # For numerical data
            if data_type == "numerical":
                numeric_data = pd.to_numeric(data[column_name], errors='coerce').dropna()
                
                if len(numeric_data) == 0:
                    raise ValueError("No valid numerical data available for chart")
                
                if actual_chart_type == "histogram":
                    bins = min(20, len(numeric_data.unique()))
                    fig = px.histogram(x=numeric_data.values, nbins=bins,
                                     title=f"{metric_name.replace('_', ' ').title()} - Distribution")
                    fig.update_traces(marker_color=color, marker_line=dict(width=1, color='black'))
                    fig.update_layout(xaxis_title=metric_name.replace('_', ' ').title(), yaxis_title="Count")
                    
                elif actual_chart_type == "box":
                    fig = px.box(y=numeric_data.values, 
                               title=f"{metric_name.replace('_', ' ').title()} - Box Plot")
                    fig.update_traces(marker_color=color)
                    fig.update_layout(yaxis_title=metric_name.replace('_', ' ').title())
                    
                elif actual_chart_type == "violin":
                    fig = px.violin(y=numeric_data.values, 
                                  title=f"{metric_name.replace('_', ' ').title()} - Violin Plot")
                    fig.update_traces(marker_color=color)
                    fig.update_layout(yaxis_title=metric_name.replace('_', ' ').title())
                    
                elif actual_chart_type == "scatter":
                    # Create index for x-axis
                    x_values = list(range(len(numeric_data)))
                    fig = px.scatter(x=x_values, y=numeric_data.values, 
                                   title=f"{metric_name.replace('_', ' ').title()} - Scatter Plot")
                    fig.update_traces(marker=dict(color=color, size=6, opacity=0.7))
                    fig.update_layout(xaxis_title="Index", yaxis_title=metric_name.replace('_', ' ').title())
                    
                elif actual_chart_type == "line":
                    # Create index for x-axis
                    x_values = list(range(len(numeric_data)))
                    fig = px.line(x=x_values, y=numeric_data.values, 
                                title=f"{metric_name.replace('_', ' ').title()} - Trend Analysis")
                    fig.update_traces(line=dict(color=color, width=2))
                    fig.update_layout(xaxis_title="Sequence", yaxis_title=metric_name.replace('_', ' ').title())
                    
                elif actual_chart_type == "bar":
                    # Create bins for numerical data in bar chart
                    value_counts = numeric_data.value_counts().sort_index().head(15)
                    fig = px.bar(x=value_counts.index, y=value_counts.values,
                               title=f"{metric_name.replace('_', ' ').title()} - Distribution")
                    fig.update_traces(marker_color=color)
                    fig.update_layout(xaxis_title=metric_name.replace('_', ' ').title(), yaxis_title="Count")
            
            # For categorical data
            elif data_type == "categorical":
                value_counts = data[column_name].value_counts().head(15)
                
                if len(value_counts) == 0:
                    raise ValueError("No valid categorical data available for chart")
                
                if actual_chart_type == "bar":
                    fig = px.bar(x=value_counts.index, y=value_counts.values,
                               title=f"{metric_name.replace('_', ' ').title()} - Distribution")
                    fig.update_traces(marker_color=color)
                    fig.update_layout(xaxis_title=metric_name.replace('_', ' ').title(), yaxis_title="Count")
                    
                elif actual_chart_type == "pie":
                    fig = px.pie(values=value_counts.values, names=value_counts.index,
                               title=f"{metric_name.replace('_', ' ').title()} - Composition")
                    fig.update_traces(marker=dict(colors=[color] * len(value_counts)))
                    
                elif actual_chart_type == "treemap":
                    fig = px.treemap(names=value_counts.index, values=value_counts.values,
                                   title=f"{metric_name.replace('_', ' ').title()} - Treemap")
                    
                elif actual_chart_type == "sunburst":
                    fig = px.sunburst(names=value_counts.index, values=value_counts.values,
                                    title=f"{metric_name.replace('_', ' ').title()} - Sunburst")
            
            # For date data
            elif data_type == "date":
                date_data = pd.to_datetime(data[column_name], errors='coerce').dropna()
                if len(date_data) == 0:
                    raise ValueError("No valid date data available for chart")
                
                # Aggregate by date for cleaner charts
                date_counts = date_data.value_counts().sort_index()
                
                if actual_chart_type == "line":
                    fig = px.line(x=date_counts.index, y=date_counts.values,
                                title=f"{metric_name.replace('_', ' ').title()} - Timeline")
                    fig.update_traces(line=dict(color=color, width=2))
                    fig.update_layout(xaxis_title="Date", yaxis_title="Count")
                    
                elif actual_chart_type == "bar":
                    # Show top dates or aggregate by month
                    if len(date_counts) > 50:
                        # Aggregate by month if too many dates
                        monthly_data = date_data.dt.to_period('M').value_counts().sort_index()
                        fig = px.bar(x=monthly_data.index.astype(str), y=monthly_data.values,
                                   title=f"{metric_name.replace('_', ' ').title()} - Monthly Distribution")
                    else:
                        fig = px.bar(x=date_counts.index.astype(str), y=date_counts.values,
                                   title=f"{metric_name.replace('_', ' ').title()} - Distribution")
                    fig.update_traces(marker_color=color)
                    fig.update_layout(xaxis_title="Date", yaxis_title="Count")
                    
                elif actual_chart_type == "area":
                    fig = px.area(x=date_counts.index, y=date_counts.values,
                                title=f"{metric_name.replace('_', ' ').title()} - Area Chart")
                    fig.update_traces(marker_color=color)
                    fig.update_layout(xaxis_title="Date", yaxis_title="Count")
            
            # Apply consistent styling
            if fig is not None:
                fig.update_layout(
                    width=1000,
                    height=600,
                    font=dict(size=12, family="Arial"),
                    plot_bgcolor='rgba(240,240,240,0.1)',
                    paper_bgcolor='white',
                    xaxis=dict(
                        showgrid=True, 
                        gridwidth=1, 
                        gridcolor='lightgray',
                        title_font=dict(size=14)
                    ),
                    yaxis=dict(
                        showgrid=True, 
                        gridwidth=1, 
                        gridcolor='lightgray',
                        title_font=dict(size=14)
                    ),
                    title_font=dict(size=16, family="Arial", color="#2c3e50"),
                    showlegend=True if actual_chart_type in ["pie", "treemap", "sunburst"] else False
                )
                
        except Exception as e:
            print(f"⚠ Chart creation error for {metric_name}: {str(e)}")
            # Create informative error chart
            fig = go.Figure()
            fig.add_annotation(
                text=f"Chart not available: {str(e)}",
                xref="paper", yref="paper",
                x=0.5, y=0.5, showarrow=False,
                font=dict(size=16, color="red")
            )
            fig.update_layout(
                title=f"Chart Error - {metric_name}",
                width=800,
                height=400
            )
            chart_selection_info = {
                "was_fallback": True,
                "reason": f"Chart creation failed: {str(e)}"
            }
            actual_chart_type = "error"
        
        return fig, color, chart_selection_info, actual_chart_type

    def create_comparison_chart(self, data, col1_name, col2_name, comparison_name, requested_chart_type):
        """Create comparison charts with proper data handling - FIXED"""
        color = self.get_unique_color()
        fig = None
        chart_selection_info = {}
        actual_chart_type = requested_chart_type
        
        try:
            # Get clean data for both columns
            col1_clean = data[col1_name].dropna()
            col2_clean = data[col2_name].dropna()
            
            # Align indices
            common_idx = col1_clean.index.intersection(col2_clean.index)
            col1_data = col1_clean.loc[common_idx]
            col2_data = col2_clean.loc[common_idx]
            
            if len(common_idx) == 0:
                raise ValueError("No common data points between the two columns")
            
            # Convert to numeric if possible
            try:
                col1_numeric = pd.to_numeric(col1_data, errors='coerce').dropna()
                col2_numeric = pd.to_numeric(col2_data, errors='coerce').dropna()
                
                # Realign after numeric conversion
                common_numeric_idx = col1_numeric.index.intersection(col2_numeric.index)
                col1_numeric = col1_numeric.loc[common_numeric_idx]
                col2_numeric = col2_numeric.loc[common_numeric_idx]
                
                if len(common_numeric_idx) > 0:
                    # Both columns are numeric - create scatter plot
                    if requested_chart_type in ["scatter", "line"]:
                        actual_chart_type = requested_chart_type
                    else:
                        actual_chart_type = "scatter"
                        chart_selection_info = {
                            "was_fallback": True,
                            "reason": f"Comparison between numerical columns - scatter chart is most appropriate"
                        }
                    
                    if actual_chart_type == "scatter":
                        fig = px.scatter(
                            x=col1_numeric.values, 
                            y=col2_numeric.values,
                            title=f"{col1_name} vs {col2_name} - Correlation Analysis"
                        )
                        fig.update_traces(marker=dict(color=color, size=6, opacity=0.7))
                        fig.update_layout(
                            xaxis_title=col1_name.replace('_', ' ').title(),
                            yaxis_title=col2_name.replace('_', ' ').title()
                        )
                    elif actual_chart_type == "line":
                        # Sort by first column for line chart
                        sorted_idx = col1_numeric.sort_index().index
                        fig = go.Figure()
                        fig.add_trace(go.Scatter(
                            x=col1_numeric.loc[sorted_idx].values, 
                            y=col2_numeric.loc[sorted_idx].values,
                            mode='lines+markers',
                            line=dict(color=color, width=2),
                            marker=dict(size=4),
                            name=f"{col1_name} vs {col2_name}"
                        ))
                        fig.update_layout(
                            title=f"{col1_name} vs {col2_name} - Trend Comparison",
                            xaxis_title=col1_name.replace('_', ' ').title(),
                            yaxis_title=col2_name.replace('_', ' ').title()
                        )
                else:
                    raise ValueError("No valid numerical data for comparison")
                    
            except (ValueError, TypeError):
                # Fallback to grouped bar chart for non-numeric comparisons
                actual_chart_type = "bar"
                chart_selection_info = {
                    "was_fallback": True,
                    "reason": "Non-numeric data comparison - using grouped bar chart"
                }
                
                # Get value counts for both columns
                val1_counts = col1_data.value_counts().head(10)
                val2_counts = col2_data.value_counts().head(10)
                
                fig = go.Figure()
                fig.add_trace(go.Bar(
                    x=val1_counts.index,
                    y=val1_counts.values,
                    name=col1_name,
                    marker_color=color
                ))
                fig.add_trace(go.Bar(
                    x=val2_counts.index,
                    y=val2_counts.values,
                    name=col2_name,
                    marker_color=self.get_unique_color()
                ))
                fig.update_layout(
                    title=f"{col1_name} vs {col2_name} - Distribution Comparison",
                    xaxis_title="Categories",
                    yaxis_title="Count",
                    barmode='group'
                )
            
            # Apply styling
            if fig is not None:
                fig.update_layout(
                    width=1000,
                    height=600,
                    font=dict(size=12, family="Arial"),
                    plot_bgcolor='rgba(240,240,240,0.1)'
                )
            else:
                raise ValueError("Could not create comparison chart")
                
        except Exception as e:
            print(f"⚠ Comparison chart error for {comparison_name}: {str(e)}")
            # Create a proper error visualization
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=[0, 1], y=[0, 1],
                mode='markers+text',
                text=['Comparison', 'Not Available'],
                textposition="middle center",
                marker=dict(size=0)
            ))
            fig.update_layout(
                title=f"Comparison Error: {comparison_name}",
                annotations=[
                    dict(
                        text=f"Error: {str(e)}",
                        xref="paper", yref="paper",
                        x=0.5, y=0.5, showarrow=False,
                        font=dict(size=14, color="red")
                    )
                ],
                width=800,
                height=400
            )
            chart_selection_info = {
                "was_fallback": True,
                "reason": f"Comparison chart creation failed: {str(e)}"
            }
            actual_chart_type = "error"
        
        return fig, color, chart_selection_info, actual_chart_type

    def generate_comparison_insights(self, data, col1_name, col2_name, metric1, metric2):
        """Generate insights comparing two columns"""
        try:
            # Get clean data for both columns
            col1_clean = data[col1_name].dropna()
            col2_clean = data[col2_name].dropna()
            
            # Align indices
            common_idx = col1_clean.index.intersection(col2_clean.index)
            col1_data = col1_clean.loc[common_idx]
            col2_data = col2_clean.loc[common_idx]
            
            if len(common_idx) == 0:
                return {"error": "No common data points for comparison"}
            
            # Try numeric correlation
            try:
                col1_numeric = pd.to_numeric(col1_data, errors='coerce').dropna()
                col2_numeric = pd.to_numeric(col2_data, errors='coerce').dropna()
                
                common_numeric_idx = col1_numeric.index.intersection(col2_numeric.index)
                if len(common_numeric_idx) > 1:
                    correlation = np.corrcoef(col1_numeric.loc[common_numeric_idx], 
                                            col2_numeric.loc[common_numeric_idx])[0,1]
                    
                    return {
                        "correlation_analysis": {
                            "correlation_coefficient": round(float(correlation), 3),
                            "relationship_strength": "strong" if abs(correlation) > 0.7 else 
                                                   "moderate" if abs(correlation) > 0.3 else "weak",
                            "relationship_direction": "positive" if correlation > 0 else "negative"
                        },
                        "business_insight": f"Correlation between {metric1} and {metric2} is {correlation:.3f} ({'positive' if correlation > 0 else 'negative'} relationship)"
                    }
            except:
                pass
            
            # Fallback analysis for non-numeric data
            return {
                "analysis": "Comparison analysis completed",
                "data_points_compared": len(common_idx),
                "business_insight": f"Successfully compared {metric1} with {metric2} using {len(common_idx)} data points"
            }
            
        except Exception as e:
            return {"error": f"Comparison analysis failed: {str(e)}"}

    def process_insights(self, insights_config):
        """Process insights-only configuration - SIMPLIFIED OUTPUT"""
        results = []
        
        for metric_name in insights_config:
            column_name = self.find_best_column_match(metric_name)
            if not column_name:
                print(f"⚠ Column not found for metric: {metric_name}")
                continue
                
            print(f"📊 Generating insights for: {metric_name} -> {column_name}")
            insights = self.generate_insights(self.dataset, column_name, metric_name)
            
            results.append({
                "metric_name": metric_name,
                "column_used": column_name,
                "insights": insights
            })
        
        # Save to single JSON file - SIMPLIFIED OUTPUT
        output_path = os.path.join(self.output_dir, "insights.json")
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Insights saved to: {output_path}")
        return results

    def process_insights_charts(self, insights_charts_config):
        """Process insights with charts configuration - SIMPLIFIED OUTPUT"""
        results = []
        
        for metric_name, requested_chart_type in insights_charts_config.items():
            column_name = self.find_best_column_match(metric_name)
            if not column_name:
                print(f"⚠ Column not found for metric: {metric_name}")
                continue
                
            print(f"📈 Generating chart for: {metric_name} -> {column_name} (requested: {requested_chart_type})")
            
            # Generate insights
            insights = self.generate_insights(self.dataset, column_name, metric_name)
            
            # Create chart with fallback mechanism
            fig, chart_color, chart_selection_info, actual_chart_type = self.create_chart(
                self.dataset, column_name, metric_name, requested_chart_type
            )
            
            # Save chart
            chart_filename = f"{metric_name}_{actual_chart_type}.png"
            chart_path = os.path.join(self.output_dir, "charts", chart_filename)
            
            try:
                fig.write_image(chart_path)
                chart_status = "success"
            except Exception as e:
                print(f"⚠ Chart saving error: {str(e)}")
                chart_status = f"error: {str(e)}"
            
            # Only include fallback info if fallback occurred
            result_item = {
                "metric_name": metric_name,
                "column_used": column_name,
                "chart_type": actual_chart_type,
                "chart_path": chart_path,
                "chart_status": chart_status,
                "insights": insights
            }
            
            # Only add chart_selection_info if fallback was used
            if chart_selection_info:
                result_item["chart_selection_info"] = chart_selection_info
            
            results.append(result_item)
        
        # Save to single JSON file - SIMPLIFIED OUTPUT
        output_path = os.path.join(self.output_dir, "insights_charts.json")
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Insights with charts saved to: {output_path}")
        return results

    def process_comparison(self, comparison_config):
        """Process comparison configuration - SIMPLIFIED OUTPUT"""
        results = []
        
        for comparison_name, requested_chart_type in comparison_config.items():
            if '_vs_' not in comparison_name:
                print(f"⚠ Invalid comparison name format: {comparison_name}")
                continue
                
            columns = comparison_name.split('_vs_')
            if len(columns) != 2:
                print(f"⚠ Invalid comparison format: {comparison_name}")
                continue
                
            col1_name = self.find_best_column_match(columns[0])
            col2_name = self.find_best_column_match(columns[1])
            
            if not col1_name or not col2_name:
                print(f"⚠ Columns not found for comparison: {comparison_name}")
                continue
            
            print(f"🔁 Processing comparison: {comparison_name} (requested: {requested_chart_type})")
            
            # Generate individual insights
            col1_insights = self.generate_insights(self.dataset, col1_name, columns[0])
            col2_insights = self.generate_insights(self.dataset, col2_name, columns[1])
            
            # Create comparison chart with fallback
            fig, chart_color, chart_selection_info, actual_chart_type = self.create_comparison_chart(
                self.dataset, col1_name, col2_name, comparison_name, requested_chart_type
            )
            
            # Save chart
            chart_filename = f"{comparison_name}_{actual_chart_type}.png"
            chart_path = os.path.join(self.output_dir, "charts", chart_filename)
            
            try:
                fig.write_image(chart_path)
                chart_status = "success"
            except Exception as e:
                print(f"⚠ Chart saving error: {str(e)}")
                chart_status = f"error: {str(e)}"
            
            # Generate comparison insights
            comparison_insights = self.generate_comparison_insights(
                self.dataset, col1_name, col2_name, columns[0], columns[1]
            )
            
            # Only include fallback info if fallback occurred
            result_item = {
                "comparison_name": comparison_name,
                "columns_compared": [col1_name, col2_name],
                "chart_type": actual_chart_type,
                "chart_path": chart_path,
                "chart_status": chart_status,
                "individual_insights": {
                    columns[0]: col1_insights,
                    columns[1]: col2_insights
                },
                "comparison_insights": comparison_insights
            }
            
            # Only add chart_selection_info if fallback was used
            if chart_selection_info:
                result_item["chart_selection_info"] = chart_selection_info
            
            results.append(result_item)
        
        # Save to single JSON file - SIMPLIFIED OUTPUT
        output_path = os.path.join(self.output_dir, "comparison.json")
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        
        print(f"✅ Comparisons saved to: {output_path}")
        return results

    def process_json_config(self, json_config_path):
        """Process the complete JSON configuration"""
        self.used_colors = set()  # Reset colors for new run
        
        if not os.path.exists(json_config_path):
            print(f"❌ Config file not found: {json_config_path}")
            # Create sample config
            return {}
        with open(json_config_path, 'r') as f:
            config = json.load(f)
        
        print("🚀 Starting data analysis with intelligent chart fallback...")
        results = {}
        
        # Process each section
        if "insights" in config:
            print("\n" + "="*50)
            print("PROCESSING INSIGHTS...")
            print("="*50)
            results["insights"] = self.process_insights(config["insights"])
        
        if "insights_charts" in config:
            print("\n" + "="*50)
            print("PROCESSING INSIGHTS WITH CHARTS...")
            print("="*50)
            results["insights_charts"] = self.process_insights_charts(config["insights_charts"])
        
        if "comparison" in config:
            print("\n" + "="*50)
            print("PROCESSING COMPARISONS...")
            print("="*50)
            results["comparison"] = self.process_comparison(config["comparison"])
        
        # Summary
        print("\n" + "="*50)
        print("ANALYSIS COMPLETED SUCCESSFULLY! ✅")
        print("="*50)
        print(f"📊 Insights generated: {len(results.get('insights', []))}")
        print(f"📈 Charts created: {len(results.get('insights_charts', []))}")
        print(f"🔁 Comparisons analyzed: {len(results.get('comparison', []))}")
        print(f"🎨 Unique colors used: {len(self.used_colors)}")
        print(f"💾 Output directory: {self.output_dir}")
        
        return results


# Main execution
# if __name__ == "__main__":
# #     # Initialize the visualization engine
#     engine = HotelBookingVisualization(dataset_path=r"input\csv\input.csv")
    
# #     # Process the JSON configuration
#     json_config_path = r"input\input.json"
#     results = engine.process_json_config(json_config_path)
    
#     print(f"\n🎯 Analysis completed! Check the 'output' folder for results.")