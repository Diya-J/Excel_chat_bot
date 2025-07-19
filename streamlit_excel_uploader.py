"""
Excel Data Assistant with LLM Integration
=========================================

A Streamlit application that allows users to upload Excel files and ask natural language
questions about their data. The app uses OpenAI's GPT model to generate Pandas code
and automatically creates visualizations when appropriate.

Features:
- Excel file upload and parsing
- Natural language query processing
- Automatic chart generation
- Data validation and cleaning
- Query history and export functionality
- Responsive and intuitive UI

Author: Diya J
Date: 19/07/2025
"""

import streamlit as st
import pandas as pd
import re
import openai
import matplotlib.pyplot as plt
import seaborn as sns
import json
from datetime import datetime
import io
import base64
import os
import streamlit as st
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get API key from environment variable
#OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")
OPENAI_API_KEY = st.secrets.get("OPENAI_API_KEY")

# Page configuration
st.set_page_config(
    page_title="Excel Data Assistant",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for clean, professional styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
        font-weight: bold;
    }
    
    .section-header {
        font-size: 1.5rem;
        color: #2c3e50;
        border-bottom: 2px solid #3498db;
        padding-bottom: 0.5rem;
        margin-top: 2rem;
        font-weight: 600;
    }
    
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
        color: #155724;
    }
    
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
        color: #0c5460;
    }
    
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        border-radius: 5px;
        padding: 1rem;
        margin: 1rem 0;
        color: #721c24;
    }
    
    /* Custom button styling */
    .stButton > button {
        background-color: #007bff;
        color: white;
        border: none;
        border-radius: 5px;
        padding: 0.5rem 1rem;
        font-weight: 500;
        transition: background-color 0.3s ease;
    }
    
    .stButton > button:hover {
        background-color: #0056b3;
    }
    
    /* Metric styling */
    .metric-container {
        background-color: #f8f9fa;
        border: 1px solid #dee2e6;
        border-radius: 5px;
        padding: 1rem;
        margin: 0.5rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'query_history' not in st.session_state:
    st.session_state.query_history = []
if 'current_df' not in st.session_state:
    st.session_state.current_df = None
if 'file_name' not in st.session_state:
    st.session_state.file_name = None

def main():
    """Main application function"""
    display_header()
    setup_sidebar()
    handle_file_upload()
    if st.session_state.current_df is not None:
        display_data_info()
        handle_query_section()
        display_query_history()

def display_header():
    """Display the main header"""
    st.markdown('<h1 class="main-header">Excel Data Assistant</h1>', unsafe_allow_html=True)
    st.markdown("""
    <div class="info-box">
        <strong>How it works:</strong> Upload your Excel file, ask questions in natural language, 
        and get instant insights with automatic visualizations!
    </div>
    """, unsafe_allow_html=True)

def setup_sidebar():
    """Setup the sidebar with app info and statistics"""
    # Check if API key is configured and store in session state
    if OPENAI_API_KEY:
        st.session_state.openai_api_key = OPENAI_API_KEY
    else:
        st.session_state.openai_api_key = None
    
    st.sidebar.header("App Statistics")
    if st.session_state.current_df is not None:
        st.sidebar.markdown('<div class="metric-container">', unsafe_allow_html=True)
        st.sidebar.metric("Rows", len(st.session_state.current_df))
        st.sidebar.metric("Columns", len(st.session_state.current_df.columns))
        st.sidebar.metric("Queries Made", len(st.session_state.query_history))
        st.sidebar.markdown('</div>', unsafe_allow_html=True)
    
    st.sidebar.markdown("---")
    st.sidebar.header("About")
    st.sidebar.info("""
    This app uses OpenAI's GPT model to understand your questions and generate 
    appropriate data analysis code. Your data stays on your device and is not stored.
    """)

def normalize_column_name(col):
    """
    Normalize column names by removing special characters and converting to lowercase
    
    Args:
        col (str): Original column name
        
    Returns:
        str: Normalized column name
    """
    if pd.isna(col):
        return "unnamed_column"
    col = str(col)
    col = re.sub(r"[^a-zA-Z0-9_]", "_", col)
    col = re.sub(r"_+", "_", col)
    col = col.strip("_").lower()
    return col if col else "unnamed_column"

def validate_dataframe(df):
    """
    Validate the uploaded DataFrame
    
    Args:
        df (pd.DataFrame): DataFrame to validate
        
    Returns:
        tuple: (is_valid, error_message)
    """
    if df.empty:
        return False, "The uploaded file contains no data."
    
    if len(df) > 1000:
        return False, "File too large. Please upload a file with 1000 rows or fewer."
    
    if len(df.columns) > 20:
        return False, "Too many columns. Please upload a file with 20 columns or fewer."
    
    return True, ""

def handle_missing_values(df):
    """
    Handle missing values in the DataFrame based on column types
    
    Args:
        df (pd.DataFrame): DataFrame to process
        
    Returns:
        pd.DataFrame: DataFrame with missing values handled
    """
    df_processed = df.copy()
    
    for col in df_processed.columns:
        if pd.api.types.is_numeric_dtype(df_processed[col]):
            # Use median for numeric columns
            median_val = df_processed[col].median()
            df_processed[col] = df_processed[col].fillna(median_val)
        elif pd.api.types.is_datetime64_any_dtype(df_processed[col]):
            # Use mode for datetime columns
            mode_val = df_processed[col].mode()
            if not mode_val.empty:
                df_processed[col] = df_processed[col].fillna(mode_val[0])
            else:
                df_processed[col] = df_processed[col].fillna(pd.NaT)
        else:
            # Use mode for categorical/text columns
            mode_val = df_processed[col].mode()
            if not mode_val.empty:
                df_processed[col] = df_processed[col].fillna(mode_val[0])
            else:
                df_processed[col] = df_processed[col].fillna("Unknown")
    
    return df_processed

def handle_file_upload():
    """Handle Excel file upload and processing"""
    st.markdown('<h2 class="section-header">Upload Your Excel File</h2>', unsafe_allow_html=True)
    
    uploaded_file = st.file_uploader(
        "Choose an Excel file (.xlsx)",
        type=["xlsx"],
        help="Upload an Excel file with up to 1000 rows and 20 columns"
    )
    
    if uploaded_file is not None:
        with st.spinner("Processing your Excel file..."):
            try:
                # Read the Excel file
                df = pd.read_excel(uploaded_file, engine="openpyxl")
                
                # Validate the data
                is_valid, error_msg = validate_dataframe(df)
                if not is_valid:
                    st.markdown(f'<div class="error-box">{error_msg}</div>', unsafe_allow_html=True)
                    return
                
                # Normalize column names
                df.columns = [normalize_column_name(col) for col in df.columns]
                
                # Handle missing values
                df_processed = handle_missing_values(df)
                
                # Store in session state
                st.session_state.current_df = df_processed
                st.session_state.file_name = uploaded_file.name
                
                st.markdown('<div class="success-box">File uploaded and processed successfully!</div>', unsafe_allow_html=True)
                
            except Exception as e:
                st.markdown(f'<div class="error-box">Error processing file: {str(e)}</div>', unsafe_allow_html=True)
                st.info("Please ensure your file is a valid Excel (.xlsx) file.")

def display_data_info():
    """Display information about the uploaded data"""
    df = st.session_state.current_df
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.subheader("Data Preview")
        st.dataframe(df.head(10), use_container_width=True)
        
        # Show data statistics
        st.subheader("Data Statistics")
        numeric_cols = df.select_dtypes(include=['number']).columns
        if len(numeric_cols) > 0:
            st.dataframe(df[numeric_cols].describe(), use_container_width=True)
    
    with col2:
        st.subheader("Column Information")
        col_info = []
        for col in df.columns:
            dtype = str(df[col].dtype)
            missing = df[col].isnull().sum()
            unique = df[col].nunique()
            col_info.append({
                "Column": col,
                "Type": dtype,
                "Missing": missing,
                "Unique Values": unique
            })
        
        col_df = pd.DataFrame(col_info)
        st.dataframe(col_df, use_container_width=True)

def get_pandas_code_from_openai(user_query, columns, api_key):
    """
    Generate Pandas code using OpenAI API
    
    Args:
        user_query (str): User's natural language question
        columns (list): List of DataFrame column names
        api_key (str): OpenAI API key
        
    Returns:
        tuple: (generated_code, error_message)
    """
    openai.api_key = api_key
    
    # Enhanced prompt for better code generation
    prompt = f"""
You are an expert Python data analyst. Given a pandas DataFrame called 'df' with the following columns: {columns}.

Your task is to write a single line of pandas code that answers the user's question.
The code should be:
- Safe and executable
- Return meaningful results
- Handle edge cases appropriately
- Use proper pandas methods

User Question: {user_query}

Write only the pandas code (no explanations, no imports, no print statements):
"""
    
    try:
        response = openai.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[
                {"role": "system", "content": "You are a helpful Python data analyst. Always return safe, executable pandas code."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=150,
            temperature=0.1,
            stop=["\n", "```"]
        )
        
        code = response.choices[0].message.content.strip()
        # Clean up the code
        code = re.sub(r'^```python\s*', '', code)
        code = re.sub(r'\s*```$', '', code)
        
        return code, None
        
    except Exception as e:
        return None, str(e)

def detect_chart_type(user_query):
    """
    Detect if the query requires a chart and what type
    
    Args:
        user_query (str): User's question
        
    Returns:
        str: Chart type or None
    """
    query_lower = user_query.lower()
    
    chart_patterns = {
        "distribution": ["distribution", "histogram", "spread", "frequency", "how many", "count of", "show the range", "distribution of"],
        "trend": ["trend", "over time", "change", "evolution", "progression", "line chart", "time series", "growth", "decline"],
        "comparison": ["compare", "comparison", "vs", "versus", "difference", "bar chart", "grouped", "by category", "across"],
        "correlation": ["correlation", "relationship", "scatter", "correlate", "related to"],
        "pie": ["pie chart", "percentage", "proportion", "share", "breakdown"]
    }
    
    for chart_type, keywords in chart_patterns.items():
        if any(keyword in query_lower for keyword in keywords):
            return chart_type
    
    return None

def generate_chart(df, chart_type, user_query):
    """
    Generate appropriate chart based on chart type and data
    
    Args:
        df (pd.DataFrame): DataFrame to visualize
        chart_type (str): Type of chart to generate
        user_query (str): Original user query
        
    Returns:
        tuple: (success, error_message)
    """
    try:
        # Set style for clean charts
        plt.style.use('default')
        fig, ax = plt.subplots(figsize=(10, 6))
        columns = list(df.columns)
        
        if chart_type == "distribution":
            num_cols = [col for col in columns if pd.api.types.is_numeric_dtype(df[col])]
            if num_cols:
                sns.histplot(df[num_cols[0]], kde=True, ax=ax, bins=20, color='#1f77b4')
                ax.set_xlabel(num_cols[0].replace('_', ' ').title())
                ax.set_ylabel("Frequency")
                ax.set_title(f"Distribution of {num_cols[0].replace('_', ' ').title()}")
            else:
                return False, "No numeric columns available for distribution plot."
                
        elif chart_type == "trend":
            num_cols = [col for col in columns if pd.api.types.is_numeric_dtype(df[col])]
            if len(num_cols) >= 1:
                x_col = columns[0] if not pd.api.types.is_numeric_dtype(df[columns[0]]) else columns[1] if len(columns) > 1 else num_cols[0]
                y_col = num_cols[0]
                sns.lineplot(data=df, x=x_col, y=y_col, ax=ax, color='#1f77b4')
                ax.set_xlabel(x_col.replace('_', ' ').title())
                ax.set_ylabel(y_col.replace('_', ' ').title())
                ax.set_title(f"Trend: {y_col.replace('_', ' ').title()} over {x_col.replace('_', ' ').title()}")
            else:
                return False, "Not enough numeric columns for trend plot."
                
        elif chart_type == "comparison":
            cat_cols = [col for col in columns if pd.api.types.is_object_dtype(df[col])]
            num_cols = [col for col in columns if pd.api.types.is_numeric_dtype(df[col])]
            if cat_cols and num_cols:
                # Group by categorical column and aggregate numeric column
                grouped_data = df.groupby(cat_cols[0])[num_cols[0]].mean().reset_index()
                sns.barplot(data=grouped_data, x=cat_cols[0], y=num_cols[0], ax=ax, color='#1f77b4')
                ax.set_xlabel(cat_cols[0].replace('_', ' ').title())
                ax.set_ylabel(f"Average {num_cols[0].replace('_', ' ').title()}")
                ax.set_title(f"Comparison: {num_cols[0].replace('_', ' ').title()} by {cat_cols[0].replace('_', ' ').title()}")
                plt.xticks(rotation=45)
            else:
                return False, "Not enough categorical/numeric columns for comparison plot."
                
        elif chart_type == "correlation":
            num_cols = [col for col in columns if pd.api.types.is_numeric_dtype(df[col])]
            if len(num_cols) >= 2:
                correlation_matrix = df[num_cols].corr()
                sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', center=0, ax=ax)
                ax.set_title("Correlation Matrix")
            else:
                return False, "Need at least 2 numeric columns for correlation analysis."
                
        elif chart_type == "pie":
            cat_cols = [col for col in columns if pd.api.types.is_object_dtype(df[col])]
            if cat_cols:
                value_counts = df[cat_cols[0]].value_counts()
                colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', '#8c564b']
                ax.pie(value_counts.values, labels=value_counts.index, autopct='%1.1f%%', colors=colors[:len(value_counts)])
                ax.set_title(f"Distribution of {cat_cols[0].replace('_', ' ').title()}")
            else:
                return False, "No categorical columns available for pie chart."
        
        plt.tight_layout()
        st.pyplot(fig, use_container_width=True)
        plt.close()
        
        return True, None
        
    except Exception as e:
        return False, str(e)

def handle_query_section():
    """Handle the query input and processing section"""
    st.markdown('<h2 class="section-header">Ask Questions About Your Data</h2>', unsafe_allow_html=True)
    
    # Get API key from session state
    openai_api_key = st.session_state.get('openai_api_key', '')
    
    with st.form("query_form"):
        user_query = st.text_area(
            "Enter your question in natural language:",
            placeholder="e.g., What is the average salary? Show me a bar chart of sales by region.",
            height=100
        )
        
        col1, col2 = st.columns([1, 4])
        with col1:
            submitted = st.form_submit_button("Ask Question", use_container_width=True)
        with col2:
            if st.form_submit_button("Example Questions", use_container_width=True):
                st.session_state.show_examples = True
    
    # Show example questions
    if st.session_state.get('show_examples', False):
        st.markdown('<div class="info-box"><strong>Example Questions:</strong></div>', unsafe_allow_html=True)
        examples = [
            "What is the average salary?",
            "How many employees are in each department?",
            "Show me a bar chart of sales by region",
            "What is the correlation between age and salary?",
            "Which department has the highest average salary?",
            "Create a pie chart showing the distribution of job titles"
        ]
        for example in examples:
            if st.button(example, key=f"example_{example}"):
                st.session_state.user_query = example
                st.session_state.show_examples = False
                st.rerun()
    
    if submitted and user_query.strip():
        if not openai_api_key:
            st.markdown('<div class="error-box">Please set OPENAI_API_KEY in your .env file.</div>', unsafe_allow_html=True)
        else:
            process_query(user_query, openai_api_key)

def process_query(user_query, api_key):
    """Process a user query and display results"""
    df = st.session_state.current_df
    
    with st.spinner("Processing your question with AI..."):
        # Detect chart type
        chart_type = detect_chart_type(user_query)
        
        # Generate Pandas code
        pandas_code, openai_error = get_pandas_code_from_openai(user_query, list(df.columns), api_key)
        
        if openai_error:
            st.markdown(f'<div class="error-box">Error with OpenAI API: {openai_error}</div>', unsafe_allow_html=True)
            return
        
        if not pandas_code:
            st.markdown('<div class="error-box">No code generated. Please try rephrasing your question.</div>', unsafe_allow_html=True)
            return
        
        # Display generated code
        st.subheader("Generated Code")
        st.code(pandas_code, language="python")
        
        # Execute the code
        with st.spinner("Running analysis..."):
            try:
                # Safe execution environment
                safe_dict = {
                    'df': df.copy(),
                    'pd': pd,
                    'len': len,
                    'sum': sum,
                    'min': min,
                    'max': max,
                    'mean': lambda x: sum(x) / len(x) if len(x) > 0 else 0,
                    'round': round
                }
                
                result = eval(pandas_code, {"__builtins__": {}}, safe_dict)
                
                # Display results
                st.subheader("Results")
                if isinstance(result, pd.DataFrame):
                    st.dataframe(result, use_container_width=True)
                elif isinstance(result, pd.Series):
                    st.dataframe(result.to_frame(), use_container_width=True)
                else:
                    st.metric("Result", str(result))
                
                # Generate chart if appropriate
                if chart_type:
                    st.subheader("Visualization")
                    success, chart_error = generate_chart(df, chart_type, user_query)
                    if not success:
                        st.markdown(f'<div class="info-box">{chart_error}</div>', unsafe_allow_html=True)
                
                # Store query in history
                query_record = {
                    'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    'query': user_query,
                    'code': pandas_code,
                    'result_type': type(result).__name__,
                    'chart_generated': chart_type is not None
                }
                st.session_state.query_history.append(query_record)
                
                # Success message
                st.markdown('<div class="success-box">Analysis completed successfully!</div>', unsafe_allow_html=True)
                
            except Exception as e:
                st.markdown(f'<div class="error-box">Error executing code: {str(e)}</div>', unsafe_allow_html=True)
                st.markdown('<div class="info-box">Try rephrasing your question or check if the data contains the columns you\'re asking about.</div>', unsafe_allow_html=True)

def display_query_history():
    """Display query history and export options"""
    if st.session_state.query_history:
        st.markdown('<h2 class="section-header">Query History</h2>', unsafe_allow_html=True)
        
        # Display recent queries
        for i, record in enumerate(reversed(st.session_state.query_history[-5:])):  # Show last 5
            with st.expander(f"Query {len(st.session_state.query_history) - i}: {record['query'][:50]}..."):
                st.write(f"**Time:** {record['timestamp']}")
                st.write(f"**Code:** `{record['code']}`")
                st.write(f"**Result Type:** {record['result_type']}")
                if record['chart_generated']:
                    st.write("Chart was generated")
        
        # Export options
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Export History as JSON"):
                history_json = json.dumps(st.session_state.query_history, indent=2)
                st.download_button(
                    label="Download JSON",
                    data=history_json,
                    file_name=f"query_history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
                    mime="application/json"
                )
        
        with col2:
            if st.button("Clear History"):
                st.session_state.query_history = []
                st.rerun()

def create_download_link(df, filename, text):
    """Create a download link for the DataFrame"""
    csv = df.to_csv(index=False)
    b64 = base64.b64encode(csv.encode()).decode()
    href = f'<a href="data:file/csv;base64,{b64}" download="{filename}">{text}</a>'
    return href

if __name__ == "__main__":
    main()
