# Excel Data Assistant

A Streamlit application that allows users to upload Excel files and ask natural language questions about their data. The app uses OpenAI's GPT model to generate Pandas code and automatically creates visualizations when appropriate.

## Features

- Excel File Upload: Support for .xlsx files with automatic data validation
- Natural Language Queries: Ask questions in plain English about your data
- Automatic Visualization: Charts are generated based on query intent (bar, line, histogram, pie, correlation)
- Data Cleaning: Automatic column normalization and missing value handling
- Query History: Track and export your analysis history
- Professional UI: Clean, modern interface without distractions
- Secure: API keys stored in environment variables, data stays on your device

## Installation

1. Clone the repository:
   
   git clone https://github.com/Diya-J/Excel_bot


2. Install dependencies:
   
   pip install -r requirements.txt
   


3. Run the application:

   streamlit run streamlit_excel_uploader.py


## Usage

1. Upload Excel File: Use the file uploader to select your .xlsx file
2. Review Data: Check the data preview and column information
3. Ask Questions: Enter natural language queries about your data
4. View Results:See both tabular results and automatic visualizations
5. Export History: Download your query history as JSON

## Example Queries

- "What is the average salary?"
- "How many employees are in each department?"
- "Show me a bar chart of sales by region"
- "What is the correlation between age and salary?"
- "Which department has the highest average salary?"
- "Create a pie chart showing the distribution of job titles"

## Supported Chart Types

- Distribution: Histograms for numeric data
- Trend: Line charts for time series or sequential data
- Comparison: Bar charts for categorical comparisons
- Correlation: Heatmaps for numeric relationships
- Pie Charts: For categorical distributions

## File Requirements

- Format: .xlsx files only
- Size: Up to 1000 rows
- Columns: Up to 20 columns
- Content: Single sheet with headers

## Architecture

- Frontend: Streamlit for responsive web interface
- Data Processing: Pandas for data manipulation and analysis
- AI Integration: OpenAI GPT-3.5-turbo for natural language understanding
- Visualization: Matplotlib and Seaborn for chart generation
- Configuration:Environment variables for secure API key management

## Security

- API keys are stored in environment variables, not in code
- Data processing happens locally on your device
- No data is stored or transmitted to external servers
- Safe code execution environment with restricted builtins

## Deployment

### Local Development

streamlit run streamlit_excel_uploader.py


## Limitations

- Requires OpenAI API key and internet connection
- Limited to 1000 rows and 20 columns per file
- Single sheet Excel files only
- Chart generation depends on data types and query intent

## Development

### Project Structure

Chatbot_ai/
├── streamlit_excel_uploader.py  # Main application
├── requirements.txt             # Python dependencies
├── .env                        # Environment variables (create this)
├── README.md                   # This file


### Key Functions
- main(): Application entry point
- handle_file_upload(): Excel file processing
- process_query(): Natural language query processing
- generate_chart(): Automatic visualization generation
- get_pandas_code_from_openai(): AI-powered code generation


### Version 1.0.0
- Initial release with core functionality
- Environment variable configuration
- Professional UI design
- Automatic chart generation
- Query history and export features
