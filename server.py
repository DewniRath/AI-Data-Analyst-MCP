from fastmcp import FastMCP
import pandas as pd
import os
import logging
from pathlib import Path
from typing import Optional, List, Dict, Any
import matplotlib.pyplot as plt
import seaborn as sns
import xlsxwriter


# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Create the MCP server
mcp = FastMCP("DataAgent")

# The central folder for all your project files
BASE_DIR = Path(r"C:\Users\dewni\Documents\my_mcp_project").resolve()

def validate_file_path(filename: str) -> Path:
    """
    Validates that the file path is within BASE_DIR (prevents directory traversal attacks).
    
    Args:
        filename: The filename to validate
        
    Returns:
        Path: Absolute path if valid
        
    Raises:
        ValueError: If path is outside BASE_DIR or file doesn't exist
    """
    try:
        file_path = (BASE_DIR / filename).resolve()
        if not str(file_path).startswith(str(BASE_DIR)):
            raise ValueError(f"Access denied: {filename} is outside project directory")
        if not file_path.exists():
            raise ValueError(f"File not found: {filename}")
        return file_path
    except Exception as e:
        logger.error(f"Path validation failed for {filename}: {str(e)}")
        raise

@mcp.tool()
def list_files() -> str:
    """Lists all data files (CSV, Excel, JSON) available in the project folder."""
    try:
        extensions = ('.csv', '.xlsx', '.xls', '.json')
        files = [f for f in os.listdir(BASE_DIR) if f.endswith(extensions)]
        logger.info(f"Listed {len(files)} files")
        return f"Files found: {', '.join(files)}" if files else "No supported data files found."
    except Exception as e:
        logger.error(f"Error listing files: {str(e)}")
        raise Exception(f"Failed to list files: {str(e)}")

@mcp.tool()
def inspect_data(filename: str) -> str:
    """Shows column names, data types, and null counts for a file (Schema)."""
    try:
        path = validate_file_path(filename)
        df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
        
        summary = []
        for col in df.columns:
            summary.append(f"Col: {col} | Type: {df[col].dtype} | Nulls: {df[col].isnull().sum()}")
        
        logger.info(f"Inspected {filename}: {len(df.columns)} columns, {len(df)} rows")
        return "\n".join(summary)
    except ValueError as e:
        logger.warning(f"Validation error for {filename}: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error inspecting {filename}: {str(e)}")
        raise Exception(f"Failed to inspect {filename}: {str(e)}")

@mcp.tool()
def read_data(filename: str, rows: int = 20) -> str:
    """Reads a file and returns the first few rows (Default 20)."""
    try:
        if rows <= 0:
            raise ValueError("rows parameter must be greater than 0")
        
        path = validate_file_path(filename)
        df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
        
        logger.info(f"Read {min(rows, len(df))} rows from {filename}")
        return df.head(rows).to_string()
    except ValueError as e:
        logger.warning(f"Validation error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error reading {filename}: {str(e)}")
        raise Exception(f"Failed to read {filename}: {str(e)}")

@mcp.tool()
def calculate_pivots(filename: str, index_col: str, value_col: str, agg_func: str = "sum") -> str:
    """
    Creates a pivot table summary.
    
    Args:
        filename: CSV/Excel file to analyze
        index_col: Column to group by
        value_col: Column to aggregate
        agg_func: Aggregation function ('sum', 'mean', 'count', 'min', 'max')
        
    Returns:
        Pivot table as string
    """
    try:
        valid_funcs = ['sum', 'mean', 'count', 'min', 'max', 'median', 'std']
        if agg_func not in valid_funcs:
            raise ValueError(f"agg_func must be one of {valid_funcs}")
        
        path = validate_file_path(filename)
        df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
        
        if index_col not in df.columns:
            raise ValueError(f"Column '{index_col}' not found in {filename}")
        if value_col not in df.columns:
            raise ValueError(f"Column '{value_col}' not found in {filename}")
        
        result = df.groupby(index_col)[value_col].agg(agg_func).reset_index()
        logger.info(f"Pivot created on {filename}: {index_col} by {value_col} ({agg_func})")
        return result.to_string()
    except ValueError as e:
        logger.warning(f"Validation error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error creating pivot: {str(e)}")
        raise Exception(f"Failed to create pivot table: {str(e)}")

@mcp.tool()
def merge_files(file1: str, file2: str, join_key: str, join_type: str = "left") -> str:
    """
    Joins two files together using a common key.
    
    Args:
        file1: First CSV/Excel file
        file2: Second CSV/Excel file
        join_key: Common column name to join on
        join_type: 'left', 'right', 'inner', or 'outer'
        
    Returns:
        Merged data (first 20 rows) as string
    """
    try:
        valid_joins = ['left', 'right', 'inner', 'outer']
        if join_type not in valid_joins:
            raise ValueError(f"join_type must be one of {valid_joins}")
        
        path1 = validate_file_path(file1)
        path2 = validate_file_path(file2)
        
        df1 = pd.read_csv(path1) if file1.endswith('.csv') else pd.read_excel(path1)
        df2 = pd.read_csv(path2) if file2.endswith('.csv') else pd.read_excel(path2)
        
        if join_key not in df1.columns:
            raise ValueError(f"Column '{join_key}' not found in {file1}")
        if join_key not in df2.columns:
            raise ValueError(f"Column '{join_key}' not found in {file2}")
        
        merged_df = pd.merge(df1, df2, on=join_key, how=join_type)
        logger.info(f"Merged {file1} and {file2} on {join_key} ({join_type} join): {len(merged_df)} rows")
        return merged_df.head(20).to_string()
    except ValueError as e:
        logger.warning(f"Validation error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error merging files: {str(e)}")
        raise Exception(f"Failed to merge files: {str(e)}")

@mcp.tool()
def clean_data(filename: str, column: Optional[str] = None, fill_value: float = 0, fill_method: str = "value") -> str:
    """
    Handles missing values in specified columns.
    
    Args:
        filename: CSV/Excel file to clean
        column: Column to clean (if None, cleans all numeric columns)
        fill_value: Value to fill NaN with (for fill_method='value')
        fill_method: 'value', 'mean', 'ffill' (forward fill), 'bfill' (backward fill)
        
    Returns:
        Success message with output path
    """
    try:
        if fill_method not in ['value', 'mean', 'ffill', 'bfill']:
            raise ValueError(f"fill_method must be one of ['value', 'mean', 'ffill', 'bfill']")
        
        path = validate_file_path(filename)
        df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
        
        if column:
            if column not in df.columns:
                raise ValueError(f"Column '{column}' not found in {filename}")
            if fill_method == 'value':
                df[column] = df[column].fillna(fill_value)
            elif fill_method == 'mean':
                df[column] = df[column].fillna(df[column].mean())
            elif fill_method == 'ffill':
                df[column] = df[column].fillna(method='ffill')
            elif fill_method == 'bfill':
                df[column] = df[column].fillna(method='bfill')
        else:
            numeric_cols = df.select_dtypes(include=['number']).columns
            for col in numeric_cols:
                if fill_method == 'value':
                    df[col] = df[col].fillna(fill_value)
                elif fill_method == 'mean':
                    df[col] = df[col].fillna(df[col].mean())
                elif fill_method == 'ffill':
                    df[col] = df[col].fillna(method='ffill')
                elif fill_method == 'bfill':
                    df[col] = df[col].fillna(method='bfill')
        
        output_path = BASE_DIR / "cleaned_data.csv"
        df.to_csv(output_path, index=False)
        logger.info(f"Data cleaned and saved to {output_path.name}")
        return f"Data cleaned! Saved as 'cleaned_data.csv' in your project folder."
    except ValueError as e:
        logger.warning(f"Validation error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error cleaning data: {str(e)}")
        raise Exception(f"Failed to clean data: {str(e)}")

@mcp.tool()
def filter_data(filename: str, column: str, operator: str, value: Any) -> str:
    """
    Filters data based on a condition.
    
    Args:
        filename: CSV/Excel file to filter
        column: Column to filter on
        operator: '==', '!=', '>', '<', '>=', '<=', 'contains', 'in'
        value: Value to compare against
        
    Returns:
        Filtered data as string
    """
    try:
        path = validate_file_path(filename)
        df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
        
        if column not in df.columns:
            raise ValueError(f"Column '{column}' not found in {filename}")
        
        if operator == '==':
            filtered = df[df[column] == value]
        elif operator == '!=':
            filtered = df[df[column] != value]
        elif operator == '>':
            filtered = df[df[column] > value]
        elif operator == '<':
            filtered = df[df[column] < value]
        elif operator == '>=':
            filtered = df[df[column] >= value]
        elif operator == '<=':
            filtered = df[df[column] <= value]
        elif operator == 'contains':
            filtered = df[df[column].astype(str).str.contains(str(value), case=False, na=False)]
        elif operator == 'in':
            filtered = df[df[column].isin(value if isinstance(value, list) else [value])]
        else:
            raise ValueError(f"Unknown operator: {operator}")
        
        logger.info(f"Filtered {filename} on {column} {operator} {value}: {len(filtered)} rows")
        return filtered.head(20).to_string()
    except ValueError as e:
        logger.warning(f"Validation error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error filtering data: {str(e)}")
        raise Exception(f"Failed to filter data: {str(e)}")

@mcp.tool()
def sort_data(filename: str, column: str, ascending: bool = True, rows: int = 20) -> str:
    """
    Sorts data by a column.
    
    Args:
        filename: CSV/Excel file to sort
        column: Column to sort by
        ascending: Sort in ascending order (True) or descending (False)
        rows: Number of rows to return
        
    Returns:
        Sorted data as string
    """
    try:
        if rows <= 0:
            raise ValueError("rows must be greater than 0")
        
        path = validate_file_path(filename)
        df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
        
        if column not in df.columns:
            raise ValueError(f"Column '{column}' not found in {filename}")
        
        sorted_df = df.sort_values(by=column, ascending=ascending)
        logger.info(f"Sorted {filename} by {column} ({'ascending' if ascending else 'descending'})")
        return sorted_df.head(rows).to_string()
    except ValueError as e:
        logger.warning(f"Validation error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error sorting data: {str(e)}")
        raise Exception(f"Failed to sort data: {str(e)}")

@mcp.tool()
def get_statistics(filename: str, column: Optional[str] = None) -> str:
    """
    Calculates statistics (mean, median, std, min, max, count) for numeric columns.
    
    Args:
        filename: CSV/Excel file to analyze
        column: Specific column to analyze (if None, analyzes all numeric columns)
        
    Returns:
        Statistics as string
    """
    try:
        path = validate_file_path(filename)
        df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
        
        if column:
            if column not in df.columns:
                raise ValueError(f"Column '{column}' not found in {filename}")
            stats = df[column].describe()
            logger.info(f"Stats for {column} in {filename}")
            return stats.to_string()
        else:
            stats = df.describe()
            logger.info(f"Stats for all numeric columns in {filename}")
            return stats.to_string()
    except ValueError as e:
        logger.warning(f"Validation error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error calculating statistics: {str(e)}")
        raise Exception(f"Failed to calculate statistics: {str(e)}")

@mcp.tool()
def find_duplicates(filename: str, column: Optional[str] = None) -> str:
    """
    Finds duplicate rows based on specific column(s).
    
    Args:
        filename: CSV/Excel file to check
        column: Specific column to check for duplicates (if None, checks all columns)
        
    Returns:
        Duplicate rows count and sample data
    """
    try:
        path = validate_file_path(filename)
        df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
        
        if column:
            if column not in df.columns:
                raise ValueError(f"Column '{column}' not found in {filename}")
            dup_count = df[column].duplicated().sum()
            duplicates = df[df.duplicated(subset=[column], keep=False)]
        else:
            dup_count = df.duplicated().sum()
            duplicates = df[df.duplicated(keep=False)]
        
        logger.info(f"Found {dup_count} duplicates in {filename}")
        if len(duplicates) == 0:
            return f"No duplicates found in {filename}"
        return f"Found {dup_count} duplicate rows:\n{duplicates.head(10).to_string()}"
    except ValueError as e:
        logger.warning(f"Validation error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error finding duplicates: {str(e)}")
        raise Exception(f"Failed to find duplicates: {str(e)}")

@mcp.tool()
def remove_duplicates(filename: str, column: Optional[str] = None) -> str:
    """
    Removes duplicate rows and saves the result.
    
    Args:
        filename: CSV/Excel file to clean
        column: Specific column to check duplicates (if None, checks all columns)
        
    Returns:
        Success message with row count before/after
    """
    try:
        path = validate_file_path(filename)
        df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
        
        original_count = len(df)
        
        if column:
            if column not in df.columns:
                raise ValueError(f"Column '{column}' not found in {filename}")
            df = df.drop_duplicates(subset=[column])
        else:
            df = df.drop_duplicates()
        
        final_count = len(df)
        removed = original_count - final_count
        
        # Save to a new file
        output_path = BASE_DIR / f"deduped_{filename.split('.')[0]}.csv"
        df.to_csv(output_path, index=False)
        
        logger.info(f"Removed {removed} duplicates from {filename}")
        return f"Removed {removed} duplicate rows ({original_count} → {final_count}). Saved to '{output_path.name}'"
    except ValueError as e:
        logger.warning(f"Validation error: {str(e)}")
        raise
    except Exception as e:
        logger.error(f"Error removing duplicates: {str(e)}")
        raise Exception(f"Failed to remove duplicates: {str(e)}")
    
@mcp.tool()
def generate_chart(filename: str, chart_type: str, x_col: str, y_col: str = None, title: str = "Data Visualization"):
    """
    Generates a chart and saves it as 'chart.png'.
    Types: 'bar', 'line', 'scatter', 'pie', 'histogram'
    """
    path = os.path.join(BASE_DIR, filename)
    df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
    
    plt.figure(figsize=(10, 6))
    
    if chart_type == 'bar':
        sns.barplot(data=df, x=x_col, y=y_col)
    elif chart_type == 'line':
        sns.lineplot(data=df, x=x_col, y=y_col)
    elif chart_type == 'scatter':
        sns.scatterplot(data=df, x=x_col, y=y_col)
    elif chart_type == 'pie':
        df.groupby(x_col).size().plot(kind='pie', autopct='%1.1f%%')
    elif chart_type == 'histogram':
        sns.histplot(data=df, x=x_col, kde=True)
        
    plt.title(title)
    output_path = os.path.join(BASE_DIR, "chart.png")
    plt.savefig(output_path)
    plt.close()
    
    return f"Chart generated and saved to {output_path}. You can now view it in your folder!"

@mcp.tool()
def create_excel_report(filename: str, report_name: str, chart_title: str):
    """Creates a native, editable Excel report with a chart."""
    path = os.path.join(BASE_DIR, filename)
    df = pd.read_csv(path) if filename.endswith('.csv') else pd.read_excel(path)
    
    output_path = os.path.join(BASE_DIR, f"{report_name}.xlsx")
    writer = pd.ExcelWriter(output_path, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Data', index=False)

    # Access the xlsxwriter workbook and worksheet objects
    workbook  = writer.book
    worksheet = writer.sheets['Data']

    # Create a chart object (e.g., Column chart)
    chart = workbook.add_chart({'type': 'column'})

    # Configure the series (using the data from the sheet)
    # Assumes Column A is Labels and Column B is Values
    chart.add_series({
        'name':       chart_title,
        'categories': ['Data', 1, 0, len(df), 0],
        'values':     ['Data', 1, 1, len(df), 1],
    })

    chart.set_title({'name': chart_title})
    worksheet.insert_chart('D2', chart)
    writer.close()
    
    return f"Report created: {output_path}. You can open this in Excel to edit the chart!"


if __name__ == "__main__":
    mcp.run()
