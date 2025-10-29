# Excel Tools for Open-WebUI - Complete Guide

## Overview

This guide provides comprehensive documentation for the Excel tools available in Open-WebUI through the Exel MCP server integration. These tools allow users to create, manipulate, and analyze Excel spreadsheets using natural language commands.

## Available Tools

### 1. create_excel_file

**Purpose**: Create new Excel spreadsheets with custom data, headers, and formatting.

**When to Use**: 
- Users want to create spreadsheets from data
- Converting tabular data to Excel format
- Creating structured data files

**Parameters**:
- `filename` (string, required): Excel filename (auto-appends .xlsx if missing)
- `headers` (array of strings, required): List of column headers
- `sheet_data` (array of arrays, required): 2D array of data rows
- `sheet_name` (string, optional): Worksheet name (default: "Sheet1")
- `formatting` (object, optional): Formatting options
  - `auto_width` (boolean): Auto-adjust column widths (default: true)
  - `header_bold` (boolean): Make headers bold (default: true)
  - `header_center` (boolean): Center headers (default: true)

**Example Usage**:
```
User: "Create a spreadsheet of employees with names, departments, and salaries"
```

**Tool Call**:
```json
{
  "filename": "employees.xlsx",
  "headers": ["Name", "Department", "Salary"],
  "sheet_data": [
    ["John Doe", "Engineering", "75000"],
    ["Jane Smith", "Sales", "65000"],
    ["Bob Johnson", "Marketing", "60000"]
  ],
  "sheet_name": "Employee Data"
}
```

### 2. get_excel_info

**Purpose**: Analyze existing Excel files for structure, content, and metadata.

**When to Use**:
- Users want to understand an existing Excel file
- Checking file properties and structure
- Analyzing spreadsheet content

**Parameters**:
- `filename` (string, required): Excel file to analyze

**Example Usage**:
```
User: "Analyze the sales.xlsx file and tell me about its structure"
```

**Tool Call**:
```json
{
  "filename": "sales.xlsx"
}
```

### 3. create_excel_chart

**Purpose**: Add charts and graphs to existing Excel files.

**When to Use**:
- Users want to visualize data with charts
- Creating data visualizations
- Adding graphs to existing spreadsheets

**Parameters**:
- `filename` (string, required): Target Excel file to add chart to
- `chart_type` (string, required): Type of chart ("bar", "line", "pie", "scatter", "area")
- `data_range` (string, required): Cell range for chart data (e.g., "A1:C10")
- `title` (string, optional): Chart title
- `sheet_name` (string, optional): Worksheet name (default: first sheet)

**Example Usage**:
```
User: "Add a bar chart to sales.xlsx showing monthly revenue from A1:B12"
```

**Tool Call**:
```json
{
  "filename": "sales.xlsx",
  "chart_type": "bar",
  "data_range": "A1:B12",
  "title": "Monthly Revenue Chart"
}
```

### 4. format_excel_cells

**Purpose**: Apply formatting to Excel cells including colors, borders, fonts, and styles.

**When to Use**:
- Users want to enhance spreadsheet appearance
- Applying consistent formatting
- Styling specific cells or ranges

**Parameters**:
- `filename` (string, required): Target Excel file
- `cell_range` (string, required): Cell range in A1:B5 format (e.g., "A1:C10")
- `formatting` (object, required): Formatting options
  - `bold` (boolean): Make text bold
  - `italic` (boolean): Make text italic
  - `underline` (boolean): Underline text
  - `font_size` (integer): Font size (6-72)
  - `font_color` (string): Font color (hex code)
  - `background_color` (string): Background color (hex code)
  - `border` (boolean): Add borders
  - `border_color` (string): Border color (hex code)
  - `alignment` (string): Text alignment ("left", "center", "right", "justify")
- `sheet_name` (string, optional): Worksheet name (default: first sheet)

**Example Usage**:
```
User: "Format the header row in employees.xlsx to be bold with blue background"
```

**Tool Call**:
```json
{
  "filename": "employees.xlsx",
  "cell_range": "A1:C1",
  "formatting": {
    "bold": true,
    "background_color": "0066CC",
    "font_color": "FFFFFF",
    "alignment": "center"
  }
}
```

### 5. import_csv_to_excel

**Purpose**: Convert CSV files to Excel format with proper formatting and structure.

**When to Use**:
- Users want to convert CSV data to Excel
- Importing external data into Excel format
- Converting delimited text files

**Parameters**:
- `csv_file` (string, required): Source CSV file path or content
- `excel_file` (string, required): Target Excel filename
- `delimiter` (string, optional): CSV delimiter character (default: ",")
- `has_headers` (boolean, optional): Whether CSV has header row (default: true)
- `sheet_name` (string, optional): Worksheet name (default: "Sheet1")

**Example Usage**:
```
User: "Convert the data.csv file to Excel format with headers"
```

**Tool Call**:
```json
{
  "csv_file": "data.csv",
  "excel_file": "converted_data.xlsx",
  "delimiter": ",",
  "has_headers": true,
  "sheet_name": "Imported Data"
}
```

### 6. export_excel_to_csv

**Purpose**: Export Excel worksheets to CSV format.

**When to Use**:
- Users want to convert Excel data to CSV
- Exporting data for other applications
- Creating delimited text files

**Parameters**:
- `excel_file` (string, required): Source Excel file
- `csv_file` (string, required): Target CSV filename
- `sheet_name` (string, optional): Worksheet name to export (default: first sheet)
- `delimiter` (string, optional): CSV delimiter character (default: ",")
- `include_headers` (boolean, optional): Include headers in CSV (default: true)

**Example Usage**:
```
User: "Export the Sales sheet from report.xlsx to CSV format"
```

**Tool Call**:
```json
{
  "excel_file": "report.xlsx",
  "csv_file": "sales_export.csv",
  "sheet_name": "Sales",
  "delimiter": ",",
  "include_headers": true
}
```

## Natural Language Examples

### Basic Excel Creation

**User Request**: "Create a spreadsheet of products with names, prices, and categories"
**LLM Response**: 
```json
{
  "filename": "products.xlsx",
  "headers": ["Product Name", "Price", "Category"],
  "sheet_data": [
    ["Laptop", "999.99", "Electronics"],
    ["Coffee Mug", "12.99", "Kitchen"],
    ["Notebook", "5.99", "Office"]
  ]
}
```

### Data Analysis

**User Request**: "Analyze the sales_data.xlsx file and tell me what's in it"
**LLM Response**:
```json
{
  "filename": "sales_data.xlsx"
}
```

### Chart Creation

**User Request**: "Add a pie chart to budget.xlsx showing expense categories from A1:B10"
**LLM Response**:
```json
{
  "filename": "budget.xlsx",
  "chart_type": "pie",
  "data_range": "A1:B10",
  "title": "Expense Categories"
}
```

### Formatting

**User Request**: "Make the headers in report.xlsx bold with green background"
**LLM Response**:
```json
{
  "filename": "report.xlsx",
  "cell_range": "A1:D1",
  "formatting": {
    "bold": true,
    "background_color": "00FF00",
    "alignment": "center"
  }
}
```

### CSV Conversion

**User Request**: "Convert my contacts.csv file to Excel format"
**LLM Response**:
```json
{
  "csv_file": "contacts.csv",
  "excel_file": "contacts.xlsx",
  "has_headers": true
}
```

## Error Handling

### Common Errors and Solutions

1. **File Not Found**
   - Error: "Excel file not found: filename.xlsx"
   - Solution: Ensure the file exists in the output directory

2. **Invalid Data Range**
   - Error: "Invalid data range 'A1:Z999': ..."
   - Solution: Use valid cell ranges (e.g., "A1:C10")

3. **Data Validation Errors**
   - Error: "Row 3 has 5 columns, expected 4"
   - Solution: Ensure all rows have the same number of columns as headers

4. **Missing Required Parameters**
   - Error: "filename, chart_type, and data_range are required"
   - Solution: Provide all required parameters

## Best Practices

### For Users
1. **Be Specific**: Provide clear filenames and data ranges
2. **Check Data**: Ensure data consistency before creating spreadsheets
3. **Use Descriptive Names**: Use meaningful filenames and sheet names
4. **Validate Ranges**: Double-check cell ranges when formatting or charting

### For LLM Integration
1. **Validate Input**: Always validate user input before calling tools
2. **Handle Errors**: Provide helpful error messages
3. **Confirm Actions**: Let users know what was accomplished
4. **Suggest Next Steps**: Recommend related actions when appropriate

## Setup Instructions

### 1. MCP Server Setup
```bash
# Start the MCP server
cd /path/to/exel_mcp
python src/main.py
```

### 2. Open-WebUI Configuration
1. Access Open-WebUI Admin Panel
2. Navigate to Tools section
3. Add new tool configuration
4. Use the enhanced JSON configuration from `config/openwebui_tools_enhanced.json`
5. Set base URL to `http://host.docker.internal:8001`
6. Test tools with sample requests

### 3. Testing
```bash
# Test basic functionality
python tests/test_openwebui_integration.py

# Test specific tools
python tests/test_new_tools.py
```

## Troubleshooting

### Connection Issues
- **Problem**: Tools not responding
- **Solution**: Check MCP server is running on port 8001
- **Command**: `curl http://localhost:8001/mcp`

### File Access Issues
- **Problem**: Cannot find created files
- **Solution**: Check output directory permissions
- **Command**: `ls -la output/`

### Performance Issues
- **Problem**: Slow response times
- **Solution**: Monitor server logs and optimize data sizes
- **Command**: Check logs for bottlenecks

## Advanced Features

### Multi-Sheet Support
Future versions will support:
- Creating multiple worksheets in one file
- Copying data between sheets
- Cross-sheet formulas and references

### Advanced Charting
Planned enhancements:
- Custom chart colors and styles
- Combination charts
- Dynamic chart ranges
- Chart templates

### Data Validation
Upcoming features:
- Cell validation rules
- Dropdown lists
- Input restrictions
- Custom validation messages

## API Reference

### Base Configuration
```json
{
  "base_url": "http://host.docker.internal:8001",
  "timeout": 30,
  "retry_attempts": 3,
  "headers": {
    "Content-Type": "application/json",
    "Accept": "application/json, text/event-stream"
  }
}
```

### Response Format
All tools return JSON responses with:
- `result`: Success message or data
- `error`: Error message (if applicable)
- `file_path`: Path to created/modified file

### Rate Limiting
- Maximum 100 requests per minute
- Large files (>10MB) may take longer
- Concurrent requests limited to 5

## Support

For issues and questions:
1. Check MCP server logs
2. Verify tool configuration
3. Test with simple examples
4. Review error messages
5. Consult this documentation

## Version History

- **v1.0**: Basic Excel creation and info tools
- **v1.1**: Added chart creation and formatting
- **v1.2**: Added CSV import/export functionality
- **v1.3**: Enhanced error handling and validation
- **v2.0**: Comprehensive tool suite with advanced features