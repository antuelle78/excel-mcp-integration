# Excel Tools - Open-WebUI Integration Guide

## 🎯 Overview

This guide provides complete instructions for integrating Excel MCP Server tools with Open-WebUI using the Python app approach.

## 📁 Files Required

### Main Integration File
- **`excel_tools_openwebui.py`** - Complete Open-WebUI compatible Python app with all Excel tools

### Test Files
- **`test_openwebui_integration.py`** - Comprehensive test suite for Open-WebUI integration

## 🚀 Quick Setup

### 1. Start MCP Server
```bash
cd /home/ghost/bin/docker/exel_mcp
source venv/bin/activate
python src/main.py
```
Server will start on: `http://localhost:8002/mcp`

### 2. Install Dependencies
```bash
source venv/bin/activate
pip install httpx
```

### 3. Test Integration
```bash
python test_openwebui_integration.py
```

## 🔧 Open-WebUI Integration

### Method 1: Direct Python App Import (Recommended)

Copy the entire `excel_tools_openwebui.py` file content into Open-WebUI's custom tool interface. The file includes:

```python
"""title: 'Excel Tools'
author: 'Exel MCP Server'
description: 'A comprehensive set of tools to create, manipulate, and analyze Excel files with charts, formatting, and CSV integration.'
version: '1.0.0'
requirements: httpx
"""
```

### Method 2: File-based Integration

1. Place `excel_tools_openwebui.py` in Open-WebUI's tools directory
2. Update the server URL in the file if needed:
   ```python
   self.mcp_server_url = "http://YOUR_SERVER_IP:8002/mcp"
   ```

## 🛠️ Available Tools

### Core Excel Operations

#### `create_excel_file`
Creates Excel files with data and formatting
```python
await tools.create_excel_file(
    filename="data.xlsx",
    headers=["Name", "Age", "City"],
    sheet_data=[["Alice", 25, "NYC"], ["Bob", 30, "LA"]],
    sheet_name="People",
    formatting={"header_bold": True}
)
```

#### `get_excel_info`
Analyzes existing Excel files
```python
await tools.get_excel_info(filename="data.xlsx")
```

### Advanced Features

#### `create_excel_chart`
Adds charts to Excel files
```python
await tools.create_excel_chart(
    filename="data.xlsx",
    chart_type="bar",  # bar, line, pie, scatter
    data_range="A1:C10",
    title="Sales Chart"
)
```

#### `format_excel_cells`
Applies formatting to cell ranges
```python
await tools.format_excel_cells(
    filename="data.xlsx",
    cell_range="A1:C1",
    formatting={
        "bold": True,
        "background_color": "4472C4",
        "font_color": "FFFFFF"
    }
)
```

### CSV Integration

#### `import_csv_to_excel`
Converts CSV to Excel
```python
await tools.import_csv_to_excel(
    csv_file="name,age,city\nAlice,25,NYC\nBob,30,LA",
    excel_file="converted.xlsx",
    delimiter=",",
    has_headers=True
)
```

#### `export_excel_to_csv`
Exports Excel to CSV
```python
await tools.export_excel_to_csv(
    excel_file="data.xlsx",
    csv_file="exported.csv",
    sheet_name="Sheet1",
    delimiter=","
)
```

### Convenience Methods

#### `create_sales_report`
Creates sales reports with charts
```python
await tools.create_sales_report(
    filename="sales.xlsx",
    sales_data=[
        ["Jan", 5000, 3000, 2000],
        ["Feb", 6000, 3500, 2500]
    ],
    include_chart=True
)
```

#### `create_employee_directory`
Creates formatted employee directories
```python
await tools.create_employee_directory(
    filename="employees.xlsx",
    employee_data=[
        ["John", "Engineering", "john@company.com", "555-0101"],
        ["Jane", "Marketing", "jane@company.com", "555-0102"]
    ]
)
```

## 📊 Formatting Options

### Header Formatting
```python
formatting = {
    "header_bold": True,
    "header_background": "4472C4",  # Blue
    "header_font_color": "FFFFFF",   # White
    "alternate_row_colors": True
}
```

### Cell Formatting
```python
formatting = {
    "bold": True,
    "italic": False,
    "font_color": "000000",        # Black
    "background_color": "FFFF00",   # Yellow
    "alignment": "center",          # left, center, right
    "font_size": 12
}
```

## 🎨 Chart Types

- **bar** - Bar charts (default)
- **line** - Line charts
- **pie** - Pie charts
- **scatter** - Scatter plots
- **area** - Area charts

## 🔍 Error Handling

The integration includes comprehensive error handling:

- Connection errors
- MCP session management
- File validation
- Data format validation
- Server response parsing

## 📝 Example Usage in Open-WebUI

Once integrated, users can ask natural language questions:

**User**: "Create a sales report for Q1 data"
**AI**: Uses `create_sales_report` tool

**User**: "Add a chart to my Excel file"
**AI**: Uses `create_excel_chart` tool

**User**: "Format the headers in blue"
**AI**: Uses `format_excel_cells` tool

## 🚀 Production Deployment

### Security Considerations
1. **Network Access**: Ensure MCP server is accessible from Open-WebUI
2. **File Paths**: Configure appropriate output directories
3. **Rate Limiting**: Implement if needed for high-usage scenarios

### Scaling
1. **Load Balancing**: Multiple MCP server instances
2. **File Storage**: Network storage for generated files
3. **Monitoring**: Log analysis and health checks

### Configuration
```python
# Update server URL for production
self.mcp_server_url = "http://your-server:8002/mcp"

# Add authentication headers if needed
self.headers["Authorization"] = "Bearer your-token"
```

## 🧪 Testing

Run the comprehensive test suite:
```bash
python test_openwebui_integration.py
```

Tests include:
- ✅ Basic Excel file creation
- ✅ Chart generation
- ✅ Cell formatting
- ✅ CSV import/export
- ✅ Convenience methods
- ✅ Error handling

## 📁 Generated Files

All generated Excel files are saved to the `output/` directory:
- `openwebui_test.xlsx` - Basic test file
- `sales_report.xlsx` - Sales report with chart
- `employees.xlsx` - Employee directory
- `products_test.xlsx` - CSV import example
- `export_test.csv` - CSV export example

## 🔧 Troubleshooting

### Connection Issues
```bash
# Check if MCP server is running
curl http://localhost:8002/mcp

# Check server logs
tail -f server.log
```

### Session Issues
The integration handles MCP session initialization automatically. If you encounter session errors:
1. Restart the MCP server
2. Check network connectivity
3. Verify server URL configuration

### File Access Issues
Ensure the `output/` directory exists and is writable:
```bash
mkdir -p output
chmod 755 output
```

## 🎉 Success Metrics

- ✅ All 6 core Excel tools working
- ✅ 2 convenience methods implemented
- ✅ Full Open-WebUI compatibility
- ✅ Comprehensive error handling
- ✅ Session management
- ✅ Production-ready code

---

**Status**: ✅ READY FOR OPEN-WEBUI DEPLOYMENT  
**Compatibility**: Open-WebUI Custom Tools Format  
**Next Step**: Import `excel_tools_openwebui.py` into Open-WebUI