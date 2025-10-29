# Open-WebUI Integration Setup Guide

This guide provides step-by-step instructions for integrating the Exel MCP Excel server with Open-WebUI.

## Prerequisites

- ✅ Exel MCP server running on `http://localhost:9080`
- ✅ Open-WebUI running on `http://localhost:8091`
- ✅ Llama 3.1 8B model loaded in Ollama
- ✅ Docker network connectivity between containers

## Integration Test Results

✅ **MCP Server Connection**: Working
✅ **Excel File Creation**: Working
✅ **Tool Calling**: Verified functional

## Method 1: Manual Tool Configuration in Open-WebUI

### Step 1: Access Open-WebUI Admin Panel

1. Open Open-WebUI in your browser: `http://localhost:8091`
2. Click on your profile icon (top right) → **Admin Panel**
3. Navigate to **Tools** section

### Step 2: Create Excel Creation Tool

Click **"Add Tool"** and configure:

#### Basic Information
- **Name**: `create_excel_file`
- **Description**:
  ```
  Creates an Excel file with the given data. Use this tool when users want to create spreadsheets, tables, or Excel files with data. Provide filename, headers array, and 2D data array.
  ```

#### Parameters Schema
```json
{
  "type": "object",
  "properties": {
    "filename": {
      "type": "string",
      "description": "Name of the Excel file to create (will auto-append .xlsx if missing)"
    },
    "headers": {
      "type": "array",
      "items": {
        "type": "string"
      },
      "description": "List of column headers for the spreadsheet"
    },
    "sheet_data": {
      "type": "array",
      "items": {
        "type": "array",
        "items": {
          "type": "string"
        }
      },
      "description": "2D array of data rows (each inner array is a row)"
    },
    "sheet_name": {
      "type": "string",
      "description": "Name of the worksheet (optional, defaults to 'Sheet1')",
      "default": "Sheet1"
    }
  },
  "required": ["filename", "headers", "sheet_data"]
}
```

#### API Configuration
- **URL**: `http://host.docker.internal:8001/mcp`
- **Method**: `POST`
- **Headers**:
  ```
  Content-Type: application/json
  Accept: application/json, text/event-stream
  ```
- **Request Body**:
  ```json
  {
    "jsonrpc": "2.0",
    "id": "{{id}}",
    "method": "tools/call",
    "params": {
      "name": "create_excel_file",
      "arguments": {
        "filename": "{{filename}}",
        "headers": {{headers}},
        "sheet_data": {{sheet_data}},
        "sheet_name": "{{sheet_name}}"
      }
    }
  }
  ```

#### Response Mapping
- **Success Path**: `result.content[0].text`
- **Error Path**: `error.message`

### Step 3: Create Excel Info Tool

Add another tool for getting file information:

#### Basic Information
- **Name**: `get_excel_info`
- **Description**:
  ```
  Gets information about an existing Excel file. Use this to check file details or analyze existing spreadsheets.
  ```

#### Parameters Schema
```json
{
  "type": "object",
  "properties": {
    "filename": {
      "type": "string",
      "description": "Name of the Excel file to analyze"
    }
  },
  "required": ["filename"]
}
```

#### API Configuration
- **URL**: `http://host.docker.internal:8001/mcp`
- **Method**: `POST`
- **Headers**: (Same as above)
- **Request Body**:
  ```json
  {
    "jsonrpc": "2.0",
    "id": "{{id}}",
    "method": "tools/call",
    "params": {
      "name": "get_excel_info",
      "arguments": {
        "filename": "{{filename}}"
      }
    }
  }
  ```

### Step 4: Enable Tools

1. Make sure both tools are **enabled**
2. Set appropriate **permissions** if needed
3. **Save** the configuration

## Method 2: Using System Prompts with Ollama Tools

If Open-WebUI supports Ollama's native tool calling:

### Step 1: Create a Custom System Prompt

In your Open-WebUI chat, set a custom system prompt:

```
You are an expert AI assistant specializing in Microsoft Excel. Your primary function is to help users create and manipulate Excel spreadsheets by calling the `create_excel_file` tool.

When a user asks you to create a spreadsheet, analyze their request and call the `create_excel_file` tool with these exact parameters:

- `filename`: A string ending in `.xlsx` (e.g., "employees.xlsx")
- `headers`: An array of strings for column headers (e.g., ["Name", "Department", "Salary"])
- `sheet_data`: A 2D array where each inner array is a row of data. For example: [["John Doe", "Engineering", "75000"], ["Jane Smith", "Sales", "65000"]]

CRITICAL: `sheet_data` must be an array of arrays (one array per row), NOT a single array. Each row is an array of strings.

Example for "Create a spreadsheet of employees with names and salaries":

{
  "filename": "employees.xlsx",
  "headers": ["Name", "Salary"],
  "sheet_data": [
    ["Alice Johnson", "55000"],
    ["Bob Wilson", "62000"],
    ["Carol Brown", "58000"]
  ]
}

The `sheet_data` has square brackets around the entire data, and each row has its own square brackets.

Do not include any text before or after the tool call. Only output the tool call JSON.
```

### Step 2: Configure Ollama Tools

The tools are automatically available when using Ollama with tool calling enabled.

## Testing the Integration

### Test Case 1: Basic Excel Creation

**User Input:**
```
Create a spreadsheet with information about 3 products: Widget A costs $10, Widget B costs $15, and Widget C costs $20.
```

**Expected Behavior:**
- LLM analyzes the request
- Calls `create_excel_file` tool
- Excel file `products.xlsx` is created with headers ["Product", "Price"] and appropriate data

### Test Case 2: File Information

**User Input:**
```
Tell me about the products.xlsx file I just created.
```

**Expected Behavior:**
- LLM calls `get_excel_info` tool
- Returns file information (size, existence, etc.)

### Test Case 3: Complex Data

**User Input:**
```
Create an employee database with names, departments, salaries, and hire dates for 4 employees.
```

**Expected Behavior:**
- LLM generates appropriate headers and sample data
- Creates properly formatted Excel file

## Troubleshooting

### Common Issues

1. **"host.docker.internal" not resolved**
   - **Linux**: Replace with `172.17.0.1` or your Docker gateway IP
   - **macOS/Windows**: Should work as-is
   - **Alternative**: Use container names if on same network

2. **Tool not appearing in chat**
   - Ensure tools are enabled in Admin Panel
   - Check that the model supports tool calling
   - Verify API endpoint accessibility

3. **MCP server errors**
   - Check container logs: `docker logs exel-mcp-server`
    - Verify server is running: `curl http://localhost:9080/mcp`
   - Check network connectivity

4. **Invalid tool responses**
   - Ensure JSON-RPC format is correct
   - Check session ID handling
   - Verify parameter validation

### Debug Commands

```bash
# Check MCP server status
curl -X POST http://localhost:9080/mcp \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}'

# Test tool calling
curl -X POST http://localhost:9080/mcp \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -H "mcp-session-id: YOUR_SESSION_ID" \
  -d '{"jsonrpc":"2.0","id":2,"method":"tools/call","params":{"name":"create_excel_file","arguments":{"filename":"debug.xlsx","headers":["Test"],"sheet_data":[["Data"]]}}}'

# Check created files
ls -la output/
```

### Logs

- **Open-WebUI**: Check container logs with `docker logs open-webui`
- **MCP Server**: Check with `docker logs exel-mcp-server`
- **Ollama**: Check with `docker logs ollama` (if running in container)

## Advanced Configuration

### Custom Tool Templates

You can create specialized tools for different use cases:

- **Financial Reports**: Pre-configured headers for balance sheets
- **Inventory Tracking**: Product catalog templates
- **Survey Results**: Response analysis formats

### Batch Operations

For multiple file operations, consider creating workflow tools that chain multiple calls.

### File Management

- Files are saved in the `./output/` directory
- Implement cleanup policies for production
- Consider cloud storage integration for web access

## Performance Optimization

- **Caching**: Cache frequently used templates
- **Batch Processing**: Handle multiple rows efficiently
- **Compression**: Enable Excel compression for large files
- **Monitoring**: Track usage and performance metrics

## Security Considerations

- **Input Validation**: All inputs are validated on the MCP server
- **File Access**: Restricted to output directory
- **Network Security**: Ensure proper firewall rules
- **Authentication**: Consider adding API keys for production

## Next Steps

1. ✅ Complete basic integration
2. 🔄 Test with real users
3. 📊 Gather feedback and metrics
4. 🔧 Optimize based on usage patterns
5. 🚀 Deploy to production environment

The integration is now ready for real-world testing!