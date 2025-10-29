# Open-WebUI Tool Integration for Exel MCP Server

This guide explains how to integrate the Exel MCP Excel creation server with Open-WebUI.

## Method 1: Using Open-WebUI's Built-in Tool System

### Step 1: Access Open-WebUI Admin Panel

1. Open Open-WebUI at `http://localhost:8091`
2. Go to **Settings** → **Admin Panel** → **Tools**

### Step 2: Add Excel Creation Tool

Create a new tool with the following configuration:

**Tool Name:** `create_excel_file`

**Description:**
```
Creates an Excel file with the given data. Use this tool when users want to create spreadsheets, tables, or Excel files with data. You must provide:
- filename: Name of the Excel file to create
- headers: Array of column headers
- sheet_data: 2D array where each inner array is a row of data
```

**Parameters Schema:**
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

**API Configuration:**

- **URL:** `http://host.docker.internal:8001/mcp`
- **Method:** `POST`
- **Headers:**
  ```
  Content-Type: application/json
  Accept: application/json, text/event-stream
  ```
- **Request Body Template:**
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

**Response Mapping:**
- **Success Path:** `result.content[0].text`
- **Error Path:** `error.message`

### Step 3: Add Excel Info Tool

Create another tool for getting Excel file information:

**Tool Name:** `get_excel_info`

**Description:**
```
Gets information about an existing Excel file. Use this to check file details or analyze existing spreadsheets.
```

**Parameters Schema:**
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

**API Configuration:** (Same as above, but change the tool name in the body)

- **URL:** `http://host.docker.internal:8001/mcp`
- **Method:** `POST`
- **Headers:** (Same as above)
- **Request Body Template:**
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

## Method 2: Using Ollama Tools Directly

If Open-WebUI supports Ollama's native tool calling, you can configure the tools directly in your chat session.

### Step 1: Create a System Prompt

Use this system prompt in your Open-WebUI chat:

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

### Step 2: Configure Tools in Ollama

The tools are already configured in Ollama when you make API calls. Open-WebUI should automatically detect and use them if the model supports tool calling.

## Testing the Integration

### Test 1: Simple Excel Creation

**User Message:** "Create a spreadsheet with 3 employees: John making $50k, Jane making $60k, and Bob making $70k."

**Expected Result:** The LLM should call the `create_excel_file` tool and create an Excel file.

### Test 2: Check File Information

**User Message:** "Tell me about the employees.xlsx file I just created."

**Expected Result:** The LLM should call the `get_excel_info` tool to get file details.

## Troubleshooting

### Common Issues

1. **"host.docker.internal" not resolved**
   - On Linux, replace `host.docker.internal` with `172.17.0.1` or your host IP
   - Or use the container name if running in the same Docker network

2. **Tool not appearing in Open-WebUI**
   - Ensure the tool is enabled in the Admin Panel
   - Check that the model supports tool calling
   - Verify the API endpoint is accessible

3. **MCP server connection errors**
   - Check that the Exel MCP server is running on port 8001
   - Verify network connectivity between containers
   - Check server logs for errors

4. **Invalid tool responses**
   - Ensure the JSON-RPC format is correct
   - Check that session IDs are handled properly
   - Verify parameter types match the schema

### Logs and Debugging

- **Open-WebUI Logs:** Check the Open-WebUI container logs
- **MCP Server Logs:** Check the exel-mcp-server container logs
- **Network:** Use `docker network ls` and `docker network inspect` to verify connectivity

## Advanced Configuration

### Custom System Prompts

You can create custom system prompts for different use cases:

- **Business Reports:** Focus on financial data and calculations
- **Data Analysis:** Include pivot table and chart creation
- **Inventory Management:** Product catalogs and stock tracking

### Multiple Worksheets

The current implementation supports single worksheets. For multiple worksheets, you would need to make multiple tool calls or extend the tool.

### File Management

- Excel files are saved in the `./output` directory
- Files are accessible via the container volume mount
- Consider implementing file cleanup for production use