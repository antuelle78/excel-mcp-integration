# Open-WebUI Pipe Function Setup for Excel Assistant

This guide provides step-by-step instructions for installing the Excel Assistant Pipe Function in Open-WebUI.

## Overview

The Excel Assistant Pipe Function creates a custom model in Open-WebUI that can generate Excel spreadsheets through natural language commands. It integrates with the Exel MCP server to provide seamless Excel creation capabilities.

## Prerequisites

- ✅ Open-WebUI installed and running
- ✅ Exel MCP server running on `http://localhost:9080` (or configured endpoint)
- ✅ Python environment with required dependencies

## Installation Methods

### Method 1: Manual Installation (Recommended)

#### Step 1: Prepare the Pipe Function

1. Copy the `excel_assistant_pipe.py` file to your Open-WebUI functions directory
2. Install required dependencies:

```bash
pip install pydantic requests
```

#### Step 2: Access Open-WebUI Admin Panel

1. Open Open-WebUI at `http://localhost:8091`
2. Navigate to **Admin Panel** → **Functions**
3. Click **"Add Function"**

#### Step 3: Configure the Pipe Function

**Basic Information:**
- **Name:** `Excel Assistant`
- **Type:** `Pipe`
- **Meta:** Leave default

**Function Code:**
Paste the entire contents of `excel_assistant_pipe.py` into the code editor.

**Valves Configuration:**
The function includes built-in valves for configuration. You can adjust:

- **MCP_BASE_URL:** URL of your Exel MCP server (default: `http://host.docker.internal:8001`)
- **MODEL_NAME:** Display name in model selector
- **DEFAULT_SHEET_NAME:** Default worksheet name
- **MAX_ROWS/MAX_COLS:** Size limits

#### Step 4: Enable and Test

1. **Enable** the function
2. The "Excel Assistant" model should now appear in your model selector
3. Test with: *"Create a spreadsheet of employees with names and salaries"*

### Method 2: Using Open-WebUI's Function Import

If Open-WebUI supports function imports:

1. Go to **Admin Panel** → **Functions**
2. Click **"Import Function"**
3. Upload the `excel_assistant_pipe.py` file
4. Configure valves as needed
5. Enable the function

## Configuration Options

### Valves Settings

| Valve | Default | Description |
|-------|---------|-------------|
| `MCP_BASE_URL` | `http://host.docker.internal:8001` | Exel MCP server endpoint |
| `MODEL_NAME` | `Excel Assistant` | Display name in UI |
| `MODEL_ID` | `excel-assistant` | Unique model identifier |
| `DEFAULT_SHEET_NAME` | `Sheet1` | Default worksheet name |
| `MAX_ROWS` | `1000` | Maximum rows per spreadsheet |
| `MAX_COLS` | `50` | Maximum columns per spreadsheet |

### Network Configuration

**For Docker environments:**
- Use `http://host.docker.internal:8001` to access MCP server from Open-WebUI container
- Ensure both containers are on the same Docker network

**For local installations:**
- Use `http://localhost:9080` if both services run locally
- Adjust firewall settings if needed

## Testing the Integration

### Test Commands

Try these natural language commands:

1. **"Create a spreadsheet of employees"**
   - Expected: Creates `employees.xlsx` with sample employee data

2. **"Make a product inventory table"**
   - Expected: Creates `products.xlsx` with product catalog

3. **"Generate a sales report"**
   - Expected: Creates `sales.xlsx` with sales transaction data

4. **"Create a data table"**
   - Expected: Creates `data.xlsx` with generic sample data

### Verification Steps

1. **Check Model Availability:**
   - "Excel Assistant" should appear in model selector

2. **Test Basic Response:**
   - Ask: *"Hello, what can you do?"*
   - Expected: Guidance on Excel creation capabilities

3. **Test Excel Creation:**
   - Ask: *"Create a spreadsheet of 3 employees"*
   - Expected: Success message with file creation confirmation

4. **Verify File Creation:**
   - Check the MCP server's `./output/` directory for created Excel files

## Troubleshooting

### Common Issues

#### 1. "Model not appearing in selector"
**Solution:**
- Ensure the function is enabled in Admin Panel
- Check function logs for errors
- Restart Open-WebUI if needed

#### 2. "Failed to initialize MCP session"
**Symptoms:** Errors about session initialization
**Solutions:**
- Verify MCP server is running: `curl http://localhost:9080/mcp`
- Check network connectivity between containers
- Update `MCP_BASE_URL` valve to correct endpoint

#### 3. "host.docker.internal not resolved"
**Symptoms:** Connection refused errors
**Solutions:**
- For Linux: Use `172.17.0.1` (Docker gateway IP)
- For macOS/Windows: `host.docker.internal` should work
- Use container names if on same Docker network

#### 4. "Excel file not created"
**Symptoms:** Success message but no file
**Solutions:**
- Check MCP server logs: `docker logs exel-mcp-server`
- Verify output directory permissions
- Ensure filename validation passes

#### 5. "Pydantic validation errors"
**Symptoms:** Errors about data types
**Solutions:**
- Check that sheet_data is properly formatted as array of arrays
- Ensure headers is an array of strings
- Verify data doesn't exceed MAX_ROWS/MAX_COLS limits

### Debug Commands

```bash
# Test MCP server connectivity
curl -X POST http://localhost:9080/mcp \
  -H "Content-Type: application/json" \
  -H "Accept: application/json, text/event-stream" \
  -d '{"jsonrpc":"2.0","id":"test","method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"test","version":"1.0"}}}'

# Check Open-WebUI function logs
docker logs open-webui 2>&1 | grep -i excel

# Test pipe function directly
cd /path/to/openwebui/functions
python3 -c "
from excel_assistant_pipe import Pipe
pipe = Pipe()
result = pipe.pipe({'messages': [{'role': 'user', 'content': 'Create employee spreadsheet'}]})
print('Result:', result)
"
```

## Advanced Configuration

### Custom Data Patterns

The pipe function includes pattern matching for common requests. You can extend `_extract_excel_request()` method to handle more specific use cases:

```python
# Add custom patterns in the Pipe class
def _extract_excel_request(self, user_message: str):
    message = user_message.lower()

    # Add your custom patterns here
    if 'budget' in message and 'department' in message:
        return {
            'filename': 'department_budget.xlsx',
            'headers': ['Department', 'Budget', 'Spent', 'Remaining'],
            'sheet_data': [
                ['Engineering', '500000', '350000', '150000'],
                ['Sales', '300000', '280000', '20000'],
                ['Marketing', '200000', '180000', '20000']
            ]
        }

    # ... existing patterns ...
```

### Multiple MCP Endpoints

For load balancing or multiple Excel servers:

```python
class Pipe:
    def __init__(self):
        self.valves = self.Valves()
        self.endpoints = [
            'http://mcp-server-1:8001',
            'http://mcp-server-2:8001',
            'http://mcp-server-3:8001'
        ]
        self.current_endpoint = 0

    def _get_next_endpoint(self):
        endpoint = self.endpoints[self.current_endpoint]
        self.current_endpoint = (self.current_endpoint + 1) % len(self.endpoints)
        return endpoint
```

### Custom System Prompts

Modify the `SYSTEM_PROMPT` valve to customize behavior:

```
You are a specialized Excel assistant for financial reporting.
When users ask for financial spreadsheets, always include:
- Date columns in YYYY-MM-DD format
- Currency values with 2 decimal places
- Summary rows with totals
- Professional formatting
```

## Performance Optimization

### Caching
- Implement session ID caching to avoid re-initialization
- Cache common data patterns
- Use connection pooling for MCP requests

### Error Handling
- Implement retry logic for network failures
- Add circuit breaker pattern for MCP server outages
- Provide fallback responses when MCP is unavailable

### Monitoring
- Add logging for request/response times
- Track success/failure rates
- Monitor Excel file generation metrics

## Security Considerations

- **Input Validation:** All user inputs are validated by the MCP server
- **File Access:** Restricted to designated output directories
- **Network Security:** Use HTTPS in production environments
- **Rate Limiting:** Implement request throttling if needed

## Production Deployment

### Docker Compose Setup

```yaml
version: '3.8'
services:
  open-webui:
    image: ghcr.io/open-webui/open-webui:main
    ports:
      - "8091:8080"
    volumes:
      - openwebui_data:/app/backend/data
    environment:
      - OLLAMA_BASE_URL=http://ollama:11434
    depends_on:
      - ollama
      - exel-mcp

  ollama:
    image: ollama/ollama:latest
    volumes:
      - ollama_data:/root/.ollama
    ports:
      - "11434:11434"

  exel-mcp:
    build: ./exel_mcp
    ports:
      - "8001:8001"
    volumes:
      - ./output:/app/output
    environment:
      - OUTPUT_DIR=/app/output

volumes:
  openwebui_data:
  ollama_data:
```

### Environment Variables

```bash
# Open-WebUI
OPENWEBUI_SECRET_KEY=your-secret-key
WEBUI_AUTH=True

# Exel MCP
MCP_HOST=0.0.0.0
MCP_PORT=8001
OUTPUT_DIR=/app/output
```

## Support and Maintenance

### Updating the Pipe Function

1. Backup current configuration
2. Update the Python code
3. Test in development environment
4. Deploy to production
5. Monitor for issues

### Monitoring

- Set up log aggregation
- Monitor MCP server health
- Track usage metrics
- Set up alerts for failures

### Backup and Recovery

- Backup Open-WebUI data volumes
- Backup generated Excel files
- Document configuration settings
- Test recovery procedures

---

## Quick Start Summary

1. **Install:** Copy `excel_assistant_pipe.py` to Open-WebUI functions
2. **Configure:** Set MCP_BASE_URL to your server endpoint
3. **Enable:** Activate the function in Admin Panel
4. **Test:** Ask to "Create a spreadsheet of employees"
5. **Verify:** Check `./output/` for created Excel files

The Excel Assistant is now ready to help users create spreadsheets through natural language! 🎉