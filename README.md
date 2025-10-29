# Exel MCP Server

A lightweight Model Context Protocol (MCP) server for Excel file manipulation using Llama 3.1 8B.

## Overview

This server allows Large Language Models (LLMs) to create and manipulate Excel spreadsheets through natural language commands. It implements the MCP protocol to provide tool calling capabilities for Excel operations.

## Architecture

- **LLM**: Llama 3.1 8B Instruct (Q4_K_M quantization)
- **MCP Framework**: FastMCP (lightweight HTTP-based)
- **Excel Processing**: openpyxl
- **Integration**: Direct HTTP API communication
- **Deployment**: Docker containerized

## Features

### Core Functionality
- ✅ Excel file creation with custom headers and data
- ✅ Input validation and security measures
- ✅ Configurable size limits and formatting
- ✅ Comprehensive error handling
- ✅ Logging and monitoring

### Security Features
- Path validation to prevent directory traversal
- Filename sanitization
- Size limits for rows and columns
- Safe file operations

### Excel Capabilities
- Multiple data types support
- Auto-adjusting column widths
- Header formatting (bold, centered)
- Custom sheet names
- UTF-8 encoding support

## Quick Start

### Current Deployment Status
✅ **Server is currently running and ready for use!**

- **MCP Server**: Running at `http://your-server-ip:9080/mcp`
- **Ollama Model**: Llama 3.1 8B Instruct (Q4_K_M) loaded and available
- **Container**: Docker container `exel-mcp-server` active
- **Integration**: LLM ↔ MCP communication tested and working

### Prerequisites
- Python 3.10+
- Ollama with Llama 3.1 8B model
- 8GB+ RAM recommended
- Docker (for containerized deployment)

### Installation

1. **Clone and setup:**
```bash
git clone <repository>
cd exel_mcp
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install fastmcp openpyxl
```

2. **Install Ollama and model:**
```bash
# Install Ollama (if not already installed)
curl -fsSL https://ollama.com/install.sh | sh

# Pull the Llama 3.1 8B model
ollama pull llama3.1:8b-instruct-q4_K_M
```

3. **Start the server:**
```bash
# Option 1: Direct Python execution
./scripts/start_server.sh

# Option 2: Docker deployment (recommended)
docker-compose up -d

# Option 3: Docker build and run
docker build -t exel-mcp .
docker run -p 8001:8001 -v $(pwd)/output:/app/output exel-mcp
```

The server will be available at `http://your-server-ip:9080/mcp`

## Configuration

### Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `HOST` | `127.0.0.1` | Server bind address |
| `PORT` | `8001` | Server port |
| `OUTPUT_DIR` | `./output` | Directory for Excel files |
| `MAX_ROWS` | `10000` | Maximum rows per sheet |
| `MAX_COLS` | `100` | Maximum columns per sheet |
| `MAX_FILENAME_LENGTH` | `255` | Maximum filename length |

### Example Configuration
```bash
export HOST=0.0.0.0
export PORT=8000
export OUTPUT_DIR=/app/output
export MAX_ROWS=50000
export MAX_COLS=200
```

## API Usage

### MCP Tools

#### `create_excel_file`
Creates an Excel file with specified data.

**Parameters:**
- `filename` (string): Excel file name (will auto-append .xlsx if missing)
- `headers` (List[str]): Column headers
- `sheet_data` (List[List[str]]): 2D array of data rows
- `sheet_name` (string, optional): Worksheet name (default: "Sheet1")
- `formatting` (dict, optional): Formatting options

**Example:**
```json
{
  "method": "tools/call",
  "params": {
    "name": "create_excel_file",
    "arguments": {
      "filename": "sales_report.xlsx",
      "headers": ["Product", "Sales", "Region"],
      "sheet_data": [
        ["Widget A", "1000", "North"],
        ["Widget B", "1500", "South"]
      ],
      "sheet_name": "Q4 Sales"
    }
  }
}
```

#### `get_excel_info`
Gets information about an existing Excel file.

**Parameters:**
- `filename` (string): Excel file name to analyze

### MCP Resources

#### `system_prompt`
Returns the system prompt for LLM guidance.

**Endpoint:** `GET /mcp/resources/system_prompt`

## LLM Integration

### Open-WebUI Integration

✅ **Open-WebUI integration is fully implemented with Python-based tools!**

The system provides multiple integration options for Open-WebUI:

#### Option 1: Python Pipe Function (Recommended)
- **Custom Model**: "Excel Assistant" appears as a selectable model
- **Natural Language**: Create spreadsheets through conversational commands
- **Smart Recognition**: Automatically detects Excel creation requests
- **Seamless Integration**: Works like any other LLM in Open-WebUI

**Files:**
- `src/excel_assistant_pipe.py` - Complete Pipe Function implementation
- `docs/OPENWEBUI_PIPE_SETUP.md` - Installation and configuration guide
- `config/pipe_requirements.txt` - Python dependencies
- `tests/test_pipe_function.py` - Testing and verification script

**Example Usage:**
```
User: "Create a spreadsheet of employees with names and salaries"
Excel Assistant: ✅ Excel file created successfully!
               Successfully created Excel file: output/employees.xlsx
```

#### Option 2: Manual Tool Configuration
- **Admin Panel Setup**: Configure tools through Open-WebUI interface
- **JSON Definitions**: Pre-built tool configurations
- **Direct API Calls**: Manual setup for custom integrations

**Setup Files:**
- `docs/OPENWEBUI_SETUP.md` - Manual configuration guide
- `config/openwebui_functions.json` - Tool definitions
- `tests/test_openwebui_integration.py` - Integration testing

**Quick Setup:**
1. Access Open-WebUI Admin Panel → Tools
2. Add the Excel tools using the provided configuration
3. Test with natural language requests

### System Prompt
The server provides a system prompt that guides LLMs on how to use the Excel creation tools:

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

### Example LLM Interaction

**User:** Create a spreadsheet of my top 3 favorite movies with their release years.

**LLM Tool Call:**
```json
{
  "filename": "favorite_movies.xlsx",
  "headers": ["Title", "Year"],
  "sheet_data": [
    ["The Shawshank Redemption", "1994"],
    ["The Godfather", "1972"],
    ["The Dark Knight", "2008"]
  ]
}
```

**Result:** Excel file created at `output/favorite_movies.xlsx`

## Testing

### Core Functionality Tests
```bash
# Test core Excel operations
python tests/test_core.py
```

### MCP Integration Tests
```bash
# Test MCP server endpoints (requires running server)
python tests/test_mcp.py
```

## Docker Deployment

### Build Image
```bash
docker build -t exel-mcp .
```

### Run Container
```bash
docker run -p 8001:8001 \
  -e HOST=0.0.0.0 \
  -e OUTPUT_DIR=/app/output \
  -v $(pwd)/output:/app/output \
  exel-mcp
```

### Docker Compose
```yaml
services:
  exel-mcp:
    build: .
    ports:
      - "8001:8001"
    environment:
      - HOST=0.0.0.0
      - OUTPUT_DIR=/app/output
    volumes:
      - ./output:/app/output
```

## Development

### Project Structure
```
exel_mcp/
├── README.md           # Main project documentation
├── pyproject.toml       # Poetry configuration
├── requirements.txt    # Python dependencies
├── Dockerfile          # Container build configuration
├── docker-compose.yml  # Local development orchestration
├── .dockerignore       # Docker ignore rules
├── src/                # Source code
│   ├── main.py         # Main MCP server implementation
│   ├── excel_assistant_pipe.py  # Open-WebUI pipe function
│   ├── web_api_wrapper.py       # Web API wrapper
│   └── simple_web_wrapper.py    # Simple web wrapper
├── tests/              # Test files
│   ├── test_core.py    # Core functionality tests
│   ├── test_mcp.py     # MCP integration tests
│   ├── test_openwebui_integration.py  # Open-WebUI tests
│   ├── test_pipe_function.py  # Pipe function tests
│   ├── test_server.py  # Server tests
│   └── integration_test.py     # Integration tests
├── scripts/            # Utility scripts
│   ├── start_server.sh # Server startup script
│   └── entrypoint.sh   # Docker entrypoint
├── config/             # Configuration files
│   ├── openwebui_functions.json  # Open-WebUI function definitions
│   ├── openwebui_tools.json      # Open-WebUI tool definitions
│   ├── pipe_requirements.txt     # Pipe-specific requirements
│   └── web_wrapper_requirements.txt  # Web wrapper requirements
├── docs/               # Documentation
│   ├── AGENTS.md       # Agent development guidelines
│   ├── GEMINI.md       # Gemini integration docs
│   ├── OPENWEBUI_SETUP.md  # Open-WebUI setup guide
│   ├── OPENWEBUI_PIPE_SETUP.md  # Pipe function setup
│   ├── project_plan.md # Implementation plan
│   └── system_prompt.txt  # LLM system prompt
└── output/             # Generated Excel files
```

### Code Quality
- **Linting**: Follow PEP 8 style guide
- **Type Hints**: All functions use type annotations
- **Error Handling**: Comprehensive exception handling
- **Logging**: Structured logging with appropriate levels
- **Security**: Input validation and sanitization

### Adding New Features
1. Define the tool/resource in `main.py`
2. Add appropriate validation
3. Update tests
4. Update documentation
5. Test integration

## Performance Considerations

### Resource Usage
- **Memory**: ~6-8GB for model + server
- **Storage**: Minimal (Excel files only)
- **CPU**: Low baseline, spikes during file creation

### Optimization Tips
- Use appropriate quantization levels
- Configure size limits based on use case
- Monitor memory usage in production
- Implement caching for frequently accessed data

## Troubleshooting

### Common Issues

**Server won't start:**
- Check if port is already in use
- Verify Python dependencies are installed
- Check file permissions for output directory

**LLM can't connect:**
- Ensure Ollama is running
- Verify model is downloaded
- Check network connectivity

**Excel files not created:**
- Check output directory permissions
- Verify filename validation
- Check server logs for errors

### Logs
Server logs are written to stdout/stderr. Key log levels:
- `INFO`: Normal operations
- `WARNING`: Potential issues
- `ERROR`: Failures requiring attention

## Contributing

1. Follow the coding standards in `docs/AGENTS.md`
2. Add tests for new functionality
3. Update documentation
4. Ensure backward compatibility
5. Test with multiple LLM configurations

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Support

For issues and questions:
- Check the troubleshooting section
- Review server logs
- Test with the provided test scripts
- Ensure all prerequisites are met