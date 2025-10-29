# Project Plan: Llama 3.1 8B + FastMCP Excel Server

## Overview
This project implements a lightweight MCP (Model Context Protocol) server for Excel file manipulation using Llama 3.1 8B instruct model via Ollama. The server allows Large Language Models to create Excel spreadsheets through natural language commands while maintaining minimal resource usage.

## Architecture
- **LLM**: Llama 3.1 8B Instruct (Q4_K_M quantization)
- **MCP Framework**: FastMCP (lightweight HTTP-based)
- **Excel Processing**: openpyxl
- **Integration**: Direct Ollama API communication
- **Deployment**: Docker containerized

## Phase 1: Environment Setup & Model Preparation

### 1.1 Ollama Installation & Model Setup
```bash
# Install Ollama (if not already installed)
curl -fsSL https://ollama.com/install.sh | sh

# Pull the recommended Llama 3.1 8B instruct model
ollama pull llama3.1:8b-instruct-q4_K_M

# Verify model installation
ollama list
```

### 1.2 Model Configuration
- **Temperature**: 0.1 (for consistent tool calling)
- **Format**: JSON (for structured tool responses)
- **Context Window**: 128K (default, sufficient for Excel operations)
- **Quantization**: Q4_K_M (optimal balance of size/speed/quality)

### 1.3 System Requirements
- Python 3.10+
- Poetry (for dependency management)
- Ollama running locally
- 8GB+ RAM recommended
- 10GB+ disk space for model

## Phase 2: FastMCP Server Architecture

### 2.1 Current Server Analysis
**Existing Components:**
- `main.py`: FastMCP server with HTTP transport
- `create_excel_file()` tool: Basic Excel generation
- `system_prompt.txt`: LLM guidance for tool usage
- Poetry configuration with FastMCP and openpyxl

**Architecture Decision:**
- ✅ Keep FastMCP (lightweight, direct HTTP integration)
- ✅ Maintain current tool structure
- ✅ Enhance error handling and validation

### 2.2 Server Enhancement Plan
**Required Improvements:**
1. **Input Validation**: Add parameter validation for Excel operations
2. **Error Handling**: Comprehensive exception handling
3. **Security**: Path validation, size limits
4. **Logging**: Request/response logging
5. **Configuration**: Environment-based settings

## Phase 3: LLM Integration Strategy

### 3.1 Tool Calling Format
**Current System Prompt Structure:**
```json
{
  "tool": "create_excel_file",
  "filename": "output.xlsx",
  "headers": ["Column1", "Column2"],
  "sheet_data": [["value1", "value2"]]
}
```

**Llama 3.1 Compatibility:**
- ✅ Native JSON tool calling support
- ✅ Structured output format
- ✅ Function parameter validation

### 3.2 Integration Flow
```
User Request → Ollama API → Llama 3.1 8B → Tool Call JSON → FastMCP Server → Excel Generation → Response
```

### 3.3 API Communication
**Ollama Integration Options:**
1. **Direct HTTP API**: `POST http://localhost:11434/api/chat`
2. **Python Client**: `ollama-python` library
3. **Streaming**: Real-time response handling

## Phase 4: Enhanced Excel Functionality

### 4.1 Core Tool Expansion
**Current: Single tool**
```
create_excel_file(filename, headers, sheet_data)
```

**Enhanced Capabilities:**
- Multiple sheet support
- Data formatting options
- Cell styling
- Formula support
- Data validation

### 4.2 Tool Parameter Enhancement
```python
@app.tool()
def create_excel_file(
    filename: str,
    headers: List[str],
    sheet_data: List[List[str]],
    sheet_name: str = "Sheet1",
    formatting: Optional[Dict] = None
) -> str:
```

## Phase 5: Testing & Validation

### 5.1 Unit Testing
- Tool parameter validation
- Excel file generation
- Error handling scenarios
- File I/O operations

### 5.2 Integration Testing
- End-to-end LLM → MCP → Excel flow
- Various Excel formats and sizes
- Error recovery scenarios
- Performance benchmarking

### 5.3 LLM Testing
- Tool calling accuracy
- Response format validation
- Context handling
- Multi-turn conversations

## Phase 6: Deployment & Production

### 6.1 Containerization
```dockerfile
FROM python:3.10-slim
RUN pip install poetry
COPY pyproject.toml poetry.lock ./
RUN poetry install --no-dev
COPY . .
EXPOSE 8000
CMD ["poetry", "run", "python", "main.py"]
```

### 6.2 Production Configuration
- Environment variables for configuration
- Health checks
- Logging configuration
- Resource limits
- Security hardening

### 6.3 Monitoring
- Request/response logging
- Performance metrics
- Error tracking
- Usage analytics

## Phase 7: Implementation Steps

### Step 1: Model Validation
```bash
# Test model loading and basic functionality
ollama run llama3.1:8b-instruct-q4_K_M "Hello, can you help me create Excel files?"
```

### Step 2: Server Enhancement
1. Add input validation to `create_excel_file` tool
2. Implement comprehensive error handling
3. Add logging and monitoring
4. Enhance Excel generation capabilities

### Step 3: Integration Testing
1. Test LLM → MCP communication
2. Validate tool calling format
3. Test various Excel scenarios
4. Performance optimization

### Step 4: Production Deployment
1. Containerize the application
2. Set up production configuration
3. Implement monitoring and logging
4. Deploy and validate

## Risk Assessment & Mitigation

### Potential Issues:
1. **Model Loading**: Ensure sufficient RAM (8GB+ recommended)
2. **Tool Calling**: Validate JSON format compatibility
3. **Performance**: Monitor inference speed for Excel operations
4. **Security**: Implement path validation and size limits

### Fallback Options:
1. **Alternative Quantization**: Q3_K_M if Q4_K_M too slow
2. **Model Size**: Llama 3.2 3B if 8B too resource-intensive
3. **Integration**: Direct Ollama API if Python client issues

## Success Criteria

### Functional Requirements:
- [ ] LLM can successfully call Excel creation tools
- [ ] Excel files generate correctly with proper formatting
- [ ] Error handling works for invalid inputs
- [ ] Server responds within 5 seconds for typical requests

### Performance Requirements:
- [ ] Model loads within 30 seconds
- [ ] Tool calls complete within 10 seconds
- [ ] Memory usage stays under 12GB
- [ ] Concurrent requests supported

### Quality Requirements:
- [ ] Tool calling accuracy >95%
- [ ] Excel files validate correctly
- [ ] Error messages are helpful
- [ ] Code follows PEP 8 standards

## Timeline & Milestones

### Week 1: Setup & Foundation
- [ ] Environment setup and model installation
- [ ] Basic server functionality verification
- [ ] Initial integration testing

### Week 2: Enhancement & Testing
- [ ] Server enhancements (validation, error handling)
- [ ] Comprehensive testing suite
- [ ] Performance optimization

### Week 3: Production & Deployment
- [ ] Containerization and production config
- [ ] Monitoring and logging implementation
- [ ] Final validation and deployment

## Dependencies & Resources

### Required Packages:
- fastmcp >= 2.13.0.2
- openpyxl >= 3.1.5
- ollama (for Python client, optional)

### External Resources:
- Ollama running locally
- Llama 3.1 8B model (4.9GB)
- Python 3.10+ environment
- Docker for containerization

### Documentation References:
- [FastMCP Documentation](https://fastmcp.com)
- [Ollama API Documentation](https://github.com/ollama/ollama)
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io)

## Notes & Considerations

- Maintain lightweight design principles throughout
- Prioritize tool calling reliability over advanced features
- Ensure backward compatibility with existing integrations
- Document all API changes and new features
- Regular performance monitoring and optimization