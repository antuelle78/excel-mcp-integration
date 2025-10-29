# Excel MCP Integration - Deployment Guide

## 🚀 Docker Hub Deployment

Your Excel MCP integration is now available on Docker Hub!

### 📦 Pull the Image
```bash
docker pull antuelle78/excel-mcp-integration:latest
```

### 🏃‍♂️ Quick Start
```bash
# Create docker-compose.yml
cat > docker-compose.yml << EOF
version: '3.8'
services:
  excel-mcp:
    image: antuelle78/excel-mcp-integration:latest
    container_name: excel-mcp-server
    ports:
      - "9080:8000"  # MCP server
      - "9081:8001"  # File server
    volumes:
      - ./output:/app/output
    environment:
      - MAX_ROWS=10000
      - MAX_COLS=100
    restart: unless-stopped
EOF

# Start the service
docker compose up -d

# Check status
docker compose ps
```

### 🔗 Access Points
- **MCP Server**: `http://localhost:9080/mcp`
- **File Downloads**: `http://localhost:9081/files/`

## 📋 Open-WebUI Integration

### Step 1: Access Open-WebUI
Navigate to: `http://localhost:8080`

### Step 2: Add Custom Tool
1. Go to **Settings** → **Tools** → **Add Custom Tool**
2. Copy the content of `excel_tools_openwebui.py` from the repository
3. Paste it into the tool configuration
4. Click **Save**

### Step 3: Test Integration
Try this prompt in Open-WebUI:
```
Create an Excel file named "test.xlsx" with headers ["Name", "Age"] and data [["John", 25], ["Jane", 30]]
```

## ✅ Features Included

- ✅ **Excel File Creation** with formatting
- ✅ **Chart Generation** (bar, line, pie, scatter)
- ✅ **CSV Import/Export** functionality
- ✅ **Session Management** for persistent connections
- ✅ **Download Links** with proper URL handling
- ✅ **Headers-Only Support** with auto-placeholder rows
- ✅ **Error Handling** and validation
- ✅ **Comprehensive Test Suite**

## 🧪 Testing

Use the provided test prompts in `test_prompts.md` to validate all functionality:

```bash
# Clone the repository for test files
git clone https://github.com/antuelle78/excel-mcp-integration.git
cd excel-mcp-integration

# View test prompts
cat test_prompts.md
```

## 🔧 Configuration

### Environment Variables
- `MAX_ROWS`: Maximum rows per Excel file (default: 10000)
- `MAX_COLS`: Maximum columns per Excel file (default: 100)
- `OUTPUT_DIR`: Output directory for generated files (default: ./output)

### Port Configuration
- **9080**: MCP Server (internal: 8000)
- **9081**: File Server (internal: 8001)

## 📊 Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   Open-WebUI   │───▶│  Excel MCP Tool │───▶│   Docker Hub    │
│   (localhost)  │    │   Integration   │    │    Image       │
└─────────────────┘    └──────────────────┘    └─────────────────┘
                              │
                              ▼
                       ┌──────────────────┐
                       │  MCP Server     │
                       │  (Port 9080)   │
                       └──────────────────┘
                              │
                              ▼
                       ┌──────────────────┐
                       │  File Server    │
                       │  (Port 9081)   │
                       └──────────────────┘
```

## 🎯 Production Ready

Your Excel MCP integration is now:
- ✅ **Containerized** and deployed on Docker Hub
- ✅ **Tested** with comprehensive test suite
- ✅ **Documented** with installation guides
- ✅ **Production Ready** for enterprise use

## 📞 Support

For issues or questions:
1. Check the troubleshooting guide in the repository
2. Review test prompts for validation
3. Check Docker logs: `docker compose logs excel-mcp`

---

**🎉 Congratulations! Your Excel MCP integration is live on Docker Hub!**