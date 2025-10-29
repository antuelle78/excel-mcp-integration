# 🔧 Open-WebUI Integration Troubleshooting

## ❌ **Current Issue**

The LLM is creating files directly on filesystem instead of using the Excel Tools MCP integration:

**What's Happening:**
```
I'm sorry for the oversight. The workbook **certificat_scolarite_summary.xlsx** was created on the system and is available in the current working directory.
You can open it directly or download it with the following link:
```file://./certificat_scolarite_summary.xlsx```
```

**What Should Happen:**
```
Successfully created Excel file: output/certificat_scolarite_summary.xlsx

📁 **File Created:** certificat_scolarite_summary.xlsx
🔗 **Download Link:** [http://localhost:9081/files/certificat_scolarite_summary.xlsx](http://localhost:9081/files/certificat_scolarite_summary.xlsx)
💡 *You can download this Excel file using the link above*
```

---

## 🔍 **Root Cause Analysis**

### **The LLM is NOT using the Excel Tools MCP integration**

**Evidence:**
1. **File Path**: `file://./certificat_scolarite_summary.xlsx` (local filesystem)
2. **No MCP Tool Calls**: No session initialization or tool invocation
3. **No Download Links**: Missing the characteristic `📁 **File Created:**` format
4. **Direct File Creation**: LLM is using filesystem access instead of MCP tools

---

## 🛠️ **Solutions**

### **Solution 1: Verify Open-WebUI Tool Import**

**Step 1: Check if Excel Tools is Loaded**
In Open-WebUI, go to **Settings → Tools** and verify:
- ✅ "Excel Tools" appears in the tools list
- ✅ Tool status shows "Active" or "Enabled"
- ✅ No error messages related to the tool

**Step 2: Re-import if Necessary**
If Excel Tools is not visible:
1. Go to **Settings → Tools → Add Custom Tool**
2. Copy the entire content of `excel_tools_openwebui.py`
3. Paste into the tool editor
4. Save and activate the tool

### **Solution 2: Check Tool Activation**

**In Open-WebUI Chat:**
1. **Start a new chat**
2. **Type**: "Create an Excel file called test.xlsx with columns Name, Age"
3. **Expected Response**: Should show Excel tool invocation with download links
4. **If you see**: Direct file creation like `file://./test.xlsx` → **Tool not activated**

### **Solution 3: Verify MCP Server Connection**

**Check Connection:**
```bash
# Test if Open-WebUI can reach MCP server
curl -X POST http://localhost:9080/mcp \
  -H "Content-Type: application/json" \
  -d '{"jsonrpc":"2.0","id":"test","method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{"tools":{}},"clientInfo":{"name":"test","version":"1.0"}}}'
```

**Expected Response:**
```
HTTP/1.1 200 OK
mcp-session-id: [session-id]
event: message
data: {"jsonrpc":"2.0","id":"test","result":{...}}
```

### **Solution 4: Check Open-WebUI Configuration**

**Verify Settings:**
1. **Model**: Should be `deepseek-r1:7b` or compatible model
2. **System Prompt**: Should include Excel tools instructions
3. **Tool Access**: Ensure custom tools are enabled
4. **Network**: Open-WebUI can reach `localhost:9080`

---

## 🧪 **Diagnostic Tests**

### **Test 1: Direct Tool Verification**
In Open-WebUI chat, type:
```
List all available Excel tools and their functions.
```

**Expected Response**: Should list the 8 Excel tools with descriptions

### **Test 2: Simple Excel Creation**
```
Create a simple Excel file called "openwebui_test.xlsx" with headers: Name, Age and 2 rows of data.
```

**Expected Response**: 
```
Successfully created Excel file: output/openwebui_test.xlsx

📁 **File Created:** openwebui_test.xlsx
🔗 **Download Link:** [http://localhost:9081/files/openwebui_test.xlsx](http://localhost:9081/files/openwebui_test.xlsx)
💡 *You can download this Excel file using the link above*
```

### **Test 3: File Download Test**
After creating a file, click the download link to verify it works.

---

## 🚨 **Troubleshooting Checklist**

### **✅ Verify These Items:**

**[ ] Excel Tools Imported in Open-WebUI**
- Check Settings → Tools
- Look for "Excel Tools" by "Exel MCP Server"
- Verify status is "Active"

**[ ] MCP Server Accessible**
- Docker container running: `docker ps`
- Port 9080 accessible: `curl -I http://localhost:9080/mcp`
- No firewall blocking connection

**[ ] Tool Activation in Chat**
- New chat session started
- Excel tools being invoked (check for tool call messages)
- Download links appearing in responses

**[ ] Model Configuration**
- Using compatible model (deepseek-r1:7b recommended)
- System prompt includes Excel tool instructions
- Tool calling enabled in model settings

---

## 🔧 **Advanced Solutions**

### **If Tool Import Fails:**

**Option 1: Manual Tool Creation**
1. In Open-WebUI, go to **Settings → Tools → Create**
2. Use this exact configuration:
```python
"""title: 'Excel Tools'
author: 'Exel MCP Server'
description: 'A comprehensive set of tools to create, manipulate, and analyze Excel files with charts, formatting, and CSV integration.'
version: '1.0.0'
requirements: httpx
"""

# [Paste the rest of excel_tools_openwebui.py content here]
```

**Option 2: Check Open-WebUI Logs**
```bash
# Check Open-WebUI container logs for tool loading errors
docker logs open-webui
```

**Option 3: Network Troubleshooting**
```bash
# Test from Open-WebUI container perspective
docker exec open-webui curl -I http://host.docker.internal:9080/mcp
```

---

## 🎯 **Expected Behavior Once Fixed**

When properly integrated, Open-WebUI should:

1. **Show Excel Tools** in available tools list
2. **Invoke MCP Tools** when Excel operations requested
3. **Provide Download Links** for all created files
4. **Use Session Management** for tool calls
5. **Display Professional Formatting** with rich responses

---

## 📞 **Support Steps**

If issues persist:

1. **Restart Open-WebUI**: `docker restart open-webui`
2. **Re-import Tool**: Delete and re-add Excel Tools
3. **Check Logs**: Both Open-WebUI and MCP server logs
4. **Verify Network**: Ensure containers can communicate
5. **Test Model**: Try with different LLM model

---

## 🎉 **Success Indicators**

**When Working Correctly:**
- ✅ Excel tools appear in Open-WebUI interface
- ✅ LLM uses MCP tools instead of direct file creation
- ✅ Download links appear for all Excel files
- ✅ Files accessible via `http://localhost:9081/files/`
- ✅ Professional formatting in responses

**The issue is not with the MCP server (which is working perfectly) but with Open-WebUI not loading/using the Excel Tools integration.**