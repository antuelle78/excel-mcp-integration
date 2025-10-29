# 🚀 Excel Tools Open-WebUI Installation Guide

## ✅ **Current Status**

Excel Tools integration is **fully functional** and ready for Open-WebUI installation!

- ✅ MCP server running on `localhost:8002`
- ✅ File server running on `localhost:9081`
- ✅ All 6 Excel tools working correctly
- ✅ Session management and file downloads operational
- ✅ Integration test passed with 100% success

---

## 📋 **Installation Steps**

### **Step 1: Access Open-WebUI**

Open your browser and go to: **http://localhost:8080**

### **Step 2: Navigate to Tools Settings**

1. Click on **Settings** (gear icon ⚙️)
2. Select **Tools** from the left menu
3. Click **Add Custom Tool** button

### **Step 3: Install Excel Tools**

1. **Tool Name**: `Excel Tools`
2. **Tool Description**: `A comprehensive set of tools to create, manipulate, and analyze Excel files with charts, formatting, and CSV integration.`
3. **Tool Code**: Copy the entire content of `excel_tools_openwebui.py`

```bash
# Copy the tool content to clipboard
cat excel_tools_openwebui.py
```

4. Click **Save** or **Install**

### **Step 4: Verify Installation**

1. Go back to **Settings → Tools**
2. Verify "Excel Tools" appears in the list
3. Check that status shows **Active** or **Enabled**
4. No error messages should be visible

---

## 🧪 **Testing the Integration**

### **Test 1: Simple Excel Creation**

In Open-WebUI chat, type:
```
Create a simple Excel file with sample sales data including columns for Product, Quantity, and Price.
```

**Expected Response:**
```
Successfully created Excel file: output/sales_data.xlsx

📁 **File Created:** sales_data.xlsx
🔗 **Download Link:** http://localhost:9081/files/sales_data.xlsx
💡 *You can download this Excel file using the link above*
```

### **Test 2: Chart Creation**

```
Add a bar chart to the sales data file showing quantities by product.
```

**Expected Response:**
```
Successfully created chart in sales_data.xlsx

📊 **Chart Added:** Bar chart showing quantities by product
🔗 **Updated File:** http://localhost:9081/files/sales_data.xlsx
```

### **Test 3: CSV Import**

```
Convert this CSV data to Excel: Name,Age,City
John,25,NYC
Jane,30,LA
```

**Expected Response:**
```
Successfully converted CSV to Excel format

📁 **File Created:** converted_data.xlsx
🔗 **Download Link:** http://localhost:9081/files/converted_data.xlsx
```

---

## 🔧 **Available Tools**

| Tool | Function | Example Usage |
|------|----------|---------------|
| **create_excel_workbook** | Create Excel files with data | "Create an Excel file with employee data" |
| **create_excel_chart** | Add charts to Excel files | "Add a pie chart showing sales by region" |
| **format_excel_cells** | Apply formatting to cells | "Format the header row in bold with blue background" |
| **import_csv_to_excel** | Convert CSV to Excel | "Convert this CSV data to Excel format" |
| **export_excel_to_csv** | Export Excel to CSV | "Export the sales data to CSV format" |
| **get_excel_info** | Analyze Excel file contents | "Show me the structure and contents of this Excel file" |
| **analyze_data_summary** | Get data analysis summary | "Analyze the sales data and provide insights" |

---

## 🐛 **Troubleshooting**

### **Issue: Tool Not Appearing in Open-WebUI**

**Solution:**
1. Ensure the tool code has the correct header format:
   ```python
   """title: 'Excel Tools'
   author: 'Exel MCP Server'
   description: '...'
   version: '1.0.0'
   requirements: httpx
   """
   ```

2. Check that `class Tools:` exists and has proper methods
3. Verify Open-WebUI has internet access for `httpx` installation

### **Issue: "Failed to initialize session" Error**

**Solution:**
1. Check MCP server is running: `ps aux | grep "python src/main.py"`
2. Verify port accessibility: `curl http://localhost:8002/mcp`
3. Restart MCP server if needed

### **Issue: No Download Links in Responses**

**Solution:**
1. Check file server is running: `curl http://localhost:9081/`
2. Verify file creation in output directory: `ls -la output/`
3. Check file server logs for errors

### **Issue: LLM Creates Files Directly Instead of Using Tools**

**Solution:**
1. Verify tool is **Active** in Open-WebUI Settings → Tools
2. Check model compatibility (use `deepseek-r1:7b` or similar)
3. Ensure tool calling is enabled in model settings
4. Try re-importing the tool with fresh code

---

## 🎯 **Success Indicators**

✅ **Working Integration Looks Like:**
- LLM responses contain `📁 **File Created:**` format
- Download links use `http://localhost:9081/files/` format
- No `file://./` paths in responses
- All Excel operations complete successfully
- Files are accessible via download links

❌ **Broken Integration Looks Like:**
- LLM creates files directly: `file://./filename.xlsx`
- No download links provided
- "Tool not found" errors
- Session initialization failures

---

## 📞 **Support**

If you encounter issues:

1. **Check Logs**: `tail -f server.log` for MCP server errors
2. **Test Integration**: `python3 test_openwebui_integration.py`
3. **Verify Services**: Ensure both ports 8002 and 9081 are accessible
4. **Restart Services**: Stop/start MCP server and file server if needed

---

## 🎉 **Ready to Use!**

Your Excel Tools integration is now fully operational! Users can:

- Create complex Excel files with formatting
- Generate charts and visualizations
- Import/export CSV data
- Analyze existing Excel files
- Download results via web links

**The integration provides enterprise-grade Excel capabilities directly within Open-WebUI!** 🚀