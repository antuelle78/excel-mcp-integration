"""title: 'Excel Tools'
author: 'Exel MCP Server'
description: 'A comprehensive set of tools to create, manipulate, and analyze Excel files with charts, formatting, and CSV integration.'
version: '1.0.0'
requirements: httpx
"""

import httpx
import json
import re
import os
from typing import Optional, List, Dict, Any


class Tools:
    # Class-level session persistence
    _shared_session_id = None
    _shared_initialized = False
    
    def __init__(self):
        # IMPORTANT: Replace with the actual IP of the machine running the MCP server
        # MCP server runs on port 9080, file server on 9081 (Docker mapped ports)
        # Use host.docker.internal when Open-WebUI runs in Docker, otherwise localhost
        self.mcp_server_url = os.getenv('MCP_SERVER_URL', "http://localhost:9080/mcp")
        self.file_server_url = os.getenv('FILE_SERVER_URL', "http://host.docker.internal:9081")  # For file downloads (port 9081)
        self.headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream"
        }
        # Use shared session state
        self.session_id = Tools._shared_session_id
        self._initialized = Tools._shared_initialized

    def _initialize_session(self) -> bool:
        """Initialize MCP session if not already done."""
        if self._initialized and self.session_id:
            return True
            
        try:
            init_payload = {
                "jsonrpc": "2.0",
                "id": "init",
                "method": "initialize",
                "params": {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {
                        "tools": {}
                    },
                    "clientInfo": {
                        "name": "openwebui-excel-tools",
                        "version": "1.0.0"
                    }
                }
            }
            
            with httpx.Client(timeout=10.0) as client:
                response = client.post(
                    self.mcp_server_url,
                    json=init_payload,
                    headers=self.headers
                )
                response.raise_for_status()
                
                # Extract session ID from response headers FIRST
                session_id = response.headers.get("mcp-session-id")
                if session_id:
                    # Store in class-level shared variables
                    Tools._shared_session_id = session_id
                    Tools._shared_initialized = True
                    self.session_id = session_id
                    self._initialized = True
                    print(f"Session initialized with ID: {session_id}")
                    return True
                
                # Parse SSE response to get session ID (fallback)
                if response.headers.get("content-type") == "text/event-stream":
                    for line in response.iter_lines():
                        if line:
                            line_str = line.decode('utf-8') if isinstance(line, bytes) else str(line)
                            if line_str.startswith('data: '):
                                data = line_str[6:]
                                if data:
                                    result_data = json.loads(data)
                                    # Session ID should be in headers, not in data
                                    break
                    
        except Exception as e:
            print(f"Session initialization failed: {e}")
            
        return False

    def _call_mcp_tool(self, tool_name: str, **kwargs) -> str:
        """Helper function to call a tool on the MCP server."""
        # Initialize session if needed
        if not self._initialize_session():
            return "Error: Failed to initialize MCP session"
        
        payload = {
            "jsonrpc": "2.0",
            "method": "tools/call",
            "params": {
                "name": tool_name,
                "arguments": kwargs,
            },
            "id": "1",
        }
        
        try:
            headers = self.headers.copy()
            if self.session_id:
                headers["mcp-session-id"] = self.session_id
                print(f"Calling {tool_name} with session ID: {self.session_id}")
            else:
                print(f"Calling {tool_name} without session ID!")
                
            with httpx.Client(timeout=30.0) as client:
                response = client.post(
                    self.mcp_server_url, 
                    json=payload, 
                    headers=headers
                )
                response.raise_for_status()
                
                # Handle SSE response format
                if response.headers.get("content-type") == "text/event-stream":
                    # Parse Server-Sent Events response
                    result_data = None
                    for line in response.iter_lines():
                        if line:
                            line_str = line.decode('utf-8') if isinstance(line, bytes) else str(line)
                            if line_str.startswith('data: '):
                                data = line_str[6:]  # Remove 'data: ' prefix
                                if data:
                                    result_data = json.loads(data)
                                    break
                    
                    if result_data and 'result' in result_data:
                        return self._process_result_with_file_access(result_data['result'])
                    elif result_data and isinstance(result_data.get('error'), dict):
                        return f"Error: {result_data['error']}"
                    else:
                        return f"Unexpected response format: {result_data}"
                else:
                    # Handle regular JSON response
                    result = response.json().get("result", {})
                    return self._process_result_with_file_access(result)
                    
        except httpx.HTTPStatusError as e:
            error_msg = f"Error: The tool server returned a status of {e.response.status_code}. Response: {e.response.text}"
            print(f"HTTP Error: {error_msg}")
            # If session error, reset and retry once
            if ("session" in str(e.response.text).lower() or 
                "missing" in str(e.response.text).lower() or 
                e.response.status_code == 400):
                print("Session error detected, resetting session...")
                self._initialized = False
                self.session_id = None
                # Retry once with new session
                if not hasattr(self, '_retry_count'):
                    self._retry_count = 0
                if self._retry_count < 1:
                    self._retry_count += 1
                    return self._call_mcp_tool(tool_name, **kwargs)
            return error_msg
        except Exception as e:
            error_msg = f"An unexpected error occurred: {str(e)}"
            print(f"Unexpected Error: {error_msg}")
            return error_msg

    def _process_result_with_file_access(self, result: dict) -> str:
        """Process result and add file download information."""
        if not isinstance(result, dict):
            return json.dumps(result, indent=2)
        
        # Extract content from result
        content = result.get("content", [])
        if not content:
            return json.dumps(result, indent=2)
        
        # Process text content to find file references
        processed_content = []
        file_info = {}
        
        for item in content:
            if isinstance(item, dict) and item.get("type") == "text":
                text = item.get("text", "")
                
                # Look for file creation messages - expanded patterns
                if ("Successfully created Excel file" in text or "has been created" in text or "created" in text) and ".xlsx" in text:
                    # Extract filename - handle multiple patterns
                    file_match = (re.search(r'output/([^\.]+\.xlsx)', text) or 
                                 re.search(r'\*\*([^\.]+\.xlsx)\*\*', text) or
                                 re.search(r'([a-zA-Z0-9_\-]+\.xlsx)', text))
                    if file_match:
                        filename = file_match.group(1)
                        file_path = file_match.group(0) if 'output/' in file_match.group(0) else f"output/{filename}"
                        
                        # Create download URL
                        download_url = f"{self.file_server_url}/files/{filename}"
                        
                        # Add file info for Open-WebUI
                        file_info[filename] = {
                            "name": filename,
                            "path": file_path,
                            "download_url": download_url,
                            "type": "excel"
                        }
                        
                        # Enhance the message with download info
                        enhanced_text = f"{text}\n\n📁 **File Created:** {filename}\n🔗 **Download Link:** [{download_url}]({download_url})\n💡 *You can download this Excel file using the link above*"
                        processed_content.append({
                            "type": "text",
                            "text": enhanced_text
                        })
                    else:
                        processed_content.append(item)
                else:
                    processed_content.append(item)
            else:
                processed_content.append(item)
        
        # Create enhanced result
        enhanced_result = result.copy()
        enhanced_result["content"] = processed_content
        
        if file_info:
            enhanced_result["files"] = file_info
            enhanced_result["download_instructions"] = "Click on the download links above to access your Excel files"
        
        return json.dumps(enhanced_result, indent=2)

    async def create_excel_file(
        self, 
        filename: str, 
        headers: List[str], 
        sheet_data: List[List[Any]], 
        sheet_name: str = "Sheet1",
        formatting: Optional[Dict[str, Any]] = None
    ) -> str:
        """Creates an Excel file with given data.
        
        Args:
            filename: Name of Excel file to create
            headers: List of column headers
            sheet_data: 2D list of data rows
            sheet_name: Name of worksheet (default: "Sheet1")
            formatting: Optional formatting options
        """
        # Handle empty sheet_data by adding placeholder row
        if not sheet_data:
            sheet_data = [["Sample"] * len(headers)]
        
        result = self._call_mcp_tool(
            "create_excel_file",
            filename=filename,
            headers=headers,
            sheet_data=sheet_data,
            sheet_name=sheet_name,
            formatting=formatting
        )
        
        # Process result for file access and download links
        processed_result = self._process_result_with_file_access({"result": result})
        
        # Add download link if file was created successfully
        if "Successfully created" in result:
            return f"{processed_result}\n\n🔗 **Download Link:** {self.file_server_url}/files/{filename}"
        else:
            return processed_result

    async def get_excel_info(self, filename: str) -> str:
        """Get information about an existing Excel file.
        
        Args:
            filename: Name of the Excel file to analyze
        """
        return self._call_mcp_tool("get_excel_info", filename=filename)

    async def create_excel_chart(
        self,
        filename: str,
        chart_type: str,
        data_range: str,
        title: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> str:
        """Add charts and graphs to existing Excel files.
        
        Args:
            filename: Target Excel file to add chart to
            chart_type: Type of chart to create (bar, line, pie, scatter, area)
            data_range: Cell range for chart data (e.g., 'A1:C10')
            title: Chart title (optional)
            sheet_name: Worksheet name (optional, defaults to first sheet)
        """
        return self._call_mcp_tool(
            "create_excel_chart",
            filename=filename,
            chart_type=chart_type,
            data_range=data_range,
            title=title,
            sheet_name=sheet_name
        )

    async def format_excel_cells(
        self,
        filename: str,
        cell_range: str,
        formatting: Dict[str, Any],
        sheet_name: Optional[str] = None
    ) -> str:
        """Apply formatting to Excel cells including colors, borders, fonts, and styles.
        
        Args:
            filename: Target Excel file
            cell_range: Cell range in A1:B5 format (e.g., 'A1:C10')
            formatting: Formatting options to apply
            sheet_name: Worksheet name (optional, defaults to first sheet)
        """
        return self._call_mcp_tool(
            "format_excel_cells",
            filename=filename,
            cell_range=cell_range,
            formatting=formatting,
            sheet_name=sheet_name
        )

    async def import_csv_to_excel(
        self,
        csv_file: str,
        excel_file: str,
        delimiter: str = ",",
        has_headers: bool = True,
        sheet_name: str = "Sheet1"
    ) -> str:
        """Convert CSV files to Excel format with proper formatting and structure.
        
        Args:
            csv_file: Source CSV file path or content
            excel_file: Target Excel filename
            delimiter: CSV delimiter character (default: ',')
            has_headers: Whether CSV has header row (default: true)
            sheet_name: Worksheet name (optional, defaults to 'Sheet1')
        """
        return self._call_mcp_tool(
            "import_csv_to_excel",
            csv_file=csv_file,
            excel_file=excel_file,
            delimiter=delimiter,
            has_headers=has_headers,
            sheet_name=sheet_name
        )

    async def export_excel_to_csv(
        self,
        excel_file: str,
        csv_file: str,
        sheet_name: Optional[str] = None,
        delimiter: str = ",",
        include_headers: bool = True
    ) -> str:
        """Export Excel worksheets to CSV format.
        
        Args:
            excel_file: Source Excel file
            csv_file: Target CSV filename
            sheet_name: Worksheet name to export (optional, defaults to first sheet)
            delimiter: CSV delimiter character (default: ',')
            include_headers: Whether to include headers in CSV (default: true)
        """
        return self._call_mcp_tool(
            "export_excel_to_csv",
            excel_file=excel_file,
            csv_file=csv_file,
            sheet_name=sheet_name,
            delimiter=delimiter,
            include_headers=include_headers
        )

    # Convenience methods for common operations

    async def create_sales_report(
        self,
        filename: str,
        sales_data: List[List[Any]],
        include_chart: bool = True
    ) -> str:
        """Create a sales report with optional chart.
        
        Args:
            filename: Output Excel filename
            sales_data: Sales data with columns [Month, Sales, Expenses, Profit]
            include_chart: Whether to include a bar chart
        """
        headers = ["Month", "Sales", "Expenses", "Profit"]
        
        # Create the Excel file
        result = await self.create_excel_file(
            filename=filename,
            headers=headers,
            sheet_data=sales_data,
            sheet_name="Sales Report",
            formatting={"header_bold": True, "header_background": "4472C4"}
        )
        
        # Add chart if requested
        if include_chart and len(sales_data) > 0:
            chart_result = await self.create_excel_chart(
                filename=filename,
                chart_type="bar",
                data_range=f"A1:D{len(sales_data) + 1}",
                title="Sales Report Chart",
                sheet_name="Sales Report"
            )
            return f"{result}\n\nChart Result: {chart_result}"
        
        return result

    async def create_employee_directory(
        self,
        filename: str,
        employee_data: List[List[Any]]
    ) -> str:
        """Create an employee directory with formatting.
        
        Args:
            filename: Output Excel filename
            employee_data: Employee data with columns [Name, Department, Email, Phone]
        """
        headers = ["Name", "Department", "Email", "Phone"]
        
        return await self.create_excel_file(
            filename=filename,
            headers=headers,
            sheet_data=employee_data,
            sheet_name="Employees",
            formatting={
                "header_bold": True,
                "header_background": "2E75B6",
                "header_font_color": "FFFFFF",
                "alternate_row_colors": True
            }
        )

    async def analyze_data_summary(self, filename: str) -> str:
        """Get a summary analysis of Excel file data.
        
        Args:
            filename: Excel file to analyze
        """
        return await self.get_excel_info(filename=filename)
    
    def list_excel_tools(self) -> dict:
        """List all available Excel tools."""
        if not self._initialize_session():
            return {"success": False, "error": "Failed to initialize session"}
            
        try:
            payload = {
                "jsonrpc": "2.0",
                "id": "list_tools",
                "method": "tools/list",
                "params": {}
            }
            
            headers = self.headers.copy()
            if self.session_id:
                headers["mcp-session-id"] = self.session_id
                
            with httpx.Client(timeout=30.0) as client:
                response = client.post(
                    self.mcp_server_url,
                    json=payload,
                    headers=headers
                )
                response.raise_for_status()
                
                # Parse SSE response
                if response.headers.get("content-type") == "text/event-stream":
                    for line in response.iter_lines():
                        if line:
                            line_str = line.decode('utf-8') if isinstance(line, bytes) else str(line)
                            if line_str.startswith('data: '):
                                data = line_str[6:]
                                if data:
                                    result_data = json.loads(data)
                                    if "result" in result_data and "tools" in result_data["result"]:
                                        tools = result_data["result"]["tools"]
                                        return {
                                            "success": True,
                                            "tools": tools,
                                            "count": len(tools)
                                        }
                                    elif "error" in result_data:
                                        return {"success": False, "error": result_data["error"]}
                                    break
                
                # Try regular JSON response
                result = response.json()
                if "result" in result and "tools" in result["result"]:
                    tools = result["result"]["tools"]
                    return {
                        "success": True,
                        "tools": tools,
                        "count": len(tools)
                    }
                elif "error" in result:
                    return {"success": False, "error": result["error"]}
                    
        except Exception as e:
            return {"success": False, "error": f"Failed to list tools: {str(e)}"}
            
        return {"success": False, "error": "Unknown error occurred"}
    
    def create_excel_workbook(self, filename: str, data: dict, sheet_name: str = "Sheet1", formatting: Optional[dict] = None) -> dict:
        """Create an Excel workbook with data.
        
        Args:
            filename: Name of the Excel file to create
            data: Dictionary with sheet names as keys and data as values
            sheet_name: Name of the worksheet (default: "Sheet1")
            formatting: Optional formatting options
        """
        if not self._initialize_session():
            return {"success": False, "error": "Failed to initialize session"}
            
        try:

            
            # Convert data dict to headers and sheet_data format
            if not data:
                return {"success": False, "error": "No data provided - data parameter is empty"}
                
            if not isinstance(data, dict):
                return {"success": False, "error": f"Data must be a dictionary, got {type(data).__name__}"}
                
            if sheet_name not in data:
                available_sheets = list(data.keys())
                return {"success": False, "error": f"Sheet '{sheet_name}' not found. Available sheets: {available_sheets}"}
                
            sheet_data = data[sheet_name]
            
            if not sheet_data:
                return {"success": False, "error": f"Sheet '{sheet_name}' is empty - no data provided"}
                
            # Extract headers from first row keys if it's a dict, or use first row if it's a list
            if isinstance(sheet_data, dict):
                if not sheet_data:
                    return {"success": False, "error": "Dictionary data is empty"}
                headers = list(sheet_data.keys())
                rows = [list(sheet_data.values())]
                # Dict format always has at least one data row (the values)
            elif isinstance(sheet_data, list):
                if not sheet_data:
                    return {"success": False, "error": "List data is empty - please provide at least headers or data"}
                    
                if isinstance(sheet_data[0], dict):
                    if not sheet_data[0]:
                        return {"success": False, "error": "First dictionary in list is empty"}
                    headers = list(sheet_data[0].keys())
                    rows = [list(row.values()) for row in sheet_data]
                    # List of dicts always has at least one data row
                else:
                    # Assume first row is headers
                    if not sheet_data[0]:
                        return {"success": False, "error": "First row (headers) is empty"}
                    headers = sheet_data[0]
                    rows = sheet_data[1:] if len(sheet_data) > 1 else []
                    
                    # If no data rows provided, add placeholder row to satisfy MCP requirements
                    if not rows:
                        placeholder_row = ["Sample"] * len(headers)
                        rows = [placeholder_row]
            else:
                return {"success": False, "error": f"Unsupported data format: {type(sheet_data).__name__}. Expected dict, list of dicts, or list of lists."}
            
            # Validate headers and rows
            if not headers:
                return {"success": False, "error": "No headers could be extracted from the data"}
                
            if not rows:
                return {"success": False, "error": "No data rows could be extracted from the data"}
                
            # Ensure all rows have the same length as headers
            for i, row in enumerate(rows):
                if len(row) != len(headers):
                    return {"success": False, "error": f"Row {i+1} has {len(row)} columns but headers have {len(headers)} columns"}
            

            
            result = self._call_mcp_tool(
                "create_excel_file",
                filename=filename,
                headers=headers,
                sheet_data=rows,
                sheet_name=sheet_name,
                formatting=formatting or {}
            )
            
            # Process result for file access
            processed_result = self._process_result_with_file_access({"result": result})
            
            return {
                "success": True,
                "filename": filename,
                "message": processed_result,
                "download_link": f"{self.file_server_url}/files/{filename}" if "Successfully created" in result else None
            }
            
        except Exception as e:
            return {"success": False, "error": f"Failed to create Excel workbook: {str(e)}"}