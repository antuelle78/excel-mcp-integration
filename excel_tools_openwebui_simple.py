"""title: 'Excel Tools'
author: 'Exel MCP Server'
description: 'Create, analyze, and manipulate Excel files with charts and formatting.'
version: '1.0.0'
requirements: httpx
"""

import httpx
import json
import re
from typing import Optional, List, Dict, Any

class Tools:
    def __init__(self):
        self.mcp_server_url = "http://localhost:9080/mcp"
        self.file_server_url = "http://localhost:9081"
        self.headers = {
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream"
        }
        self.session_id = None
        self._initialized = False

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
                    "capabilities": {"tools": {}},
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
                
                # Extract session ID from response headers
                session_id = response.headers.get("mcp-session-id")
                if session_id:
                    self.session_id = session_id
                    self._initialized = True
                    return True
                    
        except Exception as e:
            print(f"Session initialization failed: {e}")
            
        return False

    def _call_mcp_tool(self, tool_name: str, **kwargs) -> str:
        """Helper function to call a tool on the MCP server."""
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
                
            with httpx.Client(timeout=30.0) as client:
                response = client.post(
                    self.mcp_server_url, 
                    json=payload, 
                    headers=headers
                )
                response.raise_for_status()
                
                # Handle SSE response format
                if response.headers.get("content-type") == "text/event-stream":
                    result_data = None
                    for line in response.iter_lines():
                        if line:
                            line_str = line.decode('utf-8') if isinstance(line, bytes) else str(line)
                            if line_str.startswith('data: '):
                                data = line_str[6:]
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
                    result = response.json().get("result", {})
                    return self._process_result_with_file_access(result)
                    
        except Exception as e:
            return f"Error: {str(e)}"

    def _process_result_with_file_access(self, result: dict) -> str:
        """Process result and add file download information."""
        if not isinstance(result, dict):
            return json.dumps(result, indent=2)
        
        content = result.get("content", [])
        if not content:
            return json.dumps(result, indent=2)
        
        processed_content = []
        file_info = {}
        
        for item in content:
            if isinstance(item, dict) and item.get("type") == "text":
                text = item.get("text", "")
                
                # Look for file creation messages
                if "Successfully" in text and ".xlsx" in text:
                    file_match = re.search(r'output/([^\.]+\.xlsx)', text)
                    if file_match:
                        filename = file_match.group(1)
                        file_path = file_match.group(0)
                        
                        # Create download URL
                        download_url = f"{self.file_server_url}/files/{filename}"
                        
                        # Add file info
                        file_info[filename] = {
                            "name": filename,
                            "path": file_path,
                            "download_url": download_url,
                            "type": "excel"
                        }
                        
                        # Enhance message with download info
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
        """Creates an Excel file with the given data."""
        return self._call_mcp_tool(
            "create_excel_file",
            filename=filename,
            headers=headers,
            sheet_data=sheet_data,
            sheet_name=sheet_name,
            formatting=formatting
        )

    async def get_excel_info(self, filename: str) -> str:
        """Get information about an existing Excel file."""
        return self._call_mcp_tool("get_excel_info", filename=filename)

    async def create_excel_chart(
        self,
        filename: str,
        chart_type: str,
        data_range: str,
        title: Optional[str] = None,
        sheet_name: Optional[str] = None
    ) -> str:
        """Add charts and graphs to existing Excel files."""
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
        """Apply formatting to Excel cells."""
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
        """Convert CSV files to Excel format."""
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
        """Export Excel worksheets to CSV format."""
        return self._call_mcp_tool(
            "export_excel_to_csv",
            excel_file=excel_file,
            csv_file=csv_file,
            sheet_name=sheet_name,
            delimiter=delimiter,
            include_headers=include_headers
        )