"""
Excel Assistant Pipe Function for Open-WebUI

This Pipe Function creates an Excel assistant model that integrates with the Exel MCP server.
It allows users to create Excel spreadsheets through natural language commands.
"""

from pydantic import BaseModel, Field
import requests
import json
import uuid
from typing import Dict, Any, Optional


class Pipe:
    """
    Excel Assistant Pipe Function

    This pipe creates a model that can generate Excel files through natural language commands.
    It communicates with the Exel MCP server to create spreadsheets.
    """

    class Valves(BaseModel):
        """Configuration valves for the Excel Assistant"""

        # MCP Server Configuration
        MCP_BASE_URL: str = Field(
            default="http://host.docker.internal:8001",
            description="Base URL of the Exel MCP server"
        )

        # Model Configuration
        MODEL_NAME: str = Field(
            default="Excel Assistant",
            description="Display name for this model in Open-WebUI"
        )

        MODEL_ID: str = Field(
            default="excel-assistant",
            description="Unique identifier for this model"
        )

        # Excel Generation Settings
        DEFAULT_SHEET_NAME: str = Field(
            default="Sheet1",
            description="Default worksheet name for new spreadsheets"
        )

        MAX_ROWS: int = Field(
            default=1000,
            description="Maximum rows allowed per spreadsheet"
        )

        MAX_COLS: int = Field(
            default=50,
            description="Maximum columns allowed per spreadsheet"
        )

        # System Prompt
        SYSTEM_PROMPT: str = Field(
            default="""You are an expert Excel assistant. When users ask you to create spreadsheets or Excel files, you must:

1. Analyze their request to understand what data they want
2. Determine appropriate column headers
3. Generate sample data that matches their request
4. Call the create_excel_file tool with proper parameters

Always respond with a tool call when asked to create Excel files. Do not provide instructions - just call the tool.

Example: If asked "Create a spreadsheet of employees", call create_excel_file with headers like ["Name", "Department", "Salary"] and appropriate sample data.""",
            description="System prompt for the Excel assistant"
        )

    def __init__(self):
        """Initialize the Excel Assistant pipe"""
        self.valves = self.Valves()
        self.session_id = None

    def pipes(self) -> list:
        """Define the models this pipe provides"""
        return [
            {
                "id": self.valves.MODEL_ID,
                "name": self.valves.MODEL_NAME,
            }
        ]

    def _initialize_mcp_session(self) -> Optional[str]:
        """Initialize a session with the MCP server"""
        try:
            response = requests.post(
                f"{self.valves.MCP_BASE_URL}/mcp",
                json={
                    "jsonrpc": "2.0",
                    "id": str(uuid.uuid4()),
                    "method": "initialize",
                    "params": {
                        "protocolVersion": "2024-11-05",
                        "capabilities": {},
                        "clientInfo": {
                            "name": "openwebui-excel-assistant",
                            "version": "1.0"
                        }
                    }
                },
                headers={
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/event-stream"
                },
                timeout=10
            )

            if response.status_code == 200:
                # Extract session ID from response headers
                session_id = response.headers.get('mcp-session-id')
                if session_id:
                    self.session_id = session_id
                    return session_id

            print(f"Failed to initialize MCP session: {response.status_code}")
            return None

        except Exception as e:
            print(f"Error initializing MCP session: {e}")
            return None

    def _call_mcp_tool(self, tool_name: str, arguments: Dict[str, Any]) -> Optional[Dict[str, Any]]:
        """Call an MCP tool with the given arguments"""
        if not self.session_id:
            self._initialize_mcp_session()

        if not self.session_id:
            return {"error": "Failed to initialize MCP session"}

        try:
            response = requests.post(
                f"{self.valves.MCP_BASE_URL}/mcp",
                json={
                    "jsonrpc": "2.0",
                    "id": str(uuid.uuid4()),
                    "method": "tools/call",
                    "params": {
                        "name": tool_name,
                        "arguments": arguments
                    }
                },
                headers={
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/event-stream",
                    "mcp-session-id": self.session_id
                },
                timeout=30
            )

            if response.status_code == 200:
                # Parse SSE response
                response_text = response.text.strip()
                lines = response_text.split('\n')
                json_data = None
                for line in lines:
                    if line.startswith('data: '):
                        json_data = line[6:].strip()
                        break

                if json_data:
                    return json.loads(json_data)
                else:
                    return {"error": "No data in MCP response"}
            else:
                return {"error": f"MCP server error: {response.status_code}"}

        except Exception as e:
            return {"error": f"Error calling MCP tool: {str(e)}"}

    def _extract_excel_request(self, user_message: str) -> Optional[Dict[str, Any]]:
        """Extract Excel creation parameters from user message"""
        # Simple pattern matching for common Excel requests
        message = user_message.lower()

        # Employee spreadsheet
        if any(word in message for word in ['employee', 'staff', 'worker', 'team member']):
            return {
                "filename": "employees.xlsx",
                "headers": ["Name", "Department", "Position", "Salary"],
                "sheet_data": [
                    ["Alice Johnson", "Engineering", "Senior Developer", "95000"],
                    ["Bob Wilson", "Sales", "Account Manager", "75000"],
                    ["Carol Brown", "Marketing", "Content Specialist", "65000"],
                    ["David Lee", "HR", "Recruiter", "70000"],
                    ["Eva Garcia", "Finance", "Analyst", "80000"]
                ]
            }

        # Product inventory
        elif any(word in message for word in ['product', 'inventory', 'item', 'catalog']):
            return {
                "filename": "products.xlsx",
                "headers": ["Product Name", "Category", "Price", "Stock", "Supplier"],
                "sheet_data": [
                    ["Laptop Pro", "Electronics", "1299.99", "25", "TechCorp"],
                    ["Wireless Mouse", "Accessories", "29.99", "100", "GadgetPlus"],
                    ["Monitor 27\"", "Electronics", "349.99", "15", "DisplayMasters"],
                    ["Keyboard RGB", "Accessories", "89.99", "50", "KeyTech"],
                    ["USB Drive 128GB", "Storage", "24.99", "200", "DataStore"]
                ]
            }

        # Sales data
        elif any(word in message for word in ['sale', 'revenue', 'transaction', 'order']):
            return {
                "filename": "sales.xlsx",
                "headers": ["Date", "Customer", "Product", "Quantity", "Unit Price", "Total"],
                "sheet_data": [
                    ["2024-01-15", "ABC Corp", "Laptop Pro", "5", "1299.99", "6499.95"],
                    ["2024-01-16", "XYZ Ltd", "Wireless Mouse", "20", "29.99", "599.80"],
                    ["2024-01-17", "TechStart Inc", "Monitor 27\"", "3", "349.99", "1049.97"],
                    ["2024-01-18", "GlobalTech", "Keyboard RGB", "10", "89.99", "899.90"],
                    ["2024-01-19", "InnovateNow", "USB Drive 128GB", "50", "24.99", "1249.50"]
                ]
            }

        # Generic data table
        elif any(word in message for word in ['table', 'spreadsheet', 'excel', 'data']):
            return {
                "filename": "data.xlsx",
                "headers": ["Column A", "Column B", "Column C", "Column D"],
                "sheet_data": [
                    ["Row 1 Data A", "Row 1 Data B", "Row 1 Data C", "Row 1 Data D"],
                    ["Row 2 Data A", "Row 2 Data B", "Row 2 Data C", "Row 2 Data D"],
                    ["Row 3 Data A", "Row 3 Data B", "Row 3 Data C", "Row 3 Data D"],
                    ["Row 4 Data A", "Row 4 Data B", "Row 4 Data C", "Row 4 Data D"],
                    ["Row 5 Data A", "Row 5 Data B", "Row 5 Data C", "Row 5 Data D"]
                ]
            }

        return None

    def pipe(self, body: Dict[str, Any]) -> str:
        """
        Main pipe function that processes user requests

        Args:
            body: The request body containing user message and model info

        Returns:
            Response string or tool call result
        """
        try:
            # Extract user message
            user_message = ""
            if "messages" in body:
                # Find the last user message
                for message in reversed(body["messages"]):
                    if message.get("role") == "user":
                        user_message = message.get("content", "")
                        break

            if not user_message:
                return "I need a message to process. Please ask me to create a spreadsheet!"

            # Check if this is an Excel creation request
            excel_params = self._extract_excel_request(user_message)

            if excel_params:
                # Call the MCP server to create Excel file
                result = self._call_mcp_tool("create_excel_file", excel_params)

                if result and "result" in result:
                    mcp_result = result["result"]
                    if "content" in mcp_result and mcp_result["content"]:
                        content = mcp_result["content"][0]
                        if content.get("type") == "text":
                            return f"✅ Excel file created successfully!\n\n{content.get('text', '')}\n\nYou can find the file in the server's output directory."

                # Handle errors
                error_msg = result.get("error", "Unknown error occurred") if result else "Failed to communicate with Excel server"
                return f"❌ Sorry, I couldn't create the Excel file: {error_msg}"

            else:
                # Not an Excel request, provide helpful guidance
                return """I'm an Excel Assistant! I can help you create spreadsheets. Try asking me to create:

• Employee spreadsheets
• Product inventories
• Sales data tables
• Any other data tables

For example:
- "Create a spreadsheet of employees with names and salaries"
- "Make a product catalog with prices and stock levels"
- "Generate a sales report table"

What kind of spreadsheet would you like me to create?"""

        except Exception as e:
            return f"❌ An error occurred: {str(e)}. Please try again or contact support."


# For testing the pipe directly
if __name__ == "__main__":
    # Test the pipe
    pipe = Pipe()

    # Test initialization
    print("Testing MCP session initialization...")
    session_id = pipe._initialize_mcp_session()
    print(f"Session ID: {session_id}")

    # Test Excel creation
    print("\nTesting Excel creation...")
    test_params = {
        "filename": "test.xlsx",
        "headers": ["Name", "Value"],
        "sheet_data": [["Test", "123"]]
    }
    result = pipe._call_mcp_tool("create_excel_file", test_params)
    print(f"Result: {result}")

    # Test pipe function
    print("\nTesting pipe function...")
    test_body = {
        "messages": [{"role": "user", "content": "Create a spreadsheet of employees"}]
    }
    response = pipe.pipe(test_body)
    print(f"Response: {response}")