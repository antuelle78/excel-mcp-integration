#!/usr/bin/env python3
"""
Web API Wrapper for Exel MCP Server

This creates a simple REST API that Open-WebUI can call through its tool system.
Since you only have web interface access to Open-WebUI, this wrapper provides
simple HTTP endpoints that can be configured as tools in Open-WebUI.
"""

from flask import Flask, request, jsonify
from typing import Dict, Any, Tuple, Optional
import requests
import json
import uuid
import os

app = Flask(__name__)

# Configuration
MCP_BASE_URL = os.getenv("MCP_BASE_URL", "http://localhost:9080")
API_PORT = int(os.getenv("API_PORT", "8080"))
API_HOST = os.getenv("API_HOST", "0.0.0.0")


class MCPClient:
    """Client for communicating with the Exel MCP server"""

    def __init__(self, base_url: str):
        self.base_url = base_url
        self.session_id = None

    def _ensure_session(self) -> Optional[str]:
        """Ensure we have a valid MCP session"""
        if self.session_id:
            return self.session_id

        try:
            response = requests.post(
                f"{self.base_url}/mcp",
                json={
                    "jsonrpc": "2.0",
                    "id": str(uuid.uuid4()),
                    "method": "initialize",
                    "params": {
                        "protocolVersion": "2024-11-05",
                        "capabilities": {},
                        "clientInfo": {"name": "web-api-wrapper", "version": "1.0"},
                    },
                },
                headers={
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/event-stream",
                },
                timeout=10,
            )

            if response.status_code == 200:
                self.session_id = response.headers.get("mcp-session-id")
                return self.session_id

        except Exception as e:
            print(f"Session initialization error: {e}")

        return None

    def create_excel_file(self, filename: str, headers: list, sheet_data: list, sheet_name: str = "Sheet1") -> Dict[str, Any]:
        """Create an Excel file through MCP"""

        if not self._ensure_session():
            return {"error": "Failed to initialize MCP session"}

        # Validate inputs
        if not filename or not headers or not sheet_data:
            return {
                "error": "Missing required parameters: filename, headers, sheet_data"
            }

        if not isinstance(headers, list) or not isinstance(sheet_data, list):
            return {"error": "headers and sheet_data must be arrays"}

        try:
            response = requests.post(
                f"{self.base_url}/mcp",
                json={
                    "jsonrpc": "2.0",
                    "id": str(uuid.uuid4()),
                    "method": "tools/call",
                    "params": {
                        "name": "create_excel_file",
                        "arguments": {
                            "filename": filename,
                            "headers": headers,
                            "sheet_data": sheet_data,
                            "sheet_name": sheet_name,
                        },
                    },
                },
                headers={
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/event-stream",
                    "mcp-session-id": self.session_id,
                },
                timeout=30,
            )

            if response.status_code == 200:
                # Parse SSE response
                response_text = response.text.strip()
                lines = response_text.split("\n")
                json_data = None
                for line in lines:
                    if line.startswith("data: "):
                        json_data = line[6:].strip()
                        break

                if json_data:
                    result = json.loads(json_data)
                    if "result" in result:
                        return {
                            "success": True,
                            "message": "Excel file created successfully",
                            "data": result,
                        }
                    elif "error" in result:
                        return {"error": result["error"].get("message", "MCP error")}
                else:
                    return {"error": "Invalid MCP response format"}
            else:
                return {"error": f"MCP server error: {response.status_code}"}

        except Exception as e:
            return {"error": f"Request failed: {str(e)}"}

    def get_excel_info(self, filename: str) -> Dict[str, Any]:
        """Get information about an Excel file"""

        if not self._ensure_session():
            return {"error": "Failed to initialize MCP session"}

        try:
            response = requests.post(
                f"{self.base_url}/mcp",
                json={
                    "jsonrpc": "2.0",
                    "id": str(uuid.uuid4()),
                    "method": "tools/call",
                    "params": {
                        "name": "get_excel_info",
                        "arguments": {"filename": filename},
                    },
                },
                headers={
                    "Content-Type": "application/json",
                    "Accept": "application/json, text/event-stream",
                    "mcp-session-id": self.session_id,
                },
                timeout=10,
            )

            if response.status_code == 200:
                response_text = response.text.strip()
                lines = response_text.split("\n")
                json_data = None
                for line in lines:
                    if line.startswith("data: "):
                        json_data = line[6:].strip()
                        break

                if json_data:
                    result = json.loads(json_data)
                    if "result" in result:
                        return {"success": True, "data": result}
                    elif "error" in result:
                        return {"error": result["error"].get("message", "MCP error")}
                else:
                    return {"error": "Invalid MCP response format"}
            else:
                return {"error": f"MCP server error: {response.status_code}"}

        except Exception as e:
            return {"error": f"Request failed: {str(e)}"}


# Global MCP client
mcp_client = MCPClient(MCP_BASE_URL)


@app.route("/health", methods=["GET"])
def health_check() -> Tuple[Dict[str, str], int]:
    """Health check endpoint"""
    return jsonify({"status": "healthy", "service": "excel-api-wrapper"})


@app.route("/create_excel", methods=["POST"])
def create_excel() -> Tuple[Dict[str, Any], int]:
    """Create Excel file endpoint for Open-WebUI"""

    try:
        data = request.get_json()

        if not data:
            return jsonify({"error": "No JSON data provided"}), 400

        # Extract parameters
        filename = data.get("filename")
        headers = data.get("headers")
        sheet_data = data.get("sheet_data")
        sheet_name = data.get("sheet_name", "Sheet1")

        # Validate required parameters
        if not filename:
            return jsonify({"error": "filename is required"}), 400
        if not headers:
            return jsonify({"error": "headers is required"}), 400
        if not sheet_data:
            return jsonify({"error": "sheet_data is required"}), 400

        # Call MCP server
        result = mcp_client.create_excel_file(filename, headers, sheet_data, sheet_name)

        if result.get("success"):
            return jsonify(
                {
                    "success": True,
                    "message": f"Excel file '{filename}' created successfully!",
                    "details": result.get("data", {}),
                }
            )
        else:
            return jsonify({"error": result.get("error", "Unknown error")}), 500

    except Exception as e:
        return jsonify({"error": f"Server error: {str(e)}"}), 500


@app.route("/get_excel_info", methods=["POST"])
def get_excel_info() -> Tuple[Dict[str, Any], int]:
    """Get Excel file information endpoint"""

    try:
        data = request.get_json()

        if not data:
            return jsonify({"error": "No JSON data provided"}), 400

        filename = data.get("filename")
        if not filename:
            return jsonify({"error": "filename is required"}), 400

        result = mcp_client.get_excel_info(filename)

        if result.get("success"):
            return jsonify({"success": True, "info": result.get("data", {})})
        else:
            return jsonify({"error": result.get("error", "Unknown error")}), 500

    except Exception as e:
        return jsonify({"error": f"Server error: {str(e)}"}), 500


@app.route("/excel_templates", methods=["GET"])
def get_templates() -> Dict[str, Any]:
    """Get available Excel templates"""

    templates = {
        "employees": {
            "filename": "employees.xlsx",
            "headers": ["Name", "Department", "Position", "Salary"],
            "description": "Employee information spreadsheet",
        },
        "products": {
            "filename": "products.xlsx",
            "headers": ["Product Name", "Category", "Price", "Stock", "Supplier"],
            "description": "Product inventory spreadsheet",
        },
        "sales": {
            "filename": "sales.xlsx",
            "headers": [
                "Date",
                "Customer",
                "Product",
                "Quantity",
                "Unit Price",
                "Total",
            ],
            "description": "Sales transactions spreadsheet",
        },
        "budget": {
            "filename": "budget.xlsx",
            "headers": ["Category", "Budgeted", "Actual", "Difference"],
            "description": "Budget tracking spreadsheet",
        },
    }

    return jsonify({"templates": templates})


@app.route("/create_from_template", methods=["POST"])
def create_from_template() -> Tuple[Dict[str, Any], int]:
    """Create Excel file from a template with sample data"""

    try:
        data = request.get_json()

        if not data:
            return jsonify({"error": "No JSON data provided"}), 400

        template_name = data.get("template")
        if not template_name:
            return jsonify({"error": "template name is required"}), 400

        # Template definitions with sample data
        templates = {
            "employees": {
                "filename": "employees.xlsx",
                "headers": ["Name", "Department", "Position", "Salary"],
                "sheet_data": [
                    ["Alice Johnson", "Engineering", "Senior Developer", "95000"],
                    ["Bob Wilson", "Sales", "Account Manager", "75000"],
                    ["Carol Brown", "Marketing", "Content Specialist", "65000"],
                    ["David Lee", "HR", "Recruiter", "70000"],
                    ["Eva Garcia", "Finance", "Analyst", "80000"],
                ],
            },
            "products": {
                "filename": "products.xlsx",
                "headers": ["Product Name", "Category", "Price", "Stock", "Supplier"],
                "sheet_data": [
                    ["Laptop Pro", "Electronics", "1299.99", "25", "TechCorp"],
                    ["Wireless Mouse", "Accessories", "29.99", "100", "GadgetPlus"],
                    ['Monitor 27"', "Electronics", "349.99", "15", "DisplayMasters"],
                    ["Keyboard RGB", "Accessories", "89.99", "50", "KeyTech"],
                    ["USB Drive 128GB", "Storage", "24.99", "200", "DataStore"],
                ],
            },
            "sales": {
                "filename": "sales.xlsx",
                "headers": [
                    "Date",
                    "Customer",
                    "Product",
                    "Quantity",
                    "Unit Price",
                    "Total",
                ],
                "sheet_data": [
                    ["2024-01-15", "ABC Corp", "Laptop Pro", "5", "1299.99", "6499.95"],
                    [
                        "2024-01-16",
                        "XYZ Ltd",
                        "Wireless Mouse",
                        "20",
                        "29.99",
                        "599.80",
                    ],
                    [
                        "2024-01-17",
                        "TechStart Inc",
                        'Monitor 27"',
                        "3",
                        "349.99",
                        "1049.97",
                    ],
                    [
                        "2024-01-18",
                        "GlobalTech",
                        "Keyboard RGB",
                        "10",
                        "89.99",
                        "899.90",
                    ],
                    [
                        "2024-01-19",
                        "InnovateNow",
                        "USB Drive 128GB",
                        "50",
                        "24.99",
                        "1249.50",
                    ],
                ],
            },
            "budget": {
                "filename": "budget.xlsx",
                "headers": ["Category", "Budgeted", "Actual", "Difference"],
                "sheet_data": [
                    ["Salaries", "500000", "485000", "15000"],
                    ["Marketing", "100000", "95000", "5000"],
                    ["Equipment", "80000", "75000", "5000"],
                    ["Travel", "30000", "25000", "5000"],
                    ["Training", "20000", "18000", "2000"],
                ],
            },
        }

        if template_name not in templates:
            return jsonify({"error": f"Template '{template_name}' not found"}), 404

        template = templates[template_name]

        # Create the Excel file
        result = mcp_client.create_excel_file(
            template["filename"], template["headers"], template["sheet_data"]
        )

        if result.get("success"):
            return jsonify(
                {
                    "success": True,
                    "message": f"Excel file '{template['filename']}' created from {template_name} template!",
                    "template": template_name,
                    "details": result.get("data", {}),
                }
            )
        else:
            return jsonify({"error": result.get("error", "Unknown error")}), 500

    except Exception as e:
        return jsonify({"error": f"Server error: {str(e)}"}), 500


if __name__ == "__main__":
    print(f"Starting Excel API Wrapper on {API_HOST}:{API_PORT}")
    print(f"MCP Server: {MCP_BASE_URL}")
    app.run(host=API_HOST, port=API_PORT, debug=False)
