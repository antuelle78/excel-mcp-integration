#!/usr/bin/env python3
"""
Simple Web API Wrapper for Exel MCP Server

This provides simple REST endpoints that Open-WebUI can call through its tool system.
Designed for users who only have web interface access to Open-WebUI.
"""

from flask import Flask, request, jsonify
from typing import Dict, Any, Tuple, Optional
import requests
import json
import uuid
import os

app = Flask(__name__)

# Configuration
MCP_BASE_URL = os.getenv('MCP_BASE_URL', 'http://localhost:9080')
API_PORT = int(os.getenv('API_PORT', '8080'))
API_HOST = os.getenv('API_HOST', '0.0.0.0')

# Global session ID
mcp_session_id = None

def get_mcp_session() -> Optional[str]:
    """Get or create MCP session"""
    global mcp_session_id
    if mcp_session_id:
        return mcp_session_id

    try:
        response = requests.post(
            f"{MCP_BASE_URL}/mcp",
            json={
                "jsonrpc": "2.0",
                "id": str(uuid.uuid4()),
                "method": "initialize",
                "params": {
                    "protocolVersion": "2024-11-05",
                    "capabilities": {},
                    "clientInfo": {
                        "name": "simple-web-wrapper",
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
            mcp_session_id = response.headers.get('mcp-session-id')
            return mcp_session_id

    except Exception as e:
        print(f"Session initialization error: {e}")

    return None

def call_mcp_tool(tool_name: str, arguments: Dict[str, Any]) -> Dict[str, Any]:
    """Call an MCP tool and return the result"""
    session_id = get_mcp_session()
    if not session_id:
        return {"error": "Failed to initialize MCP session"}

    try:
        response = requests.post(
            f"{MCP_BASE_URL}/mcp",
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
                "mcp-session-id": session_id
            },
            timeout=30
        )

        if response.status_code == 200:
            # Parse SSE response
            response_text = response.text.strip()
            lines = response_text.split('\n')
            for line in lines:
                if line.startswith('data: '):
                    json_data = line[6:].strip()
                    result = json.loads(json_data)
                    return result
        else:
            return {"error": f"MCP server error: {response.status_code}"}

    except Exception as e:
        return {"error": f"Request failed: {str(e)}"}

@app.route('/', methods=['GET'])
def index() -> Dict[str, Any]:
    """API information"""
    return jsonify({
        "service": "Excel MCP Web Wrapper",
        "version": "1.0",
        "endpoints": {
            "/create_excel": "POST - Create Excel file",
            "/templates": "GET - Get available templates",
            "/create_from_template/<template>": "POST - Create from template",
            "/health": "GET - Health check"
        },
        "mcp_server": MCP_BASE_URL
    })

@app.route('/health', methods=['GET'])
def health() -> Tuple[Dict[str, Any], int]:
    """Health check"""
    session_id = get_mcp_session()
    return jsonify({
        "status": "healthy" if session_id else "unhealthy",
        "mcp_session": session_id is not None,
        "mcp_url": MCP_BASE_URL
    })

@app.route('/create_excel', methods=['POST'])
def create_excel() -> Tuple[Dict[str, Any], int]:
    """Create Excel file from JSON data"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No JSON data provided"}), 400

        # Extract parameters
        filename = data.get('filename')
        headers = data.get('headers')
        sheet_data = data.get('sheet_data')
        sheet_name = data.get('sheet_name', 'Sheet1')

        if not filename or not headers or not sheet_data:
            return jsonify({"error": "Missing required parameters: filename, headers, sheet_data"}), 400

        # Call MCP
        result = call_mcp_tool("create_excel_file", {
            "filename": filename,
            "headers": headers,
            "sheet_data": sheet_data,
            "sheet_name": sheet_name
        })

        if 'result' in result:
            return jsonify({
                "success": True,
                "message": f"Excel file '{filename}' created successfully!",
                "file": filename
            })
        else:
            error_msg = result.get('error', {}).get('message', 'Unknown error') if 'error' in result else 'MCP call failed'
            return jsonify({"error": error_msg}), 500

    except Exception as e:
        return jsonify({"error": f"Server error: {str(e)}"}), 500

@app.route('/templates', methods=['GET'])
def get_templates():
    """Get available Excel templates"""
    templates = {
        "employees": {
            "name": "Employee List",
            "description": "Basic employee information spreadsheet",
            "filename": "employees.xlsx",
            "headers": ["Name", "Department", "Position", "Salary"]
        },
        "products": {
            "name": "Product Inventory",
            "description": "Product catalog with pricing and stock",
            "filename": "products.xlsx",
            "headers": ["Product Name", "Category", "Price", "Stock", "Supplier"]
        },
        "sales": {
            "name": "Sales Report",
            "description": "Sales transactions and revenue tracking",
            "filename": "sales.xlsx",
            "headers": ["Date", "Customer", "Product", "Quantity", "Unit Price", "Total"]
        },
        "budget": {
            "name": "Budget Tracker",
            "description": "Department budget planning and tracking",
            "filename": "budget.xlsx",
            "headers": ["Category", "Budgeted", "Actual", "Difference"]
        },
        "tasks": {
            "name": "Task List",
            "description": "Project tasks and progress tracking",
            "filename": "tasks.xlsx",
            "headers": ["Task", "Assignee", "Status", "Priority", "Due Date"]
        }
    }

    return jsonify({"templates": templates})

@app.route('/create_from_template/<template_name>', methods=['POST'])
def create_from_template(template_name):
    """Create Excel file from a predefined template"""
    try:
        # Template data
        templates = {
            "employees": {
                "filename": "employees.xlsx",
                "headers": ["Name", "Department", "Position", "Salary"],
                "sheet_data": [
                    ["Alice Johnson", "Engineering", "Senior Developer", "95000"],
                    ["Bob Wilson", "Sales", "Account Manager", "75000"],
                    ["Carol Brown", "Marketing", "Content Specialist", "65000"],
                    ["David Lee", "HR", "Recruiter", "70000"],
                    ["Eva Garcia", "Finance", "Analyst", "80000"]
                ]
            },
            "products": {
                "filename": "products.xlsx",
                "headers": ["Product Name", "Category", "Price", "Stock", "Supplier"],
                "sheet_data": [
                    ["Laptop Pro", "Electronics", "1299.99", "25", "TechCorp"],
                    ["Wireless Mouse", "Accessories", "29.99", "100", "GadgetPlus"],
                    ["Monitor 27\"", "Electronics", "349.99", "15", "DisplayMasters"],
                    ["Keyboard RGB", "Accessories", "89.99", "50", "KeyTech"],
                    ["USB Drive 128GB", "Storage", "24.99", "200", "DataStore"]
                ]
            },
            "sales": {
                "filename": "sales.xlsx",
                "headers": ["Date", "Customer", "Product", "Quantity", "Unit Price", "Total"],
                "sheet_data": [
                    ["2024-01-15", "ABC Corp", "Laptop Pro", "5", "1299.99", "6499.95"],
                    ["2024-01-16", "XYZ Ltd", "Wireless Mouse", "20", "29.99", "599.80"],
                    ["2024-01-17", "TechStart Inc", "Monitor 27\"", "3", "349.99", "1049.97"],
                    ["2024-01-18", "GlobalTech", "Keyboard RGB", "10", "89.99", "899.90"],
                    ["2024-01-19", "InnovateNow", "USB Drive 128GB", "50", "24.99", "1249.50"]
                ]
            },
            "budget": {
                "filename": "budget.xlsx",
                "headers": ["Category", "Budgeted", "Actual", "Difference"],
                "sheet_data": [
                    ["Salaries", "500000", "485000", "15000"],
                    ["Marketing", "100000", "95000", "5000"],
                    ["Equipment", "80000", "75000", "5000"],
                    ["Travel", "30000", "25000", "5000"],
                    ["Training", "20000", "18000", "2000"]
                ]
            },
            "tasks": {
                "filename": "tasks.xlsx",
                "headers": ["Task", "Assignee", "Status", "Priority", "Due Date"],
                "sheet_data": [
                    ["Design new logo", "Alice", "In Progress", "High", "2024-02-01"],
                    ["Update website", "Bob", "Pending", "Medium", "2024-02-15"],
                    ["Write documentation", "Carol", "Completed", "Low", "2024-01-30"],
                    ["Test new features", "David", "In Progress", "High", "2024-02-05"],
                    ["Deploy to production", "Eva", "Pending", "High", "2024-02-10"]
                ]
            }
        }

        if template_name not in templates:
            return jsonify({"error": f"Template '{template_name}' not found"}), 404

        template = templates[template_name]

        # Allow custom filename override
        data = request.get_json() or {}
        filename = data.get('filename', template['filename'])

        # Call MCP
        result = call_mcp_tool("create_excel_file", {
            "filename": filename,
            "headers": template["headers"],
            "sheet_data": template["sheet_data"]
        })

        if 'result' in result:
            return jsonify({
                "success": True,
                "message": f"Excel file '{filename}' created from {template_name} template!",
                "template": template_name,
                "file": filename
            })
        else:
            error_msg = result.get('error', {}).get('message', 'Unknown error') if 'error' in result else 'MCP call failed'
            return jsonify({"error": error_msg}), 500

    except Exception as e:
        return jsonify({"error": f"Server error: {str(e)}"}), 500

if __name__ == '__main__':
    print(f"Starting Simple Excel Web Wrapper on {API_HOST}:{API_PORT}")
    print(f"MCP Server: {MCP_BASE_URL}")
    print("Available endpoints:")
    print("  GET  /              - API information")
    print("  GET  /health        - Health check")
    print("  POST /create_excel  - Create custom Excel file")
    print("  GET  /templates     - List available templates")
    print("  POST /create_from_template/<name> - Create from template")
    app.run(host=API_HOST, port=API_PORT, debug=False)