#!/usr/bin/env python3
"""
Complete integration test for the Exel MCP server.
Tests the full workflow from LLM request to Excel file creation.
"""

import subprocess
import time
import requests
import json
import signal
import os
from pathlib import Path

def start_server():
    """Start the MCP server in the background."""
    print("🚀 Starting MCP server...")

    env = os.environ.copy()
    env.update({
        'HOST': '127.0.0.1',
        'PORT': '8002',  # Use a different port to avoid conflicts
        'OUTPUT_DIR': './output'
    })

    # Start server
    process = subprocess.Popen(
        ['python', 'main.py'],
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        preexec_fn=os.setsid
    )

    # Wait for server to start
    time.sleep(3)

    # Check if server is running
    try:
        response = requests.get("http://127.0.0.1:8002/mcp/tools", timeout=5)
        if response.status_code == 200:
            print("✅ Server started successfully")
            return process
        else:
            print(f"❌ Server returned status {response.status_code}")
            stop_server(process)
            return None
    except requests.exceptions.RequestException as e:
        print(f"❌ Could not connect to server: {e}")
        stop_server(process)
        return None

def stop_server(process):
    """Stop the MCP server."""
    if process:
        try:
            os.killpg(os.getpgid(process.pid), signal.SIGTERM)
            process.wait(timeout=5)
            print("✅ Server stopped")
        except subprocess.TimeoutExpired:
            process.kill()
            print("⚠️ Server force killed")

def test_mcp_workflow():
    """Test the complete MCP workflow."""
    print("\n🔄 Testing MCP workflow...")

    # Test 1: Get available tools
    try:
        response = requests.get("http://127.0.0.1:8002/mcp/tools")
        if response.status_code == 200:
            tools = response.json().get('tools', [])
            print(f"✅ Found {len(tools)} tools")
            for tool in tools:
                print(f"   - {tool['name']}")
        else:
            print(f"❌ Tools endpoint failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Tools test failed: {e}")
        return False

    # Test 2: Get system prompt
    try:
        response = requests.get("http://127.0.0.1:8002/mcp/resources/system_prompt")
        if response.status_code == 200:
            print("✅ System prompt accessible")
        else:
            print(f"❌ System prompt failed: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ System prompt test failed: {e}")
        return False

    # Test 3: Create Excel file via tool call
    tool_call = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "create_excel_file",
            "arguments": {
                "filename": "integration_test.xlsx",
                "headers": ["Product", "Price", "Stock"],
                "sheet_data": [
                    ["Widget A", "19.99", "100"],
                    ["Widget B", "29.99", "50"],
                    ["Widget C", "9.99", "200"]
                ],
                "sheet_name": "Inventory"
            }
        }
    }

    try:
        response = requests.post(
            "http://127.0.0.1:8002/mcp",
            json=tool_call,
            headers={"Content-Type": "application/json"}
        )

        if response.status_code == 200:
            result = response.json()
            print("✅ Excel creation tool call successful")
            print(f"   Result: {result.get('result', 'No result')}")
        else:
            print(f"❌ Tool call failed: {response.status_code}")
            print(f"   Response: {response.text}")
            return False
    except Exception as e:
        print(f"❌ Tool call test failed: {e}")
        return False

    # Test 4: Verify file was created
    output_file = Path("./output/integration_test.xlsx")
    if output_file.exists():
        print(f"✅ Excel file created: {output_file}")
        print(f"   Size: {output_file.stat().st_size} bytes")

        # Clean up
        output_file.unlink()
        print("🧹 Cleaned up test file")
        return True
    else:
        print("❌ Excel file was not created")
        return False

def test_error_handling():
    """Test error handling scenarios."""
    print("\n🚨 Testing error handling...")

    # Test invalid tool call
    invalid_call = {
        "jsonrpc": "2.0",
        "id": 1,
        "method": "tools/call",
        "params": {
            "name": "create_excel_file",
            "arguments": {
                "filename": "",  # Empty filename should fail
                "headers": ["Test"],
                "sheet_data": [["Data"]]
            }
        }
    }

    try:
        response = requests.post(
            "http://127.0.0.1:8002/mcp",
            json=invalid_call,
            headers={"Content-Type": "application/json"}
        )

        if response.status_code == 200:
            result = response.json()
            if 'error' in result:
                print("✅ Error handling works correctly")
                return True
            else:
                print("❌ Expected error but got success")
                return False
        else:
            print(f"❌ Unexpected status code: {response.status_code}")
            return False
    except Exception as e:
        print(f"❌ Error handling test failed: {e}")
        return False

def main():
    """Run complete integration tests."""
    print("🧪 Starting Complete Integration Tests")
    print("=" * 60)

    # Ensure output directory exists
    Path("./output").mkdir(exist_ok=True)

    # Start server
    server_process = start_server()
    if not server_process:
        print("❌ Cannot proceed without server")
        return

    try:
        # Run tests
        workflow_success = test_mcp_workflow()
        error_success = test_error_handling()

        # Summary
        print("\n" + "=" * 60)
        if workflow_success and error_success:
            print("🎉 All integration tests PASSED!")
            print("✅ Exel MCP server is ready for production use")
        else:
            print("❌ Some tests FAILED")
            print("   Check server logs and configuration")

    finally:
        # Always stop server
        stop_server(server_process)

if __name__ == "__main__":
    main()