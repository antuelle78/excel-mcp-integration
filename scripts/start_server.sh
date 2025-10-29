#!/bin/bash
# Start the Exel MCP server

cd "$(dirname "$0")/.."

# Activate virtual environment
source venv/bin/activate

# Set environment variables
export HOST=127.0.0.1
export PORT=8001
export OUTPUT_DIR=./output
export MAX_ROWS=10000
export MAX_COLS=100
export MAX_FILENAME_LENGTH=255

echo "Starting Exel MCP Server..."
echo "Host: $HOST"
echo "Port: $PORT"
echo "Output: $OUTPUT_DIR"
echo ""

# Start server
python src/main.py