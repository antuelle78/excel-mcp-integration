# AGENTS.md - Excel MCP Server

## Commands

### Build & Development
- **Install dependencies**: `pip install -r requirements.txt` or `poetry install`
- **Build**: `poetry build` or `docker build .`
- **Run locally**: `python src/main.py`
- **Run with Poetry**: `poetry run python src/main.py`
- **Run with Docker**: `docker-compose up`

### Testing
- **Run all tests**: `python -m pytest tests/ -v`
- **Run single test file**: `python -m pytest tests/integration_test.py -v`
- **Run specific test**: `python -m pytest tests/integration_test.py::test_mcp_workflow -v`
- **Run with coverage**: `python -m pytest tests/ --cov=src --cov-report=html`

### Code Quality
- **Format code**: `black .`
- **Check formatting**: `black --check .`
- **Lint**: No dedicated linter configured (follow PEP 8 manually)
- **Type check**: No mypy configured (use type hints for documentation)

## Code Style Guidelines

### Python Standards
- **PEP 8 Compliance**: Follow Python Enhancement Proposal 8 style guide
- **Line Length**: 88 characters (Black default)
- **Indentation**: 4 spaces, no tabs
- **Naming Conventions**:
  - Variables and functions: `snake_case`
  - Classes: `PascalCase`
  - Constants: `UPPER_SNAKE_CASE`
  - Private members: `_leading_underscore`

### Type Hints & Documentation
- **Type Hints**: Required for all function parameters and return types
  ```python
  def validate_filename(filename: str) -> str:
      # implementation
  ```
- **Docstrings**: Required for all public functions using Google/NumPy style
  ```python
  def create_excel_file(
      filename: str,
      headers: List[str],
      sheet_data: List[List[str]],
      sheet_name: str = "Sheet1"
  ) -> str:
      """
      Creates an Excel file with the given data.

      Args:
          filename: Name of the Excel file to create
          headers: List of column headers
          sheet_data: 2D list of data rows
          sheet_name: Name of the worksheet (default: "Sheet1")

      Returns:
          Success message with file path
      """
  ```

### Imports
- **Import Order**:
  1. Standard library imports
  2. Third-party imports
  3. Local imports
- **Grouping**: Separate groups with blank lines
- **Example**:
  ```python
  import os
  import logging
  from pathlib import Path
  from typing import List, Optional, Dict, Any

  import openpyxl
  from fastmcp import FastMCP

  from .utils import validate_input
  ```

### Error Handling
- **Specific Exceptions**: Catch specific exceptions, avoid bare `except:`
- **Logging**: Use structured logging for errors and important events
- **Validation**: Validate inputs early and provide clear error messages
- **Example**:
  ```python
  try:
      # risky operation
      result = process_data(data)
  except ValueError as e:
      logger.error(f"Validation error: {str(e)}")
      raise ValueError(f"Invalid data: {str(e)}")
  except Exception as e:
      logger.error(f"Unexpected error: {str(e)}")
      raise Exception(f"Failed to process data: {str(e)}")
  ```

### String Formatting
- **F-strings**: Preferred for string interpolation
  ```python
  message = f"Successfully created Excel file: {safe_filename}"
  ```
- **Avoid**: Old-style % formatting and .format() unless necessary

### Security & Validation
- **Input Validation**: Always validate and sanitize user inputs
- **Path Security**: Use `pathlib.Path` and resolve paths safely
- **File Extensions**: Restrict to allowed file types
- **Size Limits**: Enforce reasonable limits on data size

### Logging
- **Configuration**: Use structured logging with appropriate levels
- **Levels**: INFO for normal operations, ERROR for failures, DEBUG for development
- **Context**: Include relevant context in log messages
- **Example**:
  ```python
  logger.info(f"Creating Excel file: {filename}")
  logger.error(f"Failed to create Excel file: {str(e)}")
  ```

### Configuration
- **Environment Variables**: Use for runtime configuration
- **Constants**: Define configuration limits at module level
- **Validation**: Validate configuration values on startup

### Best Practices
- **Path Handling**: Use `pathlib.Path` instead of string concatenation
- **Resource Management**: Properly close files and connections
- **Threading**: Be careful with shared state in multi-threaded code
- **Performance**: Consider memory usage for large datasets
- **Testing**: Write comprehensive tests for all public functions