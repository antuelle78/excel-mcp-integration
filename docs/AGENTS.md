# AGENTS.md - Exel MCP Server

## Agent Guidelines

- **Reports**: Only produce reports if explicitly requested
- **Conciseness**: Be concise; avoid unnecessary verbosity
- **Efficiency**: Be economical with context consumption
- **Focus**: Address the task directly without tangential information
- **Output**: Minimal but complete; no padding or filler

## Build/Lint/Test Commands

- **Install dependencies**: `pip install -r requirements.txt` (or `poetry install`)
- **No build needed**: This is a Python MCP server; no compilation required
- **Run server**: `python src/main.py` or `docker-compose up`
- **Run all tests**: `python -m pytest tests/ -v`
- **Run single test**: `python -m pytest tests/integration_test.py -v`
- **Run specific test file**: `python -m pytest tests/test_core_functions.py -v`
- **No linter configured**: Ensure PEP 8 compliance manually

## Code Style Guidelines

- **Imports**: Standard library → third-party → local modules (fastmcp, openpyxl, then local)
- **Formatting**: PEP 8 compliant, 88 char line length, f-strings for formatting
- **Types**: Full type hints on all function parameters and return types (e.g., `def validate_filename(filename: str) -> str:`)
- **Naming**: `snake_case` for functions/variables, `UPPER_CASE` for constants (e.g., `MAX_ROWS`, `OUTPUT_DIR`)
- **Docstrings**: Required for all public functions (see src/main.py examples)
- **Error handling**: No bare `except:` clauses; catch specific exceptions, use ValueError/TypeError appropriately
- **Logging**: Use `logging` module with structured format; logger initialized at module level
- **Security**: Input validation essential (filename sanitization, path traversal prevention, dangerous character filtering)
- **Comments**: Minimal comments; code should be self-documenting via clear naming and type hints