# AGENTS.md - Excel MCP Server

## Commands
- **Install**: `pip install -r requirements.txt` or `poetry install`
- **Run**: `python src/main.py`
- **Test all**: `python -m pytest tests/ -v`
- **Test single**: `python -m pytest tests/integration_test.py::test_mcp_workflow -v`
- **Format**: `black .`
- **Lint**: Follow PEP 8 (no dedicated linter)

## Code Style Guidelines
- **Python**: PEP 8 compliant, 88 char lines, 4-space indentation
- **Naming**: `snake_case` functions/vars, `PascalCase` classes, `UPPER_SNAKE_CASE` constants
- **Types**: Required type hints on all functions, Google-style docstrings
- **Imports**: Standard → Third-party → Local, separated by blank lines
- **Error Handling**: Specific exceptions, structured logging, early validation
- **Security**: `pathlib.Path` for paths, input sanitization, size limits
- **Strings**: F-strings preferred, avoid old-style formatting
- **Logging**: INFO/ERROR levels with context, no DEBUG in production