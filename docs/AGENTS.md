# AGENTS.md - Exel MCP Server

## Commands
- **Build**: `poetry build` or `docker build .`
- **Run**: `poetry run python main.py` or `docker-compose up`
- **Test**: `poetry run pytest`
- **Test single file**: `poetry run pytest tests/test_file.py`
- **Format**: `black .`

## Code Style
- Follow PEP 8 style guide
- Use type hints for function parameters and return types
- Write docstrings for all public functions
- Use snake_case for variables and functions
- Handle exceptions appropriately, avoid bare except clauses
- Import order: standard library, third-party, local modules
- Line length: 88 characters (Black default)
- Use f-strings for string formatting