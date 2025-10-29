# Contributing to Excel MCP Integration

Thank you for your interest in contributing to the Excel MCP Integration project! This document provides guidelines and information for contributors.

## 🚀 Ways to Contribute

### 🐛 Reporting Bugs
- Use the [GitHub Issues](https://github.com/antuelle78/excel-mcp-integration/issues) page
- Provide detailed steps to reproduce the issue
- Include your environment details (OS, Docker version, etc.)
- Attach relevant log files if available

### 💡 Suggesting Features
- Open a [GitHub Discussion](https://github.com/antuelle78/excel-mcp-integration/discussions) for feature requests
- Describe the problem you're trying to solve
- Explain why this feature would be valuable
- Consider alternative solutions

### 🛠️ Contributing Code
- Fork the repository
- Create a feature branch from `main`
- Make your changes
- Add tests for new functionality
- Ensure all tests pass
- Update documentation as needed
- Submit a pull request

## 📋 Development Setup

### Prerequisites
- Python 3.10+
- Docker and Docker Compose
- Git

### Local Development
```bash
# Clone the repository
git clone https://github.com/antuelle78/excel-mcp-integration.git
cd excel-mcp-integration

# Create virtual environment
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt

# Run tests
python -m pytest tests/

# Start development server
python src/main.py
```

## 📝 Code Guidelines

### Python Style
- Follow PEP 8 style guidelines
- Use type hints for function parameters and return values
- Write descriptive variable and function names
- Add docstrings to all functions and classes

### Commit Messages
- Use clear, descriptive commit messages
- Start with a verb (Add, Fix, Update, etc.)
- Keep the first line under 50 characters
- Add detailed description if needed

### Testing
- Write tests for all new functionality
- Ensure existing tests still pass
- Test both success and error scenarios
- Include integration tests for complex features

## 📚 Documentation

### Code Documentation
- Add docstrings to all public functions
- Include parameter descriptions and examples
- Document any side effects or exceptions

### User Documentation
- Update README.md for new features
- Add examples and usage instructions
- Update API documentation as needed

## 🤝 Code of Conduct

This project follows a code of conduct to ensure a welcoming environment for all contributors:

- Be respectful and inclusive
- Focus on constructive feedback
- Help newcomers learn and contribute
- Report any unacceptable behavior

## 📞 Getting Help

If you need help or have questions:
- Check the [README](https://github.com/antuelle78/excel-mcp-integration/blob/main/README.md) for documentation
- Search existing [issues](https://github.com/antuelle78/excel-mcp-integration/issues) and [discussions](https://github.com/antuelle78/excel-mcp-integration/discussions)
- Start a new discussion for questions

## 🙏 Recognition

Contributors will be recognized in the project documentation and GitHub's contributor insights. Thank you for helping make Excel MCP Integration better!

---

For more detailed information, see the [README](https://github.com/antuelle78/excel-mcp-integration/blob/main/README.md) and project documentation.