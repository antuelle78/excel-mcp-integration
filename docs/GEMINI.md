# Exel MCP Server

## Project Overview

This project implements an MCP server for Exel, a hypothetical platform for spreadsheet management and automation. The server allows Large Language Models (LLMs) to interact with Exel spreadsheets, enabling natural language control over data manipulation, analysis, and visualization.

**Core Functionality:**

*   **Exposes a `/mcp` API Endpoint:** The server provides a single API endpoint (`/mcp`) for LLM requests.
*   **Translates LLM Requests to Exel Actions:** It converts MCP-formatted requests into specific Exel API calls, such as creating new sheets, updating cells, and generating charts.
*   **LLM-Powered Payload Generation:** The server can use an LLM to dynamically generate data for Exel API calls, allowing for natural language-driven interactions.
*   **Context Management:** It maintains conversation state with the LLM across multiple requests using a `context_id`.
*   **Authentication and Security:** The API is secured with JWT Bearer token authentication.
*   **Audit Logging:** It logs all requests and responses for auditing and debugging.

**Technology Stack:**

*   **FastAPI:** A high-performance web framework.
*   **Pydantic:** For data validation.
*   **HTTPX:** An asynchronous HTTP client for communicating with the Exel API.
*   **Redis (optional):** For persistent context storage.
*   **Pytest:** For automated testing.

## Building and Running

### Prerequisites

*   Python 3.9+
*   Poetry
*   Docker (for running a local Redis instance)

### Installation

1.  **Clone the repository:**
    ```bash
    git clone https://github.com/example/exel-mcp.git
    cd exel-mcp
    ```

2.  **Install dependencies:**
    ```bash
    poetry install
    ```

### Running the Server

1.  **Start the Redis container:**
    ```bash
    docker run -d -p 6379:6379 redis
    ```

2.  **Run the development server:**
    ```bash
    poetry run uvicorn app.main:app --reload
    ```

The server will be available at `http://localhost:8000`.

### Running Tests

To run the test suite, use the following command:

```bash
poetry run pytest
```

## Development Conventions

*   **Code Style:** We follow the [PEP 8](https://www.python.org/dev/peps/pep-0008/) style guide and use `black` for automated code formatting.
*   **Testing:** All new features and bug fixes must be accompanied by unit tests.
*   **Commits:** We follow the [Conventional Commits](https://www.conventionalcommits.org/en/v1.0.0/) specification.
*   **Branching:** We use the [GitFlow](https://www.atlassian.com/git/tutorials/comparing-workflows/gitflow-workflow) branching model.
