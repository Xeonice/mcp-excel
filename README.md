# mcp-excel

A powerful MCP server that provides Excel file reading capabilities with advanced features like data validation, dropdown lists, and cell properties extraction.

## Features

- Read Excel files and get their content as pandas DataFrames
- Extract Excel properties including:
  - Data validation rules
  - Dropdown lists
  - Merged cells
  - Hidden rows and columns
- Comprehensive error handling
- Full test coverage

## Installation

```bash
pip install mcp-excel
```

## Configuration

Add the following configuration to your MCP config file:

```json
{
  "mcpServers": {
    "excel_access": {
      "command": "uv",
      "args": [
        "run",
        "python",
        "-m",
        "mcp_excel"
      ]
    }
  }
}
```

## Usage

### Reading Excel Files

```python
from mcp import MCPClient

client = MCPClient()
df, properties = client.excel_access.read_excel("path/to/file.xlsx")

# Access the DataFrame
print(df)

# Access Excel properties
print(properties)
```

### Getting Excel Properties Only

```python
properties = client.excel_access.get_excel_properties("path/to/file.xlsx")
print(properties)
```

### Reading Specific Sheets

```python
# Read by sheet name
df, properties = client.excel_access.read_excel("path/to/file.xlsx", sheet_name="Sheet2")

# Read by sheet index (0-based)
df, properties = client.excel_access.read_excel("path/to/file.xlsx", sheet_name=1)
```

## Requirements

- Python >= 3.12
- pandas >= 2.2.3
- openpyxl >= 3.1.5
- mcp[cli] >= 1.3.0

## Development

1. Clone the repository:
   ```bash
   git clone https://github.com/Xeonice/mcp-excel.git
   cd mcp-excel
   ```

2. Create and activate a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

3. Install development dependencies:
   ```bash
   pip install -e ".[dev]"
   ```

4. Run tests:
   ```bash
   pytest
   ```

5. Build the package:
   ```bash
   python -m build
   ```

## Project Structure

```
mcp-excel/
├── mcp_excel/           # Main package directory
│   ├── __init__.py     # Package initialization
│   └── main.py         # MCP server implementation
├── tests/              # Test directory
│   ├── __init__.py
│   ├── test_data/      # Test data directory
│   └── test_excel.py   # Test cases
├── pyproject.toml      # Project configuration
└── README.md          # This file
```

## License

MIT License
