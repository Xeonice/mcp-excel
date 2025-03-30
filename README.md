# mcp-excel

MCP server to give client the ability to read Excel files

# Usage

For this MCP server to work, add the following configuration to your MCP config file:

```json
{
  "mcpServers": {
    "excel": {
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

## Features

- Read Excel files and get their content as pandas DataFrames
- Extract Excel properties including:
  - Data validation rules
  - Dropdown lists
  - Merged cells
  - Hidden rows and columns
- Comprehensive error handling
- Full test coverage

## Requirements

- Python >= 3.12
- pandas >= 2.2.3
- openpyxl >= 3.1.5
- mcp[cli] >= 1.3.0

## Development

1. Clone the repository
2. Install development dependencies:
   ```bash
   pip install -e ".[dev]"
   ```
3. Run tests:
   ```bash
   pytest
   ```

## License

MIT License
