# mcp-excel

一个功能强大的 MCP 服务器，提供 Excel 文件读取功能，包括数据验证、下拉列表和单元格属性提取等高级特性。

[English](README.md) | 中文

## 功能特性

- 读取 Excel 文件并获取其内容为 pandas DataFrame
- 提取 Excel 属性，包括：
  - 数据验证规则
  - 下拉列表
  - 合并单元格
  - 隐藏的行和列
- 全面的错误处理
- 完整的测试覆盖

## 安装

```bash
pip install mcp-excel
```

## 配置

在 MCP 配置文件中添加以下配置：

```json
{
  "mcpServers": {
    "excel_access": {
      "command": "uvx",
      "args": [
        "mcp-excel"
      ]
    }
  }
}
```

### 本地开发配置

对于本地开发和调试，使用以下配置：

```json
{
  "mcpServers": {
    "excel_access": {
      "command": "uv",
      "args": [
        "--directory",
        "/path/to/your/mcp-excel/mcp_excel",
        "run",
        "python",
        "main.py"
      ]
    }
  }
}
```

注意：将 `/path/to/your/mcp-excel` 替换为你的实际项目路径。

## 使用方法

### 读取 Excel 文件

```python
from mcp import MCPClient

client = MCPClient()
df, properties = client.excel_access.read_excel("path/to/file.xlsx")

# 访问 DataFrame
print(df)

# 访问 Excel 属性
print(properties)
```

### 仅获取 Excel 属性

```python
properties = client.excel_access.get_excel_properties("path/to/file.xlsx")
print(properties)
```

### 读取特定工作表

```python
# 按工作表名称读取
df, properties = client.excel_access.read_excel("path/to/file.xlsx", sheet_name="Sheet2")

# 按工作表索引读取（从0开始）
df, properties = client.excel_access.read_excel("path/to/file.xlsx", sheet_name=1)
```

## 系统要求

- Python >= 3.12
- pandas >= 2.2.3
- openpyxl >= 3.1.5
- mcp[cli] >= 1.3.0

## 开发指南

1. 克隆仓库：
   ```bash
   git clone https://github.com/Xeonice/mcp-excel.git
   cd mcp-excel
   ```

2. 创建并激活虚拟环境：
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Windows: .venv\Scripts\activate
   ```

3. 安装开发依赖：
   ```bash
   pip install -e ".[dev]"
   ```

4. 运行测试：
   ```bash
   pytest
   ```

5. 构建包：
   ```bash
   python -m build
   ```

### 本地调试

1. 以开发模式安装包：
   ```bash
   pip install -e .
   ```

2. 直接启动 MCP 服务器：
   ```bash
   python -m mcp_excel
   ```

3. 在另一个终端中，使用 MCP 客户端测试服务器：
   ```python
   from mcp import MCPClient
   
   client = MCPClient()
   df, properties = client.excel_access.read_excel("path/to/your/excel/file.xlsx")
   ```

## 项目结构

```
mcp-excel/
├── mcp_excel/           # 主包目录
│   ├── __init__.py     # 包初始化
│   └── main.py         # MCP 服务器实现
├── tests/              # 测试目录
│   ├── __init__.py
│   ├── test_data/      # 测试数据目录
│   └── test_excel.py   # 测试用例
├── pyproject.toml      # 项目配置
└── README.md          # 说明文件
```

## 许可证

MIT 许可证 