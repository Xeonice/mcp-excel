# mcp-excel

データ検証、ドロップダウンリスト、セルプロパティ抽出などの高度な機能を備えたExcelファイル読み込み機能を提供する強力なMCPサーバー。

[English](README.md) | [中文](README_zh.md) | 日本語

## 機能

- Excelファイルを読み込み、pandas DataFrameとして内容を取得
- Excelプロパティの抽出：
  - データ検証ルール
  - ドロップダウンリスト
  - 結合セル
  - 非表示の行と列
- 包括的なエラー処理
- 完全なテストカバレッジ

## インストール

```bash
pip install mcp-excel
```

## 設定

MCP設定ファイルに以下の設定を追加してください：

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

### ローカル開発設定

ローカル開発とデバッグには、以下の設定を使用してください：

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

注意：`/path/to/your/mcp-excel`を実際のプロジェクトパスに置き換えてください。

## 使用方法

### Excelファイルの読み込み

```python
from mcp import MCPClient

client = MCPClient()
df, properties = client.excel_access.read_excel("path/to/file.xlsx")

# DataFrameへのアクセス
print(df)

# Excelプロパティへのアクセス
print(properties)
```

### Excelプロパティのみの取得

```python
properties = client.excel_access.get_excel_properties("path/to/file.xlsx")
print(properties)
```

### 特定のシートの読み込み

```python
# シート名による読み込み
df, properties = client.excel_access.read_excel("path/to/file.xlsx", sheet_name="Sheet2")

# シートインデックスによる読み込み（0から開始）
df, properties = client.excel_access.read_excel("path/to/file.xlsx", sheet_name=1)
```

## システム要件

- Python >= 3.12
- pandas >= 2.2.3
- openpyxl >= 3.1.5
- mcp[cli] >= 1.3.0

## 開発ガイド

1. リポジトリのクローン：
   ```bash
   git clone https://github.com/Xeonice/mcp-excel.git
   cd mcp-excel
   ```

2. 仮想環境の作成と有効化：
   ```bash
   python -m venv .venv
   source .venv/bin/activate  # Windows: .venv\Scripts\activate
   ```

3. 開発依存関係のインストール：
   ```bash
   pip install -e ".[dev]"
   ```

4. テストの実行：
   ```bash
   pytest
   ```

5. パッケージのビルド：
   ```bash
   python -m build
   ```

### ローカルデバッグ

1. 開発モードでパッケージをインストール：
   ```bash
   pip install -e .
   ```

2. MCPサーバーを直接起動：
   ```bash
   python -m mcp_excel
   ```

3. 別のターミナルで、MCPクライアントを使用してサーバーをテスト：
   ```python
   from mcp import MCPClient
   
   client = MCPClient()
   df, properties = client.excel_access.read_excel("path/to/your/excel/file.xlsx")
   ```

## プロジェクト構造

```
mcp-excel/
├── mcp_excel/           # メインパッケージディレクトリ
│   ├── __init__.py     # パッケージ初期化
│   └── main.py         # MCPサーバー実装
├── tests/              # テストディレクトリ
│   ├── __init__.py
│   ├── test_data/      # テストデータディレクトリ
│   └── test_excel.py   # テストケース
├── pyproject.toml      # プロジェクト設定
└── README.md          # このファイル
```

## ライセンス

MITライセンス 