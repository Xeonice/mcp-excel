import pytest
import pandas as pd
import os
from main import read_excel, get_excel_properties

# 测试数据目录
TEST_DATA_DIR = "test_data"

@pytest.fixture
def sample_excel_file():
    """创建测试用的Excel文件"""
    # 确保测试数据目录存在
    if not os.path.exists(TEST_DATA_DIR):
        os.makedirs(TEST_DATA_DIR)
    
    # 创建测试Excel文件
    file_path = os.path.join(TEST_DATA_DIR, "test.xlsx")
    
    # 创建测试数据
    df = pd.DataFrame({
        'A': [1, 2, 3],
        'B': ['a', 'b', 'c'],
        'C': [True, False, True]
    })
    
    # 保存到Excel文件
    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        # 创建第二个sheet
        df.to_excel(writer, sheet_name='Sheet2', index=False)
    
    yield file_path
    
    # 清理测试文件
    if os.path.exists(file_path):
        os.remove(file_path)
    if os.path.exists(TEST_DATA_DIR):
        os.rmdir(TEST_DATA_DIR)

def test_read_excel_basic(sample_excel_file):
    """测试基本的Excel读取功能"""
    # 测试读取默认sheet
    df = read_excel(sample_excel_file)
    assert isinstance(df, pd.DataFrame)
    assert len(df) == 3
    assert list(df.columns) == ['A', 'B', 'C']
    
    # 测试读取指定sheet
    df = read_excel(sample_excel_file, sheet_name='Sheet2')
    assert isinstance(df, pd.DataFrame)
    assert len(df) == 3
    assert list(df.columns) == ['A', 'B', 'C']

def test_read_excel_errors(sample_excel_file):
    """测试Excel读取的错误处理"""
    # 测试文件不存在
    with pytest.raises(FileNotFoundError):
        read_excel("nonexistent.xlsx")
    
    # 测试sheet不存在
    with pytest.raises(ValueError) as exc_info:
        read_excel(sample_excel_file, sheet_name='NonexistentSheet')
    assert "Sheet 'NonexistentSheet' not found" in str(exc_info.value)

def test_get_excel_properties_basic(sample_excel_file):
    """测试基本的Excel属性获取功能"""
    properties = get_excel_properties(sample_excel_file)
    
    # 验证返回的属性结构
    assert isinstance(properties, dict)
    assert "data_validation" in properties
    assert "dropdown_lists" in properties
    assert "merged_cells" in properties
    assert "hidden_rows" in properties
    assert "hidden_columns" in properties
    
    # 验证所有属性都是列表类型
    assert isinstance(properties["data_validation"], list)
    assert isinstance(properties["dropdown_lists"], list)
    assert isinstance(properties["merged_cells"], list)
    assert isinstance(properties["hidden_rows"], list)
    assert isinstance(properties["hidden_columns"], list)

def test_get_excel_properties_errors(sample_excel_file):
    """测试Excel属性获取的错误处理"""
    # 测试文件不存在
    with pytest.raises(FileNotFoundError):
        get_excel_properties("nonexistent.xlsx")
    
    # 测试sheet不存在
    with pytest.raises(ValueError) as exc_info:
        get_excel_properties(sample_excel_file, sheet_name='NonexistentSheet')
    assert "Sheet 'NonexistentSheet' not found" in str(exc_info.value)

def test_get_excel_properties_with_validation():
    """测试带有数据验证的Excel文件"""
    # 确保测试数据目录存在
    if not os.path.exists(TEST_DATA_DIR):
        os.makedirs(TEST_DATA_DIR)
    
    # 创建一个带有数据验证的测试文件
    test_file = os.path.join(TEST_DATA_DIR, "test_validation.xlsx")
    
    # 创建测试数据
    df = pd.DataFrame({
        'A': [1, 2, 3],
        'B': ['a', 'b', 'c'],
        'C': [True, False, True]
    })
    
    # 保存到Excel文件
    with pd.ExcelWriter(test_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
    
    # 获取属性
    properties = get_excel_properties(test_file)
    
    # 验证属性结构
    assert isinstance(properties, dict)
    assert "data_validation" in properties
    assert "dropdown_lists" in properties
    
    # 清理测试文件
    if os.path.exists(test_file):
        os.remove(test_file)
    if os.path.exists(TEST_DATA_DIR):
        os.rmdir(TEST_DATA_DIR)

if __name__ == "__main__":
    pytest.main([__file__]) 