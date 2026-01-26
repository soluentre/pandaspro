# MultiIndex Columns 支持更新总结

## 更新日期
2025年10月15日

## 问题描述

用户在使用 `putxl` 功能导出多维 pivot 表时遇到失败问题，特别是当 pivot 的 columns 使用了多个维度（MultiIndex）时无法正常导出。

## 解决方案

实现了对 MultiIndex columns 的完整支持，并允许在 `df_format` 等参数中使用 `__`（双下划线）分隔符来定位 MultiIndex 列。

例如：对于 MultiIndex 列 `('CMU', 'ACS_Staff')`，可以使用字符串 `"CMU__ACS_Staff"` 来引用。

## 修改的文件

### 1. `pandaspro/io/excel/writer.py`

#### 修改 `get_column_letter_by_name` 方法
- 增加对 `__` 分隔符的检测和处理
- 支持将 `"CMU__ACS_Staff"` 格式的字符串转换为 MultiIndex 元组 `('CMU', 'ACS_Staff')`
- 保持对普通列名的向后兼容

```python
def get_column_letter_by_name(self, colname):
    # Support MultiIndex columns with __ separator
    if isinstance(self.columns, pd.MultiIndex):
        if isinstance(colname, str) and '__' in colname:
            # Split by __ and match with MultiIndex
            colname_parts = colname.split('__')
            for i, col in enumerate(self.columns):
                if all(str(col[j]) == colname_parts[j] for j in range(len(colname_parts))):
                    # Found match
                    ...
```

#### 修改 `range_columns` 方法
- 为 MultiIndex columns 创建字符串表示（使用 `__` 连接）
- 支持通配符匹配（如 `"*__ACS_Staff"` 或 `"CMU__*"`）
- 在处理列名时自动检测和转换

#### 修改 `range_index_merge_inputs` 方法
- 支持在 `columns` 参数中使用 MultiIndex 列的字符串表示
- 支持通配符匹配

### 2. `pandaspro/io/excel/putexcel.py`

#### 修改 `putxl` 方法中的 config 处理部分
- 增加对 MultiIndex columns 的检测
- 当列名包含 `__` 时，尝试作为 MultiIndex 列处理
- 保持向后兼容性

```python
if config:
    for name, setting in config.items():
        # Support MultiIndex columns with __ separator
        if isinstance(io.columns, pd.MultiIndex) and '__' in name:
            # Try to match as MultiIndex column
            ...
```

### 3. `pandaspro/core/tools/utils.py`

#### 修改 `df_with_index_for_mask` 函数
- 为 MultiIndex columns 实现专门的处理逻辑
- 避免列名冲突和重复
- 正确处理 index 和 columns 的合并

```python
def df_with_index_for_mask(df, force: bool = False):
    # For MultiIndex columns, use a simpler approach
    if isinstance(df.columns, pd.MultiIndex):
        result = df.reset_index()
        # Set index back
        result = result.set_index(index_cols)
        return result
```

## 新增功能

### 1. MultiIndex Columns 的字符串表示

- **格式**: 使用双下划线 `__` 分隔各个层级
- **示例**: 
  - `('CMU', 'ACS_Staff')` → `"CMU__ACS_Staff"`
  - `('VP', 'Total')` → `"VP__Total"`

### 2. 在 df_format 中使用

```python
ps.putxl(
    df,
    sheet_name='WithFormat',
    df_format={
        'fill=yellow; bold=True': 'columns(c="CMU__ACS_Staff")',
        'fill=lightblue': 'columns(c="VP__Total")'
    }
)
```

### 3. 在 config 中使用

```python
ps.putxl(
    df,
    sheet_name='WithConfig',
    config={
        'CMU__ACS_Staff': {'width': 15, 'number_format': '#,##0'},
        'VP__Total': {'width': 20, 'number_format': '#,##0.00'}
    }
)
```

### 4. 通配符支持

```python
ps.putxl(
    df,
    sheet_name='WithWildcard',
    df_format={
        # 匹配所有以 ACS_Staff 结尾的列
        'fill=lightgreen': 'columns(c="*__ACS_Staff")',
        # 匹配所有 CMU 开头的列
        'bold=True': 'columns(c="CMU__*")'
    }
)
```

## 测试

### 测试文件
- `pandaspro/test/multiindex_columns_test.py`

### 测试覆盖
1. ✅ 基本导出 MultiIndex columns 的 DataFrame
2. ✅ 使用 `df_format` 参数格式化 MultiIndex 列
3. ✅ 使用 `config` 参数配置 MultiIndex 列
4. ✅ 使用通配符匹配 MultiIndex 列

### 运行测试
```bash
cd pandaspro/test
python multiindex_columns_test.py
```

### 测试结果
所有测试均通过 ✅

## 向后兼容性

- ✅ 所有现有功能保持不变
- ✅ 对于非 MultiIndex columns 的 DataFrame，行为完全一致
- ✅ 仅在检测到 MultiIndex columns 时启用新功能
- ✅ 不影响已有代码

## 使用示例

### 完整示例：导出带格式的 Pivot Table

```python
import pandas as pd
from pandaspro.io.excel.putexcel import PutxlSet

# 创建 pivot table
df = pd.DataFrame({
    'Region': ['North', 'North', 'South', 'South'],
    'Category': ['A', 'B', 'A', 'B'],
    'Sales': [100, 200, 150, 250],
    'Profit': [20, 40, 30, 50]
})

pivot = df.pivot_table(
    values=['Sales', 'Profit'],
    index='Region',
    columns='Category',
    aggfunc='sum'
)

# 导出并格式化
ps = PutxlSet('pivot_output.xlsx')
ps.putxl(
    pivot,
    sheet_name='Pivot',
    cell='A1',
    index=True,
    df_format={
        'fill=lightblue; bold=True': 'columns(c="Sales__A")',
        'fill=lightgreen': 'columns(c="Profit__*")',
        'number_format=#,##0': 'columns(c="Sales__*")'
    },
    config={
        'Sales__A': {'width': 15},
        'Sales__B': {'width': 15},
        'Profit__A': {'width': 15},
        'Profit__B': {'width': 15}
    }
)
```

## 文档

- 详细文档：`MULTIINDEX_COLUMNS_SUPPORT.md`
- 测试说明：`pandaspro/test/README_MULTIINDEX.md`

## 注意事项

1. **分隔符**: 必须使用双下划线 `__`，不要使用单下划线
2. **精确匹配**: 所有层级必须精确匹配（转换为字符串后）
3. **通配符**: 支持 `*` 通配符进行模式匹配
4. **大小写**: 区分大小写

## 总结

此次更新完全解决了用户提出的问题，现在可以：
1. ✅ 顺利导出 MultiIndex columns 的 DataFrame
2. ✅ 在 `df_format` 中使用 `__` 分隔符定位 MultiIndex 列
3. ✅ 在 `config` 中配置 MultiIndex 列的格式
4. ✅ 使用通配符批量匹配 MultiIndex 列

所有功能均已测试通过，并保持向后兼容。

