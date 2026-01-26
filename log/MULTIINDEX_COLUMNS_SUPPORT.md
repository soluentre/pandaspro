# MultiIndex Columns 支持文档

## 概述

`pandaspro` 的 `putxl` 功能现已支持导出具有 MultiIndex columns 的 DataFrame（例如 pivot 表）。

## 主要功能

### 1. 基本导出

MultiIndex columns 的 DataFrame 可以直接导出，无需特殊处理：

```python
import pandas as pd
from pandaspro.io.excel.putexcel import PutxlSet

# 创建具有 MultiIndex columns 的 DataFrame
data = {
    ('CMU', 'ACS_Staff'): [10, 20, 30],
    ('CMU', 'Total'): [50, 60, 70],
    ('VP', 'ACS_Staff'): [15, 25, 35],
    ('VP', 'Total'): [55, 65, 75],
}
df = pd.DataFrame(data, index=['Region A', 'Region B', 'Region C'])

# 导出
ps = PutxlSet('output.xlsx')
ps.putxl(df, sheet_name='Data', cell='A1', index=True, header=True)
```

### 2. 使用 __ 分隔符指定 MultiIndex Columns

在 `df_format`、`config` 和其他参数中，可以使用 `__` 分隔符来指定 MultiIndex 列：

#### 示例 1：df_format 参数

```python
ps.putxl(
    df,
    sheet_name='WithFormat',
    cell='A1',
    index=True,
    df_format={
        'fill=yellow; bold=True': 'columns(c="CMU__ACS_Staff")',
        'fill=lightblue': 'columns(c="VP__Total")'
    }
)
```

#### 示例 2：config 参数

```python
ps.putxl(
    df,
    sheet_name='WithConfig',
    cell='A1',
    index=True,
    config={
        'CMU__ACS_Staff': {'width': 15, 'number_format': '#,##0'},
        'VP__Total': {'width': 20, 'number_format': '#,##0.00'}
    }
)
```

### 3. 通配符支持

支持使用通配符（wildcard）匹配 MultiIndex columns：

```python
ps.putxl(
    df,
    sheet_name='WithWildcard',
    cell='A1',
    index=True,
    df_format={
        # 匹配所有以 ACS_Staff 结尾的列
        'fill=lightgreen': 'columns(c="*__ACS_Staff")',
        # 匹配所有 CMU 开头的列
        'bold=True': 'columns(c="CMU__*")'
    }
)
```

## 命名规则

### MultiIndex Columns 的字符串表示

对于 MultiIndex columns，例如 `('CMU', 'ACS_Staff')`：
- 使用 `__`（双下划线）分隔各个层级
- 字符串表示为：`"CMU__ACS_Staff"`

### 示例

| MultiIndex Tuple | 字符串表示 |
|-----------------|-----------|
| `('CMU', 'ACS_Staff')` | `"CMU__ACS_Staff"` |
| `('VP', 'Total')` | `"VP__Total"` |
| `('Region', 'North', 'Sales')` | `"Region__North__Sales"` |

## 支持的参数

以下参数现在都支持使用 `__` 分隔符来指定 MultiIndex columns：

1. `df_format` - 格式化参数
2. `config` - 列配置参数
3. `cd_format` - 条件格式化参数（通过 column 参数）
4. 所有使用 `columns()` 方法的地方

## 技术实现

### 修改的文件

1. **pandaspro/io/excel/writer.py**
   - `get_column_letter_by_name()`: 增加对 `__` 分隔符的支持
   - `range_columns()`: 支持 MultiIndex columns 的字符串表示
   - `range_index_merge_inputs()`: 支持 MultiIndex columns 的通配符匹配

2. **pandaspro/io/excel/putexcel.py**
   - `putxl()` 方法的 config 部分：增加对 MultiIndex columns 的检测和处理

3. **pandaspro/core/tools/utils.py**
   - `df_with_index_for_mask()`: 修复 MultiIndex columns 的处理逻辑

## 测试

完整的测试脚本位于：`pandaspro/test/multiindex_columns_test.py`

运行测试：
```bash
cd pandaspro/test
python multiindex_columns_test.py
```

## 注意事项

1. **分隔符选择**：使用 `__`（双下划线）而不是单下划线，以避免与列名中可能存在的单下划线冲突
2. **精确匹配**：所有层级都必须精确匹配（转换为字符串后）
3. **向后兼容**：所有现有功能保持不变，仅增加新功能

## 示例：完整的 Pivot Table 导出

```python
import pandas as pd
import numpy as np
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

## 更新日期

2024-10-15

