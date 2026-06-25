import pandas as pd

from pandaspro.core.tools.tab import tab


def tab_singleton_scan(data, n: int = 1) -> pd.DataFrame:
    """
    对 DataFrame 全部字段批量执行 tab 统计，筛出「仅有一个取值计数恰好等于 n」的字段。

    判断规则
    --------
    1. 字段唯一取值类别数 > 30：跳过，不参与检测。
    2. 字段唯一取值类别数 ≤ 30：复用 tab(d='brief') 统计各取值计数。
    3. 仅当该字段恰好存在 1 个 count == n 的取值类别时，纳入结果
       （即「明显的 1 个 count=n 取值 + 其他若干不同频次取值」模式）。

    Parameters
    ----------
    data : DataFrame
        待检测的数据表。
    n : int, optional
        目标计数，默认 1。例如 n=10 时捕捉「恰好一类出现 10 次、其余类别频次均不为 10」的字段。

    Returns
    -------
    DataFrame
        列 field / value / count；无命中时返回空表。
    """
    records = []

    for col in data.columns:
        # 与 tab 默认行为一致（m=False）：不计缺失值
        n_categories = data[col].value_counts().shape[0]
        if n_categories > 30:
            continue

        tab_result = tab(data, col, d='brief')
        counts = tab_result.loc[tab_result.index != 'Total', 'count']

        # 仅保留「恰好 1 个」计数为 n 的取值类别
        matched = counts[counts == n]
        if len(matched) != 1:
            continue

        matched_value = matched.index[0]
        records.append({
            'field': col,
            'value': matched_value,
            'count': int(matched.iloc[0]),
        })

    result = pd.DataFrame(records, columns=['field', 'value', 'count'])

    if result.empty:
        print(
            f'tab_singleton_scan: 未发现符合条件的字段'
            f'（唯一取值类别≤30 且仅有一个计数为 {n} 的类别）。'
        )
    else:
        print(f'tab_singleton_scan(n={n}): 发现 {len(result)} 个符合条件的字段：')
        for _, row in result.iterrows():
            print(f"  字段: {row['field']}, 取值: {row['value']}, 计数: {row['count']}")

    return result
