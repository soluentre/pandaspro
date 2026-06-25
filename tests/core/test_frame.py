from pandaspro.core.frame import FramePro


def test_cpdhelp():
    df = FramePro({'A': [1, 2, 3]})
    assert df.cpdhelp() is df
    assert df.cpdhelp('tab') is df


def test_framepro_initialization():
    data = {'A': [1, 2, 3], 'B': [4, 5, 6]}
    df = FramePro(data)
    assert isinstance(df, FramePro)
    assert df.shape == (3, 2)


def test_tab_method():
    data = {'A': [1, 2, 3], 'B': [4, 5, 6]}
    df = FramePro(data)
    result = df.tab('A', 'detail')
    assert isinstance(result, FramePro)
    assert 'A' in result.columns


def test_add_total_method():
    data = {'Category': ['A', 'B', 'C'], 'Value': [100, 200, 300]}
    df = FramePro(data)
    df_total = df.add_total(total_label_column='Category', sum_columns='Value')
    assert df_total.iloc[-1]['Category'] == 'Total'
    assert df_total.iloc[-1]['Value'] == 600


def test_tab_singleton_scan():
    data = {
        'hit': [1, 1, 1, 2, 3, 3],       # 2 出现 1 次 → 命中
        'all_unique': [1, 2, 3, 4, 5, 6],  # 6 个 count=1 → 不命中
        'uniform': [1, 1, 1, 1, 1, 1],     # 无 count=1 → 不命中
    }
    df = FramePro(data)
    result = df.tab_singleton_scan()
    assert isinstance(result, FramePro)
    assert len(result) == 1
    assert result.iloc[0]['field'] == 'hit'
    assert result.iloc[0]['value'] == 2
    assert result.iloc[0]['count'] == 1

    df_many = FramePro({'many': list(range(35))})
    assert len(df_many.tab_singleton_scan()) == 0


def test_tab_singleton_scan_with_n():
    # dept: A×10, B×5, C×3 → 仅 A 计数为 10
    df = FramePro({'dept': ['A'] * 10 + ['B'] * 5 + ['C'] * 3})
    result = df.tab_singleton_scan(n=10)
    assert len(result) == 1
    assert result.iloc[0]['field'] == 'dept'
    assert result.iloc[0]['value'] == 'A'
    assert result.iloc[0]['count'] == 10

    # 两个类别计数均为 10 → 不命中
    df2 = FramePro({'all_ten': ['X'] * 10 + ['Y'] * 10})
    assert len(df2.tab_singleton_scan(n=10)) == 0


def _ab_sample():
    return FramePro({
        'a': ['X', 'X', 'X', 'Y', 'Y', 'Z'],
        'b': ['P', 'P', 'Q', 'P', 'Q', 'P'],
        'id': range(6),
    })


def test_cpdtab2pct_total():
    df = _ab_sample()
    pct = df.cpdtab2pct_a__b
    assert pct.loc['X', 'P'] == 33.33
    assert pct.loc['Total', 'Total'] == 100.0


def test_cpdtab2pctrow():
    df = _ab_sample()
    pct = df.cpdtab2pctrow_a__b
    assert pct.loc['X', 'P'] == 66.67
    assert pct.loc['X', 'Total'] == 100.0
    assert pct.loc['Total', 'P'] == 66.67


def test_cpdtab2pctcol():
    df = _ab_sample()
    pct = df.cpdtab2pctcol_a__b
    assert pct.loc['X', 'P'] == 50.0
    assert pct.loc['Total', 'P'] == 100.0
    assert pct.loc['X', 'Total'] == 50.0


def test_cpdtab2pct_multidim_separator():
    df = FramePro({
        'region': ['东', '东', '西'],
        'grade': ['A', 'B', 'A'],
        'quarter': ['Q1', 'Q1', 'Q2'],
        'id': [1, 2, 3],
    })
    pct = df.cpdtab2pct_region___quarter
    assert pct.loc['东', 'Q1'] == 66.67
    assert pct.loc['Total', 'Total'] == 100.0


def test_cpdtab2spctrow():
    df = FramePro({
        'region': ['东', '东', '东', '西', '西', '西'],
        'dept': ['A', 'A', 'B', 'A', 'B', 'B'],
        'grade': ['G1', 'G2', 'G1', 'G1', 'G1', 'G2'],
        'id': range(6),
    })
    pct = df.cpdtab2spctrow_region__dept___grade
    assert float(pct.iloc[-1, -1]) == 100.0
    assert float(pct.loc[('东', '东 Subtotal'), 'Total']) == 100.0
    assert isinstance(pct, FramePro)
