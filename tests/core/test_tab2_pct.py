import importlib.util
import sys
from pathlib import Path

import pandas as pd
import pytest

# 直接加载 tab2 模块，避免 pandaspro 包级依赖
_ROOT = Path(__file__).resolve().parents[2]
_spec = importlib.util.spec_from_file_location(
    'tab2', _ROOT / 'pandaspro' / 'core' / 'tools' / 'tab2.py'
)
tab2 = importlib.util.module_from_spec(_spec)
sys.modules['tab2'] = tab2
_spec.loader.exec_module(tab2)

counts_to_pct = tab2.counts_to_pct


@pytest.fixture
def sample_counts():
    """a × b 计数表，grand total = 6。"""
    return pd.DataFrame(
        {
            'P': [2, 1, 1, 4],
            'Q': [1, 1, 0, 2],
            'Total': [3, 2, 1, 6],
        },
        index=['X', 'Y', 'Z', 'Total'],
    )


def test_counts_to_pct_total(sample_counts):
    pct = counts_to_pct(sample_counts, mode='total')
    assert pct.loc['X', 'P'] == pytest.approx(33.33, abs=0.01)
    assert pct.loc['Total', 'Total'] == pytest.approx(100.0)
    assert pct.loc['X', 'P'] + pct.loc['Y', 'P'] + pct.loc['Z', 'P'] == pytest.approx(
        pct.loc['Total', 'P'], abs=0.02
    )


def test_counts_to_pct_row(sample_counts):
    pct = counts_to_pct(sample_counts, mode='row')
    assert pct.loc['X', 'P'] == pytest.approx(66.67, abs=0.01)
    assert pct.loc['X', 'Q'] == pytest.approx(33.33, abs=0.01)
    assert pct.loc['X', 'Total'] == pytest.approx(100.0)
    assert pct.loc['Total', 'P'] == pytest.approx(66.67, abs=0.01)
    assert pct.loc['Total', 'Total'] == pytest.approx(100.0)


def test_counts_to_pct_col(sample_counts):
    pct = counts_to_pct(sample_counts, mode='col')
    assert pct.loc['X', 'P'] == pytest.approx(50.0)
    assert pct.loc['Y', 'P'] == pytest.approx(25.0)
    assert pct.loc['Z', 'P'] == pytest.approx(25.0)
    assert pct.loc['Total', 'P'] == pytest.approx(100.0)
    assert pct.loc['X', 'Total'] == pytest.approx(50.0)
    assert pct.loc['Total', 'Total'] == pytest.approx(100.0)


def test_counts_to_pct_zero_grand_total():
    empty = pd.DataFrame({'P': [0, 0], 'Total': [0, 0]}, index=['A', 'Total'])
    pct = counts_to_pct(empty, mode='total')
    assert (pct == 0).all().all()
