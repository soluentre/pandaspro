from pandaspro.core.frame import FramePro
from pandaspro.core.tools.askai import (
    askai,
    extract_api_mentions,
    validate_answer_against_context,
)
from pandaspro.core.tools.cpdhelp import build_retrieval_context, route_help_topics


def test_route_help_topics_tab():
    topics = route_help_topics('cpdtab2 和 cpdtab2s 有什么区别？')
    assert 'tab' in topics


def test_build_retrieval_context():
    context, topics = build_retrieval_context('多维交叉表', 'tab')
    assert 'cpdtab2' in context
    assert topics == ['tab']


def test_validate_answer_rejects_unknown_api():
    context, _ = build_retrieval_context('tab', 'tab')
    ok, invalid = validate_answer_against_context(
        '请使用 df.cpdtab3_region__grade',
        context,
    )
    assert not ok
    assert 'cpdtab3_region__grade' in invalid or any('cpdtab3' in x for x in invalid)


def test_validate_answer_accepts_doc_api():
    context, _ = build_retrieval_context('tab', 'tab')
    ok, invalid = validate_answer_against_context(
        '用 df.cpdtab2_region__grade 做两维计数',
        context,
    )
    assert ok
    assert invalid == []


def test_askai_without_api_key():
    df = FramePro({'a': [1, 2, 3]})
    result = askai(df, 'cpdtab2 怎么用？', topic='tab', use_ai=True)
    assert not result.used_ai
    assert 'cpdtab2' in result.answer
    assert result.fallback_reason == 'no_api_key'


def test_askai_use_ai_false():
    df = FramePro({'a': [1, 2]})
    result = askai(df, 'tab', use_ai=False)
    assert not result.used_ai
    assert result.validated


def test_framepro_askai_chain():
    df = FramePro({'a': [1]})
    out = df.askai('cpdtab2', topic='tab', use_ai=False)
    assert out is df
    assert hasattr(df, '_last_askai_result')
