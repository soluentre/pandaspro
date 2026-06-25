from __future__ import annotations

import json
import re
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import Any

import pandas as pd

from pandaspro.core.tools.ai_config import config_help_message, get_ai_settings
from pandaspro.core.tools.cpdhelp import build_retrieval_context

try:
    import pandaspro
    _PKG_VERSION = getattr(pandaspro, '__version__', 'unknown')
except Exception:  # pragma: no cover
    _PKG_VERSION = 'unknown'

# 文档中允许出现的 API / 前缀（用于校验 AI 回答）
_DOC_API_TOKENS = {
    'tab', 'cpdhelp', 'askai', 'tab_singleton_scan', 'inlist', 'inrange', 'dfilter',
    'indate', 'csort', 'corder', 'lowervarlist', 'varnames', 'merge', 'add_total',
    'export_build', 'er', 'dvl', 'dvlmore', 'dvlless',
}
_DOC_PREFIXES = (
    'cpdtab_', 'cpdtabd_', 'cpdtabt_', 'cpdtab2_', 'cpdtab2s_',
    'cpdtab2pct_', 'cpdtab2pctrow_', 'cpdtab2pctcol_',
    'cpdtab2spct_', 'cpdtab2spctrow_', 'cpdtab2spctcol_',
    'cpdtab2sum_', 'cpdtab2mean_', 'cpdtab2median_', 'cpdtab2min_', 'cpdtab2max_',
    'cpdtab2std_', 'cpdtab2var_', 'cpdtab2first_', 'cpdtab2last_',
    'cpdtab2s', 'cpdtab2',
    'cpdlist_', 'cpddict_', 'cpdf_', 'cpdfnot_', 'cpdisna_', 'cpdnotna_',
)

_MENTION_PATTERN = re.compile(
    r'(?:df\.)?('
    r'cpd[a-z0-9_]*|tab_singleton_scan|tab|cpdhelp|askai|'
    r'inlist|inrange|dfilter|indate|csort|corder|lowervarlist|varnames|add_total'
    r')',
    re.IGNORECASE,
)


@dataclass
class AskAiResult:
    question: str
    answer: str
    topics: list[str]
    sources: list[str]
    used_ai: bool
    validated: bool
    fallback_reason: str | None = None


def _schema_summary(data: pd.DataFrame) -> str:
    lines = [f'shape: {data.shape}']
    for col in data.columns:
        dtype = data[col].dtype
        lines.append(f'  - {col!r}: {dtype}')
    return '\n'.join(lines)


def extract_api_mentions(text: str) -> set[str]:
    mentions = set()
    for match in _MENTION_PATTERN.finditer(text or ''):
        token = match.group(1)
        if token.lower() in {'df', 'self'}:
            continue
        mentions.add(token)
    return mentions


def _token_allowed(token: str, context: str) -> bool:
    lower_ctx = context.lower()
    t = token.lower()
    if t in _DOC_API_TOKENS and t in lower_ctx:
        return True
    if t in lower_ctx:
        return True
    for prefix in _DOC_PREFIXES:
        if t.startswith(prefix.lower()) and prefix.lower() in lower_ctx:
            return True
    # cpdtab2sum 等聚合变体
    if t.startswith('cpdtab2') and 'cpdtab2' in lower_ctx:
        agg_names = ('sum', 'mean', 'median', 'min', 'max', 'std', 'var', 'first', 'last')
        pct_names = ('pct', 'pctrow', 'pctcol', 'spct', 'spctrow', 'spctcol')
        for agg in agg_names:
            if t.startswith(f'cpdtab2{agg}') or t.startswith(f'cpdtab2s{agg}'):
                return True
        for pct in pct_names:
            if t.startswith(f'cpdtab2{pct}') or t.startswith(f'cpdtab2s{pct}'):
                return True
        if t.startswith('cpdtab2_') or t.startswith('cpdtab2s_'):
            return True
    return False


def validate_answer_against_context(answer: str, context: str) -> tuple[bool, list[str]]:
    """校验回答中的 API 名是否能在检索到的文档中找到依据。"""
    invalid = []
    for token in extract_api_mentions(answer):
        if not _token_allowed(token, context):
            invalid.append(token)
    return len(invalid) == 0, invalid


def _format_doc_fallback(topics: list[str], context: str) -> str:
    topic_str = ', '.join(f"cpdhelp('{t}')" for t in topics)
    return (
        '【包内文档检索结果 — 未调用 AI 或 AI 回答未通过校验】\n'
        f'依据: {topic_str}\n'
        f'{"-" * 40}\n'
        f'{context}\n'
        f'{"-" * 40}\n'
        '提示: 以上内容为 pandaspro 内置文档原文，请以 df.cpdhelp(...) 为准。'
    )


def _build_messages(
    question: str,
    context: str,
    topics: list[str],
    schema: str | None,
) -> list[dict[str, str]]:
    topic_str = ', '.join(topics)
    schema_block = f'\n\n【当前 DataFrame 结构（不含行数据）】\n{schema}' if schema else ''
    system = (
        f'你是 pandaspro {_PKG_VERSION} 的用法助手。'
        '你只能根据用户提供的【包内文档】回答，禁止编造文档中未出现的 API 名称。'
        '若文档不足以回答，请明确说「文档中未找到相关内容」。'
        '回答末尾必须用一行列出依据，格式：依据: cpdhelp(\'topic\')'
        f'当前检索主题: {topic_str}。'
    )
    user = (
        f'【包内文档】\n{context}{schema_block}\n\n'
        f'【用户问题】\n{question}\n\n'
        '请用简洁中文回答，并给出可运行的示例（仅使用文档中出现的 API）。'
    )
    return [
        {'role': 'system', 'content': system},
        {'role': 'user', 'content': user},
    ]


def call_deepseek_chat(
    api_key: str,
    messages: list[dict[str, str]],
    *,
    base_url: str,
    model: str,
    timeout: float = 60.0,
) -> str:
    url = base_url.rstrip('/') + '/chat/completions'
    payload = json.dumps({
        'model': model,
        'messages': messages,
        'temperature': 0.2,
    }).encode('utf-8')
    req = urllib.request.Request(
        url,
        data=payload,
        headers={
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {api_key}',
        },
        method='POST',
    )
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            body = json.loads(resp.read().decode('utf-8'))
    except urllib.error.HTTPError as e:
        detail = e.read().decode('utf-8', errors='replace')
        raise RuntimeError(f'DeepSeek API 错误 ({e.code}): {detail}') from e
    except urllib.error.URLError as e:
        raise RuntimeError(f'无法连接 DeepSeek API: {e}') from e

    try:
        return body['choices'][0]['message']['content'].strip()
    except (KeyError, IndexError, TypeError) as e:
        raise RuntimeError(f'DeepSeek API 返回格式异常: {body}') from e


def askai(
    data: pd.DataFrame,
    question: str,
    *,
    topic: str | None = None,
    include_schema: bool = False,
    use_ai: bool = True,
) -> AskAiResult:
    """
    检索 cpdhelp 文档后，可选调用 DeepSeek（BYOK）做受控解释。

    无 api_key 时返回文档检索结果；AI 回答未通过 API 名校验时回退到文档原文。
    """
    question = (question or '').strip()
    if not question:
        raise ValueError('askai: question 不能为空')

    context, topics = build_retrieval_context(question, topic)
    sources = [f"cpdhelp('{t}')" for t in topics]
    schema = _schema_summary(data) if include_schema else None

    settings: dict[str, Any] = {}
    api_key = None
    if use_ai:
        try:
            settings = get_ai_settings()
            api_key = settings.get('api_key')
        except ImportError as e:
            return AskAiResult(
                question=question,
                answer=_format_doc_fallback(topics, context) + f'\n\n({e})',
                topics=topics,
                sources=sources,
                used_ai=False,
                validated=True,
                fallback_reason='pyyaml_missing',
            )

    if not use_ai or not api_key:
        msg = config_help_message() if use_ai else ''
        answer = _format_doc_fallback(topics, context)
        if msg and use_ai:
            answer += f'\n\n{msg}'
        return AskAiResult(
            question=question,
            answer=answer,
            topics=topics,
            sources=sources,
            used_ai=False,
            validated=True,
            fallback_reason='no_api_key' if use_ai else 'use_ai_false',
        )

    messages = _build_messages(question, context, topics, schema)
    try:
        ai_answer = call_deepseek_chat(
            api_key,
            messages,
            base_url=settings.get('base_url', 'https://api.deepseek.com'),
            model=settings.get('model', 'deepseek-chat'),
        )
    except RuntimeError as e:
        fallback = _format_doc_fallback(topics, context)
        return AskAiResult(
            question=question,
            answer=f'{fallback}\n\n(AI 调用失败，已回退文档)\n{e}',
            topics=topics,
            sources=sources,
            used_ai=False,
            validated=True,
            fallback_reason='api_error',
        )

    ok, invalid = validate_answer_against_context(ai_answer, context)
    if ok:
        return AskAiResult(
            question=question,
            answer=ai_answer,
            topics=topics,
            sources=sources,
            used_ai=True,
            validated=True,
        )

    fallback = _format_doc_fallback(topics, context)
    reason = ', '.join(invalid)
    return AskAiResult(
        question=question,
        answer=(
            f'{fallback}\n\n'
            f'(AI 回答含文档未出现的 API: {reason}，已回退为文档原文)\n\n'
            f'--- AI 原始回答（仅供参考，未采纳）---\n{ai_answer}'
        ),
        topics=topics,
        sources=sources,
        used_ai=True,
        validated=False,
        fallback_reason='validation_failed',
    )


def print_askai_result(result: AskAiResult) -> None:
    mode = 'AI+文档' if result.used_ai and result.validated else '仅文档'
    print(f'askai [{mode}] 依据: {", ".join(result.sources)}')
    print(result.answer)
