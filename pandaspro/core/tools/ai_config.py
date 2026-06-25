from __future__ import annotations

import os
from pathlib import Path
from typing import Any

try:
    import yaml
except ImportError:  # pragma: no cover - optional until installed
    yaml = None


DEFAULT_DEEPSEEK_BASE_URL = 'https://api.deepseek.com'
DEFAULT_DEEPSEEK_MODEL = 'deepseek-chat'


def _config_search_paths() -> list[Path]:
    paths: list[Path] = []
    env_path = os.environ.get('PANDASPRO_CONFIG')
    if env_path:
        paths.append(Path(env_path).expanduser())
    paths.append(Path.cwd() / 'local.yaml')
    paths.append(Path.home() / '.pandaspro' / 'local.yaml')
    return paths


def load_local_yaml() -> tuple[dict[str, Any], Path | None]:
    """
    按顺序查找并加载 local.yaml：
    1. 环境变量 PANDASPRO_CONFIG
    2. 当前工作目录 ./local.yaml
    3. ~/.pandaspro/local.yaml
    """
    if yaml is None:
        raise ImportError(
            '读取 local.yaml 需要 PyYAML：pip install pyyaml'
        )

    for path in _config_search_paths():
        if path.is_file():
            with path.open(encoding='utf-8') as f:
                data = yaml.safe_load(f) or {}
            if not isinstance(data, dict):
                raise ValueError(f'local.yaml 根节点必须是字典: {path}')
            return data, path
    return {}, None


def get_ai_settings() -> dict[str, Any]:
    """
    合并 local.yaml 与环境变量，返回 AI 调用配置。

    环境变量 PANDASPRO_AI_API_KEY 可覆盖 yaml 中的 api_key。
    """
    config, config_path = load_local_yaml()
    ai = dict(config.get('ai') or {})

    env_key = os.environ.get('PANDASPRO_AI_API_KEY')
    if env_key:
        ai['api_key'] = env_key

    ai.setdefault('provider', 'deepseek')
    ai.setdefault('base_url', DEFAULT_DEEPSEEK_BASE_URL)
    ai.setdefault('model', DEFAULT_DEEPSEEK_MODEL)
    ai['_config_path'] = str(config_path) if config_path else None
    return ai


def config_help_message() -> str:
    return (
        '未找到可用的 AI API Key。请在以下任一位置创建 local.yaml：\n'
        '  - ./local.yaml\n'
        '  - ~/.pandaspro/local.yaml\n'
        '  - 或设置环境变量 PANDASPRO_CONFIG 指向配置文件\n\n'
        '示例（local.yaml.example）：\n'
        '  ai:\n'
        '    provider: deepseek\n'
        '    api_key: sk-your-key\n'
        '    base_url: https://api.deepseek.com\n'
        '    model: deepseek-chat\n\n'
        '也可仅设置环境变量 PANDASPRO_AI_API_KEY=sk-...'
    )
