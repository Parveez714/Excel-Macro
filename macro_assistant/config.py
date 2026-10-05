import os
import sys
from dataclasses import dataclass
from pathlib import Path

import yaml
from dotenv import load_dotenv

from .providers.factory import get_provider_class


def app_directory() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


@dataclass(frozen=True)
class AppConfig:
    provider: str
    model: str
    api_key: str


def _read_settings(root: Path) -> dict[str, str]:
    config_path = root / "config.yaml"
    if not config_path.exists():
        return {}
    try:
        raw_config = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
    except (OSError, yaml.YAMLError) as error:
        raise ValueError("Could not read config.yaml. Check its YAML formatting.") from error
    if not isinstance(raw_config, dict):
        raise ValueError("config.yaml must contain a simple set of key-value settings.")
    return {str(key): os.path.expandvars(str(value)) for key, value in raw_config.items() if value is not None}


def configured_provider() -> str:
    root = app_directory()
    load_dotenv(root / ".env", override=True)
    values = _read_settings(root)
    return (os.getenv("LLM_PROVIDER") or values.get("provider") or "gemini").strip().lower()


def load_config(provider_override: str | None = None) -> AppConfig:
    root = app_directory()
    load_dotenv(root / ".env", override=True)
    values = _read_settings(root)

    provider = (provider_override or os.getenv("LLM_PROVIDER") or values.get("provider") or "gemini").strip().lower()
    model = (os.getenv("LLM_MODEL") or values.get("model") or "").strip()
    try:
        api_key_env = get_provider_class(provider).api_key_env
    except ValueError as error:
        raise ValueError(str(error)) from None
    api_key = os.getenv(api_key_env, "").strip()

    if not model:
        raise ValueError("Add your provider's model name to LLM_MODEL in the .env file before previewing a macro.")
    if not api_key:
        raise ValueError(f"Add your {provider.title()} API key to {api_key_env} in the .env file before previewing a macro.")
    return AppConfig(provider=provider, model=model, api_key=api_key)
