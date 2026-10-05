import importlib
import pkgutil

from . import __path__
from .base import LLMProvider


def discover_providers() -> dict[str, type[LLMProvider]]:
    for module in pkgutil.iter_modules(__path__):
        if module.name not in {"base", "errors", "factory"}:
            importlib.import_module(f"{__package__}.{module.name}")
    return dict(LLMProvider.registry)


def get_provider_class(provider_id: str) -> type[LLMProvider]:
    providers = discover_providers()
    provider_class = providers.get(provider_id.lower())
    if provider_class is None:
        available = ", ".join(sorted(providers)) or "none"
        raise ValueError(f"Unknown provider '{provider_id}'. Available providers: {available}.")
    return provider_class


def available_providers() -> tuple[str, ...]:
    return tuple(sorted(discover_providers()))


def get_provider(provider_id: str, api_key: str, model: str) -> LLMProvider:
    provider_class = get_provider_class(provider_id)
    return provider_class(api_key=api_key, model=model)
