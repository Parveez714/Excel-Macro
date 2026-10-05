from abc import ABC, abstractmethod
from typing import ClassVar


class LLMProvider(ABC):
    """Base interface for text-generation providers."""

    provider_id: ClassVar[str]
    api_key_env: ClassVar[str]
    registry: ClassVar[dict[str, type["LLMProvider"]]] = {}

    def __init_subclass__(cls, **kwargs):
        super().__init_subclass__(**kwargs)
        provider_id = getattr(cls, "provider_id", None)
        if provider_id:
            LLMProvider.registry[provider_id] = cls

    def __init__(self, api_key: str, model: str):
        self.api_key = api_key
        self.model = model

    @abstractmethod
    def generate(self, system_prompt: str, user_prompt: str) -> str:
        """Return generated text for the supplied prompts."""
        raise NotImplementedError
