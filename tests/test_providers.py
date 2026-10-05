from types import SimpleNamespace
from unittest.mock import patch

from macro_assistant.providers.errors import with_backoff
from macro_assistant.providers.gemini_provider import GeminiProvider
from macro_assistant.providers.openai_provider import OpenAIProvider


def test_openai_provider_sends_prompts_and_returns_text():
    client = SimpleNamespace(
        chat=SimpleNamespace(
            completions=SimpleNamespace(
                create=lambda **kwargs: SimpleNamespace(
                    choices=[SimpleNamespace(message=SimpleNamespace(content="generated text"))]
                )
            )
        )
    )
    provider = OpenAIProvider("test-key", "test-model", client=client)

    assert provider.generate("system", "user") == "generated text"


def test_gemini_provider_sends_prompts_and_returns_text():
    calls = []

    def generate_content(**kwargs):
        calls.append(kwargs)
        return SimpleNamespace(text="generated text")

    client = SimpleNamespace(models=SimpleNamespace(generate_content=generate_content))
    provider = GeminiProvider("test-key", "test-model", client=client)

    assert provider.generate("system", "user") == "generated text"
    assert calls[0]["model"] == "test-model"
    assert calls[0]["contents"] == "user"
    assert calls[0]["config"]["system_instruction"] == "system"


def test_provider_retries_transient_errors():
    class TemporaryError(Exception):
        status_code = 503

    calls = 0

    def operation():
        nonlocal calls
        calls += 1
        if calls < 2:
            raise TemporaryError()
        return "ok"

    with patch("macro_assistant.providers.errors.time.sleep"):
        assert with_backoff(operation) == "ok"
    assert calls == 2
