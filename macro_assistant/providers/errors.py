import random
import time
from collections.abc import Callable
from typing import TypeVar

T = TypeVar("T")


class ProviderError(RuntimeError):
    """An actionable, user-facing provider error."""


def is_retryable(error: Exception) -> bool:
    if "quota" in str(error).lower():
        return False  # An exhausted quota won't recover in seconds; retrying only burns more requests.
    status = getattr(error, "status_code", None) or getattr(error, "code", None)
    if isinstance(status, int):
        return status in (429, 499) or status >= 500
    name = type(error).__name__.lower()
    text = str(error).lower()
    retry_words = (
        "timeout",
        "connection",
        "temporar",
        "resourceexhausted",
        "disconnect",
        "protocolerror",
        "reset",
        "brokenpipe",
        "unavailable",
        "eof",
        "socket",
        "network",
    )
    return any(word in name or word in text for word in retry_words)


def with_backoff(operation: Callable[[], T], attempts: int = 4) -> T:
    for attempt in range(attempts):
        try:
            return operation()
        except Exception as error:
            if attempt == attempts - 1 or not is_retryable(error):
                raise
            time.sleep((1.0 * (2**attempt)) + random.uniform(0, 0.5))
    raise RuntimeError("Retry operation ended unexpectedly")


def friendly_error(error: Exception) -> ProviderError:
    status = getattr(error, "status_code", None) or getattr(error, "code", None)
    err_str = str(error)
    if status in (401, 403) or "API_KEY_INVALID" in err_str:
        return ProviderError("The API key was rejected. Check the provider key in your .env file.")
    if "quota" in err_str.lower():
        return ProviderError("Your Gemini API quota is used up (free tier limits are low). Wait a minute or check your plan, or switch LLM_MODEL in .env to a lighter model such as a flash-lite variant.")
    if status == 429 or "RESOURCE_EXHAUSTED" in err_str:
        return ProviderError("The provider is rate-limiting requests. Wait a few seconds and try again.")
    if status == 503 or "UNAVAILABLE" in err_str:
        return ProviderError("The AI model is experiencing high demand. Please try again in a moment.")
    if status == 499 or "CANCELLED" in err_str:
        return ProviderError("The request to the AI provider was cancelled, usually because it took too long. Try again or use a faster model.")
    if isinstance(error, (TimeoutError, ConnectionError)) or is_retryable(error):
        return ProviderError("Could not reach the AI provider after retrying. Check your internet connection and try again.")
    return ProviderError(f"The AI provider returned an error: {error}")
