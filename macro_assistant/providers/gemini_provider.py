from .base import LLMProvider
from .errors import ProviderError, friendly_error, with_backoff


class GeminiProvider(LLMProvider):
    provider_id = "gemini"
    api_key_env = "GEMINI_API_KEY"

    def __init__(self, api_key: str, model: str, client=None):
        super().__init__(api_key, model)
        if client is None:
            from google import genai

            client = genai.Client(api_key=api_key, http_options={"timeout": 180000})
        self.client = client

    def _call(self, system_prompt: str, user_prompt: str, fast: bool):
        config = {
            "system_instruction": system_prompt,
            "response_mime_type": "application/json",
            "max_output_tokens": 8192,
            "automatic_function_calling": {"disable": True},
        }
        if fast:
            # Macro writing doesn't need long reasoning; this cuts latency and quota use.
            config["thinking_level"] = "low"
        return self.client.models.generate_content(model=self.model, contents=user_prompt, config=config)



    def generate(self, system_prompt: str, user_prompt: str) -> str:
        try:
            try:
                response = with_backoff(lambda: self._call(system_prompt, user_prompt, True))
            except Exception as error:
                if "thinking" not in str(error).lower():
                    raise
                response = with_backoff(lambda: self._call(system_prompt, user_prompt, False))
            content = response.text
            if not content:
                raise ProviderError("The AI provider returned an empty response. Try rephrasing your request.")
            return content
        except ProviderError:
            raise
        except Exception as error:
            raise friendly_error(error) from None
