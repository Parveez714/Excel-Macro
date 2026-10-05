from .base import LLMProvider
from .errors import ProviderError, friendly_error, with_backoff


class OpenAIProvider(LLMProvider):
    provider_id = "openai"
    api_key_env = "OPENAI_API_KEY"

    def __init__(self, api_key: str, model: str, client=None):
        super().__init__(api_key, model)
        if client is None:
            from openai import OpenAI

            client = OpenAI(api_key=api_key)
        self.client = client

    def generate(self, system_prompt: str, user_prompt: str) -> str:
        try:
            response = with_backoff(
                lambda: self.client.chat.completions.create(
                    model=self.model,
                    messages=[
                        {"role": "system", "content": system_prompt},
                        {"role": "user", "content": user_prompt},
                    ],
                    response_format={"type": "json_object"},
                )
            )
            content = response.choices[0].message.content
            if not content:
                raise ProviderError("The AI provider returned an empty response. Try rephrasing your request.")
            return content
        except ProviderError:
            raise
        except Exception as error:
            raise friendly_error(error) from None
