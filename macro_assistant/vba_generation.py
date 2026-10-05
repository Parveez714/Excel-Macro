import json
import re

from .providers.factory import get_provider

SYSTEM_PROMPT = """You write safe, maintainable Excel VBA for a non-technical user. Return exactly one valid JSON object with exactly two string fields: "vba_code" and "summary". The vba_code value must contain raw VBA only, without Markdown fences. The summary must be 2-3 plain-English sentences explaining what the macro will do. Write code for a standard VBA module only. Do not create buttons, event hooks, Workbook_Open procedures, or modify workbook structure unless the user explicitly requests such behavior. Do not use shell commands, file deletion, registry access, or external network calls. Make the macro safe to run and include a clearly named public Sub as its entry point. Do not include the user's workbook data in your response."""


def parse_generation_response(response: str) -> dict[str, str]:
    text = response.strip()
    fence = re.fullmatch(r"```(?:json)?\s*(.*?)\s*```", text, flags=re.IGNORECASE | re.DOTALL)
    if fence:
        text = fence.group(1).strip()
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end > start:
        text = text[start : end + 1]
    try:
        payload = json.loads(text, strict=False)
    except json.JSONDecodeError as error:
        raise ValueError("The AI response was not valid JSON. Preview the request again to retry.") from error
    if not isinstance(payload, dict):
        raise ValueError("The AI response must be a JSON object with macro code and a summary.")
    code = payload.get("vba_code")
    summary = payload.get("summary")
    if not isinstance(code, str) or not code.strip() or not isinstance(summary, str) or not summary.strip():
        raise ValueError("The AI response is missing VBA code or its plain-English summary.")
    code = code.strip()
    if code.startswith("```") or code.endswith("```"):
        raise ValueError("The VBA code must not contain Markdown fences. Preview the request again to retry.")
    return {"vba_code": code, "summary": summary.strip()}


def build_user_prompt(request: str, workbook_context: str, repair_error: str | None = None) -> str:
    prompt = (
        "Create a VBA macro for this plain-language request.\n\n"
        f"Request:\n{request.strip()}\n\n"
        "Workbook metadata (sheet names, used ranges, and header rows only; no cell body data):\n"
        f"{workbook_context}"
    )
    if repair_error:
        prompt += (
            "\n\nThe previous VBA could not be applied. Correct the VBA based on this error, "
            "preserve the original request, and return the same JSON format:\n"
            f"{repair_error[:1000]}"
        )
    return prompt


def generate_macro(provider_id: str, model: str, api_key: str, request: str, context: str, repair_error: str | None = None) -> dict[str, str]:
    provider = get_provider(provider_id, api_key, model)
    response = provider.generate(SYSTEM_PROMPT, build_user_prompt(request, context, repair_error))
    return parse_generation_response(response)
