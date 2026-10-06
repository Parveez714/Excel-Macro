import json
import re

from .providers.factory import get_provider

SYSTEM_PROMPT = r'''You write safe, maintainable Excel VBA for a non-technical user.
VBA SYNTAX RULES:
1. Write EXACTLY ONE standalone Sub procedure (the entry point) containing the entire workflow. Do not generate multiple Sub procedures or concatenate separate procedures together.
2. In standard VBA string literals, use single normal quotes (e.g. Set ws = ThisWorkbook.Sheets("Sales Data")). Never use doubled quotes (""Sales Data"") for regular strings.
3. Only inside an Excel formula string literal should quotes be doubled (e.g. ws.Range("L2:L" & lastRow).Formula = "=IF(G2="""",0,G2)*H2*(1-I2)"). Never use backslash (\") anywhere.
4. When referencing worksheets and formatting ranges, declare ws As Worksheet and rng As Range (e.g., Set ws = ThisWorkbook.Sheets("SheetName") and Set rng = ws.Range("A2:L63")). Do not declare ws As Range or call .Rows("2:63") on a Range.
5. When calculations, aggregations, conditional values, or summaries are requested, write formulas directly into cells (for example using .Formula = "=SUM(B2:B10)" or .Formula2 = "=AVERAGE(...)") rather than hardcoding static calculation results, so that when users open the workbook in Excel, all dynamic Excel formulas are visible and live in the formula bar.

Return exactly one valid JSON object with exactly two string fields: "vba_code" and "summary". The vba_code value must contain raw VBA only, without Markdown fences. The summary must be 2-3 plain-English sentences explaining what the macro and formulas will do. Write code for a standard VBA module only. Do not create buttons, event hooks, Workbook_Open procedures, or modify workbook structure unless the user explicitly requests such behavior. Do not use shell commands, file deletion, registry access, or external network calls. Make the macro safe to run and include a clearly named public Sub as its entry point. Do not include the user's workbook data in your response.'''


def _clean_vba_syntax(code: str) -> str:
    # Ensure any joined End Sub / Sub statements are placed on new lines
    code = re.sub(r'(?i)End\s+Sub\s+(?:Public\s+)?Sub\s+', 'End Sub\n\nSub ', code)
    # Fix accidental doubled quotes in standard VBA object references: e.g. Sheets(""Name"") -> Sheets("Name")
    code = re.sub(r'([A-Za-z0-9_]+\s*\(\s*)""([^"=\n]+)""(\s*\))', r'\1"\2"\3', code)
    return code




def parse_generation_response(response: str) -> dict[str, str]:
    text = response.strip()
    fence = re.fullmatch(r"```(?:json)?\s*(.*?)\s*```", text, flags=re.IGNORECASE | re.DOTALL)
    if fence:
        text = fence.group(1).strip()
    start, end = text.find("{"), text.rfind("}")
    if start != -1 and end > start:
        text = text[start : end + 1]

    payload = None
    try:
        payload = json.loads(text, strict=False)
    except json.JSONDecodeError:
        # Resilient fallback: LLMs often output internal quotes like `rng.Formula = "=IF(G2="""",0,G2)"`
        # which breaks strict JSON parsers. Extract fields directly using structured regex.
        code_match = re.search(r'"vba_code"\s*:\s*"(.*?)"\s*,\s*"summary"\s*:\s*"(.*?)"\s*\}', text, re.DOTALL)
        if not code_match:
            # Try alternate key ordering
            code_match = re.search(r'"summary"\s*:\s*"(.*?)"\s*,\s*"vba_code"\s*:\s*"(.*?)"\s*\}', text, re.DOTALL)
            if code_match:
                payload = {"summary": code_match.group(1), "vba_code": code_match.group(2)}
        else:
            payload = {"vba_code": code_match.group(1), "summary": code_match.group(2)}

        if payload:
            # Unescape JSON control characters
            for k in ("vba_code", "summary"):
                val = payload[k]
                val = val.replace(r"\n", "\n").replace(r"\r", "\r").replace(r"\t", "\t")
                val = val.replace(r'\"', '"')
                payload[k] = val


    if not isinstance(payload, dict):
        raise ValueError("The AI response was not valid JSON. Preview the request again to retry.")

    code = payload.get("vba_code")
    summary = payload.get("summary")
    if not isinstance(code, str) or not code.strip() or not isinstance(summary, str) or not summary.strip():
        raise ValueError("The AI response is missing VBA code or its plain-English summary.")
    code = code.strip()
    if code.startswith("```") or code.endswith("```"):
        raise ValueError("The VBA code must not contain Markdown fences. Preview the request again to retry.")
    code = _clean_vba_syntax(code)
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
