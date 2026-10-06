import json
import re

from .providers.factory import get_provider

SYSTEM_PROMPT = r'''You write clean, safe, and maintainable Python scripts using openpyxl to transform Excel spreadsheets for users.
Your script transforms a workbook object named `wb`.

RULES FOR THE PYTHON CODE:
1. The script will be executed in an environment where `wb` is ALREADY loaded (openpyxl.Workbook).
2. Do not call openpyxl.load_workbook() or wb.save(). Simply perform the operations on `wb`.
3. When calculations, aggregations, or conditional results are requested, write LIVE EXCEL FORMULAS into the cells using formula strings, for example:
   `ws.cell(row=r, column=12, value=f'=IF(G{r}="",0,G{r})*H{r}*(1-I{r})')` or `ws['C10'] = '=SUM(C2:C9)'`
   Never hardcode static math calculations when a formula is asked.
4. When text trimming or casing is asked, update the cell values directly:
   `cell.value = str(cell.value).strip().title() if cell.value is not None else None`
5. When creating a summary table or pivot-style sheet, you can add a new worksheet with `wb.create_sheet(title='Pivot')`, calculate the aggregated data, and format it neatly with headers and cell borders/fills.
6. Use standard openpyxl modules and helpers (e.g., `openpyxl.styles.Font`, `PatternFill`, `Alignment`, `Border`, `Side`).
7. Do not use any unsafe modules (os, subprocess, sys, requests, socket, etc.).
8. Write code that executes quickly and robustly.

Return exactly one valid JSON object with exactly two string fields: "python_code" and "summary".
The python_code value must contain raw Python code only, without Markdown fences.
The summary must be 2-3 plain-English sentences explaining what changes, formulas, or summaries were added.
'''


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
        # Fallback regex extraction
        code_match = re.search(r'"python_code"\s*:\s*"(.*?)"\s*,\s*"summary"\s*:\s*"(.*?)"\s*\}', text, re.DOTALL)
        if not code_match:
            code_match = re.search(r'"summary"\s*:\s*"(.*?)"\s*,\s*"python_code"\s*:\s*"(.*?)"\s*\}', text, re.DOTALL)
            if code_match:
                payload = {"summary": code_match.group(1), "python_code": code_match.group(2)}
        else:
            payload = {"python_code": code_match.group(1), "summary": code_match.group(2)}

        if payload:
            for k in ("python_code", "summary"):
                val = payload[k]
                val = val.replace(r"\n", "\n").replace(r"\r", "\r").replace(r"\t", "\t")
                val = val.replace(r'\"', '"')
                payload[k] = val

    if not isinstance(payload, dict):
        raise ValueError("The AI response was not valid JSON. Preview the request again to retry.")

    code = payload.get("python_code") or payload.get("vba_code")
    summary = payload.get("summary")
    if not isinstance(code, str) or not code.strip() or not isinstance(summary, str) or not summary.strip():
        raise ValueError("The AI response is missing the code or its plain-English summary.")
    code = code.strip()
    if code.startswith("```") or code.endswith("```"):
        lines = code.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        code = "\n".join(lines).strip()
    return {"python_code": code, "summary": summary.strip()}


def build_user_prompt(request: str, workbook_context: str, repair_error: str | None = None) -> str:
    prompt = (
        "Create an openpyxl Python transformation script for this plain-language request.\n\n"
        f"Request:\n{request.strip()}\n\n"
        "Workbook metadata (sheet names, used ranges, and header rows only; no cell body data):\n"
        f"{workbook_context}"
    )
    if repair_error:
        prompt += (
            "\n\nThe previous code could not be applied. Correct the Python code based on this error, "
            "preserve the original request, and return the same JSON format:\n"
            f"{repair_error[:1000]}"
        )
    return prompt


def generate_transformation(provider_id: str, model: str, api_key: str, request: str, context: str, repair_error: str | None = None) -> dict[str, str]:
    provider = get_provider(provider_id, api_key, model)
    response = provider.generate(SYSTEM_PROMPT, build_user_prompt(request, context, repair_error))
    return parse_generation_response(response)


# Backwards compatibility alias
generate_macro = generate_transformation
