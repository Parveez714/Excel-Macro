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
9. Be mindful of merged cells: if a sheet contains existing title banners or merged header rows, place new tables or outputs on rows below the merged banner (e.g. starting at row 3), or explicitly unmerge with `ws.unmerge_cells(...)` before overwriting.

Return exactly one valid JSON object with exactly two string fields: "python_code" and "summary".
The python_code value must contain raw Python code only, without Markdown fences.
The summary must be 2-3 plain-English sentences explaining what changes, formulas, or summaries were added.
'''


def parse_generation_response(response: str) -> dict[str, str]:
    text = response.strip()
    if not text:
        raise ValueError("The AI response was not valid JSON. Preview the request again to retry.")

    # 1. Check for markdown-fenced json
    fence = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", text, flags=re.IGNORECASE | re.DOTALL)
    candidate = fence.group(1).strip() if fence else text

    # Extract JSON object substring
    start, end = candidate.find("{"), candidate.rfind("}")
    json_candidate = candidate[start : end + 1] if (start != -1 and end > start) else None
    if not json_candidate:
        start_orig, end_orig = text.find("{"), text.rfind("}")
        if start_orig != -1 and end_orig > start_orig:
            json_candidate = text[start_orig : end_orig + 1]

    payload = None
    if json_candidate:
        try:
            payload = json.loads(json_candidate, strict=False)
        except json.JSONDecodeError:
            pass

        # Clean invalid escapes or control characters
        if not payload:
            try:
                cleaned = re.sub(r'(?<!\\)\\(?!["\\/bfnrtu])', r'\\\\', json_candidate)
                payload = json.loads(cleaned, strict=False)
            except Exception:
                pass

        # Fallback regex extraction with delimiter matching
        if not payload:
            m = re.search(
                r'"(?:python_code|code|script)"\s*:\s*"(.*?)"\s*,\s*"(?:summary|description|explanation)"\s*:\s*"(.*?)"\s*\}',
                json_candidate,
                re.DOTALL,
            )
            if not m:
                m = re.search(
                    r'"(?:python_code|code|script)"\s*:\s*"(.*)(?:"\s*,\s*"(?:summary|description|explanation)"\s*:\s*"(.*?)"\s*\})',
                    json_candidate,
                    re.DOTALL,
                )
            if m:
                payload = {"python_code": m.group(1), "summary": m.group(2)}
            else:
                m2 = re.search(
                    r'"(?:summary|description|explanation)"\s*:\s*"(.*?)"\s*,\s*"(?:python_code|code|script)"\s*:\s*"(.*?)"\s*\}',
                    json_candidate,
                    re.DOTALL,
                )
                if not m2:
                    m2 = re.search(
                        r'"(?:summary|description|explanation)"\s*:\s*"(.*?)"\s*,\s*"(?:python_code|code|script)"\s*:\s*"(.*)"\s*\}',
                        json_candidate,
                        re.DOTALL,
                    )
                if m2:
                    payload = {"summary": m2.group(1), "python_code": m2.group(2)}

    # Fallback to raw markdown code block if model returned code without JSON
    if not isinstance(payload, dict) or not (payload.get("python_code") or payload.get("code") or payload.get("vba_code")):
        code_fence = re.search(r"```(?:python|py)?\s*\n(.*?)\n```", text, flags=re.DOTALL | re.IGNORECASE)
        if code_fence:
            code_extracted = code_fence.group(1).strip()
            before = text[: code_fence.start()].strip()
            after = text[code_fence.end() :].strip()
            summary_extracted = (before + " " + after).strip()
            summary_extracted = re.sub(
                r"^(?:Here is the (?:code|solution|transformation)[:\s]*|Summary[:\s]*)",
                "",
                summary_extracted,
                flags=re.IGNORECASE,
            ).strip()
            if not summary_extracted:
                summary_extracted = "Applied requested spreadsheet transformations."
            payload = {"python_code": code_extracted, "summary": summary_extracted}

    if not isinstance(payload, dict):
        raise ValueError("The AI response was not valid JSON. Preview the request again to retry.")

    # Unescape escaped characters in strings
    for k in ("python_code", "vba_code", "code", "script", "summary", "description", "explanation"):
        if k in payload and isinstance(payload[k], str):
            val = payload[k]
            if r"\n" in val:
                val = val.replace(r"\r", "\r").replace(r"\n", "\n").replace(r"\t", "\t")
            if r'\"' in val:
                val = val.replace(r'\"', '"')
            payload[k] = val

    code = payload.get("python_code") or payload.get("code") or payload.get("script") or payload.get("vba_code")
    summary = payload.get("summary") or payload.get("description") or payload.get("explanation")
    if not isinstance(code, str) or not code.strip() or not isinstance(summary, str) or not summary.strip():
        raise ValueError("The AI response is missing the code or its plain-English summary.")
    code = code.strip()
    if code.startswith("```"):
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
