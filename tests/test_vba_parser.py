import pytest

from macro_assistant.vba_generation import parse_generation_response


def test_parser_accepts_json_markdown_fence():
    result = parse_generation_response('```json\n{"vba_code":"Public Sub RunTask()\\nEnd Sub","summary":"This runs a task. It changes selected cells."}\n```')

    assert result["vba_code"] == "Public Sub RunTask()\nEnd Sub"
    assert result["summary"].startswith("This runs")


def test_parser_rejects_invalid_json():
    with pytest.raises(ValueError, match="valid JSON"):
        parse_generation_response("not json")


def test_parser_requires_both_fields():
    with pytest.raises(ValueError, match="missing VBA code"):
        parse_generation_response('{"vba_code":"", "summary":"A summary."}')


def test_parser_rejects_code_with_fences_inside_json():
    with pytest.raises(ValueError, match="Markdown fences"):
        parse_generation_response('{"vba_code":"```vba\\nSub RunTask()\\n```", "summary":"Runs a task."}')
