import re


_RISK_PATTERNS = (
    ("Shell command execution", re.compile(r"\bShell\s*(?:\(|\b)", re.IGNORECASE)),
    ("Windows Script Host access", re.compile(r"CreateObject\s*\(\s*[\"']WScript\.Shell[\"']", re.IGNORECASE)),
    ("File deletion", re.compile(r"\bKill\s+|\.Delete(?:File|Folder)?\b", re.IGNORECASE)),
    ("FileSystemObject deletion", re.compile(r"FileSystemObject[\s\S]{0,300}(?:DeleteFile|DeleteFolder|\.Delete)\b", re.IGNORECASE)),
    ("URL download or external network access", re.compile(r"URLDownloadToFile|XMLHTTP|WinHttpRequest|InternetOpen|InternetConnect|ADODB\.Stream|FollowHyperlink|QueryTables\.Add|Connections\.Add|WebQueries\.Add", re.IGNORECASE)),
    ("Registry modification", re.compile(r"\b(?:RegWrite|RegDelete|SaveSetting|DeleteSetting|RegCreateKey\w*|RegSetValue\w*|RegDeleteKey\w*)\b", re.IGNORECASE)),
)


def scan_vba(code: str) -> list[str]:
    """Return human-readable safety warnings found by conservative static checks."""
    findings = []
    for label, pattern in _RISK_PATTERNS:
        if pattern.search(code) and label not in findings:
            findings.append(label)
    return findings
