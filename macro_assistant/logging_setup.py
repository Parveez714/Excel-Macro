import json
import logging
import os
from datetime import datetime, timezone
from pathlib import Path


class JsonFormatter(logging.Formatter):
    def format(self, record: logging.LogRecord) -> str:
        return json.dumps(
            {
                "timestamp": datetime.now(timezone.utc).isoformat(),
                "level": record.levelname,
                "event": record.getMessage(),
            },
            ensure_ascii=True,
        )


def get_logger() -> logging.Logger:
    logger = logging.getLogger("macro_assistant")
    if logger.handlers:
        return logger
    base = Path(os.getenv("LOCALAPPDATA", Path.home())) / "MacroAssistant" / "logs"
    base.mkdir(parents=True, exist_ok=True)
    handler = logging.FileHandler(base / "app.log", encoding="utf-8")
    handler.setFormatter(JsonFormatter())
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    logger.propagate = False
    return logger
