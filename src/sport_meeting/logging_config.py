from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


ALLOWED_LOG_FIELDS = {
    "schema_version",
    "rules_version",
    "template_version",
    "random_seed",
    "input_hash",
    "sheet_count",
    "row_count",
    "validation_code",
    "validation_count",
    "stage",
    "elapsed_ms",
    "adapter",
    "error_code",
    "timeout_seconds",
    "controlled_pid",
    "output_name_hash",
    "artifact_id",
}


class AllowlistAdapter(logging.LoggerAdapter):
    def process(self, msg, kwargs):
        supplied = kwargs.pop("fields", {})
        safe = {key: value for key, value in supplied.items() if key in ALLOWED_LOG_FIELDS}
        suffix = " ".join(f"{key}={value}" for key, value in sorted(safe.items()))
        return (f"{msg} {suffix}".rstrip(), kwargs)


def configure_logging(log_directory: Path) -> AllowlistAdapter:
    log_directory.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("sport_meeting")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        handler = RotatingFileHandler(
            log_directory / "sport-meeting.log",
            maxBytes=5 * 1024 * 1024,
            backupCount=4,
            encoding="utf-8",
        )
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
        logger.addHandler(handler)
    return AllowlistAdapter(logger, {})

