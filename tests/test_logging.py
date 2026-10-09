import tempfile
import unittest
from pathlib import Path

from sport_meeting.logging_config import configure_logging


class LoggingTests(unittest.TestCase):
    def test_logger_drops_fields_outside_allowlist(self):
        with tempfile.TemporaryDirectory() as directory:
            logger = configure_logging(Path(directory))
            logger.info(
                "test",
                fields={"row_count": 2, "athlete_name": "不应记录", "absolute_path": "C:/secret.xlsx"},
            )
            for handler in logger.logger.handlers:
                handler.flush()
            content = (Path(directory) / "sport-meeting.log").read_text(encoding="utf-8")
            self.assertIn("row_count=2", content)
            self.assertNotIn("不应记录", content)
            self.assertNotIn("secret.xlsx", content)
            for handler in list(logger.logger.handlers):
                handler.close()
                logger.logger.removeHandler(handler)


if __name__ == "__main__":
    unittest.main()
