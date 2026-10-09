import unittest

from sport_meeting.core.legacy_booklet_content import (
    LAST_YEAR_SCHOOL_RECORDS,
    LATEST_SCHOOL_RECORDS,
    TENTH_MEET_RECORD_UPDATES,
    updated_school_records,
)
from sport_meeting.core.models import ProjectConfig


def record_rows(text: str) -> dict[tuple[str, str], tuple[str, str, str]]:
    rows = [line.split("\t") for line in text.splitlines()[1:]]
    return {(event, group): (score, holder, meet) for event, group, score, holder, meet in rows}


class SchoolRecordsTests(unittest.TestCase):
    def test_tenth_meet_list_updates_twenty_matching_records(self):
        old = record_rows(LAST_YEAR_SCHOOL_RECORDS)
        latest = record_rows(LATEST_SCHOOL_RECORDS)
        self.assertEqual(len(TENTH_MEET_RECORD_UPDATES), 20)
        self.assertEqual(len(old), len(latest))
        self.assertTrue(TENTH_MEET_RECORD_UPDATES.keys() <= old.keys())
        for key, (score, holder) in TENTH_MEET_RECORD_UPDATES.items():
            self.assertEqual(latest[key], (score, holder, "第十届校运会"))
        for key in old.keys() - TENTH_MEET_RECORD_UPDATES.keys():
            self.assertEqual(latest[key], old[key])

    def test_official_jump_rope_value_and_unit_normalization(self):
        rows = record_rows(LATEST_SCHOOL_RECORDS)
        self.assertEqual(rows[("1分钟跳绳", "初二女子")], ("210个", "示例选手69", "第十届校运会"))
        self.assertEqual(rows[("跳远", "高二女子")][0], "310厘米")
        self.assertEqual(rows[("立定跳远", "初二男子")][0], "310厘米")
        self.assertEqual(ProjectConfig().school_records_text, LATEST_SCHOOL_RECORDS)
        self.assertEqual(updated_school_records(), LATEST_SCHOOL_RECORDS)


if __name__ == "__main__":
    unittest.main()
