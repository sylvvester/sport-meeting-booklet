from __future__ import annotations

import re
from collections import Counter

from .models import (
    Athlete,
    ConfigurationError,
    Entry,
    EventRound,
    EventType,
    ProjectConfig,
    RoundType,
    Severity,
    ValidationIssue,
    ValidationReport,
)


def validate_project(
    config: ProjectConfig,
    athletes: list[Athlete],
    entries: list[Entry],
    rounds: list[EventRound],
) -> ValidationReport:
    report = ValidationReport()
    try:
        config.validate()
    except ConfigurationError as exc:
        for message in exc.errors:
            report.add(ValidationIssue("HEAT_001", Severity.BLOCKING, message, "PROJECT_CONFIG", "singleton"))

    bib_counts = Counter(athlete.bib for athlete in athletes)
    for bib, count in bib_counts.items():
        if count > 1:
            report.add(ValidationIssue("IMPORT_008", Severity.BLOCKING, "运动员号码不是全校唯一", "ATHLETE", bib))

    if not entries:
        report.add(
            ValidationIssue(
                "IMPORT_009", Severity.BLOCKING, "项目中尚未保存报名明细，请先导入报名表",
                "PROJECT", "singleton",
            )
        )

    entry_counts = Counter((entry.group_name, entry.event_name) for entry in entries)
    preliminary_keys = {
        (round_.group_name, round_.event_name)
        for round_ in rounds
        if round_.round_type is RoundType.PRELIMINARY
    }
    final_keys = {
        (round_.group_name, round_.event_name)
        for round_ in rounds
        if round_.round_type is RoundType.FINAL
    }
    for key in sorted(preliminary_keys - final_keys):
        report.add(ValidationIssue("SCHEDULE_001", Severity.BLOCKING, "预赛缺少对应决赛", "EVENT_ROUND", ":".join(key)))

    for round_ in rounds:
        actual = entry_counts[(round_.group_name, round_.event_name)]
        if (
            entries
            and round_.event_type is not EventType.TEAM
            and round_.round_type is not RoundType.FINAL
            and round_.declared_count is not None
            and actual != round_.declared_count
        ):
            report.add(
                ValidationIssue(
                    "SCHEDULE_002", Severity.BLOCKING,
                    f"日程人数 {round_.declared_count} 与报名实数 {actual} 不一致",
                    "EVENT_ROUND", round_.id, original_value=round_.declared_count, normalized_value=actual,
                )
            )
        if (
            entries
            and round_.event_type is EventType.TRACK
            and round_.round_type is not RoundType.FINAL
        ):
            try:
                distance_match = re.search(r"(?<!×)(\d+)\s*米", round_.event_name)
                distance = int(distance_match.group(1)) if distance_match else None
                if distance and distance >= 800:
                    continue
                if distance == 400:
                    capacity = 16
                else:
                    capacity = len(round_.effective_available_lanes(config))
                largest_heat = (actual + round_.heat_count - 1) // round_.heat_count if round_.heat_count else actual
                if round_.heat_count < 1 or largest_heat > capacity:
                    report.add(
                        ValidationIssue(
                            "HEAT_002", Severity.BLOCKING,
                            f"最大单组人数 {largest_heat} 超过本场次 {capacity} 条可用跑道",
                            "EVENT_ROUND", round_.id,
                        )
                    )
            except ConfigurationError as exc:
                for message in exc.errors:
                    report.add(ValidationIssue("HEAT_003", Severity.BLOCKING, message, "EVENT_ROUND", round_.id))
    return report
