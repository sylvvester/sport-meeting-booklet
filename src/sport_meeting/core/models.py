from __future__ import annotations

from datetime import date
from dataclasses import dataclass, field
from enum import StrEnum
from typing import Any

from .legacy_booklet_content import (
    LAST_YEAR_BOOKLET_TITLE,
    LAST_YEAR_COMPETITION_RULES,
    LAST_YEAR_OFFICIALS,
    LAST_YEAR_OPENING_CEREMONY,
    LAST_YEAR_ORGANIZING_COMMITTEE,
    LATEST_SCHOOL_RECORDS,
)


class ConfigurationError(ValueError):
    def __init__(self, errors: list[str]):
        super().__init__("；".join(errors))
        self.errors = errors


class PerformanceKind(StrEnum):
    TIME = "TIME"
    DISTANCE = "DISTANCE"
    COUNT = "COUNT"
    MANUAL_RANK = "MANUAL_RANK"


class ResultStatus(StrEnum):
    VALID = "VALID"
    DNS = "DNS"
    DNF = "DNF"
    DQ = "DQ"
    NM = "NM"


class RoundType(StrEnum):
    PRELIMINARY = "PRELIMINARY"
    FINAL = "FINAL"
    TIMED_FINAL = "TIMED_FINAL"


class EventType(StrEnum):
    TRACK = "TRACK"
    FIELD = "FIELD"
    TEAM = "TEAM"


class FieldStartOrder(StrEnum):
    NUMBER_ASC = "NUMBER_ASC"
    SEEDED_DRAW = "SEEDED_DRAW"


class Severity(StrEnum):
    BLOCKING = "BLOCKING"
    WARNING = "WARNING"
    AUTO_FIXED = "AUTO_FIXED"


DEFAULT_JUNIOR_BOYS_EVENTS = (
    "50米", "400米", "1000米", "1分钟跳绳", "立定跳远", "引体向上",
)
DEFAULT_JUNIOR_GIRLS_EVENTS = (
    "50米", "400米", "800米", "1分钟跳绳", "立定跳远", "仰卧起坐",
)
DEFAULT_SENIOR_BOYS_EVENTS = (
    "100米", "200米", "400米", "1500米", "跳远", "引体向上",
)
DEFAULT_SENIOR_GIRLS_EVENTS = (
    "100米", "200米", "400米", "800米", "跳远", "仰卧起坐",
)

ROUND_MODE_TIMED_FINAL = "TIMED_FINAL"
ROUND_MODE_PRELIMINARY_FINAL = "PRELIMINARY_FINAL"
ROUND_MODES = {ROUND_MODE_TIMED_FINAL, ROUND_MODE_PRELIMINARY_FINAL}
TEAM_ROUND_MODE_KNOCKOUT = "KNOCKOUT"
TEAM_GROUP_MEN = "MEN"
TEAM_GROUP_WOMEN = "WOMEN"
TEAM_GROUP_MIXED = "MIXED"
TEAM_GROUP_STAFF = "STAFF"
TEAM_GROUP_MODES = {TEAM_GROUP_MEN, TEAM_GROUP_WOMEN, TEAM_GROUP_MIXED, TEAM_GROUP_STAFF}
TEAM_COUNT_ONE_PER_CLASS = "ONE_PER_CLASS"
TEAM_COUNT_MANUAL = "MANUAL"
TEAM_COUNT_MODES = {TEAM_COUNT_ONE_PER_CLASS, TEAM_COUNT_MANUAL}


@dataclass(frozen=True, slots=True)
class TeamEventConfig:
    event_name: str
    group_mode: str
    count_mode: str
    round_mode: str = ROUND_MODE_TIMED_FINAL
    manual_team_count: int | None = None


DEFAULT_TEAM_EVENTS = (
    TeamEventConfig("4×100米接力", TEAM_GROUP_MEN, TEAM_COUNT_ONE_PER_CLASS),
    TeamEventConfig("4×100米接力", TEAM_GROUP_WOMEN, TEAM_COUNT_ONE_PER_CLASS),
    TeamEventConfig("1分钟集体跳绳", TEAM_GROUP_MIXED, TEAM_COUNT_ONE_PER_CLASS),
    TeamEventConfig(
        "拔河比赛", TEAM_GROUP_MIXED, TEAM_COUNT_MANUAL, TEAM_ROUND_MODE_KNOCKOUT
    ),
    TeamEventConfig("袋鼠跳接力赛", TEAM_GROUP_STAFF, TEAM_COUNT_MANUAL),
)


@dataclass(frozen=True, slots=True)
class Athlete:
    id: str
    bib: str
    name: str
    sex: str
    grade: str
    class_name: str
    campus: str = ""

    @property
    def unit(self) -> str:
        return f"{self.grade}({self.class_name})"


@dataclass(frozen=True, slots=True)
class Entry:
    id: str
    athlete_id: str
    event_name: str
    group_name: str
    source_file_id: str = ""
    source_sheet: str = ""
    source_row: int | None = None


@dataclass(frozen=True, slots=True)
class SourceFileSnapshot:
    id: str
    source_path: str
    content_hash: str
    sheet_count: int
    row_count: int
    imported_at: str


@dataclass(frozen=True, slots=True)
class Participant:
    id: str
    bib: str
    name: str
    unit: str


@dataclass(slots=True)
class ProjectConfig:
    schema_version: int = 1
    rules_version: str = "1.0"
    template_version: int = 1
    finalists_quota: int = 8
    track_lanes: int = 8
    lane_priority: tuple[int, ...] = (4, 5, 3, 6, 2, 7, 1, 8)
    preliminary_lane_start: int = 2
    field_start_order: FieldStartOrder = FieldStartOrder.NUMBER_ASC
    backup_retention: int = 10
    meeting_start_date: str = ""
    meeting_end_date: str = ""
    junior_boys_events: tuple[str, ...] = DEFAULT_JUNIOR_BOYS_EVENTS
    junior_girls_events: tuple[str, ...] = DEFAULT_JUNIOR_GIRLS_EVENTS
    senior_boys_events: tuple[str, ...] = DEFAULT_SENIOR_BOYS_EVENTS
    senior_girls_events: tuple[str, ...] = DEFAULT_SENIOR_GIRLS_EVENTS
    junior_boys_event_formats: dict[str, str] = field(default_factory=dict)
    junior_girls_event_formats: dict[str, str] = field(default_factory=dict)
    senior_boys_event_formats: dict[str, str] = field(default_factory=dict)
    senior_girls_event_formats: dict[str, str] = field(default_factory=dict)
    team_events: tuple[TeamEventConfig, ...] = DEFAULT_TEAM_EVENTS
    booklet_title: str = LAST_YEAR_BOOKLET_TITLE
    opening_ceremony_text: str = LAST_YEAR_OPENING_CEREMONY
    organizing_committee_text: str = LAST_YEAR_ORGANIZING_COMMITTEE
    officials_text: str = LAST_YEAR_OFFICIALS
    competition_rules_text: str = LAST_YEAR_COMPETITION_RULES
    school_records_text: str = LATEST_SCHOOL_RECORDS

    def validate(self) -> None:
        errors: list[str] = []
        if self.schema_version < 1:
            errors.append("schema_version 必须大于 0")
        if self.finalists_quota < 1:
            errors.append("finalists_quota 必须大于 0")
        if self.track_lanes < 1:
            errors.append("track_lanes 必须大于 0")
        expected = set(range(1, self.track_lanes + 1))
        if len(self.lane_priority) != self.track_lanes or set(self.lane_priority) != expected:
            errors.append("lane_priority 必须是 1..track_lanes 的完整无重复排列")
        if not 1 <= self.preliminary_lane_start <= self.track_lanes:
            errors.append("preliminary_lane_start 必须位于物理跑道范围内")
        if self.backup_retention < 1:
            errors.append("backup_retention 必须大于 0")
        if bool(self.meeting_start_date) != bool(self.meeting_end_date):
            errors.append("运动会开始和结束日期必须同时设置")
        elif self.meeting_start_date and self.meeting_end_date:
            try:
                start_date = date.fromisoformat(self.meeting_start_date)
                end_date = date.fromisoformat(self.meeting_end_date)
                if end_date < start_date:
                    errors.append("运动会结束日期不能早于开始日期")
            except ValueError:
                errors.append("运动会日期必须使用 YYYY-MM-DD 格式")
        if not self.booklet_title.strip():
            errors.append("秩序册封面名称不能为空")
        event_groups = {
            "初中男生项目": (self.junior_boys_events, self.junior_boys_event_formats),
            "初中女生项目": (self.junior_girls_events, self.junior_girls_event_formats),
            "高中男生项目": (self.senior_boys_events, self.senior_boys_event_formats),
            "高中女生项目": (self.senior_girls_events, self.senior_girls_event_formats),
        }
        for label, (events, formats) in event_groups.items():
            cleaned = [event.strip() for event in events if event.strip()]
            if not cleaned:
                errors.append(f"{label}至少需要一个项目")
            elif len(cleaned) != len(set(cleaned)):
                errors.append(f"{label}不能包含重复项目")
            unknown_events = set(formats) - set(cleaned)
            if unknown_events:
                errors.append(f"{label}赛制包含未配置项目：{'、'.join(sorted(unknown_events))}")
            invalid_modes = sorted({mode for mode in formats.values() if mode not in ROUND_MODES})
            if invalid_modes:
                errors.append(f"{label}包含无效赛制：{'、'.join(invalid_modes)}")
        seen_team_events: set[tuple[str, str]] = set()
        for team_event in self.team_events:
            if not team_event.event_name.strip():
                errors.append("集体项目名称不能为空")
            key = (team_event.event_name.strip(), team_event.group_mode)
            if key in seen_team_events:
                errors.append(f"集体项目不能重复：{team_event.event_name}")
            seen_team_events.add(key)
            if team_event.group_mode not in TEAM_GROUP_MODES:
                errors.append(f"{team_event.event_name}包含无效参赛组别")
            if team_event.count_mode not in TEAM_COUNT_MODES:
                errors.append(f"{team_event.event_name}包含无效队数来源")
            if team_event.round_mode not in ROUND_MODES | {TEAM_ROUND_MODE_KNOCKOUT}:
                errors.append(f"{team_event.event_name}包含无效赛制")
            if team_event.manual_team_count is not None and team_event.manual_team_count < 1:
                errors.append(f"{team_event.event_name}手动队数必须大于0或留空")
        if errors:
            raise ConfigurationError(errors)


@dataclass(frozen=True, slots=True)
class EventRound:
    id: str
    event_name: str
    group_name: str
    event_type: EventType
    round_type: RoundType
    heat_count: int
    scheduled_time: str
    declared_count: int | None = None
    available_lanes: tuple[int, ...] = ()
    performance_kind: PerformanceKind = PerformanceKind.TIME
    canonical_unit: str = "MILLISECOND"
    default_input_unit: str = "SECOND"
    attempts: int = 1

    def effective_available_lanes(self, config: ProjectConfig) -> tuple[int, ...]:
        lanes = self.available_lanes or tuple(range(1, config.track_lanes + 1))
        errors: list[str] = []
        if len(set(lanes)) != len(lanes):
            errors.append("可用道次不能重复")
        if any(lane < 1 or lane > config.track_lanes for lane in lanes):
            errors.append("可用道次必须位于物理跑道范围内")
        if not lanes:
            errors.append("至少需要一条可用跑道")
        if errors:
            raise ConfigurationError(errors)
        return tuple(sorted(lanes))

    @property
    def lanes_available(self) -> int:
        return len(self.available_lanes)


@dataclass(frozen=True, slots=True)
class HeatAssignment:
    id: str
    event_round_id: str
    participant_id: str
    heat_no: int
    lane: int | None = None
    order: int | None = None
    random_seed: int | None = None
    manually_adjusted: bool = False


@dataclass(frozen=True, slots=True)
class Result:
    participant_id: str
    heat_no: int
    raw_value: str
    performance_kind: PerformanceKind
    canonical_value: int | None
    canonical_unit: str | None
    standard_display: str
    status: ResultStatus = ResultStatus.VALID
    manual_rank: int | None = None


@dataclass(frozen=True, slots=True)
class Qualification:
    participant_id: str
    heat_no: int
    heat_rank: int
    overall_rank: int
    canonical_value: int
    reason: str


@dataclass(frozen=True, slots=True)
class ValidationIssue:
    code: str
    severity: Severity
    message: str
    entity_type: str = ""
    entity_id: str = ""
    source_file_id: str = ""
    source_sheet: str = ""
    source_row: int | None = None
    original_value: Any = None
    normalized_value: Any = None


@dataclass(slots=True)
class ValidationReport:
    issues: list[ValidationIssue] = field(default_factory=list)
    source_files: dict[str, str] = field(default_factory=dict)

    def add(self, issue: ValidationIssue) -> None:
        if "_" not in issue.code:
            raise ValueError("校验代码必须包含命名空间，例如 IMPORT_001")
        self.issues.append(issue)

    @property
    def has_blocking(self) -> bool:
        return any(issue.severity is Severity.BLOCKING for issue in self.issues)

    def by_severity(self, severity: Severity) -> list[ValidationIssue]:
        return [issue for issue in self.issues if issue.severity is severity]


@dataclass(frozen=True, slots=True)
class AuditLog:
    entity_type: str
    entity_id: str
    event_round_id: str | None
    action_type: str
    before_json: str | None
    after_json: str | None
    rules_version: str
    random_seed: int | None
    created_at: str
