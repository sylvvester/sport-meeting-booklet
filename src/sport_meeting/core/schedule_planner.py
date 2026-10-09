from __future__ import annotations

import math
import re
from collections import defaultdict
from dataclasses import dataclass, replace
from datetime import date, datetime, time, timedelta
from typing import Iterable

from .importers import normalize_event_name
from .models import (
    ROUND_MODE_PRELIMINARY_FINAL,
    ROUND_MODE_TIMED_FINAL,
    TEAM_COUNT_ONE_PER_CLASS,
    TEAM_GROUP_MEN,
    TEAM_GROUP_MIXED,
    TEAM_GROUP_STAFF,
    TEAM_GROUP_WOMEN,
    TEAM_ROUND_MODE_KNOCKOUT,
    Athlete,
    Entry,
    EventType,
    ProjectConfig,
    RoundType,
)


@dataclass(frozen=True, slots=True)
class SchedulePlanRow:
    competition_date: date
    period: str
    event_type: EventType
    sequence: int
    event_title: str
    group_name: str
    event_name: str
    declared_count: int | None
    heat_count: int | None
    start_time: time
    available_lanes: tuple[int, ...] = ()


class ScheduleCapacityError(ValueError):
    def __init__(self, missing_slots: int):
        self.missing_slots = missing_slots
        super().__init__(
            f"所选运动会日期无法容纳全部项目，还差 {missing_slots} 个比赛时段。"
            "请增加运动会日期后重新生成。"
        )


@dataclass(frozen=True, slots=True)
class _Session:
    competition_date: date
    period: str
    start: datetime
    end: datetime


@dataclass(frozen=True, slots=True)
class _Task:
    event_type: EventType
    event_name: str
    group_name: str
    round_type: RoundType
    declared_count: int | None
    heat_count: int | None
    preferred_session: int
    resource: str
    duration_minutes: int
    order: int
    title_suffix: str | None = None
    not_before: time | None = None
    required_period: str | None = None


_JUNIOR_GRADES = ("初一", "初二", "初三")
_SENIOR_GRADES = ("高一", "高二", "高三")
_SHORT_SPRINTS = {50, 100, 200}


def build_standard_schedule(
    config: ProjectConfig,
    athletes: Iterable[Athlete] = (),
    entries: Iterable[Entry] = (),
) -> list[SchedulePlanRow]:
    """根据项目设置和已导入报名数据生成可编辑的初始日程。

    只展开报名表中实际出现的年级；报名人数用于计算组数和时长。
    """
    start_date, end_date = _meeting_dates(config)
    sessions = _sessions(start_date, end_date)
    athlete_map = {athlete.id: athlete for athlete in athletes}
    active_grades = {athlete.grade for athlete in athlete_map.values()}
    if not active_grades:
        raise ValueError("请先批量导入报名表，系统将按实际参赛年级生成标准日程。")
    entry_counts: dict[tuple[str, str, str], int] = defaultdict(int)
    classes_by_grade: dict[str, set[str]] = defaultdict(set)
    for athlete in athlete_map.values():
        classes_by_grade[athlete.grade].add(athlete.class_name)
    for entry in entries:
        athlete = athlete_map.get(entry.athlete_id)
        if athlete is None:
            continue
        event_name, _ = normalize_event_name(entry.event_name)
        entry_counts[(athlete.grade, athlete.sex, event_name)] += 1

    tasks = _build_tasks(config, entry_counts, classes_by_grade, active_grades)
    cursors: dict[tuple[int, str], datetime] = {}
    preliminary_placements: dict[tuple[str, str], tuple[int, datetime]] = {}
    placed: list[tuple[_Task, int, datetime]] = []
    missing = 0
    for task in sorted(
        tasks,
        key=lambda item: (
            item.preferred_session,
            _session_phase(item),
            0 if item.event_type is EventType.TRACK else 1 if item.event_type is EventType.FIELD else 2,
            _session_event_order_key(item),
            _group_order_key(item.group_name),
            item.order,
        ),
    ):
        earliest_session = task.preferred_session
        earliest_start: datetime | None = None
        if task.round_type is RoundType.FINAL:
            preliminary = preliminary_placements.get((task.group_name, task.event_name))
            if preliminary is not None:
                preliminary_session, preliminary_end = preliminary
                distance = _distance(task.event_name)
                if distance == 400:
                    earliest_session = max(earliest_session, preliminary_session + 1)
                elif distance in _SHORT_SPRINTS:
                    earliest_session = max(earliest_session, preliminary_session)
                    earliest_start = preliminary_end + timedelta(minutes=45)
        placement = _place_task(
            task,
            sessions,
            cursors,
            earliest_session=earliest_session,
            earliest_start=earliest_start,
        )
        if placement is None:
            missing += 1
            continue
        session_index, start = placement
        placed.append((task, session_index, start))
        if task.round_type is RoundType.PRELIMINARY:
            preliminary_placements[(task.group_name, task.event_name)] = (
                session_index,
                start + timedelta(minutes=task.duration_minutes),
            )

    if missing:
        raise ScheduleCapacityError(missing)

    rows: list[SchedulePlanRow] = []
    for task, session_index, start in placed:
        session = sessions[session_index]
        suffix = task.title_suffix
        if suffix is None:
            suffix = {
                RoundType.PRELIMINARY: "预赛",
                RoundType.FINAL: "决赛",
                RoundType.TIMED_FINAL: "预决赛",
            }[task.round_type]
        title_group = _title_group(task.group_name, task.event_name)
        rows.append(
            SchedulePlanRow(
                session.competition_date,
                session.period,
                task.event_type,
                0,
                f"{title_group}{task.event_name}{suffix}",
                task.group_name,
                task.event_name,
                task.declared_count,
                task.heat_count,
                start.time(),
                tuple(range(1, config.track_lanes + 1))
                if task.event_type is EventType.TRACK or "接力" in task.event_name
                else (),
            )
        )

    rows.sort(
        key=lambda item: (
            item.competition_date,
            0 if item.period == "上午" else 1,
            0 if item.event_type is EventType.TRACK else 1 if item.event_type is EventType.FIELD else 2,
            item.start_time,
            item.event_title,
        )
    )
    counters: dict[tuple[date, str, EventType], int] = defaultdict(int)
    numbered: list[SchedulePlanRow] = []
    for row in rows:
        key = (row.competition_date, row.period, row.event_type)
        counters[key] += 1
        numbered.append(replace(row, sequence=counters[key]))
    return numbered


def _meeting_dates(config: ProjectConfig) -> tuple[date, date]:
    if not config.meeting_start_date or not config.meeting_end_date:
        raise ValueError("请先在项目首页设置并保存运动会开始和结束日期。")
    return date.fromisoformat(config.meeting_start_date), date.fromisoformat(config.meeting_end_date)


def _sessions(start_date: date, end_date: date) -> list[_Session]:
    result: list[_Session] = []
    current = start_date
    day_index = 0
    while current <= end_date:
        morning_start = time(9, 30) if day_index == 0 else time(8, 0)
        result.extend(
            (
                _Session(
                    current, "上午", datetime.combine(current, morning_start),
                    datetime.combine(current, time(12, 0)),
                ),
                _Session(
                    current, "下午", datetime.combine(current, time(13, 30)),
                    datetime.combine(current, time(17, 30)),
                ),
            )
        )
        current += timedelta(days=1)
        day_index += 1
    return result


def _build_tasks(
    config: ProjectConfig,
    entry_counts: dict[tuple[str, str, str], int],
    classes_by_grade: dict[str, set[str]],
    active_grades: set[str],
) -> list[_Task]:
    task_specs: list[tuple[str, str, str, str]] = []
    for grades, sex, events, formats in (
        (_JUNIOR_GRADES, "男", config.junior_boys_events, config.junior_boys_event_formats),
        (_JUNIOR_GRADES, "女", config.junior_girls_events, config.junior_girls_event_formats),
        (_SENIOR_GRADES, "男", config.senior_boys_events, config.senior_boys_event_formats),
        (_SENIOR_GRADES, "女", config.senior_girls_events, config.senior_girls_event_formats),
    ):
        for grade in grades:
            if grade not in active_grades:
                continue
            for raw_event in events:
                event_name, _ = normalize_event_name(raw_event)
                round_mode = formats.get(raw_event, formats.get(event_name, guess_event_round_mode(event_name)))
                task_specs.append((grade, sex, event_name, round_mode))

    # 不分性别的集体项目只生成一次，避免男、女项目列中同时配置后重复。
    seen_grade_team_events: set[tuple[str, str]] = set()
    tasks: list[_Task] = []
    order = 0
    for grade, sex, event_name, round_mode in task_specs:
        event_type = _event_type(event_name)
        if event_type is EventType.TEAM and "接力" not in event_name:
            team_key = (grade, event_name)
            if team_key in seen_grade_team_events:
                continue
            seen_grade_team_events.add(team_key)
            group_name = f"{grade}年级组"
            count = len(classes_by_grade.get(grade, set())) or None
        else:
            group_name = f"{grade}{'男子' if sex == '男' else '女子'}组"
            if event_type is EventType.TEAM:
                count = len(classes_by_grade.get(grade, set())) or None
            else:
                count = entry_counts.get((grade, sex, event_name)) or None

        rounds = _rounds_for(round_mode)
        for round_type in rounds:
            order += 1
            heat_count = _heat_count(event_name, event_type, round_type, count, config)
            tasks.append(
                _Task(
                    event_type,
                    event_name,
                    group_name,
                    round_type,
                    config.finalists_quota if round_type is RoundType.FINAL and count else count,
                    1 if round_type is RoundType.FINAL and count else heat_count,
                    _preferred_session(event_name, event_type, round_type),
                    _resource(event_name, event_type, group_name),
                    _duration_minutes(event_name, event_type, round_type, count, heat_count),
                    order,
                    not_before=_not_before(event_name),
                    required_period=_required_period(event_name, round_type),
                )
            )

    existing = {
        (task.event_name, task.group_name, task.round_type)
        for task in tasks
    }
    grade_order = (*_JUNIOR_GRADES, *_SENIOR_GRADES)
    for team_event in config.team_events:
        event_name, _ = normalize_event_name(team_event.event_name)
        if team_event.group_mode == TEAM_GROUP_STAFF:
            group_specs = (("教工组", team_event.manual_team_count),)
        else:
            group_specs_list: list[tuple[str, int | None]] = []
            for grade in grade_order:
                if grade not in active_grades:
                    continue
                group_name = {
                    TEAM_GROUP_MEN: f"{grade}男子组",
                    TEAM_GROUP_WOMEN: f"{grade}女子组",
                    TEAM_GROUP_MIXED: f"{grade}年级组",
                }[team_event.group_mode]
                count = (
                    len(classes_by_grade.get(grade, set())) or None
                    if team_event.count_mode == TEAM_COUNT_ONE_PER_CLASS
                    else team_event.manual_team_count
                )
                group_specs_list.append((group_name, count))
            group_specs = tuple(group_specs_list)

        for group_name, count in group_specs:
            for round_type in _rounds_for(team_event.round_mode):
                key = (event_name, group_name, round_type)
                if key in existing:
                    continue
                existing.add(key)
                order += 1
                onsite_count = (
                    2
                    if "拔河" in event_name
                    and team_event.count_mode == TEAM_COUNT_ONE_PER_CLASS
                    and count
                    else count
                )
                heat_count = _heat_count(
                    event_name, EventType.TEAM, round_type, onsite_count, config
                )
                tasks.append(
                    _Task(
                        EventType.TEAM,
                        event_name,
                        group_name,
                        round_type,
                        config.finalists_quota
                        if round_type is RoundType.FINAL and onsite_count
                        else onsite_count,
                        1 if round_type is RoundType.FINAL and onsite_count else heat_count,
                        _preferred_session(event_name, EventType.TEAM, round_type),
                        _resource(event_name, EventType.TEAM, group_name),
                        _duration_minutes(
                            event_name, EventType.TEAM, round_type, onsite_count, heat_count
                        ),
                        order,
                        title_suffix=(
                            "（冠亚军赛）"
                            if "拔河" in event_name
                            and team_event.count_mode == TEAM_COUNT_ONE_PER_CLASS
                            and onsite_count
                            else "（淘汰赛）"
                            if team_event.round_mode == TEAM_ROUND_MODE_KNOCKOUT
                            else None
                        ),
                        not_before=_not_before(event_name),
                        required_period=_required_period(event_name, round_type),
                    )
                )
    return tasks


def _event_type(event_name: str) -> EventType:
    if any(word in event_name for word in ("接力", "集体", "拔河", "袋鼠跳")):
        return EventType.TEAM
    if any(word in event_name for word in ("跳远", "跳高", "跳绳", "仰卧起坐", "引体向上", "铅球", "实心球")):
        return EventType.FIELD
    return EventType.TRACK


def _event_order_key(event_name: str, event_type: EventType) -> tuple[int, str]:
    if event_type is EventType.TRACK or "接力" in event_name:
        return _distance(event_name) or 9999, event_name
    preferred = (
        "立定跳远", "1分钟跳绳", "跳远", "跳高",
        "1分钟仰卧起坐", "仰卧起坐", "引体向上", "铅球", "实心球",
    )
    for index, keyword in enumerate(preferred):
        if keyword in event_name:
            return index, event_name
    return len(preferred), event_name


def _session_event_order_key(task: _Task) -> tuple[int, int, str]:
    """第二天上午先跑耐力项目，再安排从首日上午移出的 200 米预赛。"""
    distance = _distance(task.event_name)
    if task.preferred_session == 2 and task.event_type is EventType.TRACK:
        if distance and distance >= 800:
            return 0, distance, task.event_name
        if distance == 200:
            return 1, distance, task.event_name
    order, name = _event_order_key(task.event_name, task.event_type)
    return 0, order, name


def _session_phase(task: _Task) -> int:
    """保留常规预赛优先，同时让第二天耐力跑先于迁入的 200 米预赛。"""
    distance = _distance(task.event_name)
    if task.preferred_session == 2 and task.event_type is EventType.TRACK:
        if distance and distance >= 800:
            return 0
        if distance == 200:
            return 1
    return 0 if task.round_type is RoundType.PRELIMINARY else 1


def _group_order_key(group_name: str) -> tuple[int, int, int]:
    grades = ("初一", "初二", "初三", "高一", "高二", "高三")
    grade_index = next((index for index, grade in enumerate(grades) if group_name.startswith(grade)), 99)
    level = 0 if grade_index < 3 else 1
    within_level = grade_index if level == 0 else grade_index - 3
    sex = 0 if "男子" in group_name else 1 if "女子" in group_name else 2
    return level, sex, within_level


def _distance(event_name: str) -> int | None:
    match = re.search(r"(?<!×)(\d+)\s*米", event_name)
    return int(match.group(1)) if match else None


def guess_event_round_mode(event_name: str) -> str:
    """按学校历届赛程习惯预判赛制；用户保存的选择始终优先。"""
    normalized, _ = normalize_event_name(event_name)
    distance = _distance(normalized)
    if _event_type(normalized) is EventType.TRACK and distance in _SHORT_SPRINTS:
        return ROUND_MODE_PRELIMINARY_FINAL
    return ROUND_MODE_TIMED_FINAL


def _rounds_for(round_mode: str) -> tuple[RoundType, ...]:
    if round_mode == ROUND_MODE_PRELIMINARY_FINAL:
        return RoundType.PRELIMINARY, RoundType.FINAL
    return (RoundType.TIMED_FINAL,)


def _heat_count(
    event_name: str,
    event_type: EventType,
    round_type: RoundType,
    count: int | None,
    config: ProjectConfig,
) -> int | None:
    if count is None:
        return None
    if round_type is RoundType.FINAL or event_type is EventType.FIELD:
        return 1
    if event_type is EventType.TEAM:
        if "集体跳绳" in event_name:
            return count
        if "接力" in event_name:
            return max(1, math.ceil(count / config.track_lanes))
        return 1
    distance = _distance(event_name)
    if distance in _SHORT_SPRINTS:
        return max(1, math.ceil(count / config.track_lanes))
    if distance == 400:
        return max(1, math.ceil(count / 16))
    if distance and distance >= 800:
        return 1
    return 1


def _preferred_session(event_name: str, event_type: EventType, round_type: RoundType) -> int:
    distance = _distance(event_name)
    if "接力" in event_name:
        return 3
    if "拔河" in event_name:
        return 2
    if event_type is EventType.TEAM:
        return 1
    if event_type is EventType.FIELD:
        return 1 if any(word in event_name for word in ("仰卧起坐", "引体向上", "跳远")) and "立定" not in event_name else 0
    if distance == 200:
        return 3 if round_type is RoundType.FINAL else 2
    if distance and distance >= 800:
        return 2
    if distance == 400:
        return 1 if round_type is RoundType.FINAL else 0
    return 0


def _resource(event_name: str, event_type: EventType, group_name: str = "") -> str:
    if event_type is EventType.TRACK or "接力" in event_name:
        return "TRACK"
    if "拔河" in event_name:
        if group_name.startswith("初"):
            return "TUG_JUNIOR"
        if group_name.startswith("高"):
            return "TUG_SENIOR"
        return "TEAM_FIELD"
    if "立定跳远" in event_name:
        return "STANDING_JUMP"
    if any(word in event_name for word in ("跳远", "跳高")):
        return "JUMP_FIELD"
    if "跳绳" in event_name:
        return "ROPE_FIELD"
    if "仰卧起坐" in event_name:
        return "SITUP_FIELD"
    if "引体向上" in event_name:
        return "PULLUP_FIELD"
    if any(word in event_name for word in ("铅球", "实心球")):
        return "THROW_FIELD"
    return "TEAM_FIELD"


def _not_before(event_name: str) -> time | None:
    return time(9, 30) if "拔河" in event_name else None


def _required_period(event_name: str, round_type: RoundType) -> str | None:
    if _distance(event_name) != 400:
        return None
    if round_type is RoundType.PRELIMINARY:
        return "上午"
    if round_type is RoundType.FINAL:
        return "下午"
    return None


def _duration_minutes(
    event_name: str,
    event_type: EventType,
    round_type: RoundType,
    count: int | None,
    heat_count: int | None,
) -> int:
    estimated_count = count or 24
    distance = _distance(event_name)
    estimated_heats = heat_count or (1 if round_type is RoundType.FINAL else 2 if distance == 400 else 3)
    if event_type is EventType.TRACK:
        if round_type is RoundType.FINAL:
            return 5
        if distance in (50, 100):
            return _round_up(max(10, estimated_heats * 3), 5)
        if distance == 200:
            return _round_up(max(10, estimated_heats * 5), 5)
        if distance == 400:
            return 5 if estimated_heats == 1 else _round_up(estimated_heats * 7, 5)
        if distance and distance >= 1500:
            return 15
        if distance and distance >= 800:
            return 10
        return 10
    if "接力" in event_name:
        return _round_up(max(10, estimated_heats * 7), 5)
    if "拔河" in event_name:
        return 30
    if "集体跳绳" in event_name:
        return _round_up(max(20, estimated_count * 3), 5)
    if "1分钟跳绳" in event_name:
        return _round_up(max(15, estimated_count * 0.55), 5)
    if any(word in event_name for word in ("仰卧起坐", "引体向上")):
        return _round_up(max(25, estimated_count * 1.5), 5)
    if "立定跳远" in event_name:
        return _round_up(max(30, estimated_count * 1.1), 5)
    if any(word in event_name for word in ("跳远", "跳高", "铅球", "实心球")):
        return _round_up(max(35, estimated_count * 2), 5)
    return 20


def _round_up(value: float, step: int) -> int:
    return int(math.ceil(value / step) * step)


def _place_task(
    task: _Task,
    sessions: list[_Session],
    cursors: dict[tuple[int, str], datetime],
    *,
    earliest_session: int | None = None,
    earliest_start: datetime | None = None,
) -> tuple[int, datetime] | None:
    first_session = min(
        task.preferred_session if earliest_session is None else earliest_session,
        len(sessions) - 1,
    )
    for session_index in range(first_session, len(sessions)):
        session = sessions[session_index]
        if task.required_period is not None and session.period != task.required_period:
            continue
        key = (session_index, task.resource)
        start = cursors.get(key, session.start)
        if earliest_start is not None:
            start = max(start, earliest_start)
        if task.not_before is not None and session.period == "上午":
            start = max(start, datetime.combine(session.competition_date, task.not_before))
        while start < session.end:
            end = start + timedelta(minutes=task.duration_minutes)
            if end > session.end:
                break
            cursors[key] = end
            return session_index, start
    return None


def _title_group(group_name: str, event_name: str) -> str:
    if "接力" in event_name:
        return group_name.replace("男子组", "年级男子").replace("女子组", "年级女子")
    if group_name.endswith("年级组"):
        return group_name.removesuffix("组")
    return group_name
