from __future__ import annotations

import secrets
import uuid
from dataclasses import dataclass, field

from .grouping import assign_preliminary_lanes, distribute_preliminary, order_field_participants
from .models import (
    Athlete,
    Entry,
    EventRound,
    EventType,
    HeatAssignment,
    Participant,
    ProjectConfig,
    RoundType,
    Severity,
    ValidationIssue,
    ValidationReport,
)


_ID_NAMESPACE = uuid.UUID("07c0e418-8e1e-4530-9faf-e019df175c09")


@dataclass(slots=True)
class HeatGeneration:
    seed: int
    assignments: list[HeatAssignment] = field(default_factory=list)
    report: ValidationReport = field(default_factory=ValidationReport)


def generate_heat_assignments(
    config: ProjectConfig,
    athletes: list[Athlete],
    entries: list[Entry],
    rounds: list[EventRound],
    seed: int | None = None,
) -> HeatGeneration:
    """为赛前项目生成可复现的分组、道次或田赛出场顺序。"""
    config.validate()
    generation = HeatGeneration(seed if seed is not None else secrets.randbelow(2_000_000_000))
    athletes_by_id = {athlete.id: athlete for athlete in athletes}
    entries_by_event: dict[tuple[str, str], list[Entry]] = {}
    for entry in entries:
        entries_by_event.setdefault((entry.group_name, entry.event_name), []).append(entry)

    for round_ in rounds:
        if round_.round_type is RoundType.FINAL:
            continue
        if round_.event_type is EventType.TEAM:
            generation.report.add(
                ValidationIssue(
                    "HEAT_006", Severity.WARNING, "集体项目暂不自动生成人员分组",
                    "EVENT_ROUND", round_.id,
                )
            )
            continue
        round_entries = entries_by_event.get((round_.group_name, round_.event_name), [])
        participants = [
            Participant(athlete.id, athlete.bib, athlete.name, athlete.unit)
            for entry in round_entries
            if (athlete := athletes_by_id.get(entry.athlete_id)) is not None
        ]
        participants.sort(key=lambda item: (item.bib, item.id))
        if not participants:
            generation.report.add(
                ValidationIssue(
                    "HEAT_004", Severity.BLOCKING, "项目没有可编排的报名运动员",
                    "EVENT_ROUND", round_.id,
                )
            )
            continue
        if round_.heat_count < 1:
            generation.report.add(
                ValidationIssue(
                    "HEAT_005", Severity.BLOCKING, "项目组数尚未确认",
                    "EVENT_ROUND", round_.id,
                )
            )
            continue

        round_seed = generation.seed ^ int(round_.id[:8], 16)
        if round_.event_type is EventType.FIELD:
            ordered = order_field_participants(participants, config.field_start_order, round_seed)
            groups = _split_in_order(ordered, round_.heat_count)
            for heat_no, group in enumerate(groups, 1):
                for start_order, participant in enumerate(group, 1):
                    generation.assignments.append(
                        _assignment(round_, participant, heat_no, None, start_order, generation.seed)
                    )
            continue

        groups = distribute_preliminary(participants, round_.heat_count, round_seed)
        available_lanes = round_.effective_available_lanes(config)
        for heat_no, group in enumerate(groups, 1):
            if round_.round_type is RoundType.PRELIMINARY:
                lanes = assign_preliminary_lanes(
                    len(group), available_lanes, config.preliminary_lane_start
                )
            else:
                # 中长跑预决赛常为弧形起跑，序号可能超过物理跑道数。
                lanes = list(range(1, len(group) + 1))
            for participant, lane in zip(group, lanes, strict=True):
                generation.assignments.append(
                    _assignment(round_, participant, heat_no, lane, None, generation.seed)
                )
    return generation


def _split_in_order(participants: list[Participant], heat_count: int) -> list[list[Participant]]:
    base, remainder = divmod(len(participants), heat_count)
    sizes = [base + (index < remainder) for index in range(heat_count)]
    groups: list[list[Participant]] = []
    offset = 0
    for size in sizes:
        groups.append(participants[offset:offset + size])
        offset += size
    return groups


def _assignment(
    round_: EventRound,
    participant: Participant,
    heat_no: int,
    lane: int | None,
    order: int | None,
    seed: int,
) -> HeatAssignment:
    assignment_id = uuid.uuid5(
        _ID_NAMESPACE, f"{round_.id}:{participant.id}:{heat_no}:{lane}:{order}"
    ).hex
    return HeatAssignment(
        assignment_id, round_.id, participant.id, heat_no, lane, order, seed, False
    )
