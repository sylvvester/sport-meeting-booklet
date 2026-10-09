from __future__ import annotations

import uuid
from dataclasses import dataclass

from .grouping import assign_final_heats
from .models import EventRound, EventType, HeatAssignment, ProjectConfig, Qualification, Result
from .qualification import qualify_track


_ID_NAMESPACE = uuid.UUID("8dd7f9d7-6df8-43d5-9736-8ae8df74dc90")


@dataclass(frozen=True, slots=True)
class FinalComputation:
    qualifications: list[Qualification]
    assignments: list[HeatAssignment]


def compute_track_final(
    preliminary_round: EventRound,
    final_round: EventRound,
    results: list[Result],
    config: ProjectConfig,
) -> FinalComputation:
    qualifications = qualify_track(results, config.finalists_quota)
    qualifications.sort(
        key=lambda item: (item.overall_rank, item.heat_no, item.participant_id)
    )
    if final_round.event_type is EventType.FIELD:
        assignments = _assign_field_final(final_round, qualifications)
    else:
        assignments = _assign_laned_final(final_round, qualifications, config)
    return FinalComputation(qualifications, assignments)


def _assign_laned_final(
    final_round: EventRound,
    qualifications: list[Qualification],
    config: ProjectConfig,
) -> list[HeatAssignment]:
    available_lanes = final_round.effective_available_lanes(config)
    final_heats = assign_final_heats(
        [item.participant_id for item in qualifications], config, available_lanes
    )
    assignments: list[HeatAssignment] = []
    for heat_no, heat in enumerate(final_heats, 1):
        for participant_id, lane in heat:
            assignment_id = uuid.uuid5(
                _ID_NAMESPACE, f"{final_round.id}:{participant_id}:{heat_no}:{lane}"
            ).hex
            assignments.append(
                HeatAssignment(
                    assignment_id, final_round.id, participant_id, heat_no, lane,
                    random_seed=None,
                )
            )
    return assignments


def _assign_field_final(
    final_round: EventRound,
    qualifications: list[Qualification],
) -> list[HeatAssignment]:
    assignments: list[HeatAssignment] = []
    for order, qualification in enumerate(reversed(qualifications), start=1):
        assignment_id = uuid.uuid5(
            _ID_NAMESPACE, f"{final_round.id}:{qualification.participant_id}:1:{order}"
        ).hex
        assignments.append(
            HeatAssignment(
                assignment_id,
                final_round.id,
                qualification.participant_id,
                1,
                lane=None,
                order=order,
                random_seed=None,
            )
        )
    return assignments
