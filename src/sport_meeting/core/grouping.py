from __future__ import annotations

import math
import random
from collections import Counter, defaultdict

from .models import ConfigurationError, FieldStartOrder, Participant, ProjectConfig


def balanced_heat_sizes(participant_count: int, heat_count: int) -> list[int]:
    if participant_count < 0 or heat_count < 1:
        raise ValueError("人数不能为负且组数必须大于 0")
    base, remainder = divmod(participant_count, heat_count)
    return [base + (index < remainder) for index in range(heat_count)]


def distribute_preliminary(
    participants: list[Participant], heat_count: int, seed: int
) -> list[list[Participant]]:
    """人数均衡并尽可能把相同单位分散到不同组。"""
    targets = balanced_heat_sizes(len(participants), heat_count)
    rng = random.Random(seed)
    by_unit: dict[str, list[Participant]] = defaultdict(list)
    for participant in participants:
        by_unit[participant.unit].append(participant)
    units = list(by_unit)
    rng.shuffle(units)
    for values in by_unit.values():
        rng.shuffle(values)

    heats: list[list[Participant]] = [[] for _ in range(heat_count)]
    unit_counts = [Counter() for _ in range(heat_count)]
    tie_order = list(range(heat_count))
    rng.shuffle(tie_order)
    tie_rank = {heat: rank for rank, heat in enumerate(tie_order)}

    for unit in units:
        for participant in by_unit[unit]:
            available = [i for i in range(heat_count) if len(heats[i]) < targets[i]]
            chosen = min(
                available,
                key=lambda i: (unit_counts[i][unit], len(heats[i]) / max(targets[i], 1), tie_rank[i]),
            )
            heats[chosen].append(participant)
            unit_counts[chosen][unit] += 1
    return heats


def assign_preliminary_lanes(
    heat_size: int,
    available_lanes: tuple[int, ...],
    preliminary_lane_start: int,
) -> list[int]:
    lanes = sorted(available_lanes)
    if heat_size > len(lanes):
        raise ConfigurationError(["预赛单组人数超过本场次可用跑道数"])
    if heat_size == len(lanes):
        return lanes
    ordered = [lane for lane in lanes if lane >= preliminary_lane_start]
    ordered += [lane for lane in lanes if lane < preliminary_lane_start]
    return sorted(ordered[:heat_size])


def effective_final_lane_priority(
    config: ProjectConfig, available_lanes: tuple[int, ...]
) -> list[int]:
    config.validate()
    lanes = set(available_lanes)
    if not lanes or not lanes.issubset(set(range(1, config.track_lanes + 1))):
        raise ConfigurationError(["决赛可用道次配置无效"])
    priority = [lane for lane in config.lane_priority if lane in lanes]
    if set(priority) != lanes or len(priority) != len(lanes):
        raise ConfigurationError(["决赛道次优先序列无法覆盖全部可用道次"])
    return priority


def assign_final_heats(
    ranked_participant_ids: list[str],
    config: ProjectConfig,
    available_lanes: tuple[int, ...],
) -> list[list[tuple[str, int]]]:
    priority = effective_final_lane_priority(config, available_lanes)
    capacity = len(priority)
    if len(ranked_participant_ids) > capacity * 2:
        raise ConfigurationError(["晋级人数超过两组决赛的可用容量，需要人工处理"])
    groups = [ranked_participant_ids]
    if len(ranked_participant_ids) > capacity:
        groups = [ranked_participant_ids[::2], ranked_participant_ids[1::2]]
    return [[(participant_id, priority[index]) for index, participant_id in enumerate(group)] for group in groups]


def order_field_participants(
    participants: list[Participant], strategy: FieldStartOrder, seed: int | None = None
) -> list[Participant]:
    ordered = list(participants)
    if strategy is FieldStartOrder.NUMBER_ASC:
        return sorted(ordered, key=lambda item: (item.bib, item.id))
    if seed is None:
        raise ValueError("抽签出场顺序必须提供随机种子")
    random.Random(seed).shuffle(ordered)
    return ordered

