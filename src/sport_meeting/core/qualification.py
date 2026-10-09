from __future__ import annotations

import math
from collections import defaultdict

from .models import PerformanceKind, Qualification, Result, ResultStatus


def competition_ranks(results: list[Result], *, descending: bool = False) -> dict[str, int]:
    valid = [
        result
        for result in results
        if result.status is ResultStatus.VALID and result.canonical_value is not None
    ]
    ordered = sorted(valid, key=lambda item: item.canonical_value, reverse=descending)
    ranks: dict[str, int] = {}
    previous: int | None = None
    current_rank = 0
    for position, result in enumerate(ordered, start=1):
        if previous is None or result.canonical_value != previous:
            current_rank = position
            previous = result.canonical_value
        ranks[result.participant_id] = current_rank
    return ranks


def rank_final_results(
    results: list[Result], *, maximum_rank: int | None = None
) -> list[tuple[int | None, Result]]:
    """合并所有决赛组排名；无有效成绩者置于有效成绩之后且不授予名次。"""
    performance_kinds = {
        result.performance_kind
        for result in results
        if result.status is ResultStatus.VALID and result.canonical_value is not None
    }
    if len(performance_kinds) > 1:
        raise ValueError("同一决赛不能混用不同类型的成绩单位")
    performance_kind = next(iter(performance_kinds), PerformanceKind.TIME)
    descending = performance_kind in (PerformanceKind.DISTANCE, PerformanceKind.COUNT)
    ranks = competition_ranks(results, descending=descending)
    valid = [item for item in results if item.participant_id in ranks]
    invalid = [item for item in results if item.participant_id not in ranks]
    valid.sort(
        key=lambda item: item.canonical_value if not descending else -item.canonical_value
    )
    invalid.sort(key=lambda item: (item.heat_no, item.participant_id))
    ranked = [(ranks[item.participant_id], item) for item in valid] + [
        (None, item) for item in invalid
    ]
    if maximum_rank is not None:
        return [
            (rank, result)
            for rank, result in ranked
            if rank is not None and rank <= maximum_rank
        ]
    return ranked


def qualify_track(results: list[Result], finalists_quota: int) -> list[Qualification]:
    if finalists_quota < 1:
        raise ValueError("晋级人数必须大于 0")
    if any(result.manual_rank is not None for result in results):
        return _qualify_manual_ranks(results, finalists_quota)
    by_heat: dict[int, list[Result]] = defaultdict(list)
    performance_kinds = {
        result.performance_kind
        for result in results
        if result.status is ResultStatus.VALID and result.canonical_value is not None
    }
    if len(performance_kinds) > 1:
        raise ValueError("同一场次不能混用不同类型的成绩单位")
    performance_kind = next(iter(performance_kinds), PerformanceKind.TIME)
    descending = performance_kind in (PerformanceKind.DISTANCE, PerformanceKind.COUNT)
    for result in results:
        by_heat[result.heat_no]
        if (
            result.status is ResultStatus.VALID
            and result.canonical_value is not None
        ):
            by_heat[result.heat_no].append(result)
    if not by_heat:
        return []

    quota_per_heat = math.ceil(finalists_quota / len(by_heat))
    candidates: list[tuple[Result, int]] = []
    for heat_no in sorted(by_heat):
        heat_results = sorted(
            by_heat[heat_no], key=lambda item: item.canonical_value, reverse=descending
        )
        if not heat_results:
            continue
        ranks = competition_ranks(heat_results, descending=descending)
        candidates.extend(
            (result, ranks[result.participant_id])
            for result in heat_results
            if ranks[result.participant_id] <= quota_per_heat
        )

    candidates.sort(key=lambda item: item[0].canonical_value, reverse=descending)
    if len(candidates) > finalists_quota:
        cutoff = candidates[finalists_quota - 1][0].canonical_value
        candidates = [
            item for item in candidates
            if (
                item[0].canonical_value >= cutoff
                if descending else item[0].canonical_value <= cutoff
            )
        ]

    overall_ranks = competition_ranks([item[0] for item in candidates], descending=descending)
    return [
        Qualification(
            participant_id=result.participant_id,
            heat_no=result.heat_no,
            heat_rank=heat_rank,
            overall_rank=overall_ranks[result.participant_id],
            canonical_value=result.canonical_value,
            reason=f"第{result.heat_no}组第{heat_rank}名",
        )
        for result, heat_rank in candidates
    ]


def _qualify_manual_ranks(results: list[Result], finalists_quota: int) -> list[Qualification]:
    heat_numbers = sorted({result.heat_no for result in results})
    if not heat_numbers:
        return []
    quota_per_heat = math.ceil(finalists_quota / len(heat_numbers))
    candidates = sorted(
        (
            result for result in results
            if result.status is ResultStatus.VALID
            and result.manual_rank is not None
            and result.manual_rank <= quota_per_heat
        ),
        key=lambda item: (item.manual_rank, item.heat_no, item.participant_id),
    )
    if len(candidates) > finalists_quota:
        raise ValueError(
            f"手动组内名次产生 {len(candidates)} 名候选，但决赛名额只有 {finalists_quota}；"
            "请把不晋级者状态改为“无成绩”后再生成"
        )
    qualifications = []
    previous_rank = None
    overall_rank = 0
    for position, result in enumerate(candidates, start=1):
        if result.manual_rank != previous_rank:
            overall_rank = position
            previous_rank = result.manual_rank
        qualifications.append(
            Qualification(
                participant_id=result.participant_id,
                heat_no=result.heat_no,
                heat_rank=result.manual_rank,
                overall_rank=overall_rank,
                canonical_value=result.manual_rank,
                reason=f"手动指定：第{result.heat_no}组第{result.manual_rank}名",
            )
        )
    return qualifications
