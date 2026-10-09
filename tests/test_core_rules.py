import unittest

from sport_meeting.core.grouping import (
    assign_final_heats,
    assign_preliminary_lanes,
    distribute_preliminary,
    order_field_participants,
)
from sport_meeting.core.heat_generation import generate_heat_assignments
from sport_meeting.core.finals_workflow import compute_track_final
from sport_meeting.core.models import (
    Athlete,
    ConfigurationError,
    Entry,
    EventRound,
    EventType,
    FieldStartOrder,
    Participant,
    PerformanceKind,
    ProjectConfig,
    Result,
    ResultStatus,
    RoundType,
)
from sport_meeting.core.performance import (
    PerformanceParseError,
    format_for_input_unit,
    parse_result,
    recommended_input_unit,
)
from sport_meeting.core.qualification import competition_ranks, qualify_track, rank_final_results


def timed(participant_id: str, heat: int, milliseconds: int, status=ResultStatus.VALID):
    return Result(
        participant_id,
        heat,
        str(milliseconds),
        PerformanceKind.TIME,
        milliseconds if status is ResultStatus.VALID else None,
        "MILLISECOND" if status is ResultStatus.VALID else None,
        str(milliseconds),
        status,
    )


class ConfigTests(unittest.TestCase):
    def test_lane_priority_must_be_complete_permutation(self):
        config = ProjectConfig(lane_priority=(4, 5, 3, 6, 2, 7, 1, 1))
        with self.assertRaises(ConfigurationError):
            config.validate()


class PerformanceTests(unittest.TestCase):
    def test_time_uses_integer_milliseconds(self):
        self.assertEqual(parse_result("a", 1, "12.34秒", PerformanceKind.TIME).canonical_value, 12340)
        self.assertEqual(parse_result("a", 1, "1:05.32", PerformanceKind.TIME).canonical_value, 65320)
        self.assertEqual(parse_result("a", 1, "12″34", PerformanceKind.TIME).canonical_value, 12340)
        self.assertEqual(parse_result("a", 1, "1:53:29", PerformanceKind.TIME).canonical_value, 113290)
        self.assertEqual(
            parse_result(
                "a", 1, "59:32", PerformanceKind.TIME, display_unit="SECOND"
            ).canonical_value,
            59320,
        )
        self.assertLess(
            parse_result("a", 1, "1:53:29", PerformanceKind.TIME).canonical_value,
            parse_result("b", 1, "2:00:12", PerformanceKind.TIME).canonical_value,
        )
        with self.assertRaises(PerformanceParseError):
            parse_result("a", 1, "1:60:00", PerformanceKind.TIME)
        self.assertEqual(format_for_input_unit(PerformanceKind.TIME, 113290, "MINUTE"), "1:53:29")

    def test_recommended_units_follow_event_type(self):
        self.assertEqual(recommended_input_unit("50米", PerformanceKind.TIME), "SECOND")
        self.assertEqual(recommended_input_unit("800米", PerformanceKind.TIME), "MINUTE")
        self.assertEqual(recommended_input_unit("跳远", PerformanceKind.DISTANCE), "METER")
        self.assertEqual(recommended_input_unit("引体向上", PerformanceKind.COUNT), "COUNT")

    def test_distance_requires_configured_default_when_unit_missing(self):
        self.assertEqual(
            parse_result("a", 1, "244", PerformanceKind.DISTANCE, default_unit="cm").canonical_value,
            2440,
        )
        self.assertEqual(
            parse_result("a", 1, "2.44m", PerformanceKind.DISTANCE).canonical_value,
            2440,
        )

    def test_count_is_integer(self):
        self.assertEqual(parse_result("a", 1, "240个", PerformanceKind.COUNT).canonical_value, 240)
        with self.assertRaises(PerformanceParseError):
            parse_result("a", 1, "2.5", PerformanceKind.COUNT)


class GroupingTests(unittest.TestCase):
    def setUp(self):
        self.participants = [
            Participant(str(i), f"{10000+i}", f"运动员{i}", f"初一({i % 4 + 1})")
            for i in range(28)
        ]

    def test_grouping_is_reproducible_and_balanced(self):
        first = distribute_preliminary(self.participants, 4, 2025)
        second = distribute_preliminary(self.participants, 4, 2025)
        self.assertEqual(first, second)
        self.assertEqual([len(heat) for heat in first], [7, 7, 7, 7])
        for heat in first:
            units = [participant.unit for participant in heat]
            self.assertLessEqual(max(units.count(unit) for unit in set(units)), 2)

    def test_preliminary_lane_policy_is_independent(self):
        self.assertEqual(assign_preliminary_lanes(7, tuple(range(1, 9)), 2), list(range(2, 9)))
        self.assertEqual(assign_preliminary_lanes(6, (1, 2, 3, 5, 6, 7, 8), 2), [2, 3, 5, 6, 7, 8])

    def test_final_lane_priority_and_overflow_split(self):
        config = ProjectConfig()
        one_heat = assign_final_heats(list("abcdef"), config, tuple(range(1, 9)))
        self.assertEqual([lane for _, lane in one_heat[0]], [4, 5, 3, 6, 2, 7])
        two_heats = assign_final_heats([str(i) for i in range(1, 10)], config, tuple(range(1, 9)))
        self.assertEqual([p for p, _ in two_heats[0]], ["1", "3", "5", "7", "9"])
        self.assertEqual([p for p, _ in two_heats[1]], ["2", "4", "6", "8"])

    def test_field_order_defaults_to_bib(self):
        shuffled = list(reversed(self.participants[:5]))
        ordered = order_field_participants(shuffled, FieldStartOrder.NUMBER_ASC)
        self.assertEqual([p.bib for p in ordered], sorted(p.bib for p in shuffled))

    def test_full_preliminary_generation_is_reproducible(self):
        athletes = [
            Athlete(str(i), f"101{i + 1:02d}", f"运动员{i}", "男", "初一", str(i % 4 + 1))
            for i in range(17)
        ]
        entries = [Entry(f"e{i}", athlete.id, "50米", "初一男子组") for i, athlete in enumerate(athletes)]
        round_ = EventRound(
            "12345678abcdef00", "50米", "初一男子组", EventType.TRACK,
            RoundType.PRELIMINARY, 3, "2026-10-15 上午 9:30", 17, tuple(range(2, 9)),
        )
        first = generate_heat_assignments(ProjectConfig(), athletes, entries, [round_], seed=2026)
        second = generate_heat_assignments(ProjectConfig(), athletes, entries, [round_], seed=2026)
        self.assertEqual(first.assignments, second.assignments)
        self.assertEqual(len(first.assignments), 17)
        sizes = [sum(item.heat_no == heat for item in first.assignments) for heat in (1, 2, 3)]
        self.assertEqual(sizes, [6, 6, 5])
        self.assertNotIn(1, {item.lane for item in first.assignments})


class QualificationTests(unittest.TestCase):
    def test_three_heats_take_three_then_remove_slowest(self):
        results = []
        value = 10000
        for heat in range(1, 4):
            for lane in range(1, 6):
                results.append(timed(f"h{heat}p{lane}", heat, value))
                value += 10
        qualified = qualify_track(results, 8)
        self.assertEqual(len(qualified), 8)
        self.assertNotIn("h3p3", {item.participant_id for item in qualified})

    def test_boundary_tie_keeps_everyone(self):
        values = [100, 101, 102, 103, 104, 105, 106, 107, 107]
        results = [timed(str(index), (index % 3) + 1, value) for index, value in enumerate(values)]
        qualified = qualify_track(results, 8)
        self.assertEqual(len(qualified), 9)

    def test_invalid_results_are_not_candidates_and_not_replaced(self):
        results = [timed("a", 1, 100), timed("b", 1, 0, ResultStatus.DQ), timed("c", 2, 110)]
        qualified = qualify_track(results, 8)
        self.assertEqual({item.participant_id for item in qualified}, {"a", "c"})

    def test_heat_with_no_valid_result_still_counts_toward_per_heat_quota(self):
        results = [timed(f"a{i}", 1, 100 + i) for i in range(6)]
        results += [timed(f"b{i}", 2, 0, ResultStatus.DQ) for i in range(6)]
        results += [timed(f"c{i}", 3, 200 + i) for i in range(6)]
        qualified = qualify_track(results, 8)
        self.assertEqual(len(qualified), 6)
        self.assertEqual({item.heat_no for item in qualified}, {1, 3})

    def test_competition_ranking_skips_after_tie(self):
        ranks = competition_ranks([timed("a", 1, 100), timed("b", 1, 110), timed("c", 2, 110), timed("d", 2, 120)])
        self.assertEqual(ranks, {"a": 1, "b": 2, "c": 2, "d": 4})

    def test_final_ranking_merges_heats_keeps_ties_and_places_invalid_last(self):
        results = [
            timed("a", 1, 100),
            timed("b", 2, 110),
            timed("c", 1, 110),
            timed("d", 2, 0, ResultStatus.DQ),
        ]
        ranked = rank_final_results(results)
        self.assertEqual([(rank, result.participant_id) for rank, result in ranked], [
            (1, "a"), (2, "b"), (2, "c"), (None, "d")
        ])

    def test_final_ranking_report_limit_keeps_all_tied_at_eighth(self):
        results = [timed(str(index), 1, 100 + index) for index in range(10)]
        results[8] = timed("8", 1, 107)
        ranked = rank_final_results(results, maximum_rank=8)
        self.assertEqual(len(ranked), 9)
        self.assertEqual([rank for rank, _ in ranked][-2:], [8, 8])
        self.assertNotIn("9", {result.participant_id for _, result in ranked})

    def test_final_computation_assigns_best_performance_to_priority_lane(self):
        preliminary = EventRound(
            "abcdef01", "50米", "初一男子组", EventType.TRACK,
            RoundType.PRELIMINARY, 3, "9:30", 9, tuple(range(1, 9)),
        )
        final = EventRound(
            "abcdef02", "50米", "初一男子组", EventType.TRACK,
            RoundType.FINAL, 1, "11:30", 8, tuple(range(1, 9)),
        )
        results = [timed(str(index), index % 3 + 1, 7000 + index * 10) for index in range(9)]
        computed = compute_track_final(preliminary, final, results, ProjectConfig())
        self.assertEqual(len(computed.qualifications), 8)
        best_id = min(computed.qualifications, key=lambda item: item.canonical_value).participant_id
        lane_by_participant = {item.participant_id: item.lane for item in computed.assignments}
        self.assertEqual(lane_by_participant[best_id], 4)

    def test_distance_results_rank_larger_values_first(self):
        results = [
            Result("a", 1, "4.50", PerformanceKind.DISTANCE, 4500, "MILLIMETER", "4.50 m"),
            Result("b", 1, "4.20", PerformanceKind.DISTANCE, 4200, "MILLIMETER", "4.20 m"),
            Result("c", 2, "4.60", PerformanceKind.DISTANCE, 4600, "MILLIMETER", "4.60 m"),
            Result("d", 2, "4.10", PerformanceKind.DISTANCE, 4100, "MILLIMETER", "4.10 m"),
        ]
        qualified = qualify_track(results, 2)
        self.assertEqual({item.participant_id for item in qualified}, {"a", "c"})
        self.assertEqual(
            next(item.overall_rank for item in qualified if item.participant_id == "c"), 1
        )

    def test_manual_ranks_are_independent_per_heat(self):
        results = [
            Result(
                f"h{heat}p{rank}", heat, "", PerformanceKind.MANUAL_RANK,
                None, None, f"第{rank}名", ResultStatus.VALID, rank,
            )
            for heat in range(1, 5)
            for rank in range(1, 4)
        ]
        qualified = qualify_track(results, 8)
        self.assertEqual(len(qualified), 8)
        self.assertEqual(
            {(item.heat_no, item.heat_rank) for item in qualified},
            {(heat, rank) for heat in range(1, 5) for rank in (1, 2)},
        )

    def test_manual_rank_overflow_requires_human_decision(self):
        results = [
            Result(
                f"h{heat}p{rank}", heat, "", PerformanceKind.MANUAL_RANK,
                None, None, f"第{rank}名", ResultStatus.VALID, rank,
            )
            for heat in range(1, 4)
            for rank in range(1, 4)
        ]
        with self.assertRaisesRegex(ValueError, "决赛名额"):
            qualify_track(results, 8)

    def test_field_final_uses_reverse_preliminary_order(self):
        preliminary = EventRound(
            "field-prelim", "跳远", "初一女子组", EventType.FIELD,
            RoundType.PRELIMINARY, 1, "9:00", 3,
            performance_kind=PerformanceKind.DISTANCE,
        )
        final = EventRound(
            "field-final", "跳远", "初一女子组", EventType.FIELD,
            RoundType.FINAL, 1, "11:00", 3,
            performance_kind=PerformanceKind.DISTANCE,
        )
        results = [
            Result("a", 1, "4.10", PerformanceKind.DISTANCE, 4100, "MILLIMETER", "4.100 米"),
            Result("b", 1, "4.30", PerformanceKind.DISTANCE, 4300, "MILLIMETER", "4.300 米"),
            Result("c", 1, "4.20", PerformanceKind.DISTANCE, 4200, "MILLIMETER", "4.200 米"),
        ]
        computed = compute_track_final(preliminary, final, results, ProjectConfig(finalists_quota=3))
        self.assertTrue(all(item.lane is None for item in computed.assignments))
        self.assertEqual([item.participant_id for item in computed.assignments], ["a", "c", "b"])
        self.assertEqual([item.order for item in computed.assignments], [1, 2, 3])


if __name__ == "__main__":
    unittest.main()
