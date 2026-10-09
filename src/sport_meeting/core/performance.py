from __future__ import annotations

import re
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP

from .models import PerformanceKind, Result, ResultStatus


_STATUS_ALIASES = {
    "DNS": ResultStatus.DNS,
    "未出赛": ResultStatus.DNS,
    "弃权": ResultStatus.DNS,
    "DNF": ResultStatus.DNF,
    "未完成": ResultStatus.DNF,
    "DQ": ResultStatus.DQ,
    "犯规": ResultStatus.DQ,
    "取消资格": ResultStatus.DQ,
    "NM": ResultStatus.NM,
    "无成绩": ResultStatus.NM,
}


class PerformanceParseError(ValueError):
    pass


def recommended_input_unit(event_name: str, kind: PerformanceKind) -> str:
    if kind is PerformanceKind.COUNT:
        return "COUNT"
    if kind is PerformanceKind.DISTANCE:
        return "METER"
    if kind is PerformanceKind.TIME:
        distance_match = re.search(r"(\d+)\s*米", event_name)
        if distance_match and int(distance_match.group(1)) >= 800:
            return "MINUTE"
        return "SECOND"
    return "SECOND"


def _normalize_text(value: object) -> str:
    text = str(value).strip()
    return text.translate(
        str.maketrans(
            {
                "０": "0", "１": "1", "２": "2", "３": "3", "４": "4",
                "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
                "：": ":", "．": ".", "，": ",", "′": "'", "＇": "'",
                "″": '"', "＂": '"',
            }
        )
    )


def _decimal(text: str) -> Decimal:
    try:
        return Decimal(text)
    except InvalidOperation as exc:
        raise PerformanceParseError(f"无法识别成绩：{text}") from exc


def _to_int(value: Decimal) -> int:
    return int(value.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


def _parse_time_ms(text: str, input_unit: str | None = None) -> int:
    clean = text.lower().replace("秒", "").replace("s", "").strip()
    seconds_part = re.fullmatch(r"(\d+)\s*:\s*(\d{1,3})", clean)
    if input_unit == "SECOND" and seconds_part:
        fraction_text = seconds_part.group(2)
        seconds = Decimal(seconds_part.group(1))
        fraction = Decimal(fraction_text) / (Decimal(10) ** len(fraction_text))
        return _to_int((seconds + fraction) * 1000)
    three_part = re.fullmatch(r"(\d+)\s*:\s*(\d{1,2})\s*:\s*(\d{1,3})", clean)
    if three_part:
        minutes = Decimal(three_part.group(1))
        seconds = Decimal(three_part.group(2))
        fraction_text = three_part.group(3)
        if seconds >= 60:
            raise PerformanceParseError("分钟格式中的秒数必须小于 60")
        fraction = Decimal(fraction_text) / (Decimal(10) ** len(fraction_text))
        return _to_int((minutes * 60 + seconds + fraction) * 1000)
    if ":" in clean:
        minutes, seconds = clean.split(":", 1)
        if _decimal(seconds) >= 60:
            raise PerformanceParseError("分钟格式中的秒数必须小于 60")
        return _to_int((_decimal(minutes) * 60 + _decimal(seconds)) * 1000)
    match = re.fullmatch(r"(\d+)\s*(?:'|分)\s*(\d+)(?:\s*\"\s*(\d{1,3}))?", clean)
    if match:
        fraction = Decimal(0)
        if match.group(3):
            fraction = Decimal(match.group(3)) / (Decimal(10) ** len(match.group(3)))
        return _to_int((Decimal(match.group(1)) * 60 + Decimal(match.group(2)) + fraction) * 1000)
    match = re.fullmatch(r"(\d+)\s*\"\s*(\d{1,3})", clean)
    if match:
        fraction = Decimal(match.group(2)) / (Decimal(10) ** len(match.group(2)))
        return _to_int((Decimal(match.group(1)) + fraction) * 1000)
    return _to_int(_decimal(clean) * 1000)


def _parse_distance_mm(text: str, default_unit: str) -> int:
    clean = text.lower().replace(" ", "")
    match = re.fullmatch(r"([+-]?\d+(?:\.\d+)?)(毫米|mm|厘米|cm|米|m)?", clean)
    if not match:
        raise PerformanceParseError(f"无法识别距离：{text}")
    value = _decimal(match.group(1))
    unit = match.group(2) or default_unit.lower()
    factors = {"毫米": 1, "mm": 1, "厘米": 10, "cm": 10, "米": 1000, "m": 1000}
    if unit not in factors:
        raise PerformanceParseError(f"未知距离单位：{unit}")
    return _to_int(value * factors[unit])


def _parse_count(text: str) -> int:
    clean = re.sub(r"(?:个|次|下)$", "", text.strip())
    value = _decimal(clean)
    if value != value.to_integral_value():
        raise PerformanceParseError("计数成绩必须是整数")
    return int(value)


def format_canonical(kind: PerformanceKind, value: int | None) -> str:
    if value is None:
        return ""
    if kind is PerformanceKind.TIME:
        total = Decimal(value) / 1000
        if total >= 60:
            minutes = int(total // 60)
            seconds = total - minutes * 60
            return f"{minutes}:{seconds:05.2f}"
        return f"{total:.2f}"
    if kind is PerformanceKind.DISTANCE:
        return f"{Decimal(value) / 10:.1f} cm"
    return str(value)


def format_for_input_unit(kind: PerformanceKind, value: int | None, input_unit: str) -> str:
    if value is None:
        return ""
    if kind is PerformanceKind.TIME and input_unit == "MINUTE":
        total_centiseconds = (value + 5) // 10
        minutes, remainder = divmod(total_centiseconds, 6000)
        seconds, centiseconds = divmod(remainder, 100)
        return f"{minutes}:{seconds:02d}:{centiseconds:02d}"
    if kind is PerformanceKind.TIME:
        return f"{Decimal(value) / 1000:.2f} 秒"
    if kind is PerformanceKind.DISTANCE:
        return f"{Decimal(value) / 1000:.3f} 米"
    if kind is PerformanceKind.COUNT:
        return f"{value} 个"
    return str(value)


def parse_result(
    participant_id: str,
    heat_no: int,
    raw_value: object,
    kind: PerformanceKind,
    *,
    default_unit: str = "cm",
    display_unit: str | None = None,
) -> Result:
    text = _normalize_text(raw_value)
    status = _STATUS_ALIASES.get(text.upper()) or _STATUS_ALIASES.get(text)
    if status:
        return Result(participant_id, heat_no, text, kind, None, None, status.value, status)
    if not text:
        return Result(participant_id, heat_no, "", kind, None, None, "", ResultStatus.NM)

    if kind is PerformanceKind.TIME:
        value, unit = _parse_time_ms(text, display_unit), "MILLISECOND"
    elif kind is PerformanceKind.DISTANCE:
        value, unit = _parse_distance_mm(text, default_unit), "MILLIMETER"
    elif kind is PerformanceKind.COUNT:
        value, unit = _parse_count(text), "COUNT"
    else:
        rank = _parse_count(text)
        return Result(participant_id, heat_no, text, kind, None, None, str(rank), ResultStatus.VALID, rank)
    standard_display = (
        format_for_input_unit(kind, value, display_unit)
        if display_unit else format_canonical(kind, value)
    )
    return Result(participant_id, heat_no, text, kind, value, unit, standard_display)
