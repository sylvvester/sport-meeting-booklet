from __future__ import annotations

from pathlib import Path
from typing import TYPE_CHECKING, Iterable

from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from .models import EventType

if TYPE_CHECKING:
    from .schedule_planner import SchedulePlanRow


SCHEDULE_HEADERS = (
    "日期", "时间段", "比赛类型", "序号", "项目名称", "组别", "项目",
    "总人数/队数", "总组数", "比赛时间", "可用道次",
)

REGISTRATION_HEADERS = (
    "号码", "姓名", "性别", "年级", "班级", "项目1", "项目2",
)

REGISTRATION_EVENTS = (
    "50米", "100米", "200米", "400米", "800米", "1000米", "1500米",
    "立定跳远", "跳远", "跳高", "铅球", "实心球", "引体向上",
    "1分钟仰卧起坐", "1分钟跳绳", "4×100米接力",
)


def create_school_level_registration_template(
    path: str | Path,
    school_level: str,
    boys_events: tuple[str, ...],
    girls_events: tuple[str, ...],
) -> Path:
    """按学段生成男、女生分表的报名模板。"""
    if school_level not in {"初中", "高中"}:
        raise ValueError("学段必须是“初中”或“高中”")
    event_sets = {"男生": tuple(boys_events), "女生": tuple(girls_events)}
    if any(not events for events in event_sets.values()):
        raise ValueError("男生和女生都至少需要配置一个报名项目")

    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    workbook.remove(workbook.active)
    list_sheet = workbook.create_sheet("_下拉数据")
    grades = ("初一", "初二", "初三") if school_level == "初中" else ("高一", "高二", "高三")
    header_fill = PatternFill("solid", fgColor="1F4E78")
    section_fill = PatternFill("solid", fgColor="D9EAF7")
    thin = Side(style="thin", color="9CA3AF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for column, (sex_label, events) in enumerate(event_sets.items(), start=1):
        list_sheet.cell(1, column, f"{school_level}{sex_label}项目")
        for row, event in enumerate(events, start=2):
            list_sheet.cell(row, column, event)
    for row, grade in enumerate(grades, start=2):
        list_sheet.cell(row, 3, grade)
    for row in range(1, 31):
        list_sheet.cell(row + 1, 4, row)

    for event_column, (sex_label, events) in enumerate(event_sets.items(), start=1):
        sheet = workbook.create_sheet(sex_label)
        sheet.merge_cells("A1:G1")
        sheet["A1"] = f"示例学校田径运动会报名表（{school_level}年级用）"
        sheet["A1"].font = Font(size=16, bold=True)
        sheet["A1"].alignment = Alignment(horizontal="center", vertical="center")
        sheet.row_dimensions[1].height = 30
        sheet.merge_cells("A2:G2")
        sheet["A2"] = (
            "填表说明：\n"
            "      1、运动员号码无需填写,姓名输入,其它“性别\\年级\\班级\\项目1\\项目2”在单元格内选择输入。\n"
            "      2、各班每项限报两人（分男、女组），每班每项不得少于一人，可兼接力。\n"
            "      3、全部报完名后，请保存XX班级，如高二年级1班。\n"
            "      4、全校各班均以电子报名表上交体育组，报名截止时间9月11日晚上12点,交给本年级组体育老师。\n"
            "      5、集体项目班级自行安排参赛队员。（4x100米接力、集体跳绳、拔河比赛）"
        )
        sheet["A2"].alignment = Alignment(wrap_text=True, vertical="center")
        sheet.row_dimensions[2].height = 105
        sheet.merge_cells("A3:G3")
        sheet["A3"] = sex_label
        sheet["A3"].font = Font(bold=True, size=13)
        sheet["A3"].fill = section_fill
        sheet["A3"].alignment = Alignment(horizontal="center")
        sheet.merge_cells("A4:G4")
        sheet["A4"] = f"{sex_label}项目：{'、'.join(events)}"
        sheet["A4"].alignment = Alignment(wrap_text=True)

        for column, value in enumerate(REGISTRATION_HEADERS, start=1):
            cell = sheet.cell(5, column, value)
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        sheet.freeze_panes = "A6"
        sheet.auto_filter.ref = "A5:G35"
        widths = (14, 16, 10, 12, 10, 20, 20)
        for index, width in enumerate(widths, start=1):
            sheet.column_dimensions[chr(64 + index)].width = width

        sex_value = "男" if sex_label == "男生" else "女"
        for row in range(6, 36):
            sheet.cell(row, 3, sex_value)
            for column in range(1, 8):
                sheet.cell(row, column).border = border
        sex_validation = DataValidation(type="list", formula1=f'"{sex_value}"', allow_blank=False)
        grade_validation = DataValidation(
            type="list", formula1="'_下拉数据'!$C$2:$C$4", allow_blank=False
        )
        class_validation = DataValidation(
            type="list", formula1="'_下拉数据'!$D$2:$D$31", allow_blank=False
        )
        event_letter = chr(64 + event_column)
        event_validation = DataValidation(
            type="list",
            formula1=f"'_下拉数据'!${event_letter}$2:${event_letter}${len(events) + 1}",
            allow_blank=True,
        )
        for validation in (sex_validation, grade_validation, class_validation, event_validation):
            sheet.add_data_validation(validation)
        sex_validation.add("C6:C35")
        grade_validation.add("D6:D35")
        class_validation.add("E6:E35")
        event_validation.add("F6:G35")

    list_sheet.sheet_state = "hidden"
    workbook.active = workbook.sheetnames.index("男生")
    workbook.save(target)
    return target


def create_registration_template(path: str | Path) -> Path:
    """生成班主任可直接填写、且能被报名导入器读取的标准模板。"""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "报名信息"
    sheet.append(REGISTRATION_HEADERS)
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = "A1:G1"
    sheet.row_dimensions[1].height = 28
    sheet.column_dimensions["A"].width = 14
    sheet.column_dimensions["B"].width = 16
    sheet.column_dimensions["C"].width = 10
    sheet.column_dimensions["D"].width = 12
    sheet.column_dimensions["E"].width = 10
    for column in "FG":
        sheet.column_dimensions[column].width = 20
    sheet.column_dimensions["A"].number_format = "@"

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in sheet[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    sex_validation = DataValidation(type="list", formula1='"男,女"', allow_blank=False)
    grade_validation = DataValidation(
        type="list", formula1='"初一,初二,初三,高一,高二,高三"', allow_blank=False
    )
    event_validation = DataValidation(
        type="list", formula1=f'"{",".join(REGISTRATION_EVENTS)}"', allow_blank=True
    )
    sheet.add_data_validation(sex_validation)
    sheet.add_data_validation(grade_validation)
    sheet.add_data_validation(event_validation)
    sex_validation.add("C2:C500")
    grade_validation.add("D2:D500")
    event_validation.add("F2:G500")

    comments = {
        "A1": "号码按学校统一编码填写；建议把本列保持为文本，避免前导零丢失。",
        "B1": "填写学生真实姓名，不要添加班级或性别。",
        "C1": "从下拉框选择男或女。",
        "D1": "从下拉框选择年级。",
        "E1": "只填写班号，例如9，不要填写“9班”。",
        "F1": "每格只选择一个项目；没有第三个项目可留空。",
    }
    for coordinate, text in comments.items():
        sheet[coordinate].comment = Comment(text, "运动会系统")

    guide = workbook.create_sheet("填写说明（不要导入）")
    guide.append(["要求", "说明"])
    guide_rows = [
        ("一人一行", "同一名学生的项目填写在项目1和项目2，不要重复填写多行。"),
        ("号码唯一", "全校范围内每名学生使用唯一号码。"),
        ("项目名称", "优先从下拉列表选择，避免错别字；确有新增项目时可直接输入。"),
        ("不要改表头", "号码、姓名、性别、年级、班级和项目列名不能删除。"),
        ("提交格式", "保存为 .xlsx；一个班可提交一个文件，系统支持整文件夹批量导入。"),
    ]
    for row in guide_rows:
        guide.append(row)
    guide.column_dimensions["A"].width = 18
    guide.column_dimensions["B"].width = 76
    for cell in guide[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
    workbook.save(target)
    return target


def create_schedule_template(
    path: str | Path,
    plan_rows: Iterable["SchedulePlanRow"] | None = None,
) -> Path:
    """生成可直接交给体育老师填写的结构化比赛日程模板。"""
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    schedule = workbook.active
    schedule.title = "比赛日程"
    schedule.append(SCHEDULE_HEADERS)
    schedule.freeze_panes = "A2"
    schedule.auto_filter.ref = "A1:K1"
    schedule.row_dimensions[1].height = 28

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in schedule[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    widths = (18, 10, 12, 8, 30, 18, 18, 14, 10, 14, 14)
    for column, width in zip("ABCDEFGHIJK", widths, strict=True):
        schedule.column_dimensions[column].width = width

    period_validation = DataValidation(type="list", formula1='"上午,下午"', allow_blank=False)
    type_validation = DataValidation(type="list", formula1='"径赛,田赛,集体项目"', allow_blank=False)
    schedule.add_data_validation(period_validation)
    schedule.add_data_validation(type_validation)
    period_validation.add("B2:B500")
    type_validation.add("C2:C500")

    type_names = {
        EventType.TRACK: "径赛",
        EventType.FIELD: "田赛",
        EventType.TEAM: "集体项目",
    }
    input_fill = PatternFill("solid", fgColor="FFF2CC")
    input_side = Side(style="thin", color="D6C98A")
    input_border = Border(left=input_side, right=input_side, top=input_side, bottom=input_side)
    for row_index, item in enumerate(plan_rows or (), start=2):
        lanes = ""
        if item.available_lanes:
            expected = tuple(range(item.available_lanes[0], item.available_lanes[-1] + 1))
            lanes = (
                f"{item.available_lanes[0]}-{item.available_lanes[-1]}"
                if item.available_lanes == expected and len(item.available_lanes) > 1
                else ",".join(str(lane) for lane in item.available_lanes)
            )
        schedule.append([
            item.competition_date,
            item.period,
            type_names[item.event_type],
            item.sequence,
            item.event_title,
            item.group_name,
            item.event_name,
            item.declared_count,
            item.heat_count,
            item.start_time,
            lanes,
        ])
        schedule.cell(row_index, 1).number_format = "yyyy-mm-dd"
        schedule.cell(row_index, 10).number_format = "h:mm"
        for cell in schedule[row_index]:
            cell.alignment = Alignment(vertical="center")
        for column in (8, 9, 10, 11):
            cell = schedule.cell(row_index, column)
            cell.fill = input_fill
            cell.border = input_border
            cell.alignment = Alignment(horizontal="center", vertical="center")
    if schedule.max_row > 1:
        schedule.auto_filter.ref = f"A1:K{schedule.max_row}"

    comments = {
        "A1": "建议填写完整日期，例如：2026-10-15；也可填写“10月15日（星期四）”。",
        "E1": "完整名称，例如：初一男子组50米预赛。",
        "F1": "例如：初一男子组、初一年级组。",
        "G1": "基础项目名称，例如：50米、立定跳远、4×100米接力。",
        "H1": "只填数字，例如 27；团队项目也填写队数数字。",
        "I1": "必须由体育老师确认。50/100/200米预赛若留空，系统会按可用跑道给出建议。",
        "J1": "例如：9:30 或 10:20-10:40。",
        "K1": "选填。默认使用项目设置中的全部跑道；部分停用时填写 2-8 或 2-4,6,8。",
    }
    for coordinate, text in comments.items():
        schedule[coordinate].comment = Comment(text, "运动会系统")

    guide = workbook.create_sheet("填写说明")
    guide.append(["字段", "是否必填", "填写规则"])
    guide_rows = [
        ("日期", "是", "同一天每行保持一致；推荐 YYYY-MM-DD。"),
        ("时间段", "是", "从下拉框选择上午或下午。"),
        ("比赛类型", "是", "从下拉框选择径赛、田赛或集体项目。"),
        ("序号", "是", "同一日期、时间段、比赛类型内依次编号。"),
        ("项目名称", "是", "组别+项目+赛次，例如“初一男子组50米预赛”。"),
        ("组别", "是", "个人项目用“初一男子组”等；集体项目可用“初一年级组”。"),
        ("项目", "是", "不要包含组别和赛次。统一使用阿拉伯数字，如“1分钟跳绳”。"),
        ("总人数/队数", "是", "只填数字，不要附加“人”或“队”。已导入报名数据时由系统自动计算。"),
        ("总组数", "建议必填", "黄色单元格可人工调整；已导入报名数据时由系统按人数和道数自动计算。"),
        ("比赛时间", "是", "系统已按人数、组数和场地并行关系生成建议时间；黄色单元格可修改。"),
        ("可用道次", "否", "留空表示全部跑道；部分停用时填写 2-8 或 2-4,6,8。"),
    ]
    for row in guide_rows:
        guide.append(row)
    guide.freeze_panes = "A2"
    guide.column_dimensions["A"].width = 18
    guide.column_dimensions["B"].width = 14
    guide.column_dimensions["C"].width = 72
    for cell in guide[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill

    example = workbook.create_sheet("填写示例（不要导入）")
    example.append(SCHEDULE_HEADERS)
    example.append([
        "2026-10-15", "上午", "径赛", 1, "初一男子组50米预赛", "初一男子组",
        "50米", 27, 4, "9:30", "1-8",
    ])
    example.append([
        "2026-10-15", "上午", "田赛", 1, "初一女子组立定跳远预决赛", "初一女子组",
        "立定跳远", 24, 1, "9:30", "",
    ])
    example.append([
        "2026-10-15", "下午", "集体项目", 1, "初一年级女子4×100米接力预决赛", "初一女子组",
        "4×100米接力", 14, 2, "14:30", "1-8",
    ])
    for cell in example[1]:
        cell.font = Font(color="FFFFFF", bold=True)
        cell.fill = header_fill
    for column, width in zip("ABCDEFGHIJK", widths, strict=True):
        example.column_dimensions[column].width = width

    workbook.save(target)
    return target


def create_schedule_test_sample(path: str | Path) -> Path:
    """生成用于验收日程导入的最小测试文件。"""
    target = create_schedule_template(path)
    workbook = load_workbook(target)
    schedule = workbook["比赛日程"]
    schedule.append([
        "2026-10-15", "上午", "径赛", 1, "初一男子组50米预赛", "初一男子组",
        "50米", 17, None, "9:30", "2-8",
    ])
    schedule.append([
        "2026-10-15", "上午", "田赛", 1, "初一女子组立定跳远预决赛", "初一女子组",
        "立定跳远", 16, 1, "9:30", None,
    ])
    schedule.append([
        "2026-10-15", "下午", "集体项目", 1, "初一女子组4×100米接力预决赛", "初一女子组",
        "4×100米接力", 8, 1, "14:30", "1-8",
    ])
    workbook.save(target)
    return target


def create_grouping_registration_sample(path: str | Path) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    boys = workbook.active
    boys.title = "男生"
    boys.append(["号码", "姓名", "性别", "年级", "班级", "项目1", "项目2"])
    for index in range(17):
        boys.append([
            10101 + index * 2, f"男生{index + 1:02d}", "男", "初一", index % 4 + 1, "50米", None,
        ])
    girls = workbook.create_sheet("女生")
    girls.append(["号码", "姓名", "性别", "年级", "班级", "项目1", "项目2"])
    for index in range(8):
        girls.append([
            10202 + index * 2, f"女生{index + 1:02d}", "女", "初一", index % 4 + 1, "立定跳远", None,
        ])
    workbook.save(target)
    return target


def create_grouping_schedule_sample(path: str | Path) -> Path:
    target = create_schedule_template(path)
    workbook = load_workbook(target)
    schedule = workbook["比赛日程"]
    schedule.append([
        "2026-10-15", "上午", "径赛", 1, "初一男子组50米预赛", "初一男子组",
        "50米", 17, 3, "9:30", "2-8",
    ])
    schedule.append([
        "2026-10-15", "上午", "径赛", 2, "初一男子组50米决赛", "初一男子组",
        "50米", 8, 1, "11:30", "1-8",
    ])
    schedule.append([
        "2026-10-15", "上午", "田赛", 1, "初一女子组立定跳远预决赛", "初一女子组",
        "立定跳远", 8, 1, "9:30", None,
    ])
    workbook.save(target)
    return target
