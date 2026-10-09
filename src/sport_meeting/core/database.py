from __future__ import annotations

import hashlib
import json
import shutil
import sqlite3
from dataclasses import asdict
from datetime import datetime
from pathlib import Path
from typing import Callable

from .models import (
    Athlete,
    Entry,
    EventRound,
    EventType,
    FieldStartOrder,
    HeatAssignment,
    PerformanceKind,
    ProjectConfig,
    Qualification,
    Result,
    ResultStatus,
    RoundType,
    SourceFileSnapshot,
    TeamEventConfig,
)


CURRENT_SCHEMA_VERSION = 1
Migration = Callable[[sqlite3.Connection], None]
MIGRATIONS: dict[tuple[int, int], Migration] = {}


class ProjectVersionError(RuntimeError):
    pass


class ProjectDatabase:
    def __init__(self, path: str | Path, connection: sqlite3.Connection):
        self.path = Path(path)
        self.connection = connection
        self.connection.row_factory = sqlite3.Row
        self.connection.execute("PRAGMA foreign_keys = ON")

    @classmethod
    def create(cls, path: str | Path, config: ProjectConfig | None = None) -> "ProjectDatabase":
        project_path = Path(path)
        if project_path.exists():
            raise FileExistsError("项目文件已经存在，请更换名称或使用“打开项目”")
        project_path.parent.mkdir(parents=True, exist_ok=True)
        connection = sqlite3.connect(project_path)
        db = cls(project_path, connection)
        db._create_schema()
        db.save_config(config or ProjectConfig())
        return db

    @classmethod
    def open(cls, path: str | Path) -> "ProjectDatabase":
        project_path = Path(path)
        connection = sqlite3.connect(project_path)
        db = cls(project_path, connection)
        version = db.schema_version
        if version > CURRENT_SCHEMA_VERSION:
            connection.close()
            raise ProjectVersionError("项目由更高版本软件创建，请先升级软件")
        if version < CURRENT_SCHEMA_VERSION:
            db.backup()
            db._migrate(version, CURRENT_SCHEMA_VERSION)
        return db

    @property
    def schema_version(self) -> int:
        row = self.connection.execute("SELECT value FROM metadata WHERE key='schema_version'").fetchone()
        if row is None:
            raise ProjectVersionError("项目缺少 schema_version")
        return int(row[0])

    def _create_schema(self) -> None:
        self.connection.executescript(
            """
            BEGIN;
            CREATE TABLE IF NOT EXISTS metadata (key TEXT PRIMARY KEY, value TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS project_config (id INTEGER PRIMARY KEY CHECK(id=1), json TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS athletes (
                id TEXT PRIMARY KEY, bib TEXT NOT NULL UNIQUE, name TEXT NOT NULL, sex TEXT NOT NULL,
                grade TEXT NOT NULL, class_name TEXT NOT NULL, campus TEXT NOT NULL DEFAULT ''
            );
            CREATE TABLE IF NOT EXISTS entries (
                id TEXT PRIMARY KEY, athlete_id TEXT NOT NULL REFERENCES athletes(id), event_name TEXT NOT NULL,
                group_name TEXT NOT NULL, source_file_id TEXT, source_sheet TEXT, source_row INTEGER,
                UNIQUE(athlete_id, event_name, group_name)
            );
            CREATE TABLE IF NOT EXISTS source_files (
                id TEXT PRIMARY KEY, source_path TEXT NOT NULL, content_hash TEXT NOT NULL, sheet_count INTEGER NOT NULL,
                row_count INTEGER NOT NULL, imported_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS event_rounds (
                id TEXT PRIMARY KEY, event_name TEXT NOT NULL, group_name TEXT NOT NULL,
                event_type TEXT NOT NULL, round_type TEXT NOT NULL, heat_count INTEGER NOT NULL,
                scheduled_time TEXT NOT NULL, declared_count INTEGER, available_lanes_json TEXT NOT NULL,
                performance_kind TEXT NOT NULL, canonical_unit TEXT NOT NULL,
                default_input_unit TEXT NOT NULL, attempts INTEGER NOT NULL
            );
            CREATE TABLE IF NOT EXISTS heat_assignments (
                id TEXT PRIMARY KEY, event_round_id TEXT NOT NULL REFERENCES event_rounds(id),
                participant_id TEXT NOT NULL, heat_no INTEGER NOT NULL, lane INTEGER, start_order INTEGER,
                random_seed INTEGER, manually_adjusted INTEGER NOT NULL DEFAULT 0,
                UNIQUE(event_round_id, participant_id)
            );
            CREATE TABLE IF NOT EXISTS results (
                id INTEGER PRIMARY KEY AUTOINCREMENT, event_round_id TEXT NOT NULL REFERENCES event_rounds(id),
                participant_id TEXT NOT NULL, heat_no INTEGER NOT NULL, raw_value TEXT NOT NULL,
                performance_kind TEXT NOT NULL, canonical_value INTEGER, canonical_unit TEXT,
                standard_display TEXT NOT NULL, status TEXT NOT NULL, manual_rank INTEGER,
                UNIQUE(event_round_id, participant_id)
            );
            CREATE TABLE IF NOT EXISTS qualifications (
                event_round_id TEXT NOT NULL REFERENCES event_rounds(id), participant_id TEXT NOT NULL,
                heat_no INTEGER NOT NULL, heat_rank INTEGER NOT NULL, overall_rank INTEGER NOT NULL,
                canonical_value INTEGER NOT NULL, reason TEXT NOT NULL,
                PRIMARY KEY(event_round_id, participant_id)
            );
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT, entity_type TEXT NOT NULL, entity_id TEXT NOT NULL,
                event_round_id TEXT, action_type TEXT NOT NULL, before_json TEXT, after_json TEXT,
                rules_version TEXT NOT NULL, random_seed INTEGER, created_at TEXT NOT NULL
            );
            INSERT OR REPLACE INTO metadata(key,value) VALUES('schema_version','1');
            COMMIT;
            """
        )

    def _migrate(self, start: int, target: int) -> None:
        current = start
        try:
            self.connection.execute("BEGIN")
            while current < target:
                migration = MIGRATIONS.get((current, current + 1))
                if migration is None:
                    raise ProjectVersionError(f"缺少数据库迁移路径：{current} -> {current + 1}")
                migration(self.connection)
                current += 1
                self.connection.execute(
                    "UPDATE metadata SET value=? WHERE key='schema_version'", (str(current),)
                )
            self.connection.commit()
        except Exception:
            self.connection.rollback()
            raise

    def save_config(self, config: ProjectConfig) -> None:
        config.validate()
        data = asdict(config)
        data["field_start_order"] = config.field_start_order.value
        with self.connection:
            self.connection.execute(
                "INSERT OR REPLACE INTO project_config(id,json) VALUES(1,?)",
                (json.dumps(data, ensure_ascii=False),),
            )

    def load_config(self) -> ProjectConfig:
        row = self.connection.execute("SELECT json FROM project_config WHERE id=1").fetchone()
        if row is None:
            raise ProjectVersionError("项目缺少 ProjectConfig")
        data = json.loads(row[0])
        data["lane_priority"] = tuple(data["lane_priority"])
        for field_name in (
            "junior_boys_events", "junior_girls_events",
            "senior_boys_events", "senior_girls_events",
        ):
            if field_name in data:
                data[field_name] = tuple(data[field_name])
        if "team_events" in data:
            data["team_events"] = tuple(TeamEventConfig(**item) for item in data["team_events"])
        data["field_start_order"] = FieldStartOrder(data["field_start_order"])
        config = ProjectConfig(**data)
        config.validate()
        return config

    def load_result_input_unit(self, event_round_id: str, default: str) -> str:
        row = self.connection.execute(
            "SELECT value FROM metadata WHERE key=?", (f"result_input_unit:{event_round_id}",)
        ).fetchone()
        return row[0] if row else default

    def save_result_input_unit(self, event_round_id: str, unit: str) -> None:
        with self.connection:
            self.connection.execute(
                "INSERT OR REPLACE INTO metadata(key,value) VALUES(?,?)",
                (f"result_input_unit:{event_round_id}", unit),
            )

    def replace_registration(
        self,
        athletes,
        entries,
        source_snapshots: list[SourceFileSnapshot] | tuple[SourceFileSnapshot, ...] = (),
        source_root: str | Path | None = None,
    ) -> None:
        with self.connection:
            self.connection.execute("DELETE FROM qualifications")
            self.connection.execute("DELETE FROM results")
            self.connection.execute("DELETE FROM heat_assignments")
            self.connection.execute("DELETE FROM entries")
            self.connection.execute("DELETE FROM athletes")
            self.connection.execute("DELETE FROM source_files")
            self.connection.executemany(
                "INSERT INTO athletes(id,bib,name,sex,grade,class_name,campus) VALUES(?,?,?,?,?,?,?)",
                [(a.id, a.bib, a.name, a.sex, a.grade, a.class_name, a.campus) for a in athletes],
            )
            self.connection.executemany(
                """INSERT INTO entries(id,athlete_id,event_name,group_name,source_file_id,source_sheet,source_row)
                   VALUES(?,?,?,?,?,?,?)""",
                [(e.id, e.athlete_id, e.event_name, e.group_name, e.source_file_id, e.source_sheet, e.source_row) for e in entries],
            )
            self.connection.executemany(
                """INSERT INTO source_files(
                       id,source_path,content_hash,sheet_count,row_count,imported_at
                   ) VALUES(?,?,?,?,?,?)""",
                [
                    (
                        item.id, item.source_path, item.content_hash, item.sheet_count,
                        item.row_count, item.imported_at,
                    )
                    for item in source_snapshots
                ],
            )
            if source_root is None:
                self.connection.execute(
                    "DELETE FROM metadata WHERE key='registration_source_root'"
                )
            else:
                self.connection.execute(
                    "INSERT OR REPLACE INTO metadata(key,value) VALUES(?,?)",
                    ("registration_source_root", str(Path(source_root).resolve())),
                )

    def load_source_snapshots(self) -> list[SourceFileSnapshot]:
        rows = self.connection.execute(
            "SELECT * FROM source_files ORDER BY source_path"
        ).fetchall()
        return [
            SourceFileSnapshot(
                row["id"], row["source_path"], row["content_hash"],
                row["sheet_count"], row["row_count"], row["imported_at"],
            )
            for row in rows
        ]

    def registration_source_changes(self) -> list[str]:
        """Describe registration files changed since the last successful import."""
        snapshots = self.load_source_snapshots()
        if not snapshots:
            return []
        changes: list[str] = []
        imported_paths = {Path(item.source_path).resolve() for item in snapshots}
        for item in snapshots:
            path = Path(item.source_path)
            if not path.exists():
                changes.append(f"已删除：{path.name}")
                continue
            try:
                current_hash = self._sha256(path)
            except OSError as exc:
                changes.append(f"无法读取：{path.name}（{exc}）")
                continue
            if current_hash != item.content_hash:
                changes.append(f"已修改：{path.name}")

        root_row = self.connection.execute(
            "SELECT value FROM metadata WHERE key='registration_source_root'"
        ).fetchone()
        if root_row:
            root = Path(root_row[0])
            if root.is_dir():
                current_paths = {
                    path.resolve()
                    for path in root.rglob("*")
                    if path.is_file()
                    and path.suffix.lower() in {".xlsx", ".xlsm"}
                    and not path.name.startswith("~$")
                }
                for path in sorted(current_paths - imported_paths, key=lambda item: str(item).lower()):
                    changes.append(f"新增文件：{path.name}")
        return changes

    @staticmethod
    def _sha256(path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as source:
            for chunk in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(chunk)
        return digest.hexdigest()

    def load_registration(self) -> tuple[list[Athlete], list[Entry]]:
        athlete_rows = self.connection.execute("SELECT * FROM athletes ORDER BY bib").fetchall()
        entry_rows = self.connection.execute(
            "SELECT * FROM entries ORDER BY group_name,event_name,source_row,id"
        ).fetchall()
        athletes = [
            Athlete(
                row["id"], row["bib"], row["name"], row["sex"], row["grade"],
                row["class_name"], row["campus"],
            )
            for row in athlete_rows
        ]
        entries = [
            Entry(
                row["id"], row["athlete_id"], row["event_name"], row["group_name"],
                row["source_file_id"] or "", row["source_sheet"] or "", row["source_row"],
            )
            for row in entry_rows
        ]
        return athletes, entries

    def replace_schedule(self, rounds: list[EventRound]) -> None:
        with self.connection:
            self.connection.execute("DELETE FROM qualifications")
            self.connection.execute("DELETE FROM results")
            self.connection.execute("DELETE FROM heat_assignments")
            self.connection.execute("DELETE FROM event_rounds")
            self.connection.executemany(
                """INSERT INTO event_rounds(
                       id,event_name,group_name,event_type,round_type,heat_count,scheduled_time,
                       declared_count,available_lanes_json,performance_kind,canonical_unit,
                       default_input_unit,attempts
                   ) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                [
                    (
                        item.id, item.event_name, item.group_name, item.event_type.value,
                        item.round_type.value, item.heat_count, item.scheduled_time,
                        item.declared_count, json.dumps(item.available_lanes),
                        item.performance_kind.value, item.canonical_unit,
                        item.default_input_unit, item.attempts,
                    )
                    for item in rounds
                ],
            )

    def load_schedule(self) -> list[EventRound]:
        rows = self.connection.execute(
            "SELECT * FROM event_rounds ORDER BY scheduled_time, rowid"
        ).fetchall()
        return [
            EventRound(
                id=row["id"], event_name=row["event_name"], group_name=row["group_name"],
                event_type=EventType(row["event_type"]), round_type=RoundType(row["round_type"]),
                heat_count=row["heat_count"], scheduled_time=row["scheduled_time"],
                declared_count=row["declared_count"],
                available_lanes=tuple(json.loads(row["available_lanes_json"])),
                performance_kind=PerformanceKind(row["performance_kind"]),
                canonical_unit=row["canonical_unit"], default_input_unit=row["default_input_unit"],
                attempts=row["attempts"],
            )
            for row in rows
        ]

    def replace_heat_assignments(self, assignments: list[HeatAssignment], seed: int) -> None:
        previous_count = self.connection.execute("SELECT COUNT(*) FROM heat_assignments").fetchone()[0]
        config = self.load_config()
        with self.connection:
            self.connection.execute("DELETE FROM heat_assignments")
            self.connection.executemany(
                """INSERT INTO heat_assignments(
                       id,event_round_id,participant_id,heat_no,lane,start_order,random_seed,manually_adjusted
                   ) VALUES(?,?,?,?,?,?,?,?)""",
                [
                    (
                        item.id, item.event_round_id, item.participant_id, item.heat_no,
                        item.lane, item.order, item.random_seed, int(item.manually_adjusted),
                    )
                    for item in assignments
                ],
            )
            self.connection.execute(
                """INSERT INTO audit_log(
                       entity_type,entity_id,event_round_id,action_type,before_json,after_json,
                       rules_version,random_seed,created_at
                   ) VALUES(?,?,?,?,?,?,?,?,?)""",
                (
                    "HEAT_ASSIGNMENT", "ALL", None, "REGENERATE",
                    json.dumps({"count": previous_count}), json.dumps({"count": len(assignments)}),
                    config.rules_version, seed, datetime.now().isoformat(timespec="seconds"),
                ),
            )

    def load_heat_assignments(self) -> list[HeatAssignment]:
        rows = self.connection.execute(
            "SELECT * FROM heat_assignments ORDER BY event_round_id,heat_no,COALESCE(lane,start_order),id"
        ).fetchall()
        return [
            HeatAssignment(
                row["id"], row["event_round_id"], row["participant_id"], row["heat_no"],
                row["lane"], row["start_order"], row["random_seed"], bool(row["manually_adjusted"]),
            )
            for row in rows
        ]

    def save_track_final(
        self,
        preliminary_round_id: str,
        final_round_id: str,
        results: list[Result],
        qualifications: list[Qualification],
        final_assignments: list[HeatAssignment],
    ) -> None:
        config = self.load_config()
        with self.connection:
            self.connection.execute("DELETE FROM results WHERE event_round_id=?", (preliminary_round_id,))
            self.connection.executemany(
                """INSERT INTO results(
                       event_round_id,participant_id,heat_no,raw_value,performance_kind,
                       canonical_value,canonical_unit,standard_display,status,manual_rank
                   ) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                [
                    (
                        preliminary_round_id, item.participant_id, item.heat_no, item.raw_value,
                        item.performance_kind.value, item.canonical_value, item.canonical_unit,
                        item.standard_display, item.status.value, item.manual_rank,
                    )
                    for item in results
                ],
            )
            self.connection.execute(
                "DELETE FROM qualifications WHERE event_round_id=?", (preliminary_round_id,)
            )
            self.connection.executemany(
                """INSERT INTO qualifications(
                       event_round_id,participant_id,heat_no,heat_rank,overall_rank,canonical_value,reason
                   ) VALUES(?,?,?,?,?,?,?)""",
                [
                    (
                        preliminary_round_id, item.participant_id, item.heat_no, item.heat_rank,
                        item.overall_rank, item.canonical_value, item.reason,
                    )
                    for item in qualifications
                ],
            )
            self.connection.execute("DELETE FROM heat_assignments WHERE event_round_id=?", (final_round_id,))
            self.connection.executemany(
                """INSERT INTO heat_assignments(
                       id,event_round_id,participant_id,heat_no,lane,start_order,random_seed,manually_adjusted
                   ) VALUES(?,?,?,?,?,?,?,?)""",
                [
                    (
                        item.id, item.event_round_id, item.participant_id, item.heat_no,
                        item.lane, item.order, item.random_seed, int(item.manually_adjusted),
                    )
                    for item in final_assignments
                ],
            )
            self.connection.execute(
                """INSERT INTO audit_log(
                       entity_type,entity_id,event_round_id,action_type,before_json,after_json,
                       rules_version,random_seed,created_at
                   ) VALUES(?,?,?,?,?,?,?,?,?)""",
                (
                    "RESULT", preliminary_round_id, preliminary_round_id, "SAVE_AND_QUALIFY", None,
                    json.dumps(
                        {"result_count": len(results), "qualified_count": len(qualifications)},
                        ensure_ascii=False,
                    ),
                    config.rules_version, None, datetime.now().isoformat(timespec="seconds"),
                ),
            )

    def load_results(self, event_round_id: str) -> list[Result]:
        rows = self.connection.execute(
            "SELECT * FROM results WHERE event_round_id=? ORDER BY heat_no,id", (event_round_id,)
        ).fetchall()
        return [
            Result(
                row["participant_id"], row["heat_no"], row["raw_value"],
                PerformanceKind(row["performance_kind"]), row["canonical_value"],
                row["canonical_unit"], row["standard_display"], ResultStatus(row["status"]),
                row["manual_rank"],
            )
            for row in rows
        ]

    def save_round_results(self, event_round_id: str, results: list[Result]) -> None:
        config = self.load_config()
        with self.connection:
            self.connection.execute("DELETE FROM results WHERE event_round_id=?", (event_round_id,))
            self.connection.executemany(
                """INSERT INTO results(
                       event_round_id,participant_id,heat_no,raw_value,performance_kind,
                       canonical_value,canonical_unit,standard_display,status,manual_rank
                   ) VALUES(?,?,?,?,?,?,?,?,?,?)""",
                [
                    (
                        event_round_id, item.participant_id, item.heat_no, item.raw_value,
                        item.performance_kind.value, item.canonical_value, item.canonical_unit,
                        item.standard_display, item.status.value, item.manual_rank,
                    )
                    for item in results
                ],
            )
            self.connection.execute(
                """INSERT INTO audit_log(
                       entity_type,entity_id,event_round_id,action_type,before_json,after_json,
                       rules_version,random_seed,created_at
                   ) VALUES(?,?,?,?,?,?,?,?,?)""",
                (
                    "RESULT", event_round_id, event_round_id, "SAVE_FINAL_RESULTS", None,
                    json.dumps({"result_count": len(results)}, ensure_ascii=False),
                    config.rules_version, None, datetime.now().isoformat(timespec="seconds"),
                ),
            )

    def load_qualifications(self, event_round_id: str) -> list[Qualification]:
        rows = self.connection.execute(
            "SELECT * FROM qualifications WHERE event_round_id=? ORDER BY overall_rank,canonical_value,participant_id",
            (event_round_id,),
        ).fetchall()
        return [
            Qualification(
                row["participant_id"], row["heat_no"], row["heat_rank"],
                row["overall_rank"], row["canonical_value"], row["reason"],
            )
            for row in rows
        ]

    def backup(self) -> Path:
        backup_dir = self.path.parent / f"{self.path.stem}_backups"
        backup_dir.mkdir(parents=True, exist_ok=True)
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S-%f")
        target = backup_dir / f"{self.path.stem}-{stamp}{self.path.suffix}"
        sequence = 1
        while target.exists():
            target = backup_dir / f"{self.path.stem}-{stamp}-{sequence}{self.path.suffix}"
            sequence += 1
        self.connection.commit()
        shutil.copy2(self.path, target)
        retention = self.load_config().backup_retention
        backups = sorted(backup_dir.glob(f"{self.path.stem}-*{self.path.suffix}"), reverse=True)
        for obsolete in backups[retention:]:
            obsolete.unlink()
        return target

    def close(self) -> None:
        self.connection.close()

    def __enter__(self) -> "ProjectDatabase":
        return self

    def __exit__(self, *_args) -> None:
        self.close()
