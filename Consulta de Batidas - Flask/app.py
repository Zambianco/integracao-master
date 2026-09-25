import hashlib
import json
import os
import sqlite3
from datetime import datetime, timedelta

import requests
import urllib3
from flask import Flask, jsonify, render_template, request
from requests import Response

from api.event import event

try:
    import pythoncom
    import win32com.client
except ImportError:
    pythoncom = None
    win32com = None

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DB_PATH = os.path.join(BASE_DIR, "afd_history.sqlite3")
PARSER_VERSION = "afd-parser-v2"


def load_env_file(env_path: str) -> None:
    if not os.path.exists(env_path):
        return

    with open(env_path, encoding="utf-8") as env_file:
        for raw_line in env_file:
            line = raw_line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue

            key, value = line.split("=", 1)
            key = key.strip()
            value = value.strip().strip('"').strip("'")

            if key:
                os.environ.setdefault(key, value)


load_env_file(os.path.join(BASE_DIR, ".env"))


def get_controlid_config() -> dict:
    base_url = os.getenv("CONTROLID_BASE_URL", "").strip()
    username = os.getenv("CONTROLID_USERNAME", "").strip()
    password = os.getenv("CONTROLID_PASSWORD", "")

    return {
        "base_url": base_url,
        "username": username,
        "password": password,
        "configured": bool(base_url and username and password),
    }


class ControlIDClient:
    def __init__(self, base_url: str, username: str, password: str):
        self.base_url = base_url.rstrip("/")
        self.username = username
        self.password = password
        self.session = requests.Session()
        self.session_key = None

    def login(self) -> str:
        response = self.session.post(
            f"{self.base_url}/login.fcgi",
            json={"login": self.username, "password": self.password},
            headers={"Content-Type": "application/json; charset=utf-8"},
            verify=False,
            timeout=20,
        )
        response.raise_for_status()

        data = response.json()
        self.session_key = data["session"]
        return self.session_key

    def get_afd(self, initial_date: str, final_date: str | None = None) -> str:
        if not self.session_key:
            raise RuntimeError("Nao autenticado")

        start = datetime.strptime(initial_date, "%d/%m/%Y")
        payload = {
            "initial_date": {
                "day": start.day,
                "month": start.month,
                "year": start.year,
            }
        }

        if final_date:
            end = datetime.strptime(final_date, "%d/%m/%Y")
            payload["final_date"] = {
                "day": end.day,
                "month": end.month,
                "year": end.year,
            }

        response = self.session.post(
            f"{self.base_url}/get_afd.fcgi?session={self.session_key}",
            json=payload,
            headers={"Content-Type": "application/json; charset=utf-8"},
            verify=False,
            timeout=60,
        )
        response.raise_for_status()
        return response.text

    @staticmethod
    def _raise_http_error_with_details(response: Response) -> None:
        try:
            details = response.json()
        except ValueError:
            details = response.text.strip()

        message = f"HTTP {response.status_code}"
        if details:
            message = f"{message}: {details}"

        raise requests.HTTPError(message, response=response)

    def export_afd(self, initial_nsr: int | None = None, initial_date: str | None = None, mode: int = 595) -> str:
        if not self.session_key:
            raise RuntimeError("Nao autenticado")

        payload = {}
        if initial_nsr is not None:
            payload["initial_nsr"] = int(initial_nsr)

        if initial_date:
            start = datetime.strptime(initial_date, "%d/%m/%Y")
            payload["initial_date"] = {
                "day": start.day,
                "month": start.month,
                "year": start.year,
            }

        request_variants = [
            {
                "url": f"{self.base_url}/export_afd.fcgi?session={self.session_key}&mode={int(mode)}",
                "json": payload,
            },
            {
                "url": f"{self.base_url}/export_afd.fcgi?session={self.session_key}",
                "json": {**payload, "mode": int(mode)},
            },
            {
                "url": f"{self.base_url}/export_afd.fcgi?session={self.session_key}",
                "json": payload,
            },
        ]

        last_response = None
        for variant in request_variants:
            response = self.session.post(
                variant["url"],
                json=variant["json"],
                headers={"Content-Type": "application/json; charset=utf-8"},
                verify=False,
                timeout=60,
            )

            if response.ok:
                return response.text

            last_response = response
            if response.status_code != 400:
                self._raise_http_error_with_details(response)

        if last_response is not None:
            self._raise_http_error_with_details(last_response)

        raise RuntimeError("Falha inesperada ao exportar AFD.")


class RepCidClient:
    def __init__(self, host: str, username: str, password: str, port: int = 443, passcode: int = 0):
        self.host = host.replace("https://", "").replace("http://", "").split("/")[0].strip()
        self.username = username
        self.password = password
        self.port = port
        self.passcode = passcode

    @staticmethod
    def is_available() -> bool:
        return pythoncom is not None and win32com is not None

    def _create_rep(self):
        if not self.is_available():
            raise RuntimeError("Suporte COM nao disponivel no Python atual.")

        last_error = None
        for prog_id in ("RepCid.RepCid", "Controlid.RepCid", "RepCid_.RepCid_", "RepCid_"):
            try:
                return win32com.client.Dispatch(prog_id)
            except Exception as exc:
                last_error = exc

        raise RuntimeError(
            "RepCid COM nao registrada no Windows. Registre a RepCid.dll com RegAsm antes de usar busca por NSR."
        ) from last_error

    @staticmethod
    def _parse_ler_afd_result(result):
        if isinstance(result, tuple):
            if len(result) == 2:
                first, second = result
                if isinstance(first, bool):
                    return first, str(second or "")
                if isinstance(second, bool):
                    return second, str(first or "")
            if len(result) == 1:
                if isinstance(result[0], bool):
                    return result[0], ""
                return True, str(result[0] or "")
        if isinstance(result, bool):
            return result, ""
        if result is None:
            return False, ""
        return True, str(result)

    def get_afd_from_nsr(self, start_nsr: int, limit: int | None = None) -> str:
        pythoncom.CoInitialize()
        try:
            rep = self._create_rep()
            rep.iDClassLogin = self.username
            rep.iDClassPassword = self.password

            status = rep.Conectar(self.host, self.port, self.passcode)
            if int(status) != 1:
                last_log = ""
                try:
                    last_log = rep.GetLastLog()
                except Exception:
                    last_log = ""
                raise RuntimeError(f"Falha ao conectar via RepCid. Status={status}. {last_log}".strip())

            if not rep.BuscarAFD(int(start_nsr)):
                last_log = ""
                try:
                    last_log = rep.GetLastLog()
                except Exception:
                    last_log = ""
                raise RuntimeError(f"BuscarAFD falhou para NSR {start_nsr}. {last_log}".strip())

            lines = []
            while True:
                has_line, line = self._parse_ler_afd_result(rep.LerAFD())
                if not has_line:
                    break
                if line:
                    lines.append(line)
                if limit and len(lines) >= limit:
                    break

            try:
                rep.Desconectar()
            except Exception:
                pass

            return "\n".join(lines)
        finally:
            pythoncom.CoUninitialize()


class AFDRepository:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self._ensure_schema()
        self._migrate_legacy_cache()
        self._migrate_schema_extensions()
        self._backfill_integrity_tables()

    def _connect(self) -> sqlite3.Connection:
        connection = sqlite3.connect(self.db_path)
        connection.row_factory = sqlite3.Row
        connection.execute("PRAGMA foreign_keys = ON")
        self._ensure_runtime_schema(connection)
        return connection

    @staticmethod
    def _table_exists(connection: sqlite3.Connection, table_name: str) -> bool:
        row = connection.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?",
            (table_name,),
        ).fetchone()
        return row is not None

    @classmethod
    def _column_exists(cls, connection: sqlite3.Connection, table_name: str, column_name: str) -> bool:
        if not cls._table_exists(connection, table_name):
            return False
        columns = {
            str(row["name"])
            for row in connection.execute(f"PRAGMA table_info({table_name})").fetchall()
        }
        return column_name in columns

    def _ensure_runtime_schema(self, connection: sqlite3.Connection) -> None:
        if not self._table_exists(connection, "afd_records"):
            return
        if self._column_exists(connection, "afd_records", "parsed_json"):
            return
        self._migrate_afd_records_to_canonical(connection)

    def _ensure_schema(self) -> None:
        with self._connect() as connection:
            connection.executescript(
                """
                CREATE TABLE IF NOT EXISTS devices (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    device_key TEXT NOT NULL UNIQUE,
                    base_url TEXT NOT NULL,
                    description TEXT,
                    enabled INTEGER NOT NULL DEFAULT 1,
                    supports_nsr_query INTEGER,
                    supports_repcid INTEGER,
                    supports_date_query INTEGER,
                    last_successful_sync TEXT,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
                );

                CREATE TABLE IF NOT EXISTS afd_imports (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    device_id INTEGER NOT NULL,
                    query_mode TEXT NOT NULL,
                    requested_initial_nsr INTEGER,
                    requested_initial_date TEXT,
                    requested_final_date TEXT,
                    source_type TEXT NOT NULL,
                    raw_afd TEXT NOT NULL,
                    raw_hash TEXT NOT NULL,
                    first_nsr_found INTEGER,
                    last_nsr_found INTEGER,
                    total_lines INTEGER NOT NULL DEFAULT 0,
                    total_parsed_lines INTEGER NOT NULL DEFAULT 0,
                    imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    parser_version TEXT NOT NULL,
                    FOREIGN KEY(device_id) REFERENCES devices(id)
                );

                CREATE INDEX IF NOT EXISTS idx_afd_imports_device_imported_at
                ON afd_imports(device_id, imported_at);

                CREATE INDEX IF NOT EXISTS idx_afd_imports_device_last_nsr
                ON afd_imports(device_id, last_nsr_found);

                CREATE TABLE IF NOT EXISTS afd_records (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    device_id INTEGER NOT NULL,
                    afd_import_id INTEGER,
                    nsr INTEGER NOT NULL,
                    record_type TEXT NOT NULL,
                    event_timestamp TEXT,
                    employee_identifier TEXT,
                    operation_type TEXT,
                    event_code TEXT,
                    parsed_json TEXT NOT NULL,
                    raw_line TEXT NOT NULL,
                    raw_line_hash TEXT NOT NULL,
                    parser_version TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(device_id, nsr),
                    FOREIGN KEY(device_id) REFERENCES devices(id),
                    FOREIGN KEY(afd_import_id) REFERENCES afd_imports(id)
                );

                CREATE INDEX IF NOT EXISTS idx_afd_records_device_nsr
                ON afd_records(device_id, nsr);

                CREATE INDEX IF NOT EXISTS idx_afd_records_device_event_timestamp
                ON afd_records(device_id, event_timestamp);

                CREATE INDEX IF NOT EXISTS idx_afd_records_device_employee
                ON afd_records(device_id, employee_identifier);

                CREATE INDEX IF NOT EXISTS idx_afd_records_device_type
                ON afd_records(device_id, record_type);

                CREATE INDEX IF NOT EXISTS idx_afd_records_device_event_code
                ON afd_records(device_id, event_code);

                CREATE TABLE IF NOT EXISTS afd_day_summary (
                    device_id INTEGER NOT NULL,
                    event_date TEXT NOT NULL,
                    first_nsr INTEGER,
                    last_nsr INTEGER,
                    total_records INTEGER NOT NULL,
                    total_punches INTEGER NOT NULL,
                    previous_day_last_nsr INTEGER,
                    next_day_first_nsr INTEGER,
                    is_nsr_contiguous_before INTEGER,
                    is_nsr_contiguous_after INTEGER,
                    integrity_status TEXT NOT NULL DEFAULT 'UNKNOWN',
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    PRIMARY KEY(device_id, event_date),
                    FOREIGN KEY(device_id) REFERENCES devices(id)
                );

                CREATE INDEX IF NOT EXISTS idx_afd_day_summary_device_event_date
                ON afd_day_summary(device_id, event_date);

                CREATE TABLE IF NOT EXISTS afd_nsr_gaps (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    device_id INTEGER NOT NULL,
                    start_nsr INTEGER NOT NULL,
                    end_nsr INTEGER NOT NULL,
                    missing_count INTEGER NOT NULL,
                    detected_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    resolved INTEGER NOT NULL DEFAULT 0,
                    gap_status TEXT NOT NULL DEFAULT 'OPEN',
                    resolved_at TEXT,
                    detected_import_id INTEGER,
                    resolved_import_id INTEGER,
                    notes TEXT,
                    FOREIGN KEY(device_id) REFERENCES devices(id),
                    FOREIGN KEY(detected_import_id) REFERENCES afd_imports(id),
                    FOREIGN KEY(resolved_import_id) REFERENCES afd_imports(id),
                    CHECK (end_nsr >= start_nsr)
                );

                CREATE INDEX IF NOT EXISTS idx_afd_nsr_gaps_device_open
                ON afd_nsr_gaps(device_id, resolved, start_nsr, end_nsr);

                CREATE TABLE IF NOT EXISTS afd_record_conflicts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    device_id INTEGER NOT NULL,
                    afd_import_id INTEGER,
                    nsr INTEGER NOT NULL,
                    existing_record_id INTEGER,
                    existing_import_id INTEGER,
                    existing_raw_line TEXT NOT NULL,
                    incoming_raw_line TEXT NOT NULL,
                    existing_hash TEXT NOT NULL,
                    incoming_hash TEXT NOT NULL,
                    conflict_type TEXT NOT NULL,
                    detected_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    resolved_at TEXT,
                    resolution_note TEXT,
                    FOREIGN KEY(device_id) REFERENCES devices(id),
                    FOREIGN KEY(afd_import_id) REFERENCES afd_imports(id),
                    FOREIGN KEY(existing_import_id) REFERENCES afd_imports(id),
                    FOREIGN KEY(existing_record_id) REFERENCES afd_records(id)
                );

                CREATE INDEX IF NOT EXISTS idx_afd_conflicts_device_nsr
                ON afd_record_conflicts(device_id, nsr);
                """
            )

    def _migrate_legacy_cache(self) -> None:
        legacy_db_path = os.path.join(os.path.dirname(self.db_path), "afd_cache.sqlite3")
        if not os.path.exists(legacy_db_path) or os.path.abspath(legacy_db_path) == os.path.abspath(self.db_path):
            return

        remove_legacy_db = False
        legacy_connection = sqlite3.connect(legacy_db_path)
        legacy_connection.row_factory = sqlite3.Row
        try:
            tables = {
                row["name"]
                for row in legacy_connection.execute(
                    "SELECT name FROM sqlite_master WHERE type = 'table'"
                ).fetchall()
            }
            if "afd_records" not in tables:
                return

            legacy_rows = legacy_connection.execute(
                """
                SELECT device_key, nsr, tipo, data, hora, pis, detalhes, classe, comprimento, linha_original
                FROM afd_records
                ORDER BY device_key, nsr
                """
            ).fetchall()
            if not legacy_rows:
                remove_legacy_db = True
                return

            imported_rows = 0
            with self._connect() as connection:
                device_ids: dict[str, int] = {}
                for row in legacy_rows:
                    device_key = str(row["device_key"] or "").strip().lower()
                    if not device_key:
                        continue

                    device_id = device_ids.get(device_key)
                    if device_id is None:
                        existing_device = connection.execute(
                            "SELECT id FROM devices WHERE device_key = ?",
                            (device_key,),
                        ).fetchone()
                        if existing_device:
                            device_id = int(existing_device["id"])
                        else:
                            cursor = connection.execute(
                                """
                                INSERT INTO devices (device_key, base_url, description)
                                VALUES (?, ?, ?)
                                """,
                                (device_key, device_key, "Migrado de afd_cache.sqlite3"),
                            )
                            device_id = int(cursor.lastrowid)
                        device_ids[device_key] = device_id

                    raw_line = str(row["linha_original"] or "")
                    if not raw_line:
                        continue

                    parsed = parse_afd_line(raw_line)
                    if not parsed:
                        continue
                    raw_line_hash = self._hash_text(raw_line)
                    storage = self._record_storage_params(parsed, raw_line_hash, PARSER_VERSION)
                    connection.execute(
                        """
                        INSERT OR IGNORE INTO afd_records (
                            device_id, afd_import_id, nsr, record_type, event_timestamp,
                            employee_identifier, operation_type, event_code, parsed_json,
                            raw_line, raw_line_hash, parser_version
                        ) VALUES (?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            device_id,
                            int(row["nsr"]),
                            *storage,
                        ),
                    )
                    imported_rows += 1

                if imported_rows:
                    connection.commit()

            remove_legacy_db = True
        finally:
            legacy_connection.close()
            if remove_legacy_db:
                try:
                    os.remove(legacy_db_path)
                except OSError:
                    pass

    def _migrate_schema_extensions(self) -> None:
        with self._connect() as connection:
            self._ensure_column(connection, "devices", "enabled", "INTEGER NOT NULL DEFAULT 1")
            self._ensure_column(connection, "devices", "supports_nsr_query", "INTEGER")
            self._ensure_column(connection, "devices", "supports_repcid", "INTEGER")
            self._ensure_column(connection, "devices", "supports_date_query", "INTEGER")
            self._ensure_column(connection, "devices", "last_successful_sync", "TEXT")

            self._migrate_afd_records_to_canonical(connection)
            self._ensure_column(connection, "afd_records", "event_timestamp", "TEXT")
            self._ensure_column(connection, "afd_records", "employee_identifier", "TEXT")
            self._ensure_column(connection, "afd_records", "operation_type", "TEXT")
            self._ensure_column(connection, "afd_records", "event_code", "TEXT")
            self._ensure_column(connection, "afd_records", "parsed_json", "TEXT NOT NULL DEFAULT '{}'")
            self._ensure_column(connection, "afd_nsr_gaps", "gap_status", "TEXT NOT NULL DEFAULT 'OPEN'")
            self._ensure_column(connection, "afd_record_conflicts", "existing_import_id", "INTEGER")

            connection.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_afd_records_device_event_timestamp
                ON afd_records(device_id, event_timestamp)
                """
            )
            connection.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_afd_records_device_employee
                ON afd_records(device_id, employee_identifier)
                """
            )
            connection.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_afd_records_device_type
                ON afd_records(device_id, record_type)
                """
            )
            connection.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_afd_records_device_event_code
                ON afd_records(device_id, event_code)
                """
            )
            connection.execute(
                """
                CREATE INDEX IF NOT EXISTS idx_afd_nsr_gaps_device_open
                ON afd_nsr_gaps(device_id, resolved, start_nsr, end_nsr)
                """
            )

            self._backfill_event_timestamps(connection)
            self._backfill_gap_status(connection)
            self._deduplicate_import_hashes(connection)
            connection.execute(
                """
                CREATE UNIQUE INDEX IF NOT EXISTS idx_afd_imports_hash
                ON afd_imports(device_id, raw_hash)
                """
            )
            connection.commit()

    @staticmethod
    def _ensure_column(connection: sqlite3.Connection, table_name: str, column_name: str, ddl: str) -> None:
        if AFDRepository._column_exists(connection, table_name, column_name):
            return
        connection.execute(f"ALTER TABLE {table_name} ADD COLUMN {column_name} {ddl}")

    def _migrate_afd_records_to_canonical(self, connection: sqlite3.Connection) -> None:
        if not self._table_exists(connection, "afd_records"):
            return
        if self._column_exists(connection, "afd_records", "parsed_json") and not self._column_exists(
            connection, "afd_records", "date_value"
        ):
            return

        legacy_rows = connection.execute("SELECT * FROM afd_records ORDER BY device_id, nsr").fetchall()

        connection.commit()
        connection.execute("PRAGMA foreign_keys = OFF")
        try:
            connection.execute("DROP TABLE IF EXISTS afd_records_canonical")
            connection.execute(
                """
                CREATE TABLE afd_records_canonical (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    device_id INTEGER NOT NULL,
                    afd_import_id INTEGER,
                    nsr INTEGER NOT NULL,
                    record_type TEXT NOT NULL,
                    event_timestamp TEXT,
                    employee_identifier TEXT,
                    operation_type TEXT,
                    event_code TEXT,
                    parsed_json TEXT NOT NULL,
                    raw_line TEXT NOT NULL,
                    raw_line_hash TEXT NOT NULL,
                    parser_version TEXT NOT NULL,
                    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(device_id, nsr),
                    FOREIGN KEY(device_id) REFERENCES devices(id),
                    FOREIGN KEY(afd_import_id) REFERENCES afd_imports(id)
                )
                """
            )

            for row in legacy_rows:
                raw_line = str(row["raw_line"] or "")
                parsed = parse_afd_line(raw_line) if raw_line else None
                if not parsed:
                    parsed = {
                        "nsr": str(row["nsr"]),
                        "tipo": str(row["record_type"] or ""),
                        "event_timestamp": str(row["event_timestamp"] or ""),
                        "employee_identifier": str(row["pis"] or "") if "pis" in row.keys() else "",
                        "operation_type": "",
                        "event_code": "",
                        "parsed_json": {
                            "parse_status": "legacy_unparsed",
                            "raw_line_length": len(raw_line),
                        },
                        "linha_original": raw_line,
                    }

                parsed["nsr"] = str(row["nsr"])
                parsed["tipo"] = str(parsed.get("tipo") or row["record_type"] or "")
                parsed["linha_original"] = raw_line
                if not parsed.get("event_timestamp") and "event_timestamp" in row.keys():
                    parsed["event_timestamp"] = str(row["event_timestamp"] or "")
                if not parsed.get("event_timestamp") and "date_value" in row.keys() and "time_value" in row.keys():
                    parsed["event_timestamp"] = self._to_event_timestamp(
                        str(row["date_value"] or ""),
                        str(row["time_value"] or ""),
                    )

                raw_line_hash = str(row["raw_line_hash"] or self._hash_text(raw_line))
                storage = self._record_storage_params(parsed, raw_line_hash, PARSER_VERSION)
                connection.execute(
                    """
                    INSERT INTO afd_records_canonical (
                        id, device_id, afd_import_id, nsr, record_type, event_timestamp,
                        employee_identifier, operation_type, event_code, parsed_json,
                        raw_line, raw_line_hash, parser_version, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        int(row["id"]),
                        int(row["device_id"]),
                        int(row["afd_import_id"]) if row["afd_import_id"] is not None else None,
                        int(row["nsr"]),
                        *storage,
                        str(row["created_at"] or datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    ),
                )

            connection.execute("DROP TABLE afd_records")
            connection.execute("ALTER TABLE afd_records_canonical RENAME TO afd_records")
            if self._table_exists(connection, "afd_day_summary"):
                connection.execute("DELETE FROM afd_day_summary")
            if self._table_exists(connection, "afd_nsr_gaps"):
                connection.execute("DELETE FROM afd_nsr_gaps")
            connection.commit()
        finally:
            connection.execute("PRAGMA foreign_keys = ON")

    def _backfill_event_timestamps(self, connection: sqlite3.Connection) -> None:
        rows = connection.execute(
            """
            SELECT id, raw_line, event_timestamp
            FROM afd_records
            WHERE event_timestamp IS NULL
               OR event_timestamp = ''
            """
        ).fetchall()
        for row in rows:
            parsed = parse_afd_line(str(row["raw_line"] or ""))
            event_timestamp = str(parsed.get("event_timestamp") or "") if parsed else ""
            if not event_timestamp:
                continue
            connection.execute(
                "UPDATE afd_records SET event_timestamp = ? WHERE id = ?",
                (event_timestamp, int(row["id"])),
            )

    @staticmethod
    def _backfill_gap_status(connection: sqlite3.Connection) -> None:
        connection.execute(
            """
            UPDATE afd_nsr_gaps
            SET gap_status = CASE
                WHEN resolved = 1 AND (gap_status IS NULL OR gap_status = '' OR gap_status = 'OPEN') THEN 'RESOLVED'
                WHEN resolved = 0 AND (gap_status IS NULL OR gap_status = '') THEN 'OPEN'
                ELSE gap_status
            END
            """
        )

    @staticmethod
    def _deduplicate_import_hashes(connection: sqlite3.Connection) -> None:
        duplicate_groups = connection.execute(
            """
            SELECT device_id, raw_hash, MIN(id) AS kept_id, COUNT(*) AS total_rows
            FROM afd_imports
            GROUP BY device_id, raw_hash
            HAVING COUNT(*) > 1
            """
        ).fetchall()

        for group in duplicate_groups:
            device_id = int(group["device_id"])
            raw_hash = str(group["raw_hash"])
            kept_id = int(group["kept_id"])
            duplicate_ids = [
                int(row["id"])
                for row in connection.execute(
                    """
                    SELECT id
                    FROM afd_imports
                    WHERE device_id = ? AND raw_hash = ? AND id <> ?
                    ORDER BY id ASC
                    """,
                    (device_id, raw_hash, kept_id),
                ).fetchall()
            ]
            if not duplicate_ids:
                continue

            placeholders = ",".join("?" for _ in duplicate_ids)
            params = [kept_id, *duplicate_ids]
            connection.execute(
                f"UPDATE afd_records SET afd_import_id = ? WHERE afd_import_id IN ({placeholders})",
                params,
            )
            connection.execute(
                f"UPDATE afd_record_conflicts SET afd_import_id = ? WHERE afd_import_id IN ({placeholders})",
                params,
            )
            connection.execute(
                f"UPDATE afd_record_conflicts SET existing_import_id = ? WHERE existing_import_id IN ({placeholders})",
                params,
            )
            connection.execute(
                f"UPDATE afd_nsr_gaps SET detected_import_id = ? WHERE detected_import_id IN ({placeholders})",
                params,
            )
            connection.execute(
                f"UPDATE afd_nsr_gaps SET resolved_import_id = ? WHERE resolved_import_id IN ({placeholders})",
                params,
            )
            connection.execute(
                f"DELETE FROM afd_imports WHERE id IN ({placeholders})",
                duplicate_ids,
            )

    def _backfill_integrity_tables(self) -> None:
        with self._connect() as connection:
            stats = connection.execute(
                """
                SELECT
                    (SELECT COUNT(*) FROM afd_records) AS total_records,
                    (SELECT COUNT(*) FROM afd_day_summary) AS total_days,
                    (SELECT COUNT(*) FROM afd_nsr_gaps) AS total_gaps
                """
            ).fetchone()
            if not stats or int(stats["total_records"] or 0) == 0:
                return

            should_rebuild_days = int(stats["total_days"] or 0) == 0
            should_rebuild_gaps = int(stats["total_gaps"] or 0) == 0
            if not should_rebuild_days and not should_rebuild_gaps:
                return

            device_ids = [
                int(row["device_id"])
                for row in connection.execute(
                    "SELECT DISTINCT device_id FROM afd_records ORDER BY device_id"
                ).fetchall()
            ]

            for device_id in device_ids:
                if should_rebuild_gaps:
                    connection.execute(
                        "DELETE FROM afd_nsr_gaps WHERE device_id = ?",
                        (device_id,),
                    )
                    self._rebuild_all_open_gaps(connection, device_id)

                if should_rebuild_days:
                    all_dates = [
                        str(row["event_date"])
                        for row in connection.execute(
                            """
                            SELECT DISTINCT substr(event_timestamp, 1, 10) AS event_date
                            FROM afd_records
                            WHERE device_id = ?
                              AND event_timestamp IS NOT NULL
                              AND event_timestamp <> ''
                            ORDER BY event_date
                            """,
                            (device_id,),
                        ).fetchall()
                    ]
                    self._refresh_day_summaries(connection, device_id, all_dates)

            connection.commit()

    @staticmethod
    def build_device_key(base_url: str) -> str:
        return base_url.rstrip("/").lower()

    @staticmethod
    def _hash_text(value: str) -> str:
        return hashlib.sha256(value.encode("utf-8", errors="ignore")).hexdigest()

    @staticmethod
    def _json_text(value) -> str:
        if isinstance(value, str):
            normalized = value.strip()
            return normalized if normalized else "{}"
        return json.dumps(value or {}, ensure_ascii=False, sort_keys=True)

    @staticmethod
    def _record_class_from_type(record_type: str) -> str:
        return {
            "1": "cabecalho_afd",
            "2": "identificacao_empregador",
            "3": "batida_ponto",
            "4": "ajuste_relogio",
            "5": "cadastro_empregado",
            "6": "evento_tipo_6",
            "7": "marcacao_repp",
            "9": "trailer_afd",
        }.get(str(record_type or ""), "outro_evento")

    @staticmethod
    def _event_date_from_timestamp(event_timestamp: str | None) -> str | None:
        normalized = str(event_timestamp or "").strip()
        if len(normalized) >= 10:
            return normalized[:10]
        return None

    @staticmethod
    def _display_date_time_from_timestamp(event_timestamp: str | None) -> tuple[str, str]:
        normalized = str(event_timestamp or "").strip()
        if len(normalized) < 10:
            return "", ""

        display_date = ""
        display_time = ""
        try:
            display_date = datetime.strptime(normalized[:10], "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            display_date = normalized[:10]

        if len(normalized) >= 16:
            display_time = normalized[11:16]
        return display_date, display_time

    @classmethod
    def _record_storage_params(cls, record: dict, raw_line_hash: str, parser_version: str) -> tuple:
        parsed_json = record.get("parsed_json") or {}
        if isinstance(parsed_json, str):
            try:
                parsed_payload = json.loads(parsed_json)
            except ValueError:
                parsed_payload = {"raw": parsed_json}
        else:
            parsed_payload = parsed_json

        employee_identifier = (
            record.get("employee_identifier")
            or parsed_payload.get("employee_identifier")
            or parsed_payload.get("cpf")
            or parsed_payload.get("pis")
            or ""
        )
        operation_type = record.get("operation_type") or parsed_payload.get("operation_type") or parsed_payload.get("operation") or ""
        event_code = record.get("event_code") or parsed_payload.get("event_code") or ""

        return (
            str(record.get("tipo") or record.get("record_type") or ""),
            str(record.get("event_timestamp") or ""),
            str(employee_identifier or ""),
            str(operation_type or ""),
            str(event_code or ""),
            cls._json_text(parsed_payload),
            str(record.get("linha_original") or record.get("raw_line") or ""),
            raw_line_hash,
            parser_version,
        )

    @staticmethod
    def _to_iso_date(date_value: str) -> str | None:
        try:
            return datetime.strptime(date_value, "%d/%m/%Y").strftime("%Y-%m-%d")
        except ValueError:
            return None

    @staticmethod
    def _to_display_date(date_value: str | None) -> str | None:
        if not date_value:
            return None
        try:
            return datetime.strptime(date_value, "%Y-%m-%d").strftime("%d/%m/%Y")
        except ValueError:
            return date_value

    @staticmethod
    def _to_event_timestamp(date_value: str, time_value: str) -> str | None:
        normalized_date = str(date_value or "").strip()
        normalized_time = str(time_value or "").strip()
        if not normalized_date or not normalized_time:
            return None

        for pattern in ("%d/%m/%Y %H:%M", "%d/%m/%Y %H:%M:%S"):
            try:
                parsed = datetime.strptime(f"{normalized_date} {normalized_time}", pattern)
                return parsed.strftime("%Y-%m-%d %H:%M:%S")
            except ValueError:
                continue
        return None

    @staticmethod
    def _build_integrity_status(
        previous_day_last_nsr: int | None,
        next_day_first_nsr: int | None,
        is_before: int,
        is_after: int,
    ) -> str:
        has_previous = previous_day_last_nsr is not None
        has_next = next_day_first_nsr is not None

        if has_previous and has_next and is_before and is_after:
            return "COMPLETE"
        if has_previous and has_next and not is_before and not is_after:
            return "GAP_BOTH"
        if has_previous and not is_before:
            return "GAP_BEFORE"
        if has_next and not is_after:
            return "GAP_AFTER"
        if not has_previous and not has_next:
            return "ISOLATED"
        if not has_previous:
            return "OPEN_START"
        if not has_next:
            return "OPEN_END"
        return "PARTIAL"

    @staticmethod
    def _compress_nsr_ranges(nsrs: list[int]) -> list[tuple[int, int]]:
        if not nsrs:
            return []

        ranges: list[tuple[int, int]] = []
        start = end = nsrs[0]
        for nsr in nsrs[1:]:
            if nsr == end + 1:
                end = nsr
                continue
            ranges.append((start, end))
            start = end = nsr
        ranges.append((start, end))
        return ranges

    def _expand_summary_dates(self, connection: sqlite3.Connection, device_id: int, iso_dates: set[str]) -> set[str]:
        expanded = set(iso_dates)
        for event_date in list(iso_dates):
            previous = connection.execute(
                """
                SELECT event_date
                FROM afd_day_summary
                WHERE device_id = ? AND event_date < ?
                ORDER BY event_date DESC
                LIMIT 1
                """,
                (device_id, event_date),
            ).fetchone()
            next_day = connection.execute(
                """
                SELECT event_date
                FROM afd_day_summary
                WHERE device_id = ? AND event_date > ?
                ORDER BY event_date ASC
                LIMIT 1
                """,
                (device_id, event_date),
            ).fetchone()
            if previous:
                expanded.add(str(previous["event_date"]))
            if next_day:
                expanded.add(str(next_day["event_date"]))
        return expanded

    def _refresh_day_summaries(self, connection: sqlite3.Connection, device_id: int, iso_dates: list[str] | set[str]) -> None:
        if not iso_dates:
            return

        target_dates = self._expand_summary_dates(connection, device_id, {value for value in iso_dates if value})
        for event_date in sorted(target_dates):
            row = connection.execute(
                """
                SELECT
                    MIN(nsr) AS first_nsr,
                    MAX(nsr) AS last_nsr,
                    COUNT(*) AS total_records,
                    SUM(CASE WHEN record_type IN ('3', '7') THEN 1 ELSE 0 END) AS total_punches
                FROM afd_records
                WHERE device_id = ?
                  AND substr(event_timestamp, 1, 10) = ?
                """,
                (device_id, event_date),
            ).fetchone()

            total_records = int(row["total_records"] or 0) if row else 0
            if total_records == 0:
                connection.execute(
                    "DELETE FROM afd_day_summary WHERE device_id = ? AND event_date = ?",
                    (device_id, event_date),
                )
                continue

            previous = connection.execute(
                """
                SELECT MAX(nsr) AS last_nsr
                FROM afd_records
                WHERE device_id = ?
                  AND event_timestamp IS NOT NULL
                  AND event_timestamp <> ''
                  AND substr(event_timestamp, 1, 10) < ?
                """,
                (device_id, event_date),
            ).fetchone()
            next_day = connection.execute(
                """
                SELECT MIN(nsr) AS first_nsr
                FROM afd_records
                WHERE device_id = ?
                  AND event_timestamp IS NOT NULL
                  AND event_timestamp <> ''
                  AND substr(event_timestamp, 1, 10) > ?
                """,
                (device_id, event_date),
            ).fetchone()

            first_nsr = int(row["first_nsr"])
            last_nsr = int(row["last_nsr"])
            previous_day_last_nsr = int(previous["last_nsr"]) if previous and previous["last_nsr"] is not None else None
            next_day_first_nsr = int(next_day["first_nsr"]) if next_day and next_day["first_nsr"] is not None else None
            is_before = int(
                previous_day_last_nsr is not None
                and not self._has_open_gap_in_range(connection, device_id, previous_day_last_nsr + 1, first_nsr - 1)
            )
            is_after = int(
                next_day_first_nsr is not None
                and not self._has_open_gap_in_range(connection, device_id, last_nsr + 1, next_day_first_nsr - 1)
            )
            integrity_status = self._build_integrity_status(
                previous_day_last_nsr=previous_day_last_nsr,
                next_day_first_nsr=next_day_first_nsr,
                is_before=is_before,
                is_after=is_after,
            )

            connection.execute(
                """
                INSERT INTO afd_day_summary (
                    device_id, event_date, first_nsr, last_nsr, total_records, total_punches,
                    previous_day_last_nsr, next_day_first_nsr, is_nsr_contiguous_before,
                    is_nsr_contiguous_after, integrity_status, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
                ON CONFLICT(device_id, event_date) DO UPDATE SET
                    first_nsr = excluded.first_nsr,
                    last_nsr = excluded.last_nsr,
                    total_records = excluded.total_records,
                    total_punches = excluded.total_punches,
                    previous_day_last_nsr = excluded.previous_day_last_nsr,
                    next_day_first_nsr = excluded.next_day_first_nsr,
                    is_nsr_contiguous_before = excluded.is_nsr_contiguous_before,
                    is_nsr_contiguous_after = excluded.is_nsr_contiguous_after,
                    integrity_status = excluded.integrity_status,
                    updated_at = CURRENT_TIMESTAMP
                """,
                (
                    device_id,
                    event_date,
                    first_nsr,
                    last_nsr,
                    total_records,
                    int(row["total_punches"] or 0),
                    previous_day_last_nsr,
                    next_day_first_nsr,
                    is_before,
                    is_after,
                    integrity_status,
                ),
            )

    def _rebuild_all_open_gaps(self, connection: sqlite3.Connection, device_id: int) -> None:
        rows = connection.execute(
            """
            SELECT nsr
            FROM afd_records
            WHERE device_id = ?
            ORDER BY nsr ASC
            """,
            (device_id,),
        ).fetchall()
        nsrs = [int(row["nsr"]) for row in rows]
        for previous_nsr, next_nsr in zip(nsrs, nsrs[1:]):
            if next_nsr <= previous_nsr + 1:
                continue
            start_nsr = previous_nsr + 1
            end_nsr = next_nsr - 1
            connection.execute(
                """
                INSERT INTO afd_nsr_gaps (
                    device_id, start_nsr, end_nsr, missing_count, gap_status, notes
                ) VALUES (?, ?, ?, ?, ?, ?)
                """,
                (
                    device_id,
                    start_nsr,
                    end_nsr,
                    end_nsr - start_nsr + 1,
                    "OPEN",
                    "Backfill inicial de lacunas a partir da base existente.",
                ),
            )

    def _refresh_nsr_gaps(self, connection: sqlite3.Connection, device_id: int, afd_import_id: int, inserted_nsrs: list[int]) -> None:
        if not inserted_nsrs:
            return

        for range_start, range_end in self._compress_nsr_ranges(sorted(inserted_nsrs)):
            previous = connection.execute(
                """
                SELECT MAX(nsr) AS nsr
                FROM afd_records
                WHERE device_id = ? AND nsr < ?
                """,
                (device_id, range_start),
            ).fetchone()
            next_day = connection.execute(
                """
                SELECT MIN(nsr) AS nsr
                FROM afd_records
                WHERE device_id = ? AND nsr > ?
                """,
                (device_id, range_end),
            ).fetchone()

            window_min = int(previous["nsr"]) + 1 if previous and previous["nsr"] is not None else range_start
            window_max = int(next_day["nsr"]) - 1 if next_day and next_day["nsr"] is not None else range_end
            if window_min > window_max:
                continue

            rows = connection.execute(
                """
                SELECT nsr
                FROM afd_records
                WHERE device_id = ?
                  AND nsr BETWEEN ? AND ?
                ORDER BY nsr ASC
                """,
                (device_id, window_min, window_max),
            ).fetchall()
            present_nsrs = {int(row["nsr"]) for row in rows}

            current_missing: list[tuple[int, int]] = []
            missing_start = None
            for nsr in range(window_min, window_max + 1):
                if nsr not in present_nsrs:
                    if missing_start is None:
                        missing_start = nsr
                    continue
                if missing_start is not None:
                    current_missing.append((missing_start, nsr - 1))
                    missing_start = None
            if missing_start is not None:
                current_missing.append((missing_start, window_max))

            overlapping_open_gaps = connection.execute(
                """
                SELECT id, start_nsr, end_nsr
                FROM afd_nsr_gaps
                WHERE device_id = ?
                  AND resolved = 0
                  AND start_nsr <= ?
                  AND end_nsr >= ?
                """,
                (device_id, window_max, window_min),
            ).fetchall()
            open_gap_keys = {
                (int(row["start_nsr"]), int(row["end_nsr"])): int(row["id"])
                for row in overlapping_open_gaps
            }
            current_missing_keys = set(current_missing)

            for row in overlapping_open_gaps:
                key = (int(row["start_nsr"]), int(row["end_nsr"]))
                if key in current_missing_keys:
                    continue
                connection.execute(
                    """
                    UPDATE afd_nsr_gaps
                    SET resolved = 1,
                        gap_status = 'RESOLVED',
                        resolved_at = CURRENT_TIMESTAMP,
                        resolved_import_id = ?,
                        notes = COALESCE(notes || ' ', '') || ?
                    WHERE id = ?
                    """,
                    (
                        afd_import_id,
                        "Reconciliada automaticamente apos chegada de novos NSRs.",
                        int(row["id"]),
                    ),
                )

            for start_nsr, end_nsr in current_missing:
                if (start_nsr, end_nsr) in open_gap_keys:
                    continue
                connection.execute(
                    """
                    INSERT INTO afd_nsr_gaps (
                        device_id, start_nsr, end_nsr, missing_count, detected_import_id, gap_status, notes
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        device_id,
                        start_nsr,
                        end_nsr,
                        end_nsr - start_nsr + 1,
                        afd_import_id,
                        "OPEN",
                        "Lacuna detectada automaticamente durante importacao.",
                    ),
                )

    def get_or_create_device(self, base_url: str, description: str = "") -> int:
        device_key = self.build_device_key(base_url)
        with self._connect() as connection:
            row = connection.execute(
                "SELECT id FROM devices WHERE device_key = ?",
                (device_key,),
            ).fetchone()
            if row:
                connection.execute(
                    """
                    UPDATE devices
                    SET base_url = ?, description = ?, updated_at = CURRENT_TIMESTAMP
                    WHERE id = ?
                    """,
                    (base_url.rstrip("/"), description, int(row["id"])),
                )
                connection.commit()
                return int(row["id"])

            cursor = connection.execute(
                """
                INSERT INTO devices (device_key, base_url, description)
                VALUES (?, ?, ?)
                """,
                (device_key, base_url.rstrip("/"), description),
            )
            connection.commit()
            return int(cursor.lastrowid)

    def create_import(
        self,
        device_id: int,
        query_mode: str,
        source_type: str,
        raw_afd: str,
        requested_initial_nsr: int | None = None,
        requested_initial_date: str | None = None,
        requested_final_date: str | None = None,
    ) -> int:
        raw_hash = self._hash_text(raw_afd)
        with self._connect() as connection:
            try:
                cursor = connection.execute(
                    """
                    INSERT INTO afd_imports (
                        device_id, query_mode, requested_initial_nsr, requested_initial_date, requested_final_date,
                        source_type, raw_afd, raw_hash, parser_version
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        device_id,
                        query_mode,
                        requested_initial_nsr,
                        requested_initial_date,
                        requested_final_date,
                        source_type,
                        raw_afd,
                        raw_hash,
                        PARSER_VERSION,
                    ),
                )
                connection.commit()
                return int(cursor.lastrowid)
            except sqlite3.IntegrityError:
                row = connection.execute(
                    """
                    SELECT id
                    FROM afd_imports
                    WHERE device_id = ? AND raw_hash = ?
                    """,
                    (device_id, raw_hash),
                ).fetchone()
                if not row:
                    raise
                return int(row["id"])

    def finalize_import(self, afd_import_id: int, records: list[dict]) -> None:
        numeric_nsrs = sorted(
            int(record["nsr"])
            for record in records
            if str(record.get("nsr") or "").isdigit()
        )
        first_nsr = numeric_nsrs[0] if numeric_nsrs else None
        last_nsr = numeric_nsrs[-1] if numeric_nsrs else None
        with self._connect() as connection:
            connection.execute(
                """
                UPDATE afd_imports
                SET first_nsr_found = ?, last_nsr_found = ?, total_lines = ?, total_parsed_lines = ?
                WHERE id = ?
                """,
                (
                    first_nsr,
                    last_nsr,
                    len(records),
                    len(records),
                    afd_import_id,
                ),
            )
            connection.commit()

    def mark_device_sync_success(self, device_id: int) -> None:
        with self._connect() as connection:
            connection.execute(
                """
                UPDATE devices
                SET last_successful_sync = CURRENT_TIMESTAMP,
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                """,
                (device_id,),
            )
            connection.commit()

    def get_max_nsr(self, device_id: int) -> int | None:
        with self._connect() as connection:
            row = connection.execute(
                "SELECT MAX(nsr) AS max_nsr FROM afd_records WHERE device_id = ?",
                (device_id,),
            ).fetchone()
        return int(row["max_nsr"]) if row and row["max_nsr"] is not None else None

    @classmethod
    def _record_row_to_dict(cls, row: sqlite3.Row) -> dict:
        event_timestamp = str(row["event_timestamp"] or "")
        parsed_json = str(row["parsed_json"] or "{}")
        try:
            parsed = json.loads(parsed_json)
        except ValueError:
            parsed = {}
        display_date, display_time = cls._display_date_time_from_timestamp(event_timestamp)
        record_type = str(row["record_type"] or "")
        employee_identifier = str(row["employee_identifier"] or "")
        details = (
            parsed.get("details")
            or parsed.get("event_description")
            or parsed.get("name")
            or parsed.get("receipt_or_hash")
            or parsed.get("parse_status")
            or ""
        )
        return {
            "nsr": str(row["nsr"]),
            "tipo": record_type,
            "data": display_date,
            "hora": display_time,
            "event_timestamp": event_timestamp,
            "employee_identifier": employee_identifier,
            "operation_type": str(row["operation_type"] or ""),
            "event_code": str(row["event_code"] or ""),
            "parsed_json": parsed,
            "pis": parsed.get("pis") or employee_identifier,
            "detalhes": str(details or ""),
            "classe": cls._record_class_from_type(record_type),
            "comprimento": len(str(row["raw_line"] or "")),
            "linha_original": row["raw_line"],
            "source": "db",
        }

    def _event_timestamp_select_sql(self, connection: sqlite3.Connection) -> str:
        if self._column_exists(connection, "afd_records", "event_timestamp"):
            return "event_timestamp"
        return "NULL AS event_timestamp"

    def get_device_capabilities(self, device_id: int) -> dict:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT supports_nsr_query, supports_repcid, supports_date_query
                FROM devices
                WHERE id = ?
                """,
                (device_id,),
            ).fetchone()

        return {
            "supports_nsr_query": int(row["supports_nsr_query"]) if row and row["supports_nsr_query"] is not None else None,
            "supports_repcid": int(row["supports_repcid"]) if row and row["supports_repcid"] is not None else None,
            "supports_date_query": int(row["supports_date_query"]) if row and row["supports_date_query"] is not None else None,
        }

    def update_device_capabilities(
        self,
        device_id: int,
        supports_nsr_query: int | None = None,
        supports_repcid: int | None = None,
        supports_date_query: int | None = None,
    ) -> None:
        updates = []
        params: list[int | None] = []
        if supports_nsr_query is not None:
            updates.append("supports_nsr_query = ?")
            params.append(int(supports_nsr_query))
        if supports_repcid is not None:
            updates.append("supports_repcid = ?")
            params.append(int(supports_repcid))
        if supports_date_query is not None:
            updates.append("supports_date_query = ?")
            params.append(int(supports_date_query))

        if not updates:
            return

        params.append(device_id)
        with self._connect() as connection:
            connection.execute(
                f"""
                UPDATE devices
                SET {", ".join(updates)},
                    updated_at = CURRENT_TIMESTAMP
                WHERE id = ?
                """,
                params,
            )
            connection.commit()

    def get_cached_coverage(self, device_id: int) -> dict:
        with self._connect() as connection:
            coverage_row = connection.execute(
                """
                SELECT MIN(nsr) AS min_nsr, MAX(nsr) AS max_nsr
                FROM afd_records
                WHERE device_id = ?
                """,
                (device_id,),
            ).fetchone()
            gap_row = connection.execute(
                """
                SELECT COUNT(*) AS open_gap_count
                FROM afd_nsr_gaps
                WHERE device_id = ? AND resolved = 0
                """,
                (device_id,),
            ).fetchone()

        return {
            "min_nsr": int(coverage_row["min_nsr"]) if coverage_row and coverage_row["min_nsr"] is not None else None,
            "max_nsr": int(coverage_row["max_nsr"]) if coverage_row and coverage_row["max_nsr"] is not None else None,
            "open_gap_count": int(gap_row["open_gap_count"] or 0) if gap_row else 0,
        }

    def get_records_since(self, device_id: int, initial_nsr: int) -> list[dict]:
        with self._connect() as connection:
            event_timestamp_sql = self._event_timestamp_select_sql(connection)
            rows = connection.execute(
                f"""
                SELECT
                    nsr, record_type, {event_timestamp_sql}, employee_identifier,
                    operation_type, event_code, parsed_json, raw_line
                FROM afd_records
                WHERE device_id = ? AND nsr >= ?
                ORDER BY nsr ASC
                """,
                (device_id, initial_nsr),
            ).fetchall()

        return [self._record_row_to_dict(row) for row in rows]

    def get_records_in_nsr_range(self, device_id: int, start_nsr: int, end_nsr: int) -> list[dict]:
        with self._connect() as connection:
            event_timestamp_sql = self._event_timestamp_select_sql(connection)
            rows = connection.execute(
                f"""
                SELECT
                    nsr, record_type, {event_timestamp_sql}, employee_identifier,
                    operation_type, event_code, parsed_json, raw_line
                FROM afd_records
                WHERE device_id = ?
                  AND nsr BETWEEN ? AND ?
                ORDER BY nsr ASC
                """,
                (device_id, start_nsr, end_nsr),
            ).fetchall()

        return [self._record_row_to_dict(row) for row in rows]

    def get_records_between_dates(self, device_id: int, initial_date: str, final_date: str) -> list[dict]:
        with self._connect() as connection:
            event_timestamp_sql = self._event_timestamp_select_sql(connection)
            rows = connection.execute(
                f"""
                SELECT
                    nsr, record_type, {event_timestamp_sql}, employee_identifier,
                    operation_type, event_code, parsed_json, raw_line
                FROM afd_records
                WHERE device_id = ?
                  AND substr(event_timestamp, 1, 10) BETWEEN ? AND ?
                ORDER BY nsr ASC
                """,
                (
                    device_id,
                    datetime.strptime(initial_date, "%d/%m/%Y").strftime("%Y-%m-%d"),
                    datetime.strptime(final_date, "%d/%m/%Y").strftime("%Y-%m-%d"),
                ),
            ).fetchall()

        return [self._record_row_to_dict(row) for row in rows]

    def get_day_bounds(self, device_id: int, target_date: str) -> dict | None:
        iso_date = self._to_iso_date(target_date)
        if not iso_date:
            return None

        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    event_date,
                    first_nsr,
                    last_nsr,
                    total_records,
                    total_punches,
                    previous_day_last_nsr,
                    next_day_first_nsr,
                    is_nsr_contiguous_before,
                    is_nsr_contiguous_after,
                    integrity_status
                FROM afd_day_summary
                WHERE device_id = ? AND event_date = ?
                """,
                (device_id, iso_date),
            ).fetchone()

        if not row:
            return None

        return {
            "date": self._to_display_date(str(row["event_date"])),
            "first_nsr": int(row["first_nsr"]),
            "last_nsr": int(row["last_nsr"]),
            "total_records": int(row["total_records"]),
            "total_punches": int(row["total_punches"] or 0),
            "previous_day_last_nsr": int(row["previous_day_last_nsr"]) if row["previous_day_last_nsr"] is not None else None,
            "next_day_first_nsr": int(row["next_day_first_nsr"]) if row["next_day_first_nsr"] is not None else None,
            "is_nsr_contiguous_before": bool(row["is_nsr_contiguous_before"]),
            "is_nsr_contiguous_after": bool(row["is_nsr_contiguous_after"]),
            "integrity_status": str(row["integrity_status"]),
        }

    def get_previous_day_with_records(self, device_id: int, target_date: str) -> dict | None:
        iso_date = self._to_iso_date(target_date)
        if not iso_date:
            return None

        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    event_date,
                    first_nsr,
                    last_nsr,
                    total_records,
                    total_punches,
                    integrity_status
                FROM afd_day_summary
                WHERE device_id = ?
                  AND event_date < ?
                ORDER BY event_date DESC
                LIMIT 1
                """,
                (device_id, iso_date),
            ).fetchone()

        if not row:
            return None

        return {
            "date": self._to_display_date(str(row["event_date"])),
            "first_nsr": int(row["first_nsr"]),
            "last_nsr": int(row["last_nsr"]),
            "total_records": int(row["total_records"]),
            "total_punches": int(row["total_punches"] or 0),
            "integrity_status": str(row["integrity_status"]),
        }

    def get_next_day_with_records(self, device_id: int, target_date: str) -> dict | None:
        iso_date = self._to_iso_date(target_date)
        if not iso_date:
            return None

        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT
                    event_date,
                    first_nsr,
                    last_nsr,
                    total_records,
                    total_punches,
                    integrity_status
                FROM afd_day_summary
                WHERE device_id = ?
                  AND event_date > ?
                ORDER BY event_date ASC
                LIMIT 1
                """,
                (device_id, iso_date),
            ).fetchone()

        if not row:
            return None

        return {
            "date": self._to_display_date(str(row["event_date"])),
            "first_nsr": int(row["first_nsr"]),
            "last_nsr": int(row["last_nsr"]),
            "total_records": int(row["total_records"]),
            "total_punches": int(row["total_punches"] or 0),
            "integrity_status": str(row["integrity_status"]),
        }

    @staticmethod
    def _has_open_gap_in_range(connection: sqlite3.Connection, device_id: int, start_nsr: int, end_nsr: int) -> bool:
        if end_nsr < start_nsr:
            return False
        row = connection.execute(
            """
            SELECT 1
            FROM afd_nsr_gaps
            WHERE device_id = ?
              AND resolved = 0
              AND start_nsr <= ?
              AND end_nsr >= ?
            LIMIT 1
            """,
            (device_id, end_nsr, start_nsr),
        ).fetchone()
        return row is not None

    def has_open_gap_in_range(self, device_id: int, start_nsr: int, end_nsr: int) -> bool:
        if end_nsr < start_nsr:
            return False

        with self._connect() as connection:
            return self._has_open_gap_in_range(connection, device_id, start_nsr, end_nsr)

    def get_open_gaps_in_range(self, device_id: int, start_nsr: int, end_nsr: int) -> list[dict]:
        if end_nsr < start_nsr:
            return []

        with self._connect() as connection:
            rows = connection.execute(
                """
                SELECT start_nsr, end_nsr, missing_count, gap_status
                FROM afd_nsr_gaps
                WHERE device_id = ?
                  AND resolved = 0
                  AND start_nsr <= ?
                  AND end_nsr >= ?
                ORDER BY start_nsr ASC
                """,
                (device_id, end_nsr, start_nsr),
            ).fetchall()

        return [
            {
                "start_nsr": int(row["start_nsr"]),
                "end_nsr": int(row["end_nsr"]),
                "missing_count": int(row["missing_count"]),
                "gap_status": str(row["gap_status"]),
            }
            for row in rows
        ]

    def get_first_missing_nsr(self, device_id: int, start_nsr: int, end_nsr: int) -> int | None:
        if end_nsr < start_nsr:
            return None

        coverage = self.get_cached_coverage(device_id)
        min_nsr = coverage["min_nsr"]
        max_nsr = coverage["max_nsr"]

        if min_nsr is None or max_nsr is None:
            return start_nsr

        if start_nsr < min_nsr:
            return start_nsr

        open_gaps = self.get_open_gaps_in_range(device_id, start_nsr, end_nsr)
        if open_gaps:
            return max(start_nsr, int(open_gaps[0]["start_nsr"]))

        if end_nsr > max_nsr:
            return max(start_nsr, max_nsr + 1)

        return None

    def get_nsr_range_stats(self, device_id: int, start_nsr: int, end_nsr: int) -> dict:
        open_gaps = self.get_open_gaps_in_range(device_id, start_nsr, end_nsr)
        records = self.get_records_in_nsr_range(device_id, start_nsr, end_nsr)
        return {
            "total_found": len(records),
            "expected_total": max(0, end_nsr - start_nsr + 1),
            "open_gaps": open_gaps,
            "is_complete": len(open_gaps) == 0 and len(records) == max(0, end_nsr - start_nsr + 1),
        }

    def find_first_day_covering_nsr(self, device_id: int, nsr: int) -> str | None:
        with self._connect() as connection:
            row = connection.execute(
                """
                SELECT event_date
                FROM afd_day_summary
                WHERE device_id = ?
                  AND first_nsr <= ?
                  AND last_nsr >= ?
                ORDER BY event_date ASC
                LIMIT 1
                """,
                (device_id, nsr, nsr),
            ).fetchone()

        if not row:
            return None
        return self._to_display_date(str(row["event_date"]))

    def save_records(self, device_id: int, afd_import_id: int, records: list[dict]) -> dict:
        inserted = 0
        duplicates = 0
        conflicts = 0
        inserted_nsrs: list[int] = []
        affected_dates: set[str] = set()

        with self._connect() as connection:
            for record in records:
                nsr_value = str(record.get("nsr") or "").strip()
                raw_line = str(record.get("linha_original") or "")
                if not nsr_value.isdigit() or not raw_line:
                    continue

                nsr = int(nsr_value)
                raw_line_hash = self._hash_text(raw_line)
                existing = connection.execute(
                    """
                    SELECT id, afd_import_id, raw_line, raw_line_hash, parser_version
                    FROM afd_records
                    WHERE device_id = ? AND nsr = ?
                    """,
                    (device_id, nsr),
                ).fetchone()

                if existing is None:
                    storage = self._record_storage_params(record, raw_line_hash, PARSER_VERSION)
                    connection.execute(
                        """
                        INSERT INTO afd_records (
                            device_id, afd_import_id, nsr, record_type, event_timestamp,
                            employee_identifier, operation_type, event_code, parsed_json,
                            raw_line, raw_line_hash, parser_version
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """,
                        (
                            device_id,
                            afd_import_id,
                            nsr,
                            *storage,
                        ),
                    )
                    inserted += 1
                    inserted_nsrs.append(nsr)
                    iso_date = self._event_date_from_timestamp(str(record.get("event_timestamp") or ""))
                    if iso_date:
                        affected_dates.add(iso_date)
                    continue

                if existing["raw_line_hash"] == raw_line_hash and existing["raw_line"] == raw_line:
                    if str(existing["parser_version"] or "") != PARSER_VERSION:
                        storage = self._record_storage_params(record, raw_line_hash, PARSER_VERSION)
                        connection.execute(
                            """
                            UPDATE afd_records
                            SET record_type = ?,
                                event_timestamp = ?,
                                employee_identifier = ?,
                                operation_type = ?,
                                event_code = ?,
                                parsed_json = ?,
                                raw_line = ?,
                                raw_line_hash = ?,
                                parser_version = ?,
                                updated_at = CURRENT_TIMESTAMP
                            WHERE id = ?
                            """,
                            (*storage, int(existing["id"])),
                        )
                        iso_date = self._event_date_from_timestamp(str(record.get("event_timestamp") or ""))
                        if iso_date:
                            affected_dates.add(iso_date)
                    duplicates += 1
                    continue

                connection.execute(
                    """
                    INSERT INTO afd_record_conflicts (
                        device_id, afd_import_id, nsr, existing_record_id, existing_import_id, existing_raw_line,
                        incoming_raw_line, existing_hash, incoming_hash, conflict_type
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """,
                    (
                        device_id,
                        afd_import_id,
                        nsr,
                        int(existing["id"]),
                        int(existing["afd_import_id"]) if existing["afd_import_id"] is not None else None,
                        str(existing["raw_line"]),
                        raw_line,
                        str(existing["raw_line_hash"]),
                        raw_line_hash,
                        "same_nsr_different_line",
                    ),
                )
                conflicts += 1

            if inserted_nsrs:
                self._refresh_nsr_gaps(connection, device_id, afd_import_id, inserted_nsrs)
                if not affected_dates:
                    affected_dates = {
                        str(row["event_date"])
                        for row in connection.execute(
                            """
                            SELECT DISTINCT substr(event_timestamp, 1, 10) AS event_date
                            FROM afd_records
                            WHERE device_id = ?
                              AND event_timestamp IS NOT NULL
                              AND event_timestamp <> ''
                            ORDER BY event_date
                            """,
                            (device_id,),
                        ).fetchall()
                    }
            if inserted_nsrs or affected_dates:
                self._refresh_day_summaries(connection, device_id, affected_dates)
            connection.commit()

        return {
            "inserted": inserted,
            "duplicates": duplicates,
            "conflicts": conflicts,
        }


def _date_from_ddmmyyyy(value: str) -> str:
    normalized = str(value or "").strip()
    if len(normalized) != 8 or not normalized.isdigit():
        return ""
    try:
        return datetime.strptime(normalized, "%d%m%Y").strftime("%Y-%m-%d")
    except ValueError:
        return ""


def _time_from_hhmm(value: str) -> str:
    normalized = str(value or "").strip()
    if len(normalized) != 4 or not normalized.isdigit():
        return ""
    return f"{normalized[0:2]}:{normalized[2:4]}"


def _event_timestamp_from_parts(raw_date: str, raw_time: str) -> str:
    iso_date = _date_from_ddmmyyyy(raw_date)
    display_time = _time_from_hhmm(raw_time)
    if not iso_date or not display_time:
        return ""
    return f"{iso_date} {display_time}:00"


def _display_date_from_timestamp(event_timestamp: str) -> str:
    display_date, _ = AFDRepository._display_date_time_from_timestamp(event_timestamp)
    return display_date


def _display_time_from_timestamp(event_timestamp: str) -> str:
    _, display_time = AFDRepository._display_date_time_from_timestamp(event_timestamp)
    return display_time


def _base_parsed_record(line: str, record_type: str, event_timestamp: str = "") -> dict:
    return {
        "nsr": line[0:9].strip() if len(line) >= 9 else "",
        "tipo": record_type,
        "data": _display_date_from_timestamp(event_timestamp),
        "hora": _display_time_from_timestamp(event_timestamp),
        "event_timestamp": event_timestamp,
        "employee_identifier": "",
        "operation_type": "",
        "event_code": "",
        "pis": "",
        "detalhes": "",
        "classe": AFDRepository._record_class_from_type(record_type),
        "comprimento": len(line),
        "linha_original": line,
        "parsed_json": {
            "record_type": record_type,
            "parse_status": "parsed",
            "raw_line_length": len(line),
        },
    }


def parse_type_1(line: str) -> dict:
    event_timestamp = _event_timestamp_from_parts(line[204:212], line[228:232]) if len(line) >= 232 else ""
    result = _base_parsed_record(line, "1", event_timestamp)
    result["detalhes"] = line[37:187].strip() if len(line) >= 187 else ""
    result["parsed_json"].update(
        {
            "generated_at": event_timestamp,
            "details": result["detalhes"],
        }
    )
    return result


def parse_type_2(line: str) -> dict:
    event_timestamp = _event_timestamp_from_parts(line[10:18], line[18:22]) if len(line) >= 22 else ""
    result = _base_parsed_record(line, "2", event_timestamp)
    result["detalhes"] = line[49:199].strip() if len(line) >= 199 else ""
    result["parsed_json"].update(
        {
            "timestamp": event_timestamp,
            "details": result["detalhes"],
        }
    )
    return result


def parse_type_3(line: str) -> dict:
    event_timestamp = _event_timestamp_from_parts(line[10:18], line[18:22]) if len(line) >= 22 else ""
    result = _base_parsed_record(line, "3", event_timestamp)
    employee_identifier = line[22:34].strip() if len(line) >= 34 else ""
    receipt_or_hash = line[34:].strip() if len(line) > 34 else ""
    result["employee_identifier"] = employee_identifier
    result["pis"] = employee_identifier
    result["detalhes"] = receipt_or_hash
    result["parsed_json"].update(
        {
            "timestamp": event_timestamp.replace(" ", "T") if event_timestamp else "",
            "employee_identifier_type": "pis",
            "employee_identifier": employee_identifier,
            "pis": employee_identifier,
            "receipt_or_hash": receipt_or_hash,
        }
    )
    return result


def parse_type_4(line: str) -> dict:
    event_timestamp = _event_timestamp_from_parts(line[10:18], line[18:22]) if len(line) >= 22 else ""
    adjusted_timestamp = _event_timestamp_from_parts(line[22:30], line[30:34]) if len(line) >= 34 else ""
    result = _base_parsed_record(line, "4", event_timestamp)
    result["detalhes"] = f"ajustado para {adjusted_timestamp}".strip()
    result["parsed_json"].update(
        {
            "timestamp": event_timestamp,
            "adjusted_timestamp": adjusted_timestamp,
            "details": result["detalhes"],
        }
    )
    return result


def parse_type_5(line: str) -> dict:
    event_timestamp = _event_timestamp_from_parts(line[10:18], line[18:22]) if len(line) >= 22 else ""
    operation = line[22:23].strip() if len(line) >= 23 else ""
    employee_identifier = line[23:35].strip() if len(line) >= 35 else ""
    employee_name = line[35:87].strip() if len(line) >= 87 else ""
    result = _base_parsed_record(line, "5", event_timestamp)
    result["employee_identifier"] = employee_identifier
    result["operation_type"] = operation
    result["pis"] = employee_identifier
    result["detalhes"] = f"{operation} {employee_name}".strip()
    result["parsed_json"].update(
        {
            "timestamp": event_timestamp,
            "operation": operation,
            "employee_identifier_type": "pis",
            "employee_identifier": employee_identifier,
            "pis": employee_identifier,
            "name": employee_name,
        }
    )
    return result


def parse_type_6(line: str) -> dict:
    event_timestamp = _event_timestamp_from_parts(line[10:18], line[18:22]) if len(line) >= 22 else ""
    event_code = line[22:].strip() if len(line) > 22 else ""
    result = _base_parsed_record(line, "6", event_timestamp)
    result["event_code"] = event_code
    result["detalhes"] = event_code
    result["parsed_json"].update(
        {
            "timestamp": event_timestamp,
            "event_code": event_code,
            "event_description": f"evento_tipo_6_{event_code}" if event_code else "",
        }
    )
    return result


def parse_type_7(line: str) -> dict:
    result = _base_parsed_record(line, "7", "")
    result["parsed_json"].update(
        {
            "parse_status": "unsupported",
            "reason": "tipo_7_layout_not_implemented",
        }
    )
    return result


def parse_type_9(line: str) -> dict:
    result = _base_parsed_record(line, "9", "")
    totals = {
        "tipo2": line[9:18].strip() if len(line) >= 18 else "",
        "tipo3": line[18:27].strip() if len(line) >= 27 else "",
        "tipo4": line[27:36].strip() if len(line) >= 36 else "",
        "tipo5": line[36:45].strip() if len(line) >= 45 else "",
    }
    result["detalhes"] = " ".join(f"{key}={value}" for key, value in totals.items() if value)
    result["parsed_json"].update({"totals": totals})
    return result


def parse_unsupported_type(line: str, record_type: str) -> dict:
    result = _base_parsed_record(line, record_type, "")
    result["parsed_json"].update(
        {
            "parse_status": "unsupported",
            "reason": "record_type_not_implemented",
        }
    )
    return result


def parse_afd_line(line: str) -> dict | None:
    line = line.rstrip()
    if not line:
        return None

    record_type = line[9:10].strip() if len(line) >= 10 else ""
    parsers = {
        "1": parse_type_1,
        "2": parse_type_2,
        "3": parse_type_3,
        "4": parse_type_4,
        "5": parse_type_5,
        "6": parse_type_6,
        "7": parse_type_7,
        "9": parse_type_9,
    }
    parser = parsers.get(record_type)
    if parser:
        return parser(line)
    return parse_unsupported_type(line, record_type)


def parse_afd_text(afd_text: str, source: str) -> list[dict]:
    records = []
    for raw_line in afd_text.splitlines():
        parsed = parse_afd_line(raw_line)
        if parsed:
            parsed["source"] = source
            records.append(parsed)
    return records


def summarize_records(records: list[dict]) -> dict:
    summary = {
        "batida_ponto": 0,
        "cadastro_empregado": 0,
        "identificacao_empregador": 0,
        "ajuste_relogio": 0,
        "evento_tipo_6": 0,
        "cabecalho_afd": 0,
        "trailer_afd": 0,
        "outro_evento": 0,
    }

    for record in records:
        record_class = record.get("classe", "outro_evento")
        summary[record_class] = summary.get(record_class, 0) + 1

    return summary


def records_to_raw_text(records: list[dict]) -> str:
    return "\n".join(
        str(record.get("linha_original") or "")
        for record in records
        if str(record.get("linha_original") or "").strip()
    )


def sort_records(records: list[dict]) -> list[dict]:
    def sort_key(record: dict):
        nsr = str(record.get("nsr") or "")
        return int(nsr) if nsr.isdigit() else 10**18

    return sorted(records, key=sort_key)


def merge_records_by_nsr(*record_sets: list[dict]) -> list[dict]:
    merged: dict[str, dict] = {}
    for record_set in record_sets:
        for record in record_set:
            nsr = str(record.get("nsr") or "").strip()
            if not nsr.isdigit():
                continue
            # Prefer the most recent payload for the same NSR, which lets live
            # reads correct stale or partial rows already persisted in SQLite.
            merged[nsr] = record
    return sort_records(list(merged.values()))


def expand_date_window(initial_date: str, final_date: str | None = None, days: int = 1) -> tuple[str, str]:
    start_dt = datetime.strptime(initial_date, "%d/%m/%Y")
    end_dt = datetime.strptime(final_date or initial_date, "%d/%m/%Y")
    expanded_start = start_dt - timedelta(days=days)
    expanded_end = end_dt + timedelta(days=days)
    return expanded_start.strftime("%d/%m/%Y"), expanded_end.strftime("%d/%m/%Y")


def iterate_dates(initial_date: str, final_date: str) -> list[str]:
    start_dt = datetime.strptime(initial_date, "%d/%m/%Y")
    end_dt = datetime.strptime(final_date, "%d/%m/%Y")
    current = start_dt
    values = []
    while current <= end_dt:
        values.append(current.strftime("%d/%m/%Y"))
        current += timedelta(days=1)
    return values


def is_day_closed(repository: AFDRepository, device_id: int, target_date: str) -> tuple[bool, dict]:
    current = repository.get_day_bounds(device_id, target_date)
    previous = repository.get_previous_day_with_records(device_id, target_date)
    next_day = repository.get_next_day_with_records(device_id, target_date)
    target_dt = datetime.strptime(target_date, "%d/%m/%Y").date()
    today_dt = datetime.now().date()

    if current:
        backward_closed = bool(
            previous
            and not repository.has_open_gap_in_range(device_id, previous["last_nsr"] + 1, current["first_nsr"] - 1)
        )
        forward_closed = bool(
            next_day
            and not repository.has_open_gap_in_range(device_id, current["last_nsr"] + 1, next_day["first_nsr"] - 1)
        )
        completed = bool(backward_closed and forward_closed)
        has_open_gap = repository.has_open_gap_in_range(
            device_id=device_id,
            start_nsr=current["first_nsr"],
            end_nsr=current["last_nsr"],
        )
        integrity_status = current["integrity_status"]
    else:
        bridged_empty_day = bool(
            previous
            and next_day
            and not repository.has_open_gap_in_range(device_id, previous["last_nsr"] + 1, next_day["first_nsr"] - 1)
        )
        backward_closed = bridged_empty_day
        forward_closed = bridged_empty_day
        completed = bridged_empty_day
        has_open_gap = False
        integrity_status = "COMPLETE" if completed else "EMPTY"

    is_past_day = target_dt < today_dt
    closed = bool(is_past_day and completed and not has_open_gap)

    return closed, {
        "date": target_date,
        "has_records": bool(current),
        "backward_closed": backward_closed,
        "forward_closed": forward_closed,
        "completed": completed,
        "integrity_status": integrity_status,
        "previous_record_date": previous["date"] if previous else None,
        "next_record_date": next_day["date"] if next_day else None,
        "has_open_gap": has_open_gap,
        "is_past_day": is_past_day,
        "is_today": target_dt == today_dt,
        "closed": closed,
        "should_sync": not closed,
    }


def get_day_continuity_status(repository: AFDRepository, device_id: int, target_date: str) -> dict:
    _, status = is_day_closed(repository, device_id, target_date)
    return status


def determine_sync_window_for_dates(
    repository: AFDRepository,
    device_id: int,
    initial_date: str,
    final_date: str,
) -> tuple[str | None, str | None, list[dict]]:
    statuses = [get_day_continuity_status(repository, device_id, day) for day in iterate_dates(initial_date, final_date)]
    open_days = [status["date"] for status in statuses if status["should_sync"]]

    if not open_days:
        return None, None, statuses

    sync_initial, sync_final = expand_date_window(open_days[0], open_days[-1], days=1)
    return sync_initial, sync_final, statuses


def build_day_nsr_audit(repository: AFDRepository, device_id: int, initial_date: str, final_date: str) -> list[dict]:
    audit = []
    for target_date in iterate_dates(initial_date, final_date):
        current = repository.get_day_bounds(device_id, target_date)
        previous = repository.get_previous_day_with_records(device_id, target_date)
        next_day = repository.get_next_day_with_records(device_id, target_date)

        if current:
            backward_ok = bool(
                previous
                and not repository.has_open_gap_in_range(device_id, previous["last_nsr"] + 1, current["first_nsr"] - 1)
            )
            forward_ok = bool(
                next_day
                and not repository.has_open_gap_in_range(device_id, current["last_nsr"] + 1, next_day["first_nsr"] - 1)
            )
            completed = bool(backward_ok and forward_ok)
        else:
            bridged_empty_day = bool(
                previous
                and next_day
                and not repository.has_open_gap_in_range(device_id, previous["last_nsr"] + 1, next_day["first_nsr"] - 1)
            )
            backward_ok = bridged_empty_day
            forward_ok = bridged_empty_day
            completed = bridged_empty_day

        audit.append(
            {
                "date": target_date,
                "has_records": bool(current),
                "current_first_nsr": current["first_nsr"] if current else None,
                "current_last_nsr": current["last_nsr"] if current else None,
                "total_records": current["total_records"] if current else 0,
                "total_punches": current["total_punches"] if current else 0,
                "previous_date": previous["date"] if previous else None,
                "previous_last_nsr": previous["last_nsr"] if previous else None,
                "next_date": next_day["date"] if next_day else None,
                "next_first_nsr": next_day["first_nsr"] if next_day else None,
                "backward_ok": backward_ok,
                "forward_ok": forward_ok,
                "completed": completed,
                "integrity_status": current["integrity_status"] if current else ("COMPLETE" if completed else "EMPTY"),
            }
        )

    return audit

def build_live_nsr_payload(
    repository: AFDRepository,
    device_id: int,
    base_url: str,
    username: str,
    password: str,
    start_nsr: int,
    allow_empty: bool,
) -> tuple[str, str]:
    capabilities = repository.get_device_capabilities(device_id)
    client = ControlIDClient(base_url=base_url, username=username, password=password)
    if capabilities["supports_nsr_query"] != 0:
        client.login()
        try:
            payload = client.export_afd(initial_nsr=start_nsr, mode=595)
            repository.update_device_capabilities(device_id=device_id, supports_nsr_query=1)
            return payload, "api_export_afd"
        except requests.HTTPError as exc:
            response_data = None
            if exc.response is not None:
                try:
                    response_data = exc.response.json()
                except ValueError:
                    response_data = exc.response.text

            if isinstance(response_data, dict) and response_data.get("error") == "Invalid command: export_afd":
                repository.update_device_capabilities(device_id=device_id, supports_nsr_query=0)
            else:
                raise

    if capabilities["supports_repcid"] != 0:
        repcid_client = RepCidClient(host=base_url, username=username, password=password)
        try:
            payload = repcid_client.get_afd_from_nsr(start_nsr=start_nsr)
            repository.update_device_capabilities(device_id=device_id, supports_repcid=1)
            return payload, "repcid"
        except RuntimeError:
            repository.update_device_capabilities(device_id=device_id, supports_repcid=0)
            if allow_empty:
                return "", "repcid"
            raise

    if allow_empty:
        return "", "unsupported_nsr_sync"
    raise RuntimeError("Este dispositivo nao possui mecanismo disponivel para sincronizacao direta por NSR.")


def build_sync_state(
    *,
    data_source: str | None,
    equipment_consulted: bool,
    live_records: int,
    inserted: int,
) -> str:
    if data_source == "partial_cache":
        return "partial_cache"

    if not equipment_consulted:
        return "cache"

    if inserted > 0:
        return "synced_with_import"

    if live_records > 0:
        return "synced_no_changes"

    return "consulted_equipment"


def import_afd_payload(
    repository: AFDRepository,
    device_id: int,
    query_mode: str,
    source_type: str,
    raw_text: str,
    requested_initial_nsr: int | None = None,
    requested_initial_date: str | None = None,
    requested_final_date: str | None = None,
) -> tuple[list[dict], dict]:
    records = parse_afd_text(raw_text, source="live")
    afd_import_id = repository.create_import(
        device_id=device_id,
        query_mode=query_mode,
        source_type=source_type,
        raw_afd=raw_text,
        requested_initial_nsr=requested_initial_nsr,
        requested_initial_date=requested_initial_date,
        requested_final_date=requested_final_date,
    )
    repository.finalize_import(afd_import_id, records)
    stats = repository.save_records(device_id=device_id, afd_import_id=afd_import_id, records=records)
    return records, stats


def sync_nsr_range_by_date(
    repository: AFDRepository,
    device_id: int,
    base_url: str,
    username: str,
    password: str,
    start_nsr: int,
    end_nsr: int,
) -> dict:
    capabilities = repository.get_device_capabilities(device_id)
    if capabilities["supports_date_query"] == 0:
        raise RuntimeError("Este dispositivo esta marcado como sem suporte para sincronizacao por data.")

    coverage = repository.get_cached_coverage(device_id)
    anchor_nsr = repository.get_first_missing_nsr(device_id, start_nsr, end_nsr)
    if anchor_nsr is None:
        return {
            "live_records": [],
            "import_stats": {"inserted": 0, "duplicates": 0, "conflicts": 0},
            "synced_initial_date": None,
            "synced_final_date": None,
            "sync_strategy": None,
        }

    anchor_date = repository.find_first_day_covering_nsr(device_id, anchor_nsr)
    if anchor_date is None and coverage["max_nsr"] is not None:
        anchor_date = repository.find_first_day_covering_nsr(device_id, coverage["max_nsr"])

    if anchor_date is None:
        raise RuntimeError(
            "Nao foi possivel mapear o primeiro NSR faltante para uma data. Sincronize por data antes de consultar esta faixa por NSR."
        )

    client = ControlIDClient(base_url=base_url, username=username, password=password)
    client.login()

    merged_records: list[dict] = []
    import_stats = {"inserted": 0, "duplicates": 0, "conflicts": 0}
    synced_initial_date: str | None = None
    synced_final_date: str | None = None

    for target_date in iterate_dates(anchor_date, datetime.now().strftime("%d/%m/%Y")):
        raw_text = client.get_afd(initial_date=target_date, final_date=target_date)
        repository.update_device_capabilities(device_id=device_id, supports_date_query=1)
        repository.mark_device_sync_success(device_id)
        live_records, stats = import_afd_payload(
            repository=repository,
            device_id=device_id,
            query_mode="nsr_range",
            source_type="api_get_afd",
            raw_text=raw_text,
            requested_initial_date=target_date,
            requested_final_date=target_date,
        )
        merged_records.extend(live_records)
        import_stats["inserted"] += stats["inserted"]
        import_stats["duplicates"] += stats["duplicates"]
        import_stats["conflicts"] += stats["conflicts"]
        synced_initial_date = synced_initial_date or target_date
        synced_final_date = target_date

        current_missing = repository.get_first_missing_nsr(device_id, start_nsr, end_nsr)
        if current_missing is None:
            break

    return {
        "live_records": merged_records,
        "import_stats": import_stats,
        "synced_initial_date": synced_initial_date,
        "synced_final_date": synced_final_date,
        "sync_strategy": "date_backfill",
    }


def get_nsr_range(
    repository: AFDRepository,
    device_id: int,
    base_url: str,
    username: str,
    password: str,
    start_nsr: int,
    end_nsr: int,
) -> dict:
    if end_nsr < start_nsr:
        raise RuntimeError("O NSR final deve ser maior ou igual ao NSR inicial.")

    coverage = repository.get_cached_coverage(device_id)
    open_gaps = repository.get_open_gaps_in_range(device_id, start_nsr, end_nsr)
    in_coverage = (
        coverage["min_nsr"] is not None
        and coverage["max_nsr"] is not None
        and start_nsr >= coverage["min_nsr"]
        and end_nsr <= coverage["max_nsr"]
    )

    if in_coverage and not open_gaps:
        cached_records = repository.get_records_in_nsr_range(device_id, start_nsr, end_nsr)
        if len(cached_records) == (end_nsr - start_nsr + 1):
            return {
                "records": cached_records,
                "raw": records_to_raw_text(cached_records),
                "storage": {
                    "db_records": len(cached_records),
                    "live_records": 0,
                    "synced_from_nsr": None,
                    "synced_initial_date": None,
                    "synced_final_date": None,
                    "source_type": "db_only",
                    "data_source": "cache",
                    "sync_attempted": False,
                    "sync_strategy": None,
                    "range_complete": True,
                    "range_start_nsr": start_nsr,
                    "range_end_nsr": end_nsr,
                    "missing_ranges": [],
                    "coverage": coverage,
                    "inserted": 0,
                    "duplicates": 0,
                    "conflicts": 0,
                    "equipment_consulted": False,
                    "sync_state": "cache",
                },
            }

    missing_nsr = repository.get_first_missing_nsr(device_id, start_nsr, end_nsr)
    if missing_nsr is None:
        missing_nsr = start_nsr

    live_records: list[dict] = []
    import_stats = {"inserted": 0, "duplicates": 0, "conflicts": 0}
    synced_from_nsr: int | None = None
    synced_initial_date: str | None = None
    synced_final_date: str | None = None
    sync_strategy: str | None = None

    capabilities = repository.get_device_capabilities(device_id)
    if capabilities["supports_nsr_query"] != 0 or capabilities["supports_repcid"] != 0:
        live_text, source_type = build_live_nsr_payload(
            repository=repository,
            device_id=device_id,
            base_url=base_url,
            username=username,
            password=password,
            start_nsr=missing_nsr,
            allow_empty=True,
        )
        sync_strategy = "nsr_api" if source_type == "api_export_afd" else ("repcid" if source_type == "repcid" else None)
        synced_from_nsr = missing_nsr if sync_strategy else None
        if live_text:
            repository.mark_device_sync_success(device_id)
            live_records, import_stats = import_afd_payload(
                repository=repository,
                device_id=device_id,
                query_mode="nsr_range",
                source_type=source_type,
                raw_text=live_text,
                requested_initial_nsr=missing_nsr,
            )

    post_sync_missing = repository.get_first_missing_nsr(device_id, start_nsr, end_nsr)
    if post_sync_missing is not None:
        date_sync = sync_nsr_range_by_date(
            repository=repository,
            device_id=device_id,
            base_url=base_url,
            username=username,
            password=password,
            start_nsr=start_nsr,
            end_nsr=end_nsr,
        )
        live_records = merge_records_by_nsr(live_records, date_sync["live_records"])
        import_stats["inserted"] += date_sync["import_stats"]["inserted"]
        import_stats["duplicates"] += date_sync["import_stats"]["duplicates"]
        import_stats["conflicts"] += date_sync["import_stats"]["conflicts"]
        synced_initial_date = date_sync["synced_initial_date"]
        synced_final_date = date_sync["synced_final_date"]
        if date_sync["sync_strategy"]:
            sync_strategy = date_sync["sync_strategy"]

    records = repository.get_records_in_nsr_range(device_id, start_nsr, end_nsr)
    stats = repository.get_nsr_range_stats(device_id, start_nsr, end_nsr)

    return {
        "records": records,
        "raw": records_to_raw_text(records),
        "storage": {
            "db_records": len(records),
            "live_records": len(live_records),
            "synced_from_nsr": synced_from_nsr,
            "synced_initial_date": synced_initial_date,
            "synced_final_date": synced_final_date,
            "source_type": "cache_after_sync" if sync_strategy else "partial_cache",
            "data_source": "cache_after_sync" if sync_strategy else "partial_cache",
            "sync_attempted": bool(sync_strategy),
            "sync_strategy": sync_strategy,
            "range_complete": stats["is_complete"],
            "range_start_nsr": start_nsr,
            "range_end_nsr": end_nsr,
            "missing_ranges": stats["open_gaps"],
            "coverage": repository.get_cached_coverage(device_id),
            "inserted": import_stats["inserted"],
            "duplicates": import_stats["duplicates"],
            "conflicts": import_stats["conflicts"],
            "equipment_consulted": bool(sync_strategy),
            "sync_state": build_sync_state(
                data_source="cache_after_sync" if sync_strategy else "partial_cache",
                equipment_consulted=bool(sync_strategy),
                live_records=len(live_records),
                inserted=import_stats["inserted"],
            ),
        },
    }


def create_app() -> Flask:
    app = Flask(__name__)
    app.config["AFD_DB_PATH"] = os.getenv("AFD_DB_PATH", DEFAULT_DB_PATH)
    app.register_blueprint(event)

    @app.get("/")
    def index():
        controlid_config = get_controlid_config()
        return render_template(
            "index.html",
            equipment_base_url=controlid_config["base_url"],
            equipment_configured=controlid_config["configured"],
            repcid_available=RepCidClient.is_available(),
        )

    @app.post("/api/afd")
    def afd():
        payload = request.get_json(silent=True) or {}

        controlid_config = get_controlid_config()
        base_url = controlid_config["base_url"]
        username = controlid_config["username"]
        password = controlid_config["password"]
        initial_date = (payload.get("initial_date") or "").strip()
        final_date = (payload.get("final_date") or "").strip()
        initial_nsr = str(payload.get("initial_nsr") or "").strip()
        final_nsr = str(payload.get("final_nsr") or "").strip()
        pis = str(payload.get("pis") or "").strip()
        query_mode = (payload.get("query_mode") or "date").strip().lower()

        if pis and not pis.isdigit():
            return jsonify({"error": "Informe apenas numeros no PIS."}), 400

        if not controlid_config["configured"]:
            return jsonify(
                {
                    "error": (
                        "Equipamento nao configurado no servidor. "
                        "Defina CONTROLID_BASE_URL, CONTROLID_USERNAME e CONTROLID_PASSWORD."
                    )
                }
            ), 500

        try:
            repository = AFDRepository(app.config["AFD_DB_PATH"])
            device_id = repository.get_or_create_device(base_url=base_url)

            records = []
            raw_text = ""
            import_stats = {"inserted": 0, "duplicates": 0, "conflicts": 0}
            cache_meta = {
                "db_records": 0,
                "live_records": 0,
                "synced_from_nsr": None,
                "synced_initial_date": None,
                "synced_final_date": None,
                "date_days_closed": 0,
                "date_days_open": 0,
                "day_nsr_audit": [],
                "source_type": None,
                "data_source": None,
                "sync_attempted": False,
                "equipment_consulted": False,
                "sync_state": "cache",
                "sync_strategy": None,
                "range_complete": None,
                "range_start_nsr": None,
                "range_end_nsr": None,
                "missing_ranges": [],
                "coverage": None,
                "db_path": app.config["AFD_DB_PATH"],
            }

            if query_mode == "nsr":
                if not initial_nsr.isdigit():
                    return jsonify({"error": "Informe um NSR inicial numerico."}), 400

                requested_nsr = int(initial_nsr)
                if final_nsr:
                    if not final_nsr.isdigit():
                        return jsonify({"error": "Informe um NSR final numerico."}), 400

                    range_result = get_nsr_range(
                        repository=repository,
                        device_id=device_id,
                        base_url=base_url,
                        username=username,
                        password=password,
                        start_nsr=requested_nsr,
                        end_nsr=int(final_nsr),
                    )
                    records = range_result["records"]
                    raw_text = range_result["raw"]
                    cache_meta.update(range_result["storage"])
                    import_stats = {
                        "inserted": cache_meta["inserted"],
                        "duplicates": cache_meta["duplicates"],
                        "conflicts": cache_meta["conflicts"],
                    }
                else:
                    cached_records = repository.get_records_since(device_id=device_id, initial_nsr=requested_nsr)
                    cache_meta["db_records"] = len(cached_records)

                    max_cached_nsr = repository.get_max_nsr(device_id=device_id)
                    next_nsr_to_fetch = requested_nsr if max_cached_nsr is None else max(requested_nsr, max_cached_nsr + 1)

                    live_text, source_type = build_live_nsr_payload(
                        repository=repository,
                        device_id=device_id,
                        base_url=base_url,
                        username=username,
                        password=password,
                        start_nsr=next_nsr_to_fetch,
                        allow_empty=bool(cached_records),
                    )
                    cache_meta["source_type"] = source_type
                    cache_meta["equipment_consulted"] = source_type in {"api_export_afd", "repcid"}
                    cache_meta["synced_from_nsr"] = next_nsr_to_fetch if live_text else None
                    cache_meta["data_source"] = "cache_after_sync" if live_text else "cache"
                    cache_meta["sync_attempted"] = cache_meta["equipment_consulted"]
                    cache_meta["sync_strategy"] = "nsr_api" if source_type == "api_export_afd" else ("repcid" if source_type == "repcid" else None)
                    if live_text:
                        repository.mark_device_sync_success(device_id)

                    live_records = parse_afd_text(live_text, source="live")
                    cache_meta["live_records"] = len(live_records)

                    if live_text or not cached_records:
                        afd_import_id = repository.create_import(
                            device_id=device_id,
                            query_mode="nsr",
                            source_type=source_type,
                            raw_afd=live_text,
                            requested_initial_nsr=next_nsr_to_fetch,
                        )
                        repository.finalize_import(afd_import_id, live_records)
                        import_stats = repository.save_records(device_id=device_id, afd_import_id=afd_import_id, records=live_records)

                    records = sort_records(cached_records + live_records)
                    raw_text = records_to_raw_text(records)
            else:
                if not initial_date:
                    return jsonify({"error": "Informe a data inicial."}), 400

                try:
                    datetime.strptime(initial_date, "%d/%m/%Y")
                    if final_date:
                        datetime.strptime(final_date, "%d/%m/%Y")
                except ValueError:
                    return jsonify({"error": "Use o formato de data dd/mm/aaaa."}), 400

                requested_final_date = final_date or initial_date
                sync_initial_date, sync_final_date, day_statuses = determine_sync_window_for_dates(
                    repository=repository,
                    device_id=device_id,
                    initial_date=initial_date,
                    final_date=requested_final_date,
                )
                cache_meta["date_days_closed"] = sum(1 for status in day_statuses if status["closed"])
                cache_meta["date_days_open"] = sum(1 for status in day_statuses if status["should_sync"])

                live_records = []
                if sync_initial_date and sync_final_date:
                    cache_meta["synced_initial_date"] = sync_initial_date
                    cache_meta["synced_final_date"] = sync_final_date

                    client = ControlIDClient(base_url=base_url, username=username, password=password)
                    client.login()
                    raw_text = client.get_afd(initial_date=sync_initial_date, final_date=sync_final_date)
                    live_records = parse_afd_text(raw_text, source="live")
                    cache_meta["live_records"] = len(live_records)
                    cache_meta["source_type"] = "api_get_afd"
                    cache_meta["data_source"] = "cache_after_sync"
                    cache_meta["sync_attempted"] = True
                    cache_meta["equipment_consulted"] = True
                    cache_meta["sync_strategy"] = "date_window"
                    repository.update_device_capabilities(device_id=device_id, supports_date_query=1)
                    repository.mark_device_sync_success(device_id)

                    afd_import_id = repository.create_import(
                        device_id=device_id,
                        query_mode="date",
                        source_type="api_get_afd",
                        raw_afd=raw_text,
                        requested_initial_date=sync_initial_date,
                        requested_final_date=sync_final_date,
                    )
                    repository.finalize_import(afd_import_id, live_records)
                    import_stats = repository.save_records(device_id=device_id, afd_import_id=afd_import_id, records=live_records)
                else:
                    cache_meta["source_type"] = "db_only"
                    cache_meta["data_source"] = "cache"
                    cache_meta["sync_attempted"] = False
                    cache_meta["equipment_consulted"] = False
                    cache_meta["sync_strategy"] = None
                    raw_text = ""

                records = repository.get_records_between_dates(
                    device_id=device_id,
                    initial_date=initial_date,
                    final_date=requested_final_date,
                )
                cache_meta["db_records"] = len(records)
                cache_meta["day_nsr_audit"] = build_day_nsr_audit(
                    repository=repository,
                    device_id=device_id,
                    initial_date=initial_date,
                    final_date=requested_final_date,
                )

                live_nsrs = {
                    str(record.get("nsr"))
                    for record in live_records
                    if str(record.get("classe") or "") == "batida_ponto" and str(record.get("nsr") or "").isdigit()
                }
                for record in records:
                    if str(record.get("nsr") or "") in live_nsrs:
                        record["source"] = "live"

            filtered_records = [
                record for record in records
                if record.get("classe") == "batida_ponto"
                and (
                    not pis
                    or str(record.get("pis") or "").strip().startswith(pis)
                    or str(record.get("pis") or "").strip().endswith(pis)
                )
            ]
            summary = summarize_records(records)
            cache_meta["sync_state"] = build_sync_state(
                data_source=cache_meta["data_source"],
                equipment_consulted=cache_meta["equipment_consulted"],
                live_records=cache_meta["live_records"],
                inserted=import_stats["inserted"],
            )
        except requests.HTTPError as exc:
            return jsonify({"error": f"Erro HTTP no equipamento: {exc}"}), 502
        except requests.RequestException as exc:
            return jsonify({"error": f"Falha de comunicacao com o equipamento: {exc}"}), 502
        except Exception as exc:
            return jsonify({"error": f"Erro ao consultar AFD: {exc}"}), 500

        return jsonify(
            {
                "records": records,
                "filtered_records": filtered_records,
                "total": len(records),
                "total_filtered": len(filtered_records),
                "summary": summary,
                "raw": raw_text,
                "storage": {
                    "db_path": app.config["AFD_DB_PATH"],
                    "db_records": cache_meta["db_records"],
                    "live_records": cache_meta["live_records"],
                    "synced_from_nsr": cache_meta["synced_from_nsr"],
                    "synced_initial_date": cache_meta["synced_initial_date"],
                    "synced_final_date": cache_meta["synced_final_date"],
                    "date_days_closed": cache_meta["date_days_closed"],
                    "date_days_open": cache_meta["date_days_open"],
                    "day_nsr_audit": cache_meta["day_nsr_audit"],
                    "source_type": cache_meta["source_type"],
                    "data_source": cache_meta["data_source"],
                    "sync_attempted": cache_meta["sync_attempted"],
                    "equipment_consulted": cache_meta["equipment_consulted"],
                    "sync_state": cache_meta["sync_state"],
                    "sync_strategy": cache_meta["sync_strategy"],
                    "range_complete": cache_meta["range_complete"],
                    "range_start_nsr": cache_meta["range_start_nsr"],
                    "range_end_nsr": cache_meta["range_end_nsr"],
                    "missing_ranges": cache_meta["missing_ranges"],
                    "coverage": cache_meta["coverage"],
                    "inserted": import_stats["inserted"],
                    "duplicates": import_stats["duplicates"],
                    "conflicts": import_stats["conflicts"],
                },
            }
        )

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
