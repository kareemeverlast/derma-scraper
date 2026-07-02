"""
storage.py — SQLite checkpointing layer for Derma Scope.

Every clinic and email result is written to disk the instant it's found, so a
crash, timeout, dropped connection, or a closed browser tab never loses
already-scraped (and already-paid-for) data. A run can always be resumed:
completed (city, keyword) query pairs and already-checked website emails are
skipped, so resuming never re-queries Google or re-probes a site for no
reason.

This app may be used by multiple employees against a shared, hosted
instance at once, so every write opens its own short-lived connection
(WAL journal mode + a busy timeout) rather than holding one connection open
across a whole scrape.

Set DERMA_DB_PATH to point this at persistent storage in a hosted deployment
— on ephemeral container storage this file (and all in-progress runs) will
be lost on restart.
"""

import json
import os
import sqlite3
from datetime import datetime, timezone

DB_PATH = os.environ.get(
    "DERMA_DB_PATH",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "scrape_runs.db"),
)

# Columns persisted per clinic row, beyond the internal (run_id, place_id) key.
_CLINIC_FIELDS = [
    "name", "address", "city", "country", "website", "phone",
    "email", "rating", "review_count", "scraped_at",
]


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def get_conn() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=30)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=30000")
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                label TEXT NOT NULL,
                owner TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'running',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL,
                params_json TEXT NOT NULL,
                total_queries INTEGER NOT NULL,
                api_calls INTEGER NOT NULL DEFAULT 0
            );

            CREATE TABLE IF NOT EXISTS clinics (
                run_id INTEGER NOT NULL,
                place_id TEXT NOT NULL,
                name TEXT, address TEXT, city TEXT, country TEXT,
                website TEXT, phone TEXT, email TEXT DEFAULT '',
                email_checked INTEGER NOT NULL DEFAULT 0,
                rating REAL, review_count INTEGER, scraped_at TEXT,
                PRIMARY KEY (run_id, place_id)
            );

            CREATE TABLE IF NOT EXISTS completed_queries (
                run_id INTEGER NOT NULL,
                city TEXT NOT NULL,
                keyword TEXT NOT NULL,
                completed_at TEXT NOT NULL,
                PRIMARY KEY (run_id, city, keyword)
            );

            CREATE TABLE IF NOT EXISTS run_errors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id INTEGER NOT NULL,
                message TEXT NOT NULL,
                created_at TEXT NOT NULL
            );
            """
        )


def create_run(owner: str, label: str, params: dict, total_queries: int) -> int:
    now = _now()
    with get_conn() as conn:
        cur = conn.execute(
            "INSERT INTO runs (label, owner, status, created_at, updated_at, "
            "params_json, total_queries) VALUES (?, ?, 'running', ?, ?, ?, ?)",
            (label, owner, now, now, json.dumps(params), total_queries),
        )
        return cur.lastrowid


def save_clinics_batch(run_id: int, rows: list) -> None:
    if not rows:
        return
    with get_conn() as conn:
        conn.executemany(
            f"""INSERT OR IGNORE INTO clinics
                (run_id, place_id, {', '.join(_CLINIC_FIELDS)})
                VALUES (?, ?, {', '.join('?' for _ in _CLINIC_FIELDS)})""",
            [
                (run_id, row["place_id"], *(row.get(f) for f in _CLINIC_FIELDS))
                for row in rows
            ],
        )


def mark_query_complete(run_id: int, city: str, keyword: str) -> None:
    with get_conn() as conn:
        conn.execute(
            "INSERT OR IGNORE INTO completed_queries (run_id, city, keyword, completed_at) "
            "VALUES (?, ?, ?, ?)",
            (run_id, city, keyword, _now()),
        )


def get_completed_queries(run_id: int) -> set:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT city, keyword FROM completed_queries WHERE run_id = ?", (run_id,)
        ).fetchall()
    return {(r["city"], r["keyword"]) for r in rows}


def get_existing_place_ids(run_id: int) -> set:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT place_id FROM clinics WHERE run_id = ?", (run_id,)
        ).fetchall()
    return {r["place_id"] for r in rows}


def get_pending_email_rows(run_id: int) -> list:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT place_id, website FROM clinics "
            "WHERE run_id = ? AND website != '' AND email_checked = 0",
            (run_id,),
        ).fetchall()
    return [dict(r) for r in rows]


def update_email_result(run_id: int, place_id: str, email: str) -> None:
    with get_conn() as conn:
        conn.execute(
            "UPDATE clinics SET email = ?, email_checked = 1 "
            "WHERE run_id = ? AND place_id = ?",
            (email, run_id, place_id),
        )


def log_error(run_id: int, message: str) -> None:
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO run_errors (run_id, message, created_at) VALUES (?, ?, ?)",
            (run_id, message, _now()),
        )


def touch_run(run_id: int, api_calls_delta: int = 0) -> None:
    with get_conn() as conn:
        conn.execute(
            "UPDATE runs SET updated_at = ?, api_calls = api_calls + ? WHERE id = ?",
            (_now(), api_calls_delta, run_id),
        )


def mark_run_completed(run_id: int) -> None:
    with get_conn() as conn:
        conn.execute(
            "UPDATE runs SET status = 'completed', updated_at = ? WHERE id = ?",
            (_now(), run_id),
        )


def cancel_run(run_id: int) -> None:
    """Drop a run out of the resumable list without deleting its data."""
    with get_conn() as conn:
        conn.execute(
            "UPDATE runs SET status = 'cancelled', updated_at = ? WHERE id = ?",
            (_now(), run_id),
        )


def list_resumable_runs() -> list:
    with get_conn() as conn:
        rows = conn.execute(
            """
            SELECT r.id, r.label, r.owner, r.updated_at, r.total_queries,
                   (SELECT COUNT(*) FROM completed_queries cq WHERE cq.run_id = r.id) AS queries_done,
                   (SELECT COUNT(*) FROM clinics c WHERE c.run_id = r.id) AS clinics_found,
                   (SELECT COUNT(*) FROM clinics c WHERE c.run_id = r.id AND c.website != '' AND c.email_checked = 1) AS emails_done,
                   (SELECT COUNT(*) FROM clinics c WHERE c.run_id = r.id AND c.website != '' AND c.email_checked = 0) AS emails_pending
            FROM runs r
            WHERE r.status = 'running'
            ORDER BY r.updated_at DESC
            """
        ).fetchall()
    return [dict(r) for r in rows]


def get_run(run_id: int) -> dict:
    with get_conn() as conn:
        row = conn.execute("SELECT * FROM runs WHERE id = ?", (run_id,)).fetchone()
    if row is None:
        return None
    run = dict(row)
    run["params"] = json.loads(run["params_json"])
    return run


def get_run_clinics(run_id: int) -> list:
    with get_conn() as conn:
        rows = conn.execute(
            f"SELECT {', '.join(_CLINIC_FIELDS)} FROM clinics WHERE run_id = ?",
            (run_id,),
        ).fetchall()
    return [dict(r) for r in rows]


def get_run_errors(run_id: int) -> list:
    with get_conn() as conn:
        rows = conn.execute(
            "SELECT message FROM run_errors WHERE run_id = ? ORDER BY id", (run_id,)
        ).fetchall()
    return [r["message"] for r in rows]
