# =============================================================================
#  22-17 LABEL FINANCE BOT — Bothost edition
#  Управление финансами лейбла: артисты, доходы/расходы, Excel-импорт,
#  месячные отчёты, выплаты и поиск аномалий через Claude.
#
#  Интерфейс: ТОЛЬКО inline-кнопки, без slash-команд.
#
#  Сборка под Bothost:
#   * точка входа — main.py, зависимости — requirements.txt;
#   * БД по умолчанию SQLite в /app/data (единственная папка, которая
#     переживает передеплой). Если задан DATABASE_URL с postgres:// —
#     автоматически используется PostgreSQL (psycopg2), код общий;
#   * все секреты — из переменных окружения панели Bothost;
#   * long polling, один процесс, без вебхуков и открытых портов.
# =============================================================================
from __future__ import annotations

import asyncio
import json
import logging
import os
import re
import sqlite3
import statistics
import threading
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from pathlib import Path
from typing import Any, Iterator, Sequence

from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import ChatAction, ParseMode
from telegram.error import BadRequest
from telegram.ext import (
    ApplicationBuilder,
    CallbackQueryHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

try:
    from telegram.ext import AIORateLimiter
except ImportError:
    AIORateLimiter = None

# -----------------------------------------------------------------------------
# Конфигурация и логирование
# -----------------------------------------------------------------------------
logging.basicConfig(
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    level=os.getenv("LOG_LEVEL", "INFO"),
)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("finbot")


def _env_decimal(name: str, default: str) -> Decimal:
    try:
        return Decimal(os.getenv(name, default).replace(",", "."))
    except InvalidOperation:
        return Decimal(default)


class Config:
    BOT_TOKEN: str = os.getenv("BOT_TOKEN", "").strip()
    ANTHROPIC_API_KEY: str = os.getenv("ANTHROPIC_API_KEY", "").strip()
    # Имя модели вынесено в переменную окружения, чтобы не хардкодить версию.
    CLAUDE_MODEL: str = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-5")
    CLAUDE_MAX_TOKENS: int = int(os.getenv("CLAUDE_MAX_TOKENS", "4096"))

    # Пусто => SQLite в DATA_DIR. postgres://... => PostgreSQL.
    DATABASE_URL: str = os.getenv("DATABASE_URL", "").strip()
    DATA_DIR: str = os.getenv("DATA_DIR", "/app/data")
    SQLITE_NAME: str = os.getenv("SQLITE_NAME", "finbot.db")
    DB_POOL_MIN: int = int(os.getenv("DB_POOL_MIN", "1"))
    DB_POOL_MAX: int = int(os.getenv("DB_POOL_MAX", "5"))

    ALLOWED_USER_IDS: set[int] = {
        int(x)
        for x in os.getenv("ALLOWED_USER_IDS", "").replace(" ", "").split(",")
        if x.strip().lstrip("-").isdigit()
    }
    DEFAULT_CURRENCY: str = os.getenv("DEFAULT_CURRENCY", "RUB")
    DEFAULT_RATE: Decimal = _env_decimal("DEFAULT_RATE", "0.20")
    MAX_EXCEL_BYTES: int = int(os.getenv("MAX_EXCEL_MB", "10")) * 1024 * 1024
    MAX_EXCEL_ROWS: int = int(os.getenv("MAX_EXCEL_ROWS", "500"))
    ANOMALY_SIGMA: Decimal = _env_decimal("ANOMALY_SIGMA", "2")
    ANOMALY_MONTHS: int = int(os.getenv("ANOMALY_MONTHS", "6"))
    LABEL_NAME: str = os.getenv("LABEL_NAME", "22-17")


# -----------------------------------------------------------------------------
# Деньги: храним в копейках (целые числа), считаем в Decimal
# -----------------------------------------------------------------------------
CENTS = Decimal("0.01")


def to_cents(value: Any) -> int:
    return int(
        (Decimal(str(value)).quantize(CENTS, rounding=ROUND_HALF_UP) * 100)
        .to_integral_value(rounding=ROUND_HALF_UP)
    )


def from_cents(value: Any) -> Decimal:
    return (Decimal(int(value or 0)) / 100).quantize(CENTS)


def as_date(value: Any) -> date:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()


def json_dumps(obj: Any) -> str:
    def default(o: Any) -> Any:
        if isinstance(o, Decimal):
            return float(o)
        if isinstance(o, (date, datetime)):
            return o.isoformat()
        return str(o)

    return json.dumps(obj, ensure_ascii=False, default=default)


# -----------------------------------------------------------------------------
# Схемы БД
# -----------------------------------------------------------------------------
SCHEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS artists (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    name            TEXT NOT NULL UNIQUE,
    rate            REAL NOT NULL DEFAULT 0.2,
    tg_username     TEXT,
    is_active       INTEGER NOT NULL DEFAULT 1,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS transactions (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    artist_id       INTEGER REFERENCES artists(id) ON DELETE SET NULL,
    kind            TEXT NOT NULL CHECK (kind IN ('income', 'expense')),
    amount_cents    INTEGER NOT NULL CHECK (amount_cents > 0),
    currency        TEXT NOT NULL DEFAULT 'RUB',
    category        TEXT,
    description     TEXT,
    occurred_on     TEXT NOT NULL,
    source          TEXT NOT NULL DEFAULT 'manual',
    external_key    TEXT,
    created_by      INTEGER,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE UNIQUE INDEX IF NOT EXISTS transactions_external_key_uniq
    ON transactions (external_key) WHERE external_key IS NOT NULL;
CREATE INDEX IF NOT EXISTS transactions_occurred_idx ON transactions (occurred_on);
CREATE INDEX IF NOT EXISTS transactions_artist_idx ON transactions (artist_id);

CREATE TABLE IF NOT EXISTS payments (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    artist_id       INTEGER NOT NULL REFERENCES artists(id) ON DELETE CASCADE,
    period_start    TEXT NOT NULL,
    period_end      TEXT NOT NULL,
    revenue_cents   INTEGER NOT NULL DEFAULT 0,
    expenses_cents  INTEGER NOT NULL DEFAULT 0,
    rate            REAL NOT NULL,
    amount_cents    INTEGER NOT NULL,
    status          TEXT NOT NULL DEFAULT 'pending'
                    CHECK (status IN ('pending', 'paid', 'canceled')),
    paid_at         TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    UNIQUE (artist_id, period_start, period_end)
);

CREATE TABLE IF NOT EXISTS ai_reports (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    kind            TEXT NOT NULL,
    period_start    TEXT,
    period_end      TEXT,
    payload         TEXT,
    body            TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);

CREATE TABLE IF NOT EXISTS audit_log (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    tg_user_id      INTEGER,
    action          TEXT NOT NULL,
    details         TEXT,
    created_at      TEXT NOT NULL DEFAULT (datetime('now'))
);
"""

SCHEMA_POSTGRES = """
CREATE TABLE IF NOT EXISTS artists (
    id              SERIAL PRIMARY KEY,
    name            TEXT NOT NULL UNIQUE,
    rate            DOUBLE PRECISION NOT NULL DEFAULT 0.2,
    tg_username     TEXT,
    is_active       SMALLINT NOT NULL DEFAULT 1,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS transactions (
    id              BIGSERIAL PRIMARY KEY,
    artist_id       INTEGER REFERENCES artists(id) ON DELETE SET NULL,
    kind            TEXT NOT NULL CHECK (kind IN ('income', 'expense')),
    amount_cents    BIGINT NOT NULL CHECK (amount_cents > 0),
    currency        TEXT NOT NULL DEFAULT 'RUB',
    category        TEXT,
    description     TEXT,
    occurred_on     DATE NOT NULL,
    source          TEXT NOT NULL DEFAULT 'manual',
    external_key    TEXT,
    created_by      BIGINT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE UNIQUE INDEX IF NOT EXISTS transactions_external_key_uniq
    ON transactions (external_key) WHERE external_key IS NOT NULL;
CREATE INDEX IF NOT EXISTS transactions_occurred_idx ON transactions (occurred_on);
CREATE INDEX IF NOT EXISTS transactions_artist_idx ON transactions (artist_id);

CREATE TABLE IF NOT EXISTS payments (
    id              BIGSERIAL PRIMARY KEY,
    artist_id       INTEGER NOT NULL REFERENCES artists(id) ON DELETE CASCADE,
    period_start    DATE NOT NULL,
    period_end      DATE NOT NULL,
    revenue_cents   BIGINT NOT NULL DEFAULT 0,
    expenses_cents  BIGINT NOT NULL DEFAULT 0,
    rate            DOUBLE PRECISION NOT NULL,
    amount_cents    BIGINT NOT NULL,
    status          TEXT NOT NULL DEFAULT 'pending'
                    CHECK (status IN ('pending', 'paid', 'canceled')),
    paid_at         TIMESTAMPTZ,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    UNIQUE (artist_id, period_start, period_end)
);

CREATE TABLE IF NOT EXISTS ai_reports (
    id              BIGSERIAL PRIMARY KEY,
    kind            TEXT NOT NULL,
    period_start    DATE,
    period_end      DATE,
    payload         TEXT,
    body            TEXT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS audit_log (
    id              BIGSERIAL PRIMARY KEY,
    tg_user_id      BIGINT,
    action          TEXT NOT NULL,
    details         TEXT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);
"""


class Database:
    """Единый слой доступа: SQLite (по умолчанию) или PostgreSQL.

    SQL пишется с плейсхолдерами `?`; для PostgreSQL они заменяются на `%s`.
    Все вызовы синхронные и идут из хендлеров через asyncio.to_thread.
    """

    def __init__(self) -> None:
        url = Config.DATABASE_URL
        self.is_postgres = url.startswith(("postgres://", "postgresql://"))

        if self.is_postgres:
            import psycopg2
            import psycopg2.extras
            from psycopg2.pool import ThreadedConnectionPool

            self._extras = psycopg2.extras
            self._pool = ThreadedConnectionPool(
                Config.DB_POOL_MIN, Config.DB_POOL_MAX, dsn=url
            )
            log.info("DB backend: PostgreSQL")
        else:
            directory = Path(Config.DATA_DIR)
            try:
                directory.mkdir(parents=True, exist_ok=True)
                probe = directory / ".write_test"
                probe.write_text("ok", encoding="utf-8")
                probe.unlink()
            except OSError:
                directory = Path("./data")
                directory.mkdir(parents=True, exist_ok=True)
                log.warning(
                    "%s недоступна для записи, использую %s "
                    "(данные не переживут передеплой!)",
                    Config.DATA_DIR,
                    directory.resolve(),
                )
            self.sqlite_path = str(directory / Config.SQLITE_NAME)
            self._lock = threading.RLock()
            self._conn = sqlite3.connect(
                self.sqlite_path, check_same_thread=False, timeout=30
            )
            self._conn.row_factory = sqlite3.Row
            self._conn.execute("PRAGMA journal_mode=WAL;")
            self._conn.execute("PRAGMA foreign_keys=ON;")
            log.info("DB backend: SQLite (%s)", self.sqlite_path)

    # ---------- утилиты диалекта ----------
    def _sql(self, sql: str) -> str:
        return sql.replace("?", "%s") if self.is_postgres else sql

    def d(self, value: date) -> Any:
        return value if self.is_postgres else value.isoformat()

    def now(self) -> Any:
        return (
            datetime.now()
            if self.is_postgres
            else datetime.now().isoformat(" ", "seconds")
        )

    @contextmanager
    def _pg_cursor(self, commit: bool) -> Iterator[Any]:
        conn = self._pool.getconn()
        try:
            with conn:
                with conn.cursor(cursor_factory=self._extras.RealDictCursor) as cur:
                    yield cur
                if not commit:
                    conn.rollback()
        except Exception:
            conn.rollback()
            raise
        finally:
            self._pool.putconn(conn)

    # ---------- операции ----------
    def query(self, sql: str, params: Sequence[Any] = ()) -> list[dict]:
        if self.is_postgres:
            with self._pg_cursor(commit=False) as cur:
                cur.execute(self._sql(sql), tuple(params))
                return [dict(r) for r in cur.fetchall()]
        with self._lock:
            cur = self._conn.execute(sql, tuple(params))
            return [dict(r) for r in cur.fetchall()]

    def scalar(self, sql: str, params: Sequence[Any] = ()) -> Any:
        rows = self.query(sql, params)
        if not rows:
            return None
        return next(iter(rows[0].values()))

    def execute(self, sql: str, params: Sequence[Any] = ()) -> int:
        if self.is_postgres:
            with self._pg_cursor(commit=True) as cur:
                cur.execute(self._sql(sql), tuple(params))
                return cur.rowcount
        with self._lock:
            cur = self._conn.execute(sql, tuple(params))
            self._conn.commit()
            return cur.rowcount

    def executemany(self, sql: str, rows: Sequence[Sequence[Any]]) -> int:
        if not rows:
            return 0
        if self.is_postgres:
            with self._pg_cursor(commit=True) as cur:
                self._extras.execute_batch(
                    cur, self._sql(sql), [tuple(r) for r in rows], page_size=200
                )
                return len(rows)
        with self._lock:
            self._conn.executemany(sql, [tuple(r) for r in rows])
            self._conn.commit()
            return len(rows)

    def init_schema(self) -> None:
        schema = SCHEMA_POSTGRES if self.is_postgres else SCHEMA_SQLITE
        if self.is_postgres:
            with self._pg_cursor(commit=True) as cur:
                cur.execute(schema)
        else:
            with self._lock:
                self._conn.executescript(schema)
                self._conn.commit()
        log.info("DB schema ready")

    def close(self) -> None:
        if self.is_postgres:
            self._pool.closeall()
        else:
            with self._lock:
                self._conn.close()


# -----------------------------------------------------------------------------
# Репозиторий
# -----------------------------------------------------------------------------
NO_ARTIST = "— без артиста —"
NO_CATEGORY = "без категории"


class Repo:
    def __init__(self, db: Database) -> None:
        self.db = db

    # ---------- артисты ----------
    def add_artist(self, name: str, rate: Decimal, username: str | None = None) -> dict:
        self.db.execute(
            """
            INSERT INTO artists (name, rate, tg_username, is_active)
            VALUES (?, ?, ?, 1)
            ON CONFLICT (name) DO UPDATE SET
                rate = excluded.rate,
                tg_username = COALESCE(excluded.tg_username, artists.tg_username),
                is_active = 1
            """,
            (name, float(rate), username),
        )
        rows = self.db.query(
            "SELECT id, name, rate, tg_username FROM artists WHERE name = ?", (name,)
        )
        return rows[0]

    def list_artists(self, only_active: bool = True) -> list[dict]:
        sql = "SELECT id, name, rate, tg_username, is_active FROM artists"
        if only_active:
            sql += " WHERE is_active = 1"
        return self.db.query(sql + " ORDER BY name")

    def get_artist(self, artist_id: int) -> dict | None:
        rows = self.db.query("SELECT * FROM artists WHERE id = ?", (artist_id,))
        return rows[0] if rows else None

    def resolve_artist_ids(self, names: list[str]) -> dict[str, int]:
        if not names:
            return {}
        placeholders = ", ".join(["?"] * len(names))
        rows = self.db.query(
            f"SELECT id, name FROM artists WHERE lower(name) IN ({placeholders})",
            [n.lower() for n in names],
        )
        return {str(r["name"]).lower(): int(r["id"]) for r in rows}

    # ---------- транзакции ----------
    INSERT_TX = """
        INSERT INTO transactions
            (artist_id, kind, amount_cents, currency, category, description,
             occurred_on, source, external_key, created_by)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT DO NOTHING
    """

    def add_transaction(
        self,
        artist_id: int | None,
        kind: str,
        amount: Decimal,
        currency: str,
        category: str | None,
        description: str | None,
        occurred_on: date,
        source: str = "manual",
        external_key: str | None = None,
        created_by: int | None = None,
    ) -> int:
        return self.db.execute(
            self.INSERT_TX,
            (
                artist_id,
                kind,
                to_cents(amount),
                currency,
                category,
                description,
                self.db.d(occurred_on),
                source,
                external_key,
                created_by,
            ),
        )

    def bulk_insert_transactions(self, rows: list[dict], created_by: int | None) -> int:
        if not rows:
            return 0
        before = int(self.db.scalar("SELECT COUNT(*) AS c FROM transactions") or 0)
        payload = [
            (
                r.get("artist_id"),
                r["kind"],
                to_cents(r["amount"]),
                (r.get("currency") or Config.DEFAULT_CURRENCY),
                r.get("category"),
                r.get("description"),
                self.db.d(r["occurred_on"]),
                "excel",
                r.get("external_key"),
                created_by,
            )
            for r in rows
        ]
        self.db.executemany(self.INSERT_TX, payload)
        after = int(self.db.scalar("SELECT COUNT(*) AS c FROM transactions") or 0)
        return max(after - before, 0)

    # ---------- аналитика ----------
    @staticmethod
    def month_bounds(year: int, month: int) -> tuple[date, date]:
        start = date(year, month, 1)
        end = date(year + (month == 12), (month % 12) + 1, 1)
        return start, end

    def monthly_breakdown(self, year: int, month: int) -> list[dict]:
        start, end = self.month_bounds(year, month)
        rows = self.db.query(
            f"""
            SELECT COALESCE(a.name, '{NO_ARTIST}') AS artist,
                   COALESCE(a.rate, 0) AS rate,
                   SUM(CASE WHEN t.kind = 'income'  THEN t.amount_cents ELSE 0 END) AS revenue_cents,
                   SUM(CASE WHEN t.kind = 'expense' THEN t.amount_cents ELSE 0 END) AS expenses_cents,
                   COUNT(*) AS tx_count
            FROM transactions t
            LEFT JOIN artists a ON a.id = t.artist_id
            WHERE t.occurred_on >= ? AND t.occurred_on < ?
            GROUP BY COALESCE(a.name, '{NO_ARTIST}'), COALESCE(a.rate, 0)
            ORDER BY 3 DESC
            """,
            (self.db.d(start), self.db.d(end)),
        )
        result: list[dict] = []
        for r in rows:
            revenue = from_cents(r["revenue_cents"])
            expenses = from_cents(r["expenses_cents"])
            rate = Decimal(str(r["rate"] or 0))
            gross = revenue - expenses
            # Прибыль артиста = (Выручка - Расходы) * Ставка
            share = (gross * rate).quantize(CENTS, rounding=ROUND_HALF_UP)
            result.append(
                {
                    "artist": r["artist"],
                    "rate": rate,
                    "revenue": revenue,
                    "expenses": expenses,
                    "gross": gross,
                    "artist_share": share,
                    "label_profit": (gross - share).quantize(CENTS),
                    "tx_count": int(r["tx_count"]),
                }
            )
        return result

    def category_breakdown(self, year: int, month: int) -> list[dict]:
        start, end = self.month_bounds(year, month)
        rows = self.db.query(
            f"""
            SELECT kind,
                   COALESCE(category, '{NO_CATEGORY}') AS category,
                   SUM(amount_cents) AS total_cents,
                   COUNT(*) AS cnt
            FROM transactions
            WHERE occurred_on >= ? AND occurred_on < ?
            GROUP BY kind, COALESCE(category, '{NO_CATEGORY}')
            ORDER BY 3 DESC
            """,
            (self.db.d(start), self.db.d(end)),
        )
        return [
            {
                "kind": r["kind"],
                "category": r["category"],
                "total": from_cents(r["total_cents"]),
                "count": int(r["cnt"]),
            }
            for r in rows
        ]

    def anomaly_candidates(self, months: int | None = None) -> dict[str, Any]:
        """Выбросы, дубли и тренды считаются в Python — одинаково для
        SQLite и PostgreSQL, без диалектных функций."""
        months = months or Config.ANOMALY_MONTHS
        since = date.today() - timedelta(days=31 * months)
        rows = self.db.query(
            f"""
            SELECT t.id,
                   COALESCE(a.name, '{NO_ARTIST}') AS artist,
                   t.kind, t.amount_cents, t.currency,
                   COALESCE(t.category, '{NO_CATEGORY}') AS category,
                   t.description, t.occurred_on, t.artist_id
            FROM transactions t
            LEFT JOIN artists a ON a.id = t.artist_id
            WHERE t.occurred_on >= ?
            ORDER BY t.occurred_on
            """,
            (self.db.d(since),),
        )

        groups: dict[tuple[str, str], list[dict]] = {}
        monthly: dict[str, dict[str, Any]] = {}
        dupe_counter: dict[tuple[Any, str, int, str], int] = {}
        orphans = 0

        for r in rows:
            occurred = as_date(r["occurred_on"])
            groups.setdefault((r["kind"], r["category"]), []).append(r)

            ym = occurred.strftime("%Y-%m")
            bucket = monthly.setdefault(
                ym,
                {"ym": ym, "revenue": Decimal(0), "expenses": Decimal(0), "tx_count": 0},
            )
            amount = from_cents(r["amount_cents"])
            if r["kind"] == "income":
                bucket["revenue"] += amount
            else:
                bucket["expenses"] += amount
            bucket["tx_count"] += 1

            key = (
                r["artist_id"],
                r["kind"],
                int(r["amount_cents"]),
                occurred.isoformat(),
            )
            dupe_counter[key] = dupe_counter.get(key, 0) + 1

            if r["artist_id"] is None:
                orphans += 1

        sigma = float(Config.ANOMALY_SIGMA)
        outliers: list[dict] = []
        for (kind, category), items in groups.items():
            if len(items) < 3:
                continue
            amounts = [int(i["amount_cents"]) for i in items]
            mean = statistics.fmean(amounts)
            sd = statistics.pstdev(amounts)
            if sd <= 0:
                continue
            for item in items:
                deviation = abs(int(item["amount_cents"]) - mean)
                if deviation > sigma * sd:
                    outliers.append(
                        {
                            "id": item["id"],
                            "artist": item["artist"],
                            "kind": kind,
                            "category": category,
                            "amount": from_cents(item["amount_cents"]),
                            "currency": item["currency"],
                            "description": item["description"],
                            "occurred_on": as_date(item["occurred_on"]),
                            "avg_amount": from_cents(round(mean)),
                            "sd_amount": from_cents(round(sd)),
                            "deviation_sigmas": round(deviation / sd, 2),
                        }
                    )
        outliers.sort(key=lambda x: x["deviation_sigmas"], reverse=True)

        duplicates = [
            {
                "artist_id": k[0],
                "kind": k[1],
                "amount": from_cents(k[2]),
                "occurred_on": k[3],
                "count": v,
            }
            for k, v in sorted(dupe_counter.items(), key=lambda kv: kv[1], reverse=True)
            if v > 1
        ][:20]

        return {
            "outliers": outliers[:40],
            "duplicates": duplicates,
            "monthly": [monthly[k] for k in sorted(monthly)],
            "integrity": {"orphan_tx": orphans, "total_tx": len(rows)},
            "sigma": str(Config.ANOMALY_SIGMA),
            "period_from": since,
        }

    # ---------- выплаты ----------
    def upsert_payments_for_month(self, year: int, month: int) -> list[dict]:
        start, end = self.month_bounds(year, month)
        agg_rows = self.db.query(
            """
            SELECT artist_id,
                   SUM(CASE WHEN kind = 'income'  THEN amount_cents ELSE 0 END) AS revenue_cents,
                   SUM(CASE WHEN kind = 'expense' THEN amount_cents ELSE 0 END) AS expenses_cents
            FROM transactions
            WHERE occurred_on >= ? AND occurred_on < ? AND artist_id IS NOT NULL
            GROUP BY artist_id
            """,
            (self.db.d(start), self.db.d(end)),
        )
        agg = {int(r["artist_id"]): r for r in agg_rows}
        period_end = end - timedelta(days=1)
        created: list[dict] = []

        for artist in self.list_artists():
            data = agg.get(int(artist["id"]), {})
            revenue_cents = int(data.get("revenue_cents") or 0)
            expenses_cents = int(data.get("expenses_cents") or 0)
            rate = Decimal(str(artist["rate"] or 0))
            gross = max(
                from_cents(revenue_cents) - from_cents(expenses_cents), Decimal(0)
            )
            amount_cents = to_cents(
                (gross * rate).quantize(CENTS, rounding=ROUND_HALF_UP)
            )
            self.db.execute(
                """
                INSERT INTO payments
                    (artist_id, period_start, period_end, revenue_cents,
                     expenses_cents, rate, amount_cents, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'pending')
                ON CONFLICT (artist_id, period_start, period_end) DO UPDATE SET
                    revenue_cents = excluded.revenue_cents,
                    expenses_cents = excluded.expenses_cents,
                    rate = excluded.rate,
                    amount_cents = excluded.amount_cents
                WHERE payments.status = 'pending'
                """,
                (
                    artist["id"],
                    self.db.d(start),
                    self.db.d(period_end),
                    revenue_cents,
                    expenses_cents,
                    float(rate),
                    amount_cents,
                ),
            )
            created.append(
                {"artist": artist["name"], "amount": from_cents(amount_cents)}
            )
        return created

    def list_payments(self, limit: int = 20, status: str | None = None) -> list[dict]:
        sql = """
            SELECT p.id, a.name AS artist, p.period_start, p.period_end,
                   p.revenue_cents, p.expenses_cents, p.rate, p.amount_cents,
                   p.status, p.paid_at
            FROM payments p
            JOIN artists a ON a.id = p.artist_id
        """
        params: list[Any] = []
        if status:
            sql += " WHERE p.status = ?"
            params.append(status)
        sql += " ORDER BY p.period_start DESC, p.amount_cents DESC LIMIT ?"
        params.append(limit)

        rows = self.db.query(sql, params)
        return [
            {
                "id": r["id"],
                "artist": r["artist"],
                "period_start": as_date(r["period_start"]),
                "period_end": as_date(r["period_end"]),
                "revenue": from_cents(r["revenue_cents"]),
                "expenses": from_cents(r["expenses_cents"]),
                "rate": Decimal(str(r["rate"] or 0)),
                "amount": from_cents(r["amount_cents"]),
                "status": r["status"],
            }
            for r in rows
        ]

    def mark_payment_paid(self, payment_id: int) -> None:
        self.db.execute(
            "UPDATE payments SET status = 'paid', paid_at = ? WHERE id = ?",
            (self.db.now(), payment_id),
        )

    # ---------- служебное ----------
    def save_report(
        self,
        kind: str,
        body: str,
        payload: dict,
        period: tuple[date, date] | None = None,
    ) -> None:
        self.db.execute(
            """
            INSERT INTO ai_reports (kind, period_start, period_end, payload, body)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                kind,
                self.db.d(period[0]) if period else None,
                self.db.d(period[1]) if period else None,
                json_dumps(payload),
                body,
            ),
        )

    def audit(self, user_id: int | None, action: str, details: dict) -> None:
        self.db.execute(
            "INSERT INTO audit_log (tg_user_id, action, details) VALUES (?, ?, ?)",
            (user_id, action, json_dumps(details)),
        )


# -----------------------------------------------------------------------------
# Claude
# -----------------------------------------------------------------------------
SYSTEM_ANALYST = (
    f"Ты — финансовый аналитик музыкального лейбла {Config.LABEL_NAME}. "
    "Работаешь строго с переданными данными, ничего не выдумываешь. "
    "Формула прибыли артиста: (Выручка - Расходы) * Ставка; "
    "остаток после выплаты артисту — прибыль лейбла. "
    "Пишешь по-русски, коротко, по делу, суммы с разделителями разрядов. "
    "Если данных недостаточно — прямо говоришь об этом."
)

EXCEL_PARSE_PROMPT = """Ниже — сырые строки из Excel-выгрузки лейбла (первая строка обычно заголовки).
Преобразуй их в нормализованные финансовые операции.

Верни ТОЛЬКО валидный JSON без markdown-обёртки:
{
  "rows": [
    {
      "artist": "строка или null",
      "kind": "income" | "expense",
      "amount": число (положительное, точка как разделитель),
      "currency": "RUB" | "USD" | "EUR",
      "category": "строка или null",
      "description": "строка или null",
      "occurred_on": "YYYY-MM-DD"
    }
  ],
  "errors": ["человекочитаемые проблемы: пустые суммы, битые даты, отрицательные доходы, дубли"],
  "summary": "1-2 предложения о содержимом файла"
}

Правила:
- отрицательная сумма в колонке дохода => kind="expense", amount = abs(value);
- расходы (реклама, студия, мастеринг, дистрибуция, аванс) => kind="expense";
- роялти, стриминг, концерты, синхронизации, мерч => kind="income";
- если дата неполная (например "07.2026") — используй первое число месяца;
- строки-итоги ("Итого", "Total", "Сумма") пропускай и упоминай в errors;
- сомнительные строки не выбрасывай молча — пиши в errors.
"""


class ClaudeService:
    def __init__(self, api_key: str, model: str, max_tokens: int) -> None:
        from anthropic import Anthropic

        self.client = Anthropic(api_key=api_key, max_retries=3, timeout=120.0)
        self.model = model
        self.max_tokens = max_tokens

    def _complete(
        self, prompt: str, system: str = SYSTEM_ANALYST, max_tokens: int | None = None
    ) -> str:
        try:
            msg = self.client.messages.create(
                model=self.model,
                max_tokens=max_tokens or self.max_tokens,
                system=system,
                messages=[{"role": "user", "content": prompt}],
            )
        except Exception as exc:  # сеть, таймаут и прочее
            log.exception("Claude call failed")
            status = getattr(exc, "status_code", None)
            if status:
                raise RuntimeError(f"Claude недоступен (HTTP {status}).") from exc
            raise RuntimeError(f"Claude недоступен: {exc}") from exc
        return "".join(
            b.text for b in msg.content if getattr(b, "type", "") == "text"
        ).strip()

    async def complete(self, prompt: str, **kw: Any) -> str:
        return await asyncio.to_thread(self._complete, prompt, **kw)

    @staticmethod
    def _extract_json(text: str) -> dict:
        text = re.sub(
            r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE
        ).strip()
        start, end = text.find("{"), text.rfind("}")
        if start == -1 or end == -1:
            raise ValueError("Claude вернул ответ без JSON")
        return json.loads(text[start : end + 1])

    async def parse_excel(self, raw_rows: list[list[Any]]) -> dict:
        payload = json_dumps({"rows": raw_rows})
        text = await self.complete(
            f"{EXCEL_PARSE_PROMPT}\n\nДанные:\n{payload}", max_tokens=8000
        )
        return await asyncio.to_thread(self._extract_json, text)

    async def monthly_report(
        self, year: int, month: int, breakdown: list[dict], categories: list[dict]
    ) -> str:
        prompt = (
            f"Составь финансовый отчёт лейбла за {month:02d}.{year}.\n\n"
            f"Данные по артистам (JSON):\n{json_dumps(breakdown)}\n\n"
            f"Разрез по категориям (JSON):\n{json_dumps(categories)}\n\n"
            "Структура ответа:\n"
            "1. Итоги месяца: выручка, расходы, маржа, прибыль лейбла, к выплате артистам.\n"
            "2. Топ-3 артиста по прибыли и кто убыточен.\n"
            "3. Главные статьи расходов.\n"
            "4. Замеченные ошибки/подозрительные данные.\n"
            "5. 2-3 конкретные рекомендации.\n"
            "Формат: HTML-теги Telegram (<b>, <i>, <code>), без markdown, до 3500 символов."
        )
        return await self.complete(prompt)

    async def anomaly_report(self, data: dict) -> str:
        prompt = (
            "Проанализируй данные на аномалии и ошибки учёта.\n\n"
            f"{json_dumps(data)}\n\n"
            "Верни:\n"
            "1. КРИТИЧНО — что требует проверки сегодня (с суммами и датами).\n"
            "2. ВНИМАНИЕ — подозрительные тренды и выбросы.\n"
            "3. ОШИБКИ ДАННЫХ — дубли, операции без артиста, нереальные значения.\n"
            "4. Что проверить руками — чек-лист из 3-5 пунктов.\n"
            "Если аномалий нет — скажи это прямо. "
            "Формат: HTML-теги Telegram, без markdown, до 3500 символов."
        )
        return await self.complete(prompt)


# -----------------------------------------------------------------------------
# Форматирование и парсинг
# -----------------------------------------------------------------------------
def money(value: Any, currency: str = "") -> str:
    d = Decimal(str(value or 0)).quantize(CENTS)
    s = f"{d:,.2f}".replace(",", " ").replace(".", ",")
    return f"{s} {currency}".strip()


def parse_amount(text: str) -> Decimal:
    cleaned = re.sub(r"[^\d,.\-]", "", text or "").replace(",", ".")
    if cleaned.count(".") > 1:  # 1.234.567 -> 1234567
        head, _, tail = cleaned.rpartition(".")
        cleaned = head.replace(".", "") + "." + tail
    try:
        value = Decimal(cleaned).quantize(CENTS, rounding=ROUND_HALF_UP)
    except (InvalidOperation, ValueError) as exc:
        raise ValueError(
            "Не смог распознать сумму. Пример: 125000 или 125 000,50"
        ) from exc
    if value <= 0:
        raise ValueError("Сумма должна быть больше нуля.")
    return value


def parse_rate(text: str) -> Decimal:
    raw = parse_amount(text)
    rate = raw / Decimal(100) if raw > 1 else raw
    if not (Decimal(0) < rate <= Decimal(1)):
        raise ValueError(
            "Ставка должна быть в диапазоне 1-100% (например 20 или 0.2)."
        )
    return rate.quantize(Decimal("0.0001"))


def parse_date(text: str) -> date:
    text = (text or "").strip().lower()
    if text in {"", "сегодня", "-"}:
        return date.today()
    for fmt in ("%d.%m.%Y", "%d.%m.%y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    raise ValueError("Не смог распознать дату. Формат: ДД.ММ.ГГГГ или «сегодня».")


def chunks(text: str, size: int = 3800) -> list[str]:
    return [text[i : i + size] for i in range(0, len(text), size)] or ["—"]


# -----------------------------------------------------------------------------
# Клавиатуры (интерфейс только на inline-кнопках)
# -----------------------------------------------------------------------------
def kb_main() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(
        [
            [InlineKeyboardButton("👤 Добавить артиста", callback_data="artist:add")],
            [
                InlineKeyboardButton("➕ Доход", callback_data="tx:new:income"),
                InlineKeyboardButton("➖ Расход", callback_data="tx:new:expense"),
            ],
            [InlineKeyboardButton("📊 Отчёт за месяц", callback_data="report:menu")],
            [InlineKeyboardButton("📥 Загрузить Excel", callback_data="excel:wait")],
            [InlineKeyboardButton("🚨 Анализ аномалий", callback_data="anomaly:run")],
            [InlineKeyboardButton("💸 Список платежей", callback_data="pay:list:all")],
            [InlineKeyboardButton("🎤 Артисты", callback_data="artist:list")],
        ]
    )


def kb_back(
    extra: list[list[InlineKeyboardButton]] | None = None,
) -> InlineKeyboardMarkup:
    rows = list(extra or [])
    rows.append([InlineKeyboardButton("⬅️ Главное меню", callback_data="nav:main")])
    return InlineKeyboardMarkup(rows)


def kb_artists(
    artists: list[dict], prefix: str, allow_none: bool = False
) -> InlineKeyboardMarkup:
    rows: list[list[InlineKeyboardButton]] = []
    for i in range(0, len(artists), 2):
        rows.append(
            [
                InlineKeyboardButton(
                    str(a["name"])[:24], callback_data=f"{prefix}:{a['id']}"
                )
                for a in artists[i : i + 2]
            ]
        )
    if allow_none:
        rows.append([InlineKeyboardButton(NO_ARTIST, callback_data=f"{prefix}:0")])
    rows.append([InlineKeyboardButton("⬅️ Главное меню", callback_data="nav:main")])
    return InlineKeyboardMarkup(rows)


def kb_months() -> InlineKeyboardMarkup:
    today = date.today()
    rows, row = [], []
    for offset in range(6):
        month = today.month - offset
        year = today.year
        while month <= 0:
            month += 12
            year -= 1
        row.append(
            InlineKeyboardButton(
                f"{month:02d}.{year}", callback_data=f"report:run:{year}:{month}"
            )
        )
        if len(row) == 3:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    rows.append([InlineKeyboardButton("⬅️ Главное меню", callback_data="nav:main")])
    return InlineKeyboardMarkup(rows)


def kb_payments(payments: list[dict]) -> InlineKeyboardMarkup:
    rows = [
        [
            InlineKeyboardButton(
                f"✅ Оплатить #{p['id']} · {str(p['artist'])[:14]}",
                callback_data=f"pay:paid:{p['id']}",
            )
        ]
        for p in payments
        if p["status"] == "pending"
    ][:8]
    rows.append(
        [
            InlineKeyboardButton("⏳ Ожидают", callback_data="pay:list:pending"),
            InlineKeyboardButton("✅ Оплачены", callback_data="pay:list:paid"),
        ]
    )
    rows.append(
        [InlineKeyboardButton("🔄 Пересчитать за месяц", callback_data="pay:recalc")]
    )
    rows.append([InlineKeyboardButton("⬅️ Главное меню", callback_data="nav:main")])
    return InlineKeyboardMarkup(rows)


# -----------------------------------------------------------------------------
# Доступ и вспомогательные ответы
# -----------------------------------------------------------------------------
def authorized(update: Update) -> bool:
    user = update.effective_user
    if not user:
        return False
    if not Config.ALLOWED_USER_IDS:  # пустой список = открытый режим (dev)
        return True
    return user.id in Config.ALLOWED_USER_IDS


async def send(
    update: Update, text: str, keyboard: InlineKeyboardMarkup | None = None
) -> None:
    parts = chunks(text)
    chat = update.effective_chat
    for idx, part in enumerate(parts):
        markup = keyboard if idx == len(parts) - 1 else None
        try:
            await chat.send_message(part, reply_markup=markup, parse_mode=ParseMode.HTML)
        except BadRequest:  # некорректный HTML от модели — отправляем как текст
            await chat.send_message(part, reply_markup=markup)


async def show_main_menu(update: Update, prefix: str = "") -> None:
    head = f"{prefix}\n\n" if prefix else ""
    await send(
        update,
        f"{head}<b>💼 Финансы лейбла {Config.LABEL_NAME}</b>\nВыберите действие:",
        kb_main(),
    )


def reset_state(context: ContextTypes.DEFAULT_TYPE) -> None:
    for key in ("state", "draft"):
        context.user_data.pop(key, None)


# -----------------------------------------------------------------------------
# Текстовые сообщения (конечный автомат, без slash-команд)
# -----------------------------------------------------------------------------
async def on_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not authorized(update):
        await update.effective_chat.send_message("⛔ Доступ запрещён.")
        return

    repo: Repo = context.bot_data["repo"]
    state: str | None = context.user_data.get("state")
    text = (update.effective_message.text or "").strip()

    if not state:
        await show_main_menu(update)
        return

    try:
        if state == "artist:name":
            context.user_data["draft"] = {"name": text[:120]}
            context.user_data["state"] = "artist:rate"
            await send(
                update,
                f"Ставка артиста <b>{text[:120]}</b> в процентах (например <code>20</code>):",
                kb_back(),
            )

        elif state == "artist:rate":
            rate = parse_rate(text)
            draft = context.user_data.get("draft", {})
            artist = await asyncio.to_thread(
                repo.add_artist, draft["name"], rate, update.effective_user.username
            )
            await asyncio.to_thread(
                repo.audit, update.effective_user.id, "artist_add", artist
            )
            reset_state(context)
            await show_main_menu(
                update,
                f"✅ Артист <b>{artist['name']}</b> сохранён. "
                f"Ставка: {Decimal(str(artist['rate'])) * 100:.2f}%",
            )

        elif state == "tx:amount":
            amount = parse_amount(text)
            context.user_data["draft"]["amount"] = amount
            context.user_data["state"] = "tx:category"
            await send(
                update,
                f"Сумма: <b>{money(amount, Config.DEFAULT_CURRENCY)}</b>\n"
                "Теперь категория и описание одной строкой\n"
                "(например: <code>Стриминг | Spotify июль</code>):",
                kb_back(),
            )

        elif state == "tx:category":
            category, _, description = (x.strip() for x in text.partition("|"))
            context.user_data["draft"]["category"] = category[:80] or None
            context.user_data["draft"]["description"] = description[:400] or None
            context.user_data["state"] = "tx:date"
            await send(
                update, "Дата операции (ДД.ММ.ГГГГ) или «сегодня»:", kb_back()
            )

        elif state == "tx:date":
            draft = context.user_data["draft"]
            occurred = parse_date(text)
            await asyncio.to_thread(
                repo.add_transaction,
                draft.get("artist_id"),
                draft["kind"],
                draft["amount"],
                Config.DEFAULT_CURRENCY,
                draft.get("category"),
                draft.get("description"),
                occurred,
                "manual",
                None,
                update.effective_user.id,
            )
            label = "Доход" if draft["kind"] == "income" else "Расход"
            amount = draft["amount"]
            reset_state(context)
            await show_main_menu(
                update,
                f"✅ {label} {money(amount, Config.DEFAULT_CURRENCY)} "
                f"от {occurred:%d.%m.%Y} записан.",
            )

        else:
            reset_state(context)
            await show_main_menu(update)

    except ValueError as exc:
        await send(update, f"⚠️ {exc}\nПопробуйте ещё раз:", kb_back())


# -----------------------------------------------------------------------------
# Inline-кнопки
# -----------------------------------------------------------------------------
async def on_callback(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()

    if not authorized(update):
        await query.edit_message_text("⛔ Доступ запрещён.")
        return

    repo: Repo = context.bot_data["repo"]
    claude: ClaudeService = context.bot_data["claude"]
    parts = (query.data or "").split(":")
    domain, action = parts[0], parts[1] if len(parts) > 1 else ""
    args = parts[2:]

    # ---------- навигация ----------
    if domain == "nav":
        reset_state(context)
        await show_main_menu(update)
        return

    # ---------- артисты ----------
    if domain == "artist":
        if action == "add":
            context.user_data["state"] = "artist:name"
            await send(update, "👤 Введите имя (или название) артиста:", kb_back())
        elif action == "list":
            artists = await asyncio.to_thread(repo.list_artists)
            if not artists:
                await send(update, "Пока нет ни одного артиста.", kb_main())
                return
            lines = [
                f"• <b>{a['name']}</b> — {Decimal(str(a['rate'])) * 100:.2f}%"
                + (f" (@{a['tg_username']})" if a["tg_username"] else "")
                for a in artists
            ]
            await send(
                update, "🎤 <b>Артисты лейбла</b>\n" + "\n".join(lines), kb_main()
            )
        return

    # ---------- транзакции ----------
    if domain == "tx":
        if action == "new":
            kind = args[0]
            context.user_data["draft"] = {"kind": kind}
            context.user_data["state"] = None
            artists = await asyncio.to_thread(repo.list_artists)
            title = "➕ Доход" if kind == "income" else "➖ Расход"
            await send(
                update,
                f"{title}: выберите артиста",
                kb_artists(artists, f"tx:artist:{kind}", allow_none=True),
            )
        elif action == "artist":
            kind, artist_id = args[0], int(args[1])
            context.user_data["draft"] = {"kind": kind, "artist_id": artist_id or None}
            context.user_data["state"] = "tx:amount"
            await send(
                update,
                f"Введите сумму в {Config.DEFAULT_CURRENCY} (например <code>125000</code>):",
                kb_back(),
            )
        return

    # ---------- отчёты ----------
    if domain == "report":
        if action == "menu":
            await send(update, "📊 Выберите месяц отчёта:", kb_months())
            return
        if action == "run":
            year, month = int(args[0]), int(args[1])
            await update.effective_chat.send_action(ChatAction.TYPING)
            breakdown, categories = await asyncio.gather(
                asyncio.to_thread(repo.monthly_breakdown, year, month),
                asyncio.to_thread(repo.category_breakdown, year, month),
            )
            if not breakdown:
                await send(update, f"За {month:02d}.{year} операций нет.", kb_main())
                return

            revenue = sum((r["revenue"] for r in breakdown), Decimal(0))
            expenses = sum((r["expenses"] for r in breakdown), Decimal(0))
            shares = sum((r["artist_share"] for r in breakdown), Decimal(0))
            profit = sum((r["label_profit"] for r in breakdown), Decimal(0))

            head = [
                f"📊 <b>Отчёт за {month:02d}.{year}</b>",
                f"Выручка: <b>{money(revenue, Config.DEFAULT_CURRENCY)}</b>",
                f"Расходы: <b>{money(expenses, Config.DEFAULT_CURRENCY)}</b>",
                f"К выплате артистам: <b>{money(shares, Config.DEFAULT_CURRENCY)}</b>",
                f"Прибыль лейбла: <b>{money(profit, Config.DEFAULT_CURRENCY)}</b>",
                "",
                "<b>По артистам</b>",
            ]
            for r in breakdown[:15]:
                head.append(
                    f"• {r['artist']}: доход {money(r['revenue'])} / "
                    f"расход {money(r['expenses'])} → "
                    f"артисту {money(r['artist_share'])}, лейблу {money(r['label_profit'])}"
                )
            await send(update, "\n".join(head))

            try:
                ai_text = await claude.monthly_report(year, month, breakdown, categories)
            except RuntimeError as exc:
                await send(update, f"⚠️ {exc}", kb_main())
                return

            start, end = repo.month_bounds(year, month)
            await asyncio.to_thread(
                repo.save_report,
                "monthly_report",
                ai_text,
                {"breakdown": breakdown, "categories": categories},
                (start, end),
            )
            await send(update, f"🤖 <b>Аналитика Claude</b>\n\n{ai_text}", kb_main())
        return

    # ---------- Excel ----------
    if domain == "excel":
        context.user_data["state"] = None
        await send(
            update,
            "📥 Пришлите файл <code>.xlsx</code> или <code>.xls</code> как документ.\n\n"
            "Ожидаемые колонки (в любом порядке и на любом языке): дата, артист, "
            "тип операции, сумма, категория, описание. Claude сам разберёт "
            "структуру и подсветит ошибки.",
            kb_back(),
        )
        return

    # ---------- аномалии ----------
    if domain == "anomaly":
        await update.effective_chat.send_action(ChatAction.TYPING)
        data = await asyncio.to_thread(repo.anomaly_candidates)
        if not data["monthly"]:
            await send(update, "Недостаточно данных для анализа аномалий.", kb_main())
            return
        try:
            text = await claude.anomaly_report(data)
        except RuntimeError as exc:
            await send(update, f"⚠️ {exc}", kb_main())
            return
        await asyncio.to_thread(repo.save_report, "anomaly", text, data, None)
        await send(update, f"🚨 <b>Анализ аномалий</b>\n\n{text}", kb_main())
        return

    # ---------- платежи ----------
    if domain == "pay":
        if action == "list":
            status = None if not args or args[0] == "all" else args[0]
            payments = await asyncio.to_thread(repo.list_payments, 20, status)
            if not payments:
                await send(update, "Платежей не найдено.", kb_payments([]))
                return
            icons = {"pending": "⏳", "paid": "✅", "canceled": "❌"}
            lines = ["💸 <b>Платежи артистам</b>"]
            for p in payments:
                lines.append(
                    f"{icons.get(p['status'], '•')} #{p['id']} <b>{p['artist']}</b> · "
                    f"{p['period_start']:%m.%Y} · "
                    f"{money(p['amount'], Config.DEFAULT_CURRENCY)} "
                    f"(ставка {p['rate'] * 100:.0f}%)"
                )
            total_pending = sum(
                (p["amount"] for p in payments if p["status"] == "pending"), Decimal(0)
            )
            lines.append(
                f"\nК выплате: <b>{money(total_pending, Config.DEFAULT_CURRENCY)}</b>"
            )
            await send(update, "\n".join(lines), kb_payments(payments))
        elif action == "recalc":
            today = date.today()
            created = await asyncio.to_thread(
                repo.upsert_payments_for_month, today.year, today.month
            )
            payments = await asyncio.to_thread(repo.list_payments, 20, "pending")
            await send(
                update,
                f"🔄 Пересчитано начислений: <b>{len(created)}</b> за {today:%m.%Y}.",
                kb_payments(payments),
            )
        elif action == "paid":
            payment_id = int(args[0])
            await asyncio.to_thread(repo.mark_payment_paid, payment_id)
            await asyncio.to_thread(
                repo.audit, update.effective_user.id, "payment_paid", {"id": payment_id}
            )
            payments = await asyncio.to_thread(repo.list_payments, 20, "pending")
            await send(
                update,
                f"✅ Платёж #{payment_id} отмечен как оплаченный.",
                kb_payments(payments),
            )
        return

    await show_main_menu(update)


# -----------------------------------------------------------------------------
# Excel-документы
# -----------------------------------------------------------------------------
def read_workbook(blob: bytes, max_rows: int | None = None) -> list[list[Any]]:
    max_rows = max_rows or Config.MAX_EXCEL_ROWS
    from openpyxl import load_workbook

    wb = load_workbook(BytesIO(blob), data_only=True, read_only=True)
    rows: list[list[Any]] = []
    try:
        for sheet in wb.worksheets:
            for row in sheet.iter_rows(values_only=True):
                if row is None or all(cell in (None, "") for cell in row):
                    continue
                rows.append(
                    [
                        cell.strftime("%Y-%m-%d")
                        if isinstance(cell, (datetime, date))
                        else (float(cell) if isinstance(cell, Decimal) else cell)
                        for cell in row
                    ]
                )
                if len(rows) >= max_rows:
                    return rows
    finally:
        wb.close()
    return rows


async def on_document(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    if not authorized(update):
        return

    repo: Repo = context.bot_data["repo"]
    claude: ClaudeService = context.bot_data["claude"]
    doc = update.effective_message.document

    if not (doc.file_name or "").lower().endswith((".xlsx", ".xlsm", ".xls")):
        await send(update, "⚠️ Нужен файл Excel (.xlsx / .xls).", kb_main())
        return
    if (doc.file_size or 0) > Config.MAX_EXCEL_BYTES:
        await send(update, "⚠️ Файл слишком большой.", kb_main())
        return

    await update.effective_chat.send_action(ChatAction.TYPING)
    tg_file = await doc.get_file()
    blob = bytes(await tg_file.download_as_bytearray())

    try:
        raw_rows = await asyncio.to_thread(read_workbook, blob)
    except Exception as exc:
        log.exception("Excel read failed")
        await send(update, f"⚠️ Не смог прочитать файл: {exc}", kb_main())
        return

    if not raw_rows:
        await send(update, "⚠️ Файл пустой.", kb_main())
        return

    await send(update, f"🤖 Отправляю {len(raw_rows)} строк в Claude на разбор…")

    try:
        parsed = await claude.parse_excel(raw_rows)
    except (RuntimeError, ValueError, json.JSONDecodeError) as exc:
        log.exception("Claude parse failed")
        await send(update, f"⚠️ Claude не смог разобрать файл: {exc}", kb_main())
        return

    candidates = parsed.get("rows", []) or []
    errors: list[str] = list(parsed.get("errors", []) or [])

    names = sorted(
        {str(r.get("artist")).strip() for r in candidates if r.get("artist")}
    )
    mapping = await asyncio.to_thread(repo.resolve_artist_ids, names)

    prepared: list[dict] = []
    for idx, r in enumerate(candidates):
        try:
            amount = Decimal(str(r["amount"])).quantize(CENTS, rounding=ROUND_HALF_UP)
            occurred = datetime.strptime(str(r["occurred_on"]), "%Y-%m-%d").date()
            kind = r["kind"]
            if kind not in ("income", "expense") or amount <= 0:
                raise ValueError("некорректный тип или сумма")
        except Exception as exc:
            errors.append(f"строка {idx + 1}: {exc}")
            continue

        artist_name = str(r.get("artist")).strip() if r.get("artist") else None
        artist_id = mapping.get(artist_name.lower()) if artist_name else None
        if artist_name and artist_id is None:
            errors.append(
                f"артист «{artist_name}» не найден в базе — операция без привязки"
            )

        prepared.append(
            {
                "artist_id": artist_id,
                "kind": kind,
                "amount": amount,
                "currency": (r.get("currency") or Config.DEFAULT_CURRENCY)[:3].upper(),
                "category": r.get("category") or None,
                "description": r.get("description") or None,
                "occurred_on": occurred,
                "external_key": (
                    f"xlsx:{doc.file_unique_id}:{idx}:{kind}:{to_cents(amount)}:{occurred}"
                ),
            }
        )

    inserted = await asyncio.to_thread(
        repo.bulk_insert_transactions, prepared, update.effective_user.id
    )
    await asyncio.to_thread(
        repo.save_report,
        "excel_import",
        parsed.get("summary", ""),
        {"file": doc.file_name, "rows": len(prepared), "errors": errors},
        None,
    )

    income = sum((r["amount"] for r in prepared if r["kind"] == "income"), Decimal(0))
    expense = sum((r["amount"] for r in prepared if r["kind"] == "expense"), Decimal(0))

    lines = [
        f"📥 <b>Импорт «{doc.file_name}» завершён</b>",
        f"Распознано строк: <b>{len(prepared)}</b>, записано новых: <b>{inserted}</b>",
        f"Доходы: {money(income, Config.DEFAULT_CURRENCY)} · "
        f"Расходы: {money(expense, Config.DEFAULT_CURRENCY)}",
    ]
    if parsed.get("summary"):
        lines += ["", f"<i>{parsed['summary']}</i>"]
    if errors:
        lines += ["", "<b>⚠️ Проблемы в данных:</b>"] + [
            f"• {e}" for e in errors[:15]
        ]
        if len(errors) > 15:
            lines.append(f"…и ещё {len(errors) - 15}")

    reset_state(context)
    await send(update, "\n".join(lines), kb_main())


# -----------------------------------------------------------------------------
# Глобальный обработчик ошибок
# -----------------------------------------------------------------------------
async def on_error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.exception("Unhandled error", exc_info=context.error)
    if isinstance(update, Update) and update.effective_chat:
        try:
            await update.effective_chat.send_message(
                "⚠️ Внутренняя ошибка. Событие записано в лог, попробуйте ещё раз.",
                reply_markup=kb_main(),
            )
        except Exception:
            pass


# -----------------------------------------------------------------------------
# Точка входа (Bothost запускает именно этот файл)
# -----------------------------------------------------------------------------
def main() -> None:
    missing = [
        name for name in ("BOT_TOKEN", "ANTHROPIC_API_KEY") if not getattr(Config, name)
    ]
    if missing:
        raise SystemExit(
            "Не заданы переменные окружения: "
            + ", ".join(missing)
            + ". Добавьте их в разделе «Переменные окружения» панели Bothost."
        )

    db = Database()
    db.init_schema()
    repo = Repo(db)
    claude = ClaudeService(
        Config.ANTHROPIC_API_KEY, Config.CLAUDE_MODEL, Config.CLAUDE_MAX_TOKENS
    )

    builder = ApplicationBuilder().token(Config.BOT_TOKEN).concurrent_updates(True)
    if AIORateLimiter is not None:
        builder = builder.rate_limiter(AIORateLimiter())
    app = builder.build()
    app.bot_data["repo"] = repo
    app.bot_data["claude"] = claude

    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))
    # Slash-команд нет: любой текст (включая /start) открывает главное меню
    app.add_handler(MessageHandler(filters.TEXT | filters.COMMAND, on_message))
    app.add_error_handler(on_error)

    log.info(
        "Bot started | label=%s | model=%s | backend=%s",
        Config.LABEL_NAME,
        Config.CLAUDE_MODEL,
        "postgres" if db.is_postgres else "sqlite",
    )
    try:
        app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)
    finally:
        db.close()


if __name__ == "__main__":
    main()
