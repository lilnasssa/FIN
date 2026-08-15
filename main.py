# =============================================================================
#  22-17 LABEL FINANCE BOT — Bothost edition
#  Управление финансами лейбла: артисты, доходы/расходы, Excel-импорт,
#  месячные отчёты, выплаты и поиск аномалий через Gemini.
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
import csv
import json
import logging
import os
import re
import secrets
import sqlite3
import statistics
import threading
from contextlib import contextmanager
from contextvars import ContextVar
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO, StringIO
from pathlib import Path
from typing import Any, Iterator, Sequence

from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    InputFile,
    Update,
)
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
    # Ключ Google AI Studio. Поддерживаются оба привычных имени переменной.
    GEMINI_API_KEY: str = (
        os.getenv("GEMINI_API_KEY", "").strip()
        or os.getenv("GOOGLE_API_KEY", "").strip()
    )
    # auto = бот сам спросит у API список моделей и выберет свежую доступную.
    # Можно жёстко задать имя, например GEMINI_MODEL=gemini-flash-latest.
    GEMINI_MODEL: str = os.getenv("GEMINI_MODEL", "auto").strip() or "auto"
    # Что предпочитать в авторежиме: flash (быстро/дёшево) или pro (точнее).
    GEMINI_PREFER: str = os.getenv("GEMINI_PREFER", "flash").strip().lower()
    GEMINI_MAX_TOKENS: int = int(os.getenv("GEMINI_MAX_TOKENS", "4096"))

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
    # Владельцы: полные права, включая управление доступом. Задаются один раз.
    OWNER_IDS: set[int] = {
        int(x)
        for x in os.getenv("OWNER_IDS", "").replace(" ", "").split(",")
        if x.strip().lstrip("-").isdigit()
    }
    # Срок жизни кода-приглашения в часах (0 = бессрочно) и число активаций.
    INVITE_TTL_HOURS: int = int(os.getenv("INVITE_TTL_HOURS", "72"))
    INVITE_DEFAULT_USES: int = int(os.getenv("INVITE_DEFAULT_USES", "1"))

    DEFAULT_CURRENCY: str = os.getenv("DEFAULT_CURRENCY", "RUB")
    DEFAULT_RATE: Decimal = _env_decimal("DEFAULT_RATE", "0.20")
    MAX_EXCEL_BYTES: int = int(os.getenv("MAX_EXCEL_MB", "10")) * 1024 * 1024
    MAX_EXCEL_ROWS: int = int(os.getenv("MAX_EXCEL_ROWS", "500"))
    ANOMALY_SIGMA: Decimal = _env_decimal("ANOMALY_SIGMA", "2")
    ANOMALY_MONTHS: int = int(os.getenv("ANOMALY_MONTHS", "6"))
    LABEL_NAME: str = os.getenv("LABEL_NAME", "22-17")
    # Ставка налога для резерва в сводке (0.06 = УСН 6% с дохода).
    TAX_RATE: Decimal = _env_decimal("TAX_RATE", "0.06")
    # Автоматическая сводка 1-го числа владельцам и админам.
    MONTHLY_DIGEST: bool = os.getenv("MONTHLY_DIGEST", "1").strip() not in (
        "0",
        "false",
        "False",
        "",
    )
    DIGEST_HOUR: int = int(os.getenv("DIGEST_HOUR", "10"))


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

CREATE TABLE IF NOT EXISTS bot_users (
    tg_user_id      INTEGER PRIMARY KEY,
    username        TEXT,
    full_name       TEXT,
    role            TEXT NOT NULL DEFAULT 'viewer'
                    CHECK (role IN ('owner', 'admin', 'viewer')),
    is_active       INTEGER NOT NULL DEFAULT 1,
    invited_by      INTEGER,
    created_at      TEXT NOT NULL DEFAULT (datetime('now')),
    last_seen_at    TEXT
);

CREATE TABLE IF NOT EXISTS invites (
    code            TEXT PRIMARY KEY,
    role            TEXT NOT NULL DEFAULT 'viewer'
                    CHECK (role IN ('admin', 'viewer')),
    uses_left       INTEGER NOT NULL DEFAULT 1,
    expires_at      TEXT,
    created_by      INTEGER,
    is_revoked      INTEGER NOT NULL DEFAULT 0,
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

CREATE TABLE IF NOT EXISTS bot_users (
    tg_user_id      BIGINT PRIMARY KEY,
    username        TEXT,
    full_name       TEXT,
    role            TEXT NOT NULL DEFAULT 'viewer'
                    CHECK (role IN ('owner', 'admin', 'viewer')),
    is_active       SMALLINT NOT NULL DEFAULT 1,
    invited_by      BIGINT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    last_seen_at    TEXT
);

CREATE TABLE IF NOT EXISTS invites (
    code            TEXT PRIMARY KEY,
    role            TEXT NOT NULL DEFAULT 'viewer'
                    CHECK (role IN ('admin', 'viewer')),
    uses_left       INTEGER NOT NULL DEFAULT 1,
    expires_at      TEXT,
    created_by      BIGINT,
    is_revoked      SMALLINT NOT NULL DEFAULT 0,
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

    # ---------- аналитика и работа с операциями ----------
    def totals(self, start: date, end: date) -> dict:
        row = self.db.query(
            """
            SELECT
                COALESCE(SUM(CASE WHEN kind = 'income'  THEN amount_cents END), 0) AS rev,
                COALESCE(SUM(CASE WHEN kind = 'expense' THEN amount_cents END), 0) AS exp,
                COUNT(*) AS cnt
            FROM transactions
            WHERE occurred_on >= ? AND occurred_on <= ?
            """,
            (self.db.d(start), self.db.d(end)),
        )[0]
        revenue = from_cents(row["rev"] or 0)
        expenses = from_cents(row["exp"] or 0)
        return {
            "revenue": revenue,
            "expenses": expenses,
            "net": revenue - expenses,
            "count": int(row["cnt"] or 0),
        }

    def recent_transactions(self, limit: int = 10) -> list[dict]:
        rows = self.db.query(
            """
            SELECT t.id, t.kind, t.amount_cents, t.currency, t.category,
                   t.description, t.occurred_on, t.source, a.name AS artist
              FROM transactions t
              LEFT JOIN artists a ON a.id = t.artist_id
             ORDER BY t.id DESC
             LIMIT ?
            """,
            (int(limit),),
        )
        for r in rows:
            r["amount"] = from_cents(r["amount_cents"])
        return rows

    def search_transactions(self, term: str, limit: int = 15) -> list[dict]:
        like = f"%{term.lower()}%"
        rows = self.db.query(
            """
            SELECT t.id, t.kind, t.amount_cents, t.currency, t.category,
                   t.description, t.occurred_on, a.name AS artist
              FROM transactions t
              LEFT JOIN artists a ON a.id = t.artist_id
             WHERE LOWER(COALESCE(t.category, '')) LIKE ?
                OR LOWER(COALESCE(t.description, '')) LIKE ?
                OR LOWER(COALESCE(a.name, '')) LIKE ?
             ORDER BY t.occurred_on DESC, t.id DESC
             LIMIT ?
            """,
            (like, like, like, int(limit)),
        )
        for r in rows:
            r["amount"] = from_cents(r["amount_cents"])
        return rows

    def period_transactions(self, start: date, end: date) -> list[dict]:
        rows = self.db.query(
            """
            SELECT t.id, t.occurred_on, t.kind, t.amount_cents, t.currency,
                   t.category, t.description, t.source, a.name AS artist
              FROM transactions t
              LEFT JOIN artists a ON a.id = t.artist_id
             WHERE t.occurred_on >= ? AND t.occurred_on <= ?
             ORDER BY t.occurred_on, t.id
            """,
            (self.db.d(start), self.db.d(end)),
        )
        for r in rows:
            r["amount"] = from_cents(r["amount_cents"])
        return rows

    def last_transaction_by(self, user_id: int) -> dict | None:
        rows = self.db.query(
            """
            SELECT t.id, t.kind, t.amount_cents, t.category, t.description,
                   t.occurred_on, a.name AS artist
              FROM transactions t
              LEFT JOIN artists a ON a.id = t.artist_id
             WHERE t.created_by = ?
             ORDER BY t.id DESC
             LIMIT 1
            """,
            (user_id,),
        )
        if not rows:
            return None
        rows[0]["amount"] = from_cents(rows[0]["amount_cents"])
        return rows[0]

    def delete_transaction(self, tx_id: int) -> int:
        return self.db.execute("DELETE FROM transactions WHERE id = ?", (int(tx_id),))

    # ---------- доступ: пользователи ----------
    def users_count(self) -> int:
        return int(self.db.scalar("SELECT COUNT(*) FROM bot_users") or 0)

    def owners_count(self) -> int:
        return int(
            self.db.scalar(
                "SELECT COUNT(*) FROM bot_users WHERE role = 'owner' AND is_active = 1"
            )
            or 0
        )

    def get_user(self, tg_user_id: int) -> dict | None:
        rows = self.db.query(
            "SELECT * FROM bot_users WHERE tg_user_id = ?", (tg_user_id,)
        )
        return rows[0] if rows else None

    def list_users(self) -> list[dict]:
        return self.db.query(
            """
            SELECT * FROM bot_users
            ORDER BY CASE role
                        WHEN 'owner' THEN 0
                        WHEN 'admin' THEN 1
                        ELSE 2
                     END,
                     tg_user_id
            """
        )

    def grant(
        self,
        tg_user_id: int,
        role: str,
        invited_by: int | None = None,
        username: str | None = None,
        full_name: str | None = None,
    ) -> dict | None:
        self.db.execute(
            """
            INSERT INTO bot_users
                (tg_user_id, username, full_name, role, is_active, invited_by)
            VALUES (?, ?, ?, ?, 1, ?)
            ON CONFLICT (tg_user_id) DO UPDATE SET
                role = excluded.role,
                is_active = 1,
                username = COALESCE(excluded.username, bot_users.username),
                full_name = COALESCE(excluded.full_name, bot_users.full_name)
            """,
            (tg_user_id, username, full_name, role, invited_by),
        )
        return self.get_user(tg_user_id)

    def set_role(self, tg_user_id: int, role: str) -> None:
        self.db.execute(
            "UPDATE bot_users SET role = ? WHERE tg_user_id = ?", (role, tg_user_id)
        )

    def set_active(self, tg_user_id: int, active: bool) -> None:
        self.db.execute(
            "UPDATE bot_users SET is_active = ? WHERE tg_user_id = ?",
            (1 if active else 0, tg_user_id),
        )

    def delete_user(self, tg_user_id: int) -> None:
        self.db.execute("DELETE FROM bot_users WHERE tg_user_id = ?", (tg_user_id,))

    def touch_user(
        self,
        tg_user_id: int,
        username: str | None = None,
        full_name: str | None = None,
    ) -> dict | None:
        """Возвращает запись доступа и обновляет профиль, либо None если доступа нет."""
        row = self.get_user(tg_user_id)
        if row is None:
            return None
        self.db.execute(
            """
            UPDATE bot_users
               SET username = COALESCE(?, username),
                   full_name = COALESCE(?, full_name),
                   last_seen_at = ?
             WHERE tg_user_id = ?
            """,
            (
                username,
                full_name,
                datetime.now().isoformat(timespec="seconds"),
                tg_user_id,
            ),
        )
        row["username"] = username or row["username"]
        row["full_name"] = full_name or row["full_name"]
        return row

    def ensure_owners(self, ids: set[int]) -> None:
        """Владельцы из OWNER_IDS всегда существуют и всегда владельцы."""
        for uid in sorted(ids):
            row = self.get_user(uid)
            if row is None:
                self.grant(uid, "owner")
            elif row["role"] != "owner" or not int(row["is_active"]):
                self.db.execute(
                    "UPDATE bot_users SET role = 'owner', is_active = 1 "
                    "WHERE tg_user_id = ?",
                    (uid,),
                )

    # ---------- доступ: приглашения ----------
    @staticmethod
    def new_code() -> str:
        alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"  # без похожих символов
        return "".join(secrets.choice(alphabet) for _ in range(8))

    def create_invite(
        self, role: str, uses: int, ttl_hours: int, created_by: int | None
    ) -> dict:
        uses = max(1, int(uses))
        expires = (
            (datetime.now() + timedelta(hours=ttl_hours)).isoformat(timespec="seconds")
            if ttl_hours > 0
            else None
        )
        for _ in range(5):  # на случай коллизии кода
            code = self.new_code()
            if not self.db.query("SELECT code FROM invites WHERE code = ?", (code,)):
                break
        self.db.execute(
            """
            INSERT INTO invites (code, role, uses_left, expires_at, created_by)
            VALUES (?, ?, ?, ?, ?)
            """,
            (code, role, uses, expires, created_by),
        )
        return {
            "code": code,
            "role": role,
            "uses_left": uses,
            "expires_at": expires,
        }

    def list_invites(self, only_active: bool = True) -> list[dict]:
        rows = self.db.query("SELECT * FROM invites ORDER BY created_at DESC")
        if not only_active:
            return rows
        now_iso = datetime.now().isoformat(timespec="seconds")
        active = []
        for r in rows:
            if int(r["is_revoked"]) or int(r["uses_left"]) <= 0:
                continue
            if r["expires_at"] and str(r["expires_at"]) < now_iso:
                continue
            active.append(r)
        return active

    def revoke_invite(self, code: str) -> None:
        self.db.execute(
            "UPDATE invites SET is_revoked = 1 WHERE code = ?", ((code or "").upper(),)
        )

    def redeem_invite(
        self,
        code: str,
        tg_user_id: int,
        username: str | None = None,
        full_name: str | None = None,
    ) -> dict | None:
        code = (code or "").strip().upper()
        if not code:
            return None
        rows = self.db.query("SELECT * FROM invites WHERE code = ?", (code,))
        if not rows:
            return None
        inv = rows[0]
        if int(inv["is_revoked"]) or int(inv["uses_left"]) <= 0:
            return None
        if inv["expires_at"] and str(inv["expires_at"]) < datetime.now().isoformat(
            timespec="seconds"
        ):
            return None
        used = self.db.execute(
            "UPDATE invites SET uses_left = uses_left - 1 "
            "WHERE code = ? AND uses_left > 0 AND is_revoked = 0",
            (code,),
        )
        if not used:
            return None
        user = self.grant(
            tg_user_id, inv["role"], inv["created_by"], username, full_name
        )
        self.audit(
            tg_user_id, "access_invite_redeem", {"code": code, "role": inv["role"]}
        )
        return user

    def audit(self, user_id: int | None, action: str, details: dict) -> None:
        self.db.execute(
            "INSERT INTO audit_log (tg_user_id, action, details) VALUES (?, ?, ?)",
            (user_id, action, json_dumps(details)),
        )


# -----------------------------------------------------------------------------
# Gemini
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


class GeminiService:
    """Обёртка над Google Gemini (пакет google-genai). Ключ — из GEMINI_API_KEY."""

    # Модели, которые не умеют в обычный текстовый ответ — никогда не выбираем.
    BAD_MODEL_MARKERS = (
        "embedding",
        "embed",
        "aqa",
        "imagen",
        "image",
        "veo",
        "video",
        "tts",
        "audio",
        "live",
        "realtime",
        "native-audio",
        "computer-use",
        "robotics",
        "guard",
    )

    def __init__(self, api_key: str, model: str, max_tokens: int) -> None:
        from google import genai  # ленивый импорт
        from google.genai import types as genai_types

        self.client = genai.Client(api_key=api_key)
        self.types = genai_types
        self.requested_model = (model or "auto").strip()
        self.model = "" if self.requested_model.lower() == "auto" else self.requested_model
        self.max_tokens = max_tokens
        self._dead_models: set[str] = set()
        self._model_lock = threading.Lock()

    # ---------- выбор модели ----------
    @classmethod
    def pick_model(
        cls, available: list[str], prefer: str = "flash", exclude: set[str] | None = None
    ) -> str | None:
        """Выбрать самую свежую текстовую модель из тех, что вернул API."""
        exclude = exclude or set()
        best: tuple[tuple[float, ...], str] | None = None
        for raw in available:
            name = (raw or "").replace("models/", "").strip()
            low = name.lower()
            if not low or name in exclude:
                continue
            if not low.startswith("gemini"):
                continue
            if any(bad in low for bad in cls.BAD_MODEL_MARKERS):
                continue
            match = re.search(r"gemini-(\d+(?:\.\d+)?)", low)
            is_latest = low.endswith("-latest")
            # алиасы вида gemini-flash-latest всегда указывают на свежее поколение
            version = float(match.group(1)) if match else (99.0 if is_latest else 0.0)
            score = (
                0.0 if ("preview" in low or "-exp" in low or low.endswith("exp")) else 1.0,
                1.0 if prefer in low else 0.0,
                -0.5 if "lite" in low else 0.0,
                version,
                1.0 if is_latest else 0.0,
                -len(low) / 1000.0,
            )
            if best is None or score > best[0]:
                best = (score, name)
        return best[1] if best else None

    def _available_models(self) -> list[str]:
        names: list[str] = []
        for item in self.client.models.list():
            name = (getattr(item, "name", "") or "").replace("models/", "")
            if not name:
                continue
            actions = (
                getattr(item, "supported_actions", None)
                or getattr(item, "supported_generation_methods", None)
                or []
            )
            if actions and not any(
                "generatecontent" in str(a).lower() for a in actions
            ):
                continue
            names.append(name)
        return names

    def resolve_model(self, force: bool = False) -> str:
        """Вернуть рабочее имя модели, при необходимости спросив список у API."""
        with self._model_lock:
            if self.model and not force and self.model not in self._dead_models:
                return self.model
            try:
                available = self._available_models()
            except Exception as exc:
                log.warning("Не удалось получить список моделей Gemini: %s", exc)
                available = []
            chosen = self.pick_model(
                available, Config.GEMINI_PREFER, exclude=self._dead_models
            )
            if not chosen:
                raise RuntimeError(
                    "Ни одна текстовая модель Gemini недоступна для этого ключа. "
                    "Проверьте GEMINI_API_KEY в Google AI Studio или задайте GEMINI_MODEL вручную."
                )
            if chosen != self.model:
                log.info("Gemini: выбрана модель %s (из %d доступных)", chosen, len(available))
            self.model = chosen
            return chosen

    def _complete(
        self,
        prompt: str,
        system: str = SYSTEM_ANALYST,
        max_tokens: int | None = None,
        json_mode: bool = False,
    ) -> str:
        config: dict[str, Any] = {
            "system_instruction": system,
            "max_output_tokens": max_tokens or self.max_tokens,
            "temperature": 0.2,
        }
        if json_mode:  # строгий JSON для разбора Excel
            config["response_mime_type"] = "application/json"

        resp = None
        last_exc: Exception | None = None
        # Две попытки: если модель закрыли (404), берём следующую доступную.
        for attempt in range(2):
            model = self.resolve_model(force=attempt > 0)
            try:
                resp = self.client.models.generate_content(
                    model=model,
                    contents=prompt,
                    config=self.types.GenerateContentConfig(**config),
                )
                break
            except Exception as exc:
                last_exc = exc
                status = getattr(exc, "code", None) or getattr(exc, "status_code", None)
                message = str(exc)
                model_gone = status == 404 or "NOT_FOUND" in message or (
                    "no longer available" in message
                )
                if model_gone and attempt == 0:
                    log.warning(
                        "Модель %s больше недоступна, ищу замену", model
                    )
                    self._dead_models.add(model)
                    continue
                log.exception("Gemini call failed")
                if status == 429:
                    raise RuntimeError(
                        "Лимит запросов Gemini исчерпан. Подождите минуту и повторите."
                    ) from exc
                if status in (401, 403):
                    raise RuntimeError(
                        "Ключ GEMINI_API_KEY неверный или без доступа к Gemini API."
                    ) from exc
                if model_gone:
                    raise RuntimeError(
                        f"Модель {model} недоступна, а замену найти не удалось. "
                        "Задайте GEMINI_MODEL вручную в переменных окружения."
                    ) from exc
                if status:
                    raise RuntimeError(f"Gemini недоступен (код {status}).") from exc
                raise RuntimeError(f"Gemini недоступен: {exc}") from exc

        if resp is None:  # обе попытки не дали ответа
            raise RuntimeError(f"Gemini недоступен: {last_exc}")

        text = (getattr(resp, "text", None) or "").strip()
        if not text:  # запасной путь: собираем текст из частей ответа
            chunks_out = []
            for cand in getattr(resp, "candidates", None) or []:
                content = getattr(cand, "content", None)
                for part in getattr(content, "parts", None) or []:
                    if getattr(part, "text", None):
                        chunks_out.append(part.text)
            text = "".join(chunks_out).strip()
        if not text:
            raise RuntimeError(
                "Gemini вернул пустой ответ (возможно, сработали фильтры безопасности)."
            )
        return text

    async def complete(self, prompt: str, **kw: Any) -> str:
        return await asyncio.to_thread(self._complete, prompt, **kw)

    @staticmethod
    def _extract_json(text: str) -> dict:
        text = re.sub(
            r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE
        ).strip()
        start, end = text.find("{"), text.rfind("}")
        if start == -1 or end == -1:
            raise ValueError("Gemini вернул ответ без JSON")
        return json.loads(text[start : end + 1])

    async def parse_excel(self, raw_rows: list[list[Any]]) -> dict:
        payload = json_dumps({"rows": raw_rows})
        text = await self.complete(
            f"{EXCEL_PARSE_PROMPT}\n\nДанные:\n{payload}",
            max_tokens=8000,
            json_mode=True,
        )
        return await asyncio.to_thread(self._extract_json, text)

    async def quick_parse(
        self, text: str, artists: list[str], today: date
    ) -> dict:
        """Разбор операции, написанной человеческим языком."""
        prompt = (
            "Пользователь диктует бухгалтеру операцию в свободной форме.\n"
            f"Сегодня: {today.isoformat()}.\n"
            f"Артисты в базе: {json_dumps(artists)}.\n\n"
            f"Сообщение: {text!r}\n\n"
            "Верни СТРОГО JSON без пояснений:\n"
            "{\"understood\": true|false, \"kind\": \"income\"|\"expense\", "
            "\"amount\": число, \"category\": строка, \"description\": строка, "
            "\"artist\": строка|null, \"occurred_on\": \"YYYY-MM-DD\", "
            "\"confidence\": 0..1, \"note\": короткий комментарий бухгалтера}\n\n"
            "Правила:\n"
            "• если это не финансовая операция — understood=false;\n"
            "• «зашло», «пришло», «роялти», «выплата от дистрибьютора» — income;\n"
            "• «потратил», «заплатил», «минус», «реклама», «студия» — expense;\n"
            "• понимай «вчера», «позавчера», «5 числа», «в прошлом месяце»;\n"
            "• без даты — ставь сегодняшнюю;\n"
            "• имя артиста бери только из списка выше (ближайшее совпадение), иначе null;\n"
            "• категория — короткое слово: роялти, концерт, мерч, реклама, студия, "
            "клип, продвижение, налоги, прочее;\n"
            "• сумма — число без пробелов и валюты; «15к» = 15000."
        )
        raw = await self.complete(prompt, max_tokens=800, json_mode=True)
        return await asyncio.to_thread(self._extract_json, raw)

    async def ask(self, question: str, data: dict) -> str:
        prompt = (
            "Ты главный бухгалтер музыкального лейбла. Ответь на вопрос владельца, "
            "опираясь ТОЛЬКО на данные ниже.\n\n"
            f"Вопрос: {question}\n\n"
            f"Данные (JSON):\n{json_dumps(data)}\n\n"
            "Правила ответа:\n"
            "• сначала короткий прямой ответ с цифрами;\n"
            "• потом 2-4 пункта расшифровки и что с этим делать;\n"
            "• если данных не хватает — честно скажи, чего не хватает, не выдумывай;\n"
            "• все суммы — в рублях с разделителями разрядов;\n"
            "• формат: HTML-теги Telegram (<b>, <i>, <code>), без markdown, до 3000 символов."
        )
        return await self.complete(prompt, max_tokens=2500)

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
# -----------------------------------------------------------------------------
# Роли и права
#   owner  — всё + управление доступом и ролями
#   admin  — всё по финансам + приглашение наблюдателей
#   viewer — только чтение: сводки, отчёты, аномалии, вопросы ИИ
# -----------------------------------------------------------------------------
ROLE_OWNER = "owner"
ROLE_ADMIN = "admin"
ROLE_VIEWER = "viewer"
ROLE_TITLES = {
    ROLE_OWNER: "👑 владелец",
    ROLE_ADMIN: "🛠 админ",
    ROLE_VIEWER: "👀 наблюдатель",
}
# Роль текущего апдейта: меню рисуется по правам без передачи аргументов.
CURRENT_ROLE: ContextVar[str] = ContextVar("current_role", default=ROLE_VIEWER)

# Действия, меняющие данные — недоступны наблюдателю.
WRITE_GATED = {
    ("artist", "add"),
    ("tx", "new"),
    ("tx", "artist"),
    ("tx", "undo"),
    ("tx", "drop"),
    ("excel", "wait"),
    ("pay", "paid"),
    ("quick", "hint"),
    ("quick", "save"),
}


def can_write(role: str | None = None) -> bool:
    return (role or CURRENT_ROLE.get()) in (ROLE_OWNER, ROLE_ADMIN)


def can_manage(role: str | None = None) -> bool:
    return (role or CURRENT_ROLE.get()) in (ROLE_OWNER, ROLE_ADMIN)


def user_title(u: dict) -> str:
    if u.get("full_name"):
        return str(u["full_name"])
    if u.get("username"):
        return "@" + str(u["username"])
    return f"id {u['tg_user_id']}"


def kb_main() -> InlineKeyboardMarkup:
    role = CURRENT_ROLE.get()
    write = can_write(role)
    rows: list[list[InlineKeyboardButton]] = [
        [InlineKeyboardButton("💬 Спросить бухгалтера", callback_data="ai:ask")]
    ]
    if write:
        rows.append(
            [
                InlineKeyboardButton("➕ Доход", callback_data="tx:new:income"),
                InlineKeyboardButton("➖ Расход", callback_data="tx:new:expense"),
            ]
        )
    rows.append(
        [
            InlineKeyboardButton("📈 Сводка", callback_data="dash:now"),
            InlineKeyboardButton("📊 Отчёт за месяц", callback_data="report:menu"),
        ]
    )
    rows.append(
        [
            InlineKeyboardButton("🚨 Аномалии", callback_data="anomaly:run"),
            InlineKeyboardButton("💸 Платежи", callback_data="pay:list:all"),
        ]
    )
    if write:
        rows.append(
            [
                InlineKeyboardButton("📥 Загрузить Excel", callback_data="excel:wait"),
                InlineKeyboardButton("📤 Экспорт CSV", callback_data="export:menu"),
            ]
        )
    else:
        rows.append(
            [InlineKeyboardButton("📤 Экспорт CSV", callback_data="export:menu")]
        )
    artist_row = [InlineKeyboardButton("🎤 Артисты", callback_data="artist:list")]
    if write:
        artist_row.append(
            InlineKeyboardButton("👤 Новый артист", callback_data="artist:add")
        )
    rows.append(artist_row)
    if write:
        rows.append(
            [
                InlineKeyboardButton(
                    "↩️ Отменить последнюю операцию", callback_data="tx:undo"
                )
            ]
        )
    if can_manage(role):
        rows.append(
            [InlineKeyboardButton("🔑 Доступ и люди", callback_data="access:menu")]
        )
    return InlineKeyboardMarkup(rows)


def kb_access(role: str) -> InlineKeyboardMarkup:
    rows = [
        [InlineKeyboardButton("👥 Кто имеет доступ", callback_data="access:users")],
        [
            InlineKeyboardButton(
                "🎫 Код для наблюдателя", callback_data="access:invite:viewer"
            )
        ],
    ]
    if role == ROLE_OWNER:
        rows.append(
            [
                InlineKeyboardButton(
                    "🎫 Код для админа", callback_data="access:invite:admin"
                )
            ]
        )
    rows.append(
        [InlineKeyboardButton("🔗 Активные коды", callback_data="access:codes")]
    )
    rows.append(
        [
            InlineKeyboardButton(
                "➕ Выдать по Telegram ID", callback_data="access:grant"
            )
        ]
    )
    rows.append(
        [InlineKeyboardButton("⬅️ Главное меню", callback_data="nav:main")]
    )
    return InlineKeyboardMarkup(rows)


def kb_users(users: list[dict]) -> InlineKeyboardMarkup:
    rows = []
    for u in users[:40]:
        mark = "" if int(u["is_active"]) else "🚫 "
        short = ROLE_TITLES[u["role"]].split()[-1]
        rows.append(
            [
                InlineKeyboardButton(
                    f"{mark}{user_title(u)} — {short}",
                    callback_data=f"access:user:{u['tg_user_id']}",
                )
            ]
        )
    rows.append([InlineKeyboardButton("⬅️ Назад", callback_data="access:menu")])
    return InlineKeyboardMarkup(rows)


def manageable(row: dict, viewer_role: str, me: int) -> bool:
    if int(row["tg_user_id"]) == me:
        return False
    if row["role"] == ROLE_OWNER:
        return False
    return viewer_role == ROLE_OWNER or row["role"] == ROLE_VIEWER


def kb_user(row: dict, viewer_role: str, me: int) -> InlineKeyboardMarkup:
    uid = int(row["tg_user_id"])
    rows: list[list[InlineKeyboardButton]] = []
    if manageable(row, viewer_role, me):
        if viewer_role == ROLE_OWNER:
            if row["role"] == ROLE_ADMIN:
                rows.append(
                    [
                        InlineKeyboardButton(
                            "👀 Сделать наблюдателем",
                            callback_data=f"access:role:{uid}:viewer",
                        )
                    ]
                )
            else:
                rows.append(
                    [
                        InlineKeyboardButton(
                            "🛠 Сделать админом",
                            callback_data=f"access:role:{uid}:admin",
                        )
                    ]
                )
        if int(row["is_active"]):
            rows.append(
                [
                    InlineKeyboardButton(
                        "🚫 Забрать доступ", callback_data=f"access:block:{uid}"
                    )
                ]
            )
        else:
            rows.append(
                [
                    InlineKeyboardButton(
                        "✅ Вернуть доступ", callback_data=f"access:unblock:{uid}"
                    )
                ]
            )
        if viewer_role == ROLE_OWNER:
            rows.append(
                [
                    InlineKeyboardButton(
                        "🗑 Удалить из списка",
                        callback_data=f"access:delete:{uid}",
                    )
                ]
            )
    rows.append([InlineKeyboardButton("⬅️ К списку", callback_data="access:users")])
    return InlineKeyboardMarkup(rows)


def kb_codes(invites: list[dict]) -> InlineKeyboardMarkup:
    rows = []
    for inv in invites[:20]:
        rows.append(
            [
                InlineKeyboardButton(
                    f"❌ Отозвать {inv['code']}",
                    callback_data=f"access:kill:{inv['code']}",
                )
            ]
        )
    rows.append([InlineKeyboardButton("⬅️ Назад", callback_data="access:menu")])
    return InlineKeyboardMarkup(rows)


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
def _role_from_env(user_id: int) -> str | None:
    """Поддержка старых переменных: OWNER_IDS и ALLOWED_USER_IDS."""
    if user_id in Config.OWNER_IDS:
        return ROLE_OWNER
    if user_id in Config.ALLOWED_USER_IDS:
        return ROLE_ADMIN
    return None


async def resolve_access(update: Update, repo: "Repo") -> dict | None:
    """Кто пишет боту и с какой ролью. None — доступа нет."""
    user = update.effective_user
    if not user:
        return None

    row = await asyncio.to_thread(
        repo.touch_user, user.id, user.username, user.full_name
    )

    if row is None:
        role = _role_from_env(user.id)
        if role is None and not Config.OWNER_IDS and not Config.ALLOWED_USER_IDS:
            # Первый запуск без OWNER_IDS: первый человек становится владельцем.
            if await asyncio.to_thread(repo.users_count) == 0:
                role = ROLE_OWNER
                log.warning(
                    "Первый пользователь %s назначен владельцем бота", user.id
                )
        if role is None:
            return None
        row = await asyncio.to_thread(
            repo.grant, user.id, role, None, user.username, user.full_name
        )

    if row is None or not int(row["is_active"]):
        return None

    CURRENT_ROLE.set(row["role"])
    return row


async def handle_locked_message(
    update: Update, repo: "Repo", text: str
) -> None:
    """Сообщение от человека без доступа: пробуем текст как код-приглашение."""
    user = update.effective_user
    candidate = text.strip()
    if candidate.lower().startswith("/start"):  # диплинк t.me/<bot>?start=CODE
        candidate = candidate[6:].strip()
    candidate = candidate.replace("-", "").replace(" ", "").upper()

    if 6 <= len(candidate) <= 12 and candidate.isalnum():
        row = await asyncio.to_thread(
            repo.redeem_invite, candidate, user.id, user.username, user.full_name
        )
        if row:
            CURRENT_ROLE.set(row["role"])
            await show_main_menu(
                update,
                f"✅ Доступ открыт. Ваша роль: {ROLE_TITLES[row['role']]}.",
            )
            return
        await update.effective_chat.send_message(
            "⚠️ Код не подходит: неверный, просроченный или уже использованный."
        )
        return

    await update.effective_chat.send_message(
        "🔒 <b>Закрытый бот</b>\n\n"
        "Попросите у владельца код-приглашение и пришлите его сюда одним "
        "сообщением, например <code>K7P2M9QA</code>.\n\n"
        f"Ваш Telegram ID: <code>{user.id}</code>",
        parse_mode=ParseMode.HTML,
    )


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


def prev_month(year: int, month: int) -> tuple[int, int]:
    return (year - 1, 12) if month == 1 else (year, month - 1)


def kb_export_months() -> InlineKeyboardMarkup:
    today = date.today()
    rows, row = [], []
    for offset in range(6):
        month, year = today.month - offset, today.year
        while month <= 0:
            month += 12
            year -= 1
        row.append(
            InlineKeyboardButton(
                f"{month:02d}.{year}", callback_data=f"export:month:{year}:{month}"
            )
        )
        if len(row) == 3:
            rows.append(row)
            row = []
    if row:
        rows.append(row)
    rows.append(
        [InlineKeyboardButton("⬅️ Главное меню", callback_data="nav:main")]
    )
    return InlineKeyboardMarkup(rows)


def pct_delta(now: Decimal, before: Decimal) -> str:
    if before == 0:
        return "новое" if now else "—"
    change = (now - before) / before * 100
    arrow = "📈" if change >= 0 else "📉"
    return f"{arrow} {change:+.0f}%"


def dashboard_text(
    year: int,
    month: int,
    cur: dict,
    before: dict,
    breakdown: list[dict],
    cats: list[dict],
) -> str:
    cur_currency = Config.DEFAULT_CURRENCY
    to_artists = sum(
        (Decimal(str(r.get("artist_share") or 0)) for r in breakdown), Decimal("0")
    )
    label_profit = sum(
        (Decimal(str(r.get("label_profit") or 0)) for r in breakdown), Decimal("0")
    )
    margin = (cur["net"] / cur["revenue"] * 100) if cur["revenue"] else Decimal("0")
    tax = (cur["revenue"] * Config.TAX_RATE).quantize(Decimal("0.01"))
    day = date.today().day
    burn = (cur["expenses"] / day) if day else Decimal("0")

    top_expense = sorted(
        [c for c in cats if str(c.get("kind")) == "expense"],
        key=lambda c: Decimal(str(c.get("total") or 0)),
        reverse=True,
    )[:3]
    top_artists = sorted(
        breakdown,
        key=lambda r: Decimal(str(r.get("label_profit") or 0)),
        reverse=True,
    )[:3]

    lines = [
        f"📈 <b>Сводка за {month:02d}.{year}</b>",
        "",
        f"Выручка: <b>{money(cur['revenue'], cur_currency)}</b> "
        f"({pct_delta(cur['revenue'], before['revenue'])} к прошлому месяцу)",
        f"Расходы: <b>{money(cur['expenses'], cur_currency)}</b> "
        f"({pct_delta(cur['expenses'], before['expenses'])})",
        f"Чистыми: <b>{money(cur['net'], cur_currency)}</b> • маржа {margin:.0f}%",
        "",
        f"К выплате артистам: <b>{money(to_artists, cur_currency)}</b>",
        f"Остаётся лейблу: <b>{money(label_profit, cur_currency)}</b>",
        f"Налоговый резерв ({Config.TAX_RATE * 100:.0f}% с дохода): "
        f"<b>{money(tax, cur_currency)}</b>",
        f"Средний расход в день: {money(burn.quantize(Decimal('0.01')), cur_currency)}",
        f"Операций за месяц: {cur['count']}",
    ]
    if top_expense:
        lines += ["", "<b>Куда уходят деньги:</b>"]
        lines += [
            f"• {c.get('category') or 'без категории'} — "
            f"{money(c.get('total'), cur_currency)}"
            for c in top_expense
        ]
    if top_artists:
        lines += ["", "<b>Лучшие по прибыли лейбла:</b>"]
        lines += [
            f"• {r.get('artist')} — {money(r.get('label_profit'), cur_currency)}"
            for r in top_artists
        ]
    return "\n".join(lines)


def transactions_csv(rows: list[dict]) -> bytes:
    buf = StringIO()
    writer = csv.writer(buf, delimiter=";")
    writer.writerow(
        ["id", "дата", "тип", "сумма", "валюта", "артист", "категория", "описание", "источник"]
    )
    for r in rows:
        writer.writerow(
            [
                r["id"],
                str(r["occurred_on"])[:10],
                "доход" if r["kind"] == "income" else "расход",
                f"{r['amount']:.2f}".replace(".", ","),
                r.get("currency") or Config.DEFAULT_CURRENCY,
                r.get("artist") or "",
                r.get("category") or "",
                r.get("description") or "",
                r.get("source") or "",
            ]
        )
    return buf.getvalue().encode("utf-8-sig")


async def collect_books(repo: "Repo", months: int = 3) -> dict:
    """Компактный срез базы для вопросов ИИ-бухгалтеру."""
    today = date.today()
    year, month = today.year, today.month
    py, pm = prev_month(year, month)
    cur_start, cur_end = Repo.month_bounds(year, month)
    prev_start, prev_end = Repo.month_bounds(py, pm)

    (
        cur,
        before,
        breakdown,
        cats,
        artists,
        recent,
        anomalies,
    ) = await asyncio.gather(
        asyncio.to_thread(repo.totals, cur_start, cur_end),
        asyncio.to_thread(repo.totals, prev_start, prev_end),
        asyncio.to_thread(repo.monthly_breakdown, year, month),
        asyncio.to_thread(repo.category_breakdown, year, month),
        asyncio.to_thread(repo.list_artists),
        asyncio.to_thread(repo.recent_transactions, 25),
        asyncio.to_thread(repo.anomaly_candidates, months),
    )
    return {
        "сегодня": today.isoformat(),
        "валюта": Config.DEFAULT_CURRENCY,
        "текущий_месяц": {"период": f"{month:02d}.{year}", **cur},
        "прошлый_месяц": {"период": f"{pm:02d}.{py}", **before},
        "по_артистам": breakdown,
        "по_категориям": cats,
        "артисты": [
            {"имя": a["name"], "ставка": str(a["rate"])} for a in artists
        ],
        "последние_операции": [
            {
                "дата": str(r["occurred_on"])[:10],
                "тип": r["kind"],
                "сумма": str(r["amount"]),
                "артист": r.get("artist"),
                "категория": r.get("category"),
                "описание": r.get("description"),
            }
            for r in recent
        ],
        "аномалии_и_целостность": anomalies,
        "налоговая_ставка": str(Config.TAX_RATE),
    }


async def quick_capture(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE,
    repo: "Repo",
    ai: GeminiService,
    text: str,
) -> None:
    """Любое сообщение в свободной форме → черновик операции с подтверждением."""
    artists = await asyncio.to_thread(repo.list_artists)
    await update.effective_chat.send_action(ChatAction.TYPING)
    try:
        parsed = await ai.quick_parse(text, [a["name"] for a in artists], date.today())
    except Exception as exc:
        log.warning("quick_parse failed: %s", exc)
        await show_main_menu(update, f"⚠️ Не смог разобрать сообщение: {exc}")
        return

    if not parsed.get("understood") or not parsed.get("amount"):
        await show_main_menu(
            update,
            "🤔 Это не похоже на операцию. Напишите, например: "
            "<code>вчера зашло 120к роялти Spotify за MACAN</code> "
            "или <code>минус 15000 реклама в телеге</code>.",
        )
        return

    try:
        amount = parse_amount(str(parsed["amount"]))
    except ValueError as exc:
        await show_main_menu(update, f"⚠️ Сумма непонятна: {exc}")
        return

    try:
        occurred = parse_date(str(parsed.get("occurred_on") or ""))
    except ValueError:
        occurred = date.today()

    kind = "income" if str(parsed.get("kind")) == "income" else "expense"
    artist_name = parsed.get("artist") or None
    context.user_data["quick"] = {
        "kind": kind,
        "amount": str(amount),
        "category": parsed.get("category") or None,
        "description": (parsed.get("description") or text)[:400],
        "artist": artist_name,
        "occurred_on": occurred.isoformat(),
    }

    sign = "➕ Доход" if kind == "income" else "➖ Расход"
    note = parsed.get("note")
    confidence = parsed.get("confidence")
    tail = f"\n\n<i>{note}</i>" if note else ""
    if isinstance(confidence, (int, float)) and confidence < 0.6:
        tail += "\n⚠️ Не всё однозначно — проверьте перед сохранением."

    await send(
        update,
        f"🧾 <b>Проверьте операцию</b>\n\n"
        f"{sign}: <b>{money(amount, Config.DEFAULT_CURRENCY)}</b>\n"
        f"Дата: {occurred:%d.%m.%Y}\n"
        f"Артист: {artist_name or '—'}\n"
        f"Категория: {parsed.get('category') or '—'}\n"
        f"Описание: {parsed.get('description') or '—'}{tail}",
        InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton("✅ Сохранить", callback_data="quick:save"),
                    InlineKeyboardButton("❌ Отмена", callback_data="nav:main"),
                ]
            ]
        ),
    )


# -----------------------------------------------------------------------------
# Текстовые сообщения (конечный автомат, без slash-команд)
# -----------------------------------------------------------------------------
async def on_message(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    repo: Repo = context.bot_data["repo"]
    ai: GeminiService = context.bot_data["ai"]
    text = (update.effective_message.text or "").strip()

    access = await resolve_access(update, repo)
    if access is None:
        await handle_locked_message(update, repo, text)
        return

    state: str | None = context.user_data.get("state")

    if not state:
        # Без активного диалога: пробуем понять сообщение как операцию или вопрос.
        if len(text) < 4:
            await show_main_menu(update)
            return
        if not can_write(access["role"]):
            await update.effective_chat.send_action(ChatAction.TYPING)
            try:
                books = await collect_books(repo)
                answer = await ai.ask(text, books)
            except Exception as exc:
                await show_main_menu(update, f"⚠️ Не смог ответить: {exc}")
                return
            await send(update, answer, kb_main())
            return
        if any(ch.isdigit() for ch in text):
            await quick_capture(update, context, repo, ai, text)
            return
        await update.effective_chat.send_action(ChatAction.TYPING)
        try:
            books = await collect_books(repo)
            answer = await ai.ask(text, books)
        except Exception as exc:
            await show_main_menu(update, f"⚠️ Не смог ответить: {exc}")
            return
        await send(update, answer, kb_main())
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

        elif state == "ai:ask":
            reset_state(context)
            await update.effective_chat.send_action(ChatAction.TYPING)
            books = await collect_books(repo)
            answer = await ai.ask(text, books)
            await send(
                update,
                f"💬 <b>Вопрос:</b> {text}\n\n{answer}",
                kb_main(),
            )

        elif state == "access:grant":
            if not can_manage(access["role"]):
                raise ValueError("Нет прав на управление доступом.")
            fields = text.replace(",", " ").split()
            if not fields or not fields[0].lstrip("-").isdigit():
                raise ValueError(
                    "Нужен числовой Telegram ID, например 123456789 или 123456789 admin."
                )
            target_id = int(fields[0])
            new_role = fields[1].lower() if len(fields) > 1 else ROLE_VIEWER
            if new_role not in (ROLE_ADMIN, ROLE_VIEWER):
                raise ValueError("Роль может быть только admin или viewer.")
            if new_role == ROLE_ADMIN and access["role"] != ROLE_OWNER:
                raise ValueError("Админов назначает только владелец.")
            existing = await asyncio.to_thread(repo.get_user, target_id)
            if existing and existing["role"] == ROLE_OWNER:
                raise ValueError("Роль владельца менять нельзя.")
            await asyncio.to_thread(
                repo.grant, target_id, new_role, update.effective_user.id
            )
            await asyncio.to_thread(
                repo.audit,
                update.effective_user.id,
                "access_grant",
                {"target": target_id, "role": new_role},
            )
            reset_state(context)
            await show_main_menu(
                update,
                f"✅ Доступ выдан: <code>{target_id}</code> — {ROLE_TITLES[new_role]}.\n"
                "Попросите человека написать боту любое сообщение.",
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

    repo: Repo = context.bot_data["repo"]
    ai: GeminiService = context.bot_data["ai"]

    access = await resolve_access(update, repo)
    if access is None:
        await query.edit_message_text(
            "⛔ Доступ закрыт. Попросите код-приглашение и пришлите его сообщением."
        )
        return
    role: str = access["role"]
    me: int = int(access["tg_user_id"])

    parts = (query.data or "").split(":")
    domain, action = parts[0], parts[1] if len(parts) > 1 else ""
    args = parts[2:]

    if (domain, action) in WRITE_GATED and not can_write(role):
        await send(update, "⛔ У вас доступ только для просмотра.", kb_main())
        return
    if domain == "access" and not can_manage(role):
        await send(
            update, "⛔ Управлять доступом могут владелец и админ.", kb_main()
        )
        return

    # ---------- навигация ----------
    if domain == "nav":
        reset_state(context)
        await show_main_menu(update)
        return

    # ---------- ИИ-бухгалтер: свободный вопрос ----------
    if domain == "ai":
        context.user_data["state"] = "ai:ask"
        await send(
            update,
            "💬 <b>Спросите ИИ-бухгалтера</b>\n\n"
            "Он видит все ваши операции, артистов, ставки и аномалии. Примеры:\n"
            "• <i>Сколько я должен артистам в этом месяце?</i>\n"
            "• <i>Почему расходы выросли?</i>\n"
            "• <i>Какой артист убыточен и почему?</i>\n"
            "• <i>Сколько отложить на налоги?</i>\n\n"
            "Напишите вопрос обычным сообщением.",
            kb_back(),
        )
        return

    # ---------- быстрый ввод из свободного текста ----------
    if domain == "quick":
        draft = context.user_data.get("quick")
        if action != "save" or not draft:
            reset_state(context)
            await show_main_menu(update, "Черновик устарел, напишите операцию заново.")
            return
        artist_id = None
        if draft.get("artist"):
            mapping = await asyncio.to_thread(
                repo.resolve_artist_ids, [draft["artist"]]
            )
            artist_id = mapping.get(str(draft["artist"]).strip().lower()) or next(
                iter(mapping.values()), None
            )
        amount = Decimal(draft["amount"])
        occurred = date.fromisoformat(draft["occurred_on"])
        await asyncio.to_thread(
            repo.add_transaction,
            artist_id,
            draft["kind"],
            amount,
            Config.DEFAULT_CURRENCY,
            draft.get("category"),
            draft.get("description"),
            occurred,
            "quick",
            None,
            me,
        )
        context.user_data.pop("quick", None)
        label = "Доход" if draft["kind"] == "income" else "Расход"
        await show_main_menu(
            update,
            f"✅ {label} {money(amount, Config.DEFAULT_CURRENCY)} "
            f"от {occurred:%d.%m.%Y} записан.",
        )
        return

    # ---------- сводка ----------
    if domain == "dash":
        today = date.today()
        year, month = today.year, today.month
        py, pm = prev_month(year, month)
        cur_start, cur_end = Repo.month_bounds(year, month)
        prev_start, prev_end = Repo.month_bounds(py, pm)
        cur, before, breakdown, cats = await asyncio.gather(
            asyncio.to_thread(repo.totals, cur_start, cur_end),
            asyncio.to_thread(repo.totals, prev_start, prev_end),
            asyncio.to_thread(repo.monthly_breakdown, year, month),
            asyncio.to_thread(repo.category_breakdown, year, month),
        )
        await send(
            update,
            dashboard_text(year, month, cur, before, breakdown, cats),
            kb_back(
                [
                    [
                        InlineKeyboardButton(
                            "🤖 Подробный разбор от ИИ",
                            callback_data=f"report:run:{year}:{month}",
                        )
                    ],
                    [
                        InlineKeyboardButton(
                            "💬 Спросить бухгалтера", callback_data="ai:ask"
                        )
                    ],
                ]
            ),
        )
        return

    # ---------- экспорт в CSV ----------
    if domain == "export":
        if action == "menu":
            await send(
                update,
                "📤 <b>Экспорт операций</b>\n"
                "Выберите месяц — пришлю файл CSV для Excel или бухгалтерии.",
                kb_export_months(),
            )
            return
        if action == "month":
            year, month = int(args[0]), int(args[1])
            start, end = Repo.month_bounds(year, month)
            rows = await asyncio.to_thread(repo.period_transactions, start, end)
            if not rows:
                await send(update, "За этот месяц операций нет.", kb_export_months())
                return
            payload = await asyncio.to_thread(transactions_csv, rows)
            await update.effective_chat.send_document(
                document=InputFile(
                    BytesIO(payload), filename=f"finbot-{year}-{month:02d}.csv"
                ),
                caption=f"📤 Операции за {month:02d}.{year}: {len(rows)} шт.",
            )
            await show_main_menu(update)
            return
        return

    # ---------- доступ и люди ----------
    if domain == "access":
        if action == "menu":
            users, invites = await asyncio.gather(
                asyncio.to_thread(repo.list_users),
                asyncio.to_thread(repo.list_invites),
            )
            active = [u for u in users if int(u["is_active"])]
            owners = [u for u in active if u["role"] == ROLE_OWNER]
            admins = [u for u in active if u["role"] == ROLE_ADMIN]
            viewers = [u for u in active if u["role"] == ROLE_VIEWER]
            await send(
                update,
                "🔑 <b>Доступ и люди</b>\n"
                f"Владельцы: <b>{len(owners)}</b>, админы: <b>{len(admins)}</b>, "
                f"наблюдатели: <b>{len(viewers)}</b>\n"
                f"Активных кодов-приглашений: <b>{len(invites)}</b>\n\n"
                "Кто что может:\n"
                "👑 <b>владелец</b> — всё, включая роли и удаление людей\n"
                "🛠 <b>админ</b> — вести финансы и звать наблюдателей\n"
                "👀 <b>наблюдатель</b> — только смотреть цифры и спрашивать ИИ",
                kb_access(role),
            )
            return

        if action == "users":
            users = await asyncio.to_thread(repo.list_users)
            if not users:
                await send(update, "Пока никого нет.", kb_access(role))
                return
            lines = []
            for u in users:
                mark = "" if int(u["is_active"]) else " 🚫 заблокирован"
                lines.append(
                    f"• <b>{user_title(u)}</b> — {ROLE_TITLES[u['role']]}{mark}\n"
                    f"  ID: <code>{u['tg_user_id']}</code>"
                )
            await send(
                update,
                "👥 <b>Кто имеет доступ</b>\n" + "\n".join(lines),
                kb_users(users),
            )
            return

        if action == "user":
            target = await asyncio.to_thread(repo.get_user, int(args[0]))
            if not target:
                await send(update, "Пользователь не найден.", kb_access(role))
                return
            status = "активен" if int(target["is_active"]) else "заблокирован"
            seen = target.get("last_seen_at") or "—"
            note = (
                ""
                if manageable(target, role, me)
                else "\n\n<i>Этого человека вы изменить не можете.</i>"
            )
            await send(
                update,
                f"👤 <b>{user_title(target)}</b>\n"
                f"Роль: {ROLE_TITLES[target['role']]}\n"
                f"Статус: {status}\n"
                f"ID: <code>{target['tg_user_id']}</code>\n"
                f"Последняя активность: {seen}{note}",
                kb_user(target, role, me),
            )
            return

        if action in ("role", "block", "unblock", "delete"):
            target = await asyncio.to_thread(repo.get_user, int(args[0]))
            if not target:
                await send(update, "Пользователь не найден.", kb_access(role))
                return
            if not manageable(target, role, me):
                await send(update, "⛔ Недостаточно прав для этого действия.", kb_access(role))
                return
            uid = int(target["tg_user_id"])

            if action == "role":
                new_role = args[1]
                if new_role not in (ROLE_ADMIN, ROLE_VIEWER) or role != ROLE_OWNER:
                    await send(update, "⛔ Роли меняет только владелец.", kb_access(role))
                    return
                await asyncio.to_thread(repo.set_role, uid, new_role)
                text_out = f"✅ Роль изменена: {ROLE_TITLES[new_role]}"
            elif action == "block":
                await asyncio.to_thread(repo.set_active, uid, False)
                text_out = "🚫 Доступ забран"
            elif action == "unblock":
                await asyncio.to_thread(repo.set_active, uid, True)
                text_out = "✅ Доступ возвращён"
            else:
                await asyncio.to_thread(repo.delete_user, uid)
                text_out = "🗑 Удалён из списка"

            await asyncio.to_thread(
                repo.audit, me, f"access_{action}", {"target": uid, "args": args}
            )
            users = await asyncio.to_thread(repo.list_users)
            await send(
                update,
                f"{text_out}: <b>{user_title(target)}</b>",
                kb_users(users),
            )
            return

        if action == "invite":
            invite_role = args[0] if args else ROLE_VIEWER
            if invite_role == ROLE_ADMIN and role != ROLE_OWNER:
                await send(update, "⛔ Код админа выдаёт только владелец.", kb_access(role))
                return
            inv = await asyncio.to_thread(
                repo.create_invite,
                invite_role,
                Config.INVITE_DEFAULT_USES,
                Config.INVITE_TTL_HOURS,
                me,
            )
            await asyncio.to_thread(repo.audit, me, "access_invite_create", inv)
            bot_username = getattr(context.bot, "username", None)
            link = (
                f"\nСсылка: <code>https://t.me/{bot_username}?start={inv['code']}</code>"
                if bot_username
                else ""
            )
            ttl = (
                f"действует {Config.INVITE_TTL_HOURS} ч"
                if Config.INVITE_TTL_HOURS > 0
                else "без срока"
            )
            await send(
                update,
                "🎫 <b>Код-приглашение готов</b>\n\n"
                f"<code>{inv['code']}</code>\n\n"
                f"Роль: {ROLE_TITLES[invite_role]}\n"
                f"Активаций: {inv['uses_left']}, {ttl}{link}\n\n"
                "Перешлите код человеку — он пришлёт его боту сообщением и сразу получит доступ.",
                kb_access(role),
            )
            return

        if action == "codes":
            invites = await asyncio.to_thread(repo.list_invites)
            if not invites:
                await send(update, "Активных кодов нет.", kb_access(role))
                return
            lines = [
                f"• <code>{i['code']}</code> — {ROLE_TITLES[i['role']]}, "
                f"активаций осталось: {i['uses_left']}"
                + (f", до {str(i['expires_at'])[:16].replace('T', ' ')}" if i["expires_at"] else "")
                for i in invites
            ]
            await send(
                update,
                "🔗 <b>Активные коды</b>\n" + "\n".join(lines),
                kb_codes(invites),
            )
            return

        if action == "kill":
            code = args[0] if args else ""
            await asyncio.to_thread(repo.revoke_invite, code)
            await asyncio.to_thread(repo.audit, me, "access_invite_revoke", {"code": code})
            invites = await asyncio.to_thread(repo.list_invites)
            await send(update, f"❌ Код <code>{code}</code> отозван.", kb_codes(invites))
            return

        if action == "grant":
            context.user_data["state"] = "access:grant"
            await send(
                update,
                "➕ <b>Выдать доступ по Telegram ID</b>\n\n"
                "Пришлите числовой ID, например <code>123456789</code> "
                "— по умолчанию роль наблюдателя.\n"
                "Чтобы сразу сделать админом: <code>123456789 admin</code>\n\n"
                "Свой ID человек увидит, если напишет этому боту любое сообщение.",
                kb_back(),
            )
            return
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
        elif action == "undo":
            last = await asyncio.to_thread(repo.last_transaction_by, me)
            if not last:
                await send(update, "Вы ещё не добавляли операций.", kb_main())
                return
            label = "Доход" if last["kind"] == "income" else "Расход"
            await send(
                update,
                "↩️ <b>Отмена последней операции</b>\n\n"
                f"#{last['id']} • {label} "
                f"{money(last['amount'], Config.DEFAULT_CURRENCY)}\n"
                f"Дата: {str(last['occurred_on'])[:10]}\n"
                f"Артист: {last.get('artist') or '—'}\n"
                f"Категория: {last.get('category') or '—'}\n\n"
                "Удалить её?",
                InlineKeyboardMarkup(
                    [
                        [
                            InlineKeyboardButton(
                                "🗑 Удалить",
                                callback_data=f"tx:drop:{last['id']}",
                            ),
                            InlineKeyboardButton(
                                "❌ Оставить", callback_data="nav:main"
                            ),
                        ]
                    ]
                ),
            )
        elif action == "drop":
            tx_id = int(args[0])
            last = await asyncio.to_thread(repo.last_transaction_by, me)
            if not last or int(last["id"]) != tx_id:
                await show_main_menu(
                    update, "⚠️ Можно удалить только свою последнюю операцию."
                )
                return
            await asyncio.to_thread(repo.delete_transaction, tx_id)
            await asyncio.to_thread(repo.audit, me, "tx_delete", {"id": tx_id})
            await show_main_menu(update, f"🗑 Операция #{tx_id} удалена.")
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
                ai_text = await ai.monthly_report(year, month, breakdown, categories)
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
            await send(update, f"🤖 <b>Аналитика Gemini</b>\n\n{ai_text}", kb_main())
        return

    # ---------- Excel ----------
    if domain == "excel":
        context.user_data["state"] = None
        await send(
            update,
            "📥 Пришлите файл <code>.xlsx</code> или <code>.xls</code> как документ.\n\n"
            "Ожидаемые колонки (в любом порядке и на любом языке): дата, артист, "
            "тип операции, сумма, категория, описание. Gemini сам разберёт "
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
            text = await ai.anomaly_report(data)
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
    repo: Repo = context.bot_data["repo"]
    ai: GeminiService = context.bot_data["ai"]

    access = await resolve_access(update, repo)
    if access is None:
        return
    if not can_write(access["role"]):
        await send(update, "⛔ Загружать файлы могут владелец и админ.", kb_main())
        return
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

    await send(update, f"🤖 Отправляю {len(raw_rows)} строк в Gemini на разбор…")

    try:
        parsed = await ai.parse_excel(raw_rows)
    except (RuntimeError, ValueError, json.JSONDecodeError) as exc:
        log.exception("Gemini parse failed")
        await send(update, f"⚠️ Gemini не смог разобрать файл: {exc}", kb_main())
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
        name for name in ("BOT_TOKEN", "GEMINI_API_KEY") if not getattr(Config, name)
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

    if Config.OWNER_IDS:
        repo.ensure_owners(Config.OWNER_IDS)
        log.info("Владельцы из OWNER_IDS: %s", sorted(Config.OWNER_IDS))
    elif repo.users_count() == 0:
        log.warning(
            "OWNER_IDS не задан: первый, кто напишет боту, станет владельцем"
        )
    ai = GeminiService(
        Config.GEMINI_API_KEY, Config.GEMINI_MODEL, Config.GEMINI_MAX_TOKENS
    )
    try:
        ai.resolve_model()
    except Exception as exc:  # не валим бота: кнопки без ИИ продолжают работать
        log.warning("Модель Gemini пока не выбрана: %s", exc)

    builder = ApplicationBuilder().token(Config.BOT_TOKEN).concurrent_updates(True)
    if AIORateLimiter is not None:
        builder = builder.rate_limiter(AIORateLimiter())
    app = builder.build()
    app.bot_data["repo"] = repo
    app.bot_data["ai"] = ai

    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))
    # Slash-команд нет: любой текст (включая /start) открывает главное меню
    app.add_handler(MessageHandler(filters.TEXT | filters.COMMAND, on_message))
    app.add_error_handler(on_error)

    log.info(
        "Bot started | label=%s | model=%s | backend=%s",
        Config.LABEL_NAME,
        ai.model or Config.GEMINI_MODEL,
        "postgres" if db.is_postgres else "sqlite",
    )
    try:
        app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)
    finally:
        db.close()


if __name__ == "__main__":
    main()
