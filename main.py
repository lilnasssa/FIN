# =============================================================================
#  22-17 LABEL FINANCE BOT
#  Управление финансами лейбла: артисты, доходы/расходы, Excel-импорт,
#  месячные отчёты, выплаты и поиск аномалий через Claude.
#  Интерфейс: ТОЛЬКО inline-кнопки, без slash-команд.
# =============================================================================
from __future__ import annotations

import asyncio
import json
import logging
import os
import re
from contextlib import contextmanager
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from typing import Any, Iterator

import psycopg2
import psycopg2.extras
from anthropic import Anthropic, APIStatusError
from openpyxl import load_workbook
from psycopg2.pool import ThreadedConnectionPool
from telegram import InlineKeyboardButton, InlineKeyboardMarkup, Update
from telegram.constants import ChatAction, ParseMode
from telegram.error import BadRequest
from telegram.ext import (
    AIORateLimiter,
    ApplicationBuilder,
    CallbackQueryHandler,
    ContextTypes,
    MessageHandler,
    filters,
)

# -----------------------------------------------------------------------------
# Конфигурация и логирование
# -----------------------------------------------------------------------------
logging.basicConfig(
    format="%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
    level=os.getenv("LOG_LEVEL", "INFO"),
)
logging.getLogger("httpx").setLevel(logging.WARNING)
log = logging.getLogger("finbot")


class Config:
    BOT_TOKEN: str = os.environ["BOT_TOKEN"]
    ANTHROPIC_API_KEY: str = os.environ["ANTHROPIC_API_KEY"]
    # Имя модели Claude Sonnet задаётся снаружи, чтобы не хардкодить версию.
    CLAUDE_MODEL: str = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-5")
    CLAUDE_MAX_TOKENS: int = int(os.getenv("CLAUDE_MAX_TOKENS", "4096"))
    DATABASE_URL: str = os.environ["DATABASE_URL"]
    DB_POOL_MIN: int = int(os.getenv("DB_POOL_MIN", "1"))
    DB_POOL_MAX: int = int(os.getenv("DB_POOL_MAX", "8"))
    ALLOWED_USER_IDS: set[int] = {
        int(x)
        for x in os.getenv("ALLOWED_USER_IDS", "").replace(" ", "").split(",")
        if x.strip().lstrip("-").isdigit()
    }
    DEFAULT_CURRENCY: str = os.getenv("DEFAULT_CURRENCY", "RUB")
    DEFAULT_RATE: Decimal = Decimal(os.getenv("DEFAULT_RATE", "0.20"))
    MAX_EXCEL_BYTES: int = int(os.getenv("MAX_EXCEL_MB", "10")) * 1024 * 1024
    ANOMALY_SIGMA: Decimal = Decimal(os.getenv("ANOMALY_SIGMA", "2"))
    LABEL_NAME: str = os.getenv("LABEL_NAME", "22-17")


# -----------------------------------------------------------------------------
# Схема БД
# -----------------------------------------------------------------------------
SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS artists (
    id              SERIAL PRIMARY KEY,
    name            TEXT NOT NULL,
    rate            NUMERIC(5,4) NOT NULL DEFAULT 0.20 CHECK (rate >= 0 AND rate <= 1),
    tg_username     TEXT,
    is_active       BOOLEAN NOT NULL DEFAULT TRUE,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    CONSTRAINT artists_name_uniq UNIQUE (name)
);

CREATE TABLE IF NOT EXISTS transactions (
    id              BIGSERIAL PRIMARY KEY,
    artist_id       INTEGER REFERENCES artists(id) ON DELETE SET NULL,
    kind            TEXT NOT NULL CHECK (kind IN ('income', 'expense')),
    amount          NUMERIC(14,2) NOT NULL CHECK (amount > 0),
    currency        TEXT NOT NULL DEFAULT 'RUB',
    category        TEXT,
    description     TEXT,
    occurred_on     DATE NOT NULL DEFAULT CURRENT_DATE,
    source          TEXT NOT NULL DEFAULT 'manual',   -- manual | excel | api
    external_key    TEXT,                             -- ключ дедупликации импорта
    created_by      BIGINT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE UNIQUE INDEX IF NOT EXISTS transactions_external_key_uniq
    ON transactions (external_key) WHERE external_key IS NOT NULL;
CREATE INDEX IF NOT EXISTS transactions_occurred_idx ON transactions (occurred_on);
CREATE INDEX IF NOT EXISTS transactions_artist_idx   ON transactions (artist_id);

CREATE TABLE IF NOT EXISTS payments (
    id              BIGSERIAL PRIMARY KEY,
    artist_id       INTEGER NOT NULL REFERENCES artists(id) ON DELETE CASCADE,
    period_start    DATE NOT NULL,
    period_end      DATE NOT NULL,
    revenue         NUMERIC(14,2) NOT NULL DEFAULT 0,
    expenses        NUMERIC(14,2) NOT NULL DEFAULT 0,
    rate            NUMERIC(5,4) NOT NULL,
    amount          NUMERIC(14,2) NOT NULL,
    status          TEXT NOT NULL DEFAULT 'pending'
                    CHECK (status IN ('pending', 'paid', 'canceled')),
    paid_at         TIMESTAMPTZ,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now(),
    CONSTRAINT payments_period_uniq UNIQUE (artist_id, period_start, period_end)
);

CREATE TABLE IF NOT EXISTS ai_reports (
    id              BIGSERIAL PRIMARY KEY,
    kind            TEXT NOT NULL,          -- monthly_report | anomaly | excel_import
    period_start    DATE,
    period_end      DATE,
    payload         JSONB,
    body            TEXT,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);

CREATE TABLE IF NOT EXISTS audit_log (
    id              BIGSERIAL PRIMARY KEY,
    tg_user_id      BIGINT,
    action          TEXT NOT NULL,
    details         JSONB,
    created_at      TIMESTAMPTZ NOT NULL DEFAULT now()
);
"""


class Database:
    """Пул синхронных соединений psycopg2 + транзакционный курсор."""

    def __init__(self, dsn: str, minconn: int, maxconn: int) -> None:
        self._pool = ThreadedConnectionPool(minconn, maxconn, dsn=dsn)

    @contextmanager
    def cursor(self, commit: bool = False) -> Iterator[psycopg2.extras.RealDictCursor]:
        conn = self._pool.getconn()
        try:
            with conn:
                with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
                    yield cur
                if not commit:
                    conn.rollback()
        except Exception:
            conn.rollback()
            raise
        finally:
            self._pool.putconn(conn)

    def init_schema(self) -> None:
        with self.cursor(commit=True) as cur:
            cur.execute(SCHEMA_SQL)
        log.info("DB schema ready")

    def close(self) -> None:
        self._pool.closeall()


# -----------------------------------------------------------------------------
# Репозиторий (синхронные функции, вызываются через asyncio.to_thread)
# -----------------------------------------------------------------------------
class Repo:
    def __init__(self, db: Database) -> None:
        self.db = db

    # ---------- артисты ----------
    def add_artist(self, name: str, rate: Decimal, username: str | None = None) -> dict:
        with self.db.cursor(commit=True) as cur:
            cur.execute(
                """
                INSERT INTO artists (name, rate, tg_username)
                VALUES (%s, %s, %s)
                ON CONFLICT (name) DO UPDATE
                    SET rate = EXCLUDED.rate,
                        tg_username = COALESCE(EXCLUDED.tg_username, artists.tg_username),
                        is_active = TRUE
                RETURNING id, name, rate, tg_username;
                """,
                (name, rate, username),
            )
            return dict(cur.fetchone())

    def list_artists(self, only_active: bool = True) -> list[dict]:
        with self.db.cursor() as cur:
            cur.execute(
                """
                SELECT id, name, rate, tg_username, is_active
                FROM artists
                WHERE (%s IS FALSE OR is_active)
                ORDER BY name;
                """,
                (only_active,),
            )
            return [dict(r) for r in cur.fetchall()]

    def get_artist(self, artist_id: int) -> dict | None:
        with self.db.cursor() as cur:
            cur.execute("SELECT * FROM artists WHERE id = %s;", (artist_id,))
            row = cur.fetchone()
            return dict(row) if row else None

    # ---------- транзакции ----------
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
    ) -> dict:
        with self.db.cursor(commit=True) as cur:
            cur.execute(
                """
                INSERT INTO transactions
                    (artist_id, kind, amount, currency, category, description,
                     occurred_on, source, external_key, created_by)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s)
                ON CONFLICT (external_key) DO NOTHING
                RETURNING id;
                """,
                (
                    artist_id, kind, amount, currency, category, description,
                    occurred_on, source, external_key, created_by,
                ),
            )
            row = cur.fetchone()
            return dict(row) if row else {}

    def bulk_insert_transactions(self, rows: list[dict], created_by: int | None) -> int:
        """Массовая вставка из Excel с дедупликацией по external_key."""
        if not rows:
            return 0
        values = [
            (
                r.get("artist_id"),
                r["kind"],
                r["amount"],
                r.get("currency") or Config.DEFAULT_CURRENCY,
                r.get("category"),
                r.get("description"),
                r["occurred_on"],
                "excel",
                r.get("external_key"),
                created_by,
            )
            for r in rows
        ]
        with self.db.cursor(commit=True) as cur:
            psycopg2.extras.execute_values(
                cur,
                """
                INSERT INTO transactions
                    (artist_id, kind, amount, currency, category, description,
                     occurred_on, source, external_key, created_by)
                VALUES %s
                ON CONFLICT (external_key) DO NOTHING;
                """,
                values,
                page_size=200,
            )
            return cur.rowcount

    def resolve_artist_ids(self, names: list[str]) -> dict[str, int]:
        if not names:
            return {}
        with self.db.cursor() as cur:
            cur.execute(
                "SELECT id, name FROM artists WHERE lower(name) = ANY(%s);",
                ([n.lower() for n in names],),
            )
            return {r["name"].lower(): r["id"] for r in cur.fetchall()}

    # ---------- аналитика ----------
    def month_bounds(self, year: int, month: int) -> tuple[date, date]:
        start = date(year, month, 1)
        end = date(year + (month == 12), (month % 12) + 1, 1)
        return start, end

    def monthly_breakdown(self, year: int, month: int) -> list[dict]:
        start, end = self.month_bounds(year, month)
        with self.db.cursor() as cur:
            cur.execute(
                """
                SELECT
                    COALESCE(a.name, '— без артиста —')                          AS artist,
                    COALESCE(a.rate, 0)                                          AS rate,
                    COALESCE(SUM(t.amount) FILTER (WHERE t.kind = 'income'), 0)  AS revenue,
                    COALESCE(SUM(t.amount) FILTER (WHERE t.kind = 'expense'), 0) AS expenses,
                    COUNT(*)                                                     AS tx_count
                FROM transactions t
                LEFT JOIN artists a ON a.id = t.artist_id
                WHERE t.occurred_on >= %s AND t.occurred_on < %s
                GROUP BY a.name, a.rate
                ORDER BY revenue DESC;
                """,
                (start, end),
            )
            rows = [dict(r) for r in cur.fetchall()]
        for r in rows:
            gross = Decimal(r["revenue"]) - Decimal(r["expenses"])
            # Прибыль = (Выручка - Расходы) * Ставка
            r["artist_share"] = (gross * Decimal(r["rate"])).quantize(Decimal("0.01"))
            r["label_profit"] = (gross - r["artist_share"]).quantize(Decimal("0.01"))
            r["gross"] = gross.quantize(Decimal("0.01"))
        return rows

    def category_breakdown(self, year: int, month: int) -> list[dict]:
        start, end = self.month_bounds(year, month)
        with self.db.cursor() as cur:
            cur.execute(
                """
                SELECT kind, COALESCE(category, 'без категории') AS category,
                       SUM(amount) AS total, COUNT(*) AS cnt
                FROM transactions
                WHERE occurred_on >= %s AND occurred_on < %s
                GROUP BY kind, category
                ORDER BY total DESC
                LIMIT 40;
                """,
                (start, end),
            )
            return [dict(r) for r in cur.fetchall()]

    def anomaly_candidates(self, months: int = 6) -> dict[str, Any]:
        """Статистика по категориям + выбросы за пределами N сигм + дубли."""
        with self.db.cursor() as cur:
            cur.execute(
                """
                WITH scoped AS (
                    SELECT t.*, COALESCE(a.name, '—') AS artist
                    FROM transactions t
                    LEFT JOIN artists a ON a.id = t.artist_id
                    WHERE t.occurred_on >= (CURRENT_DATE - (%s || ' months')::interval)
                ),
                stats AS (
                    SELECT kind, COALESCE(category, 'без категории') AS category,
                           AVG(amount) AS avg_amount,
                           COALESCE(STDDEV_POP(amount), 0) AS sd_amount,
                           COUNT(*) AS cnt
                    FROM scoped
                    GROUP BY kind, COALESCE(category, 'без категории')
                )
                SELECT s.id, s.artist, s.kind, s.amount, s.currency,
                       COALESCE(s.category, 'без категории') AS category,
                       s.description, s.occurred_on,
                       st.avg_amount, st.sd_amount
                FROM scoped s
                JOIN stats st
                  ON st.kind = s.kind
                 AND st.category = COALESCE(s.category, 'без категории')
                WHERE st.cnt >= 3
                  AND st.sd_amount > 0
                  AND ABS(s.amount - st.avg_amount) > %s * st.sd_amount
                ORDER BY ABS(s.amount - st.avg_amount) DESC
                LIMIT 40;
                """,
                (months, Config.ANOMALY_SIGMA),
            )
            outliers = [dict(r) for r in cur.fetchall()]

            cur.execute(
                """
                SELECT artist_id, kind, amount, occurred_on, COUNT(*) AS cnt
                FROM transactions
                WHERE occurred_on >= (CURRENT_DATE - (%s || ' months')::interval)
                GROUP BY artist_id, kind, amount, occurred_on
                HAVING COUNT(*) > 1
                ORDER BY cnt DESC
                LIMIT 20;
                """,
                (months,),
            )
            duplicates = [dict(r) for r in cur.fetchall()]

            cur.execute(
                """
                SELECT to_char(occurred_on, 'YYYY-MM') AS ym,
                       SUM(amount) FILTER (WHERE kind = 'income')  AS revenue,
                       SUM(amount) FILTER (WHERE kind = 'expense') AS expenses,
                       COUNT(*) AS tx_count
                FROM transactions
                WHERE occurred_on >= (CURRENT_DATE - (%s || ' months')::interval)
                GROUP BY 1 ORDER BY 1;
                """,
                (months,),
            )
            monthly = [dict(r) for r in cur.fetchall()]

            cur.execute(
                """
                SELECT COUNT(*) AS orphan_tx
                FROM transactions
                WHERE artist_id IS NULL
                  AND occurred_on >= (CURRENT_DATE - (%s || ' months')::interval);
                """,
                (months,),
            )
            integrity = dict(cur.fetchone())

        return {
            "outliers": outliers,
            "duplicates": duplicates,
            "monthly": monthly,
            "integrity": integrity,
            "sigma": str(Config.ANOMALY_SIGMA),
        }

    # ---------- выплаты ----------
    def upsert_payments_for_month(self, year: int, month: int) -> list[dict]:
        start, end = self.month_bounds(year, month)
        with self.db.cursor(commit=True) as cur:
            cur.execute(
                """
                WITH agg AS (
                    SELECT a.id AS artist_id, a.rate,
                           COALESCE(SUM(t.amount) FILTER (WHERE t.kind = 'income'), 0)  AS revenue,
                           COALESCE(SUM(t.amount) FILTER (WHERE t.kind = 'expense'), 0) AS expenses
                    FROM artists a
                    LEFT JOIN transactions t
                           ON t.artist_id = a.id
                          AND t.occurred_on >= %(start)s AND t.occurred_on < %(end)s
                    WHERE a.is_active
                    GROUP BY a.id, a.rate
                )
                INSERT INTO payments
                    (artist_id, period_start, period_end, revenue, expenses, rate, amount)
                SELECT artist_id, %(start)s, %(end)s - 1, revenue, expenses, rate,
                       ROUND(GREATEST(revenue - expenses, 0) * rate, 2)
                FROM agg
                ON CONFLICT (artist_id, period_start, period_end) DO UPDATE
                    SET revenue = EXCLUDED.revenue,
                        expenses = EXCLUDED.expenses,
                        rate = EXCLUDED.rate,
                        amount = EXCLUDED.amount
                WHERE payments.status = 'pending'
                RETURNING artist_id, amount;
                """,
                {"start": start, "end": end},
            )
            return [dict(r) for r in cur.fetchall()]

    def list_payments(self, limit: int = 20, status: str | None = None) -> list[dict]:
        with self.db.cursor() as cur:
            cur.execute(
                """
                SELECT p.id, a.name AS artist, p.period_start, p.period_end,
                       p.revenue, p.expenses, p.rate, p.amount, p.status, p.paid_at
                FROM payments p
                JOIN artists a ON a.id = p.artist_id
                WHERE (%s IS NULL OR p.status = %s)
                ORDER BY p.period_start DESC, p.amount DESC
                LIMIT %s;
                """,
                (status, status, limit),
            )
            return [dict(r) for r in cur.fetchall()]

    def mark_payment_paid(self, payment_id: int) -> None:
        with self.db.cursor(commit=True) as cur:
            cur.execute(
                "UPDATE payments SET status = 'paid', paid_at = now() WHERE id = %s;",
                (payment_id,),
            )

    # ---------- служебное ----------
    def save_report(self, kind: str, body: str, payload: dict,
                    period: tuple[date, date] | None = None) -> None:
        with self.db.cursor(commit=True) as cur:
            cur.execute(
                """
                INSERT INTO ai_reports (kind, period_start, period_end, payload, body)
                VALUES (%s, %s, %s, %s, %s);
                """,
                (
                    kind,
                    period[0] if period else None,
                    period[1] if period else None,
                    psycopg2.extras.Json(payload, dumps=json_dumps),
                    body,
                ),
            )

    def audit(self, user_id: int | None, action: str, details: dict) -> None:
        with self.db.cursor(commit=True) as cur:
            cur.execute(
                "INSERT INTO audit_log (tg_user_id, action, details) VALUES (%s, %s, %s);",
                (user_id, action, psycopg2.extras.Json(details, dumps=json_dumps)),
            )


def json_dumps(obj: Any) -> str:
    def default(o: Any) -> Any:
        if isinstance(o, Decimal):
            return float(o)
        if isinstance(o, (date, datetime)):
            return o.isoformat()
        return str(o)

    return json.dumps(obj, ensure_ascii=False, default=default)


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
  "errors": ["проблемы: пустые суммы, битые даты, отрицательные доходы, дубли"],
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
        self.client = Anthropic(api_key=api_key, max_retries=3, timeout=120.0)
        self.model = model
        self.max_tokens = max_tokens

    # ---- низкий уровень ----
    def _complete(self, prompt: str, system: str = SYSTEM_ANALYST,
                  max_tokens: int | None = None) -> str:
        try:
            msg = self.client.messages.create(
                model=self.model,
                max_tokens=max_tokens or self.max_tokens,
                system=system,
                messages=[{"role": "user", "content": prompt}],
            )
        except APIStatusError as exc:
            log.exception("Claude API error: %s", exc)
            raise RuntimeError(f"Claude недоступен (HTTP {exc.status_code}).") from exc
        return "".join(b.text for b in msg.content if getattr(b, "type", "") == "text").strip()

    async def complete(self, prompt: str, **kw: Any) -> str:
        return await asyncio.to_thread(self._complete, prompt, **kw)

    @staticmethod
    def _extract_json(text: str) -> dict:
        text = re.sub(r"^```(?:json)?|```$", "", text.strip(), flags=re.MULTILINE).strip()
        start, end = text.find("{"), text.rfind("}")
        if start == -1 or end == -1:
            raise ValueError("Claude вернул ответ без JSON")
        return json.loads(text[start : end + 1])

    # ---- прикладные методы ----
    async def parse_excel(self, raw_rows: list[list[Any]]) -> dict:
        payload = json_dumps({"rows": raw_rows})
        text = await self.complete(
            f"{EXCEL_PARSE_PROMPT}\n\nДанные:\n{payload}",
            system=SYSTEM_ANALYST,
            max_tokens=8000,
        )
        return await asyncio.to_thread(self._extract_json, text)

    async def monthly_report(self, year: int, month: int, breakdown: list[dict],
                             categories: list[dict]) -> str:
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
# Утилиты форматирования и парсинга
# -----------------------------------------------------------------------------
def money(value: Any, currency: str = "") -> str:
    d = Decimal(str(value or 0)).quantize(Decimal("0.01"))
    s = f"{d:,.2f}".replace(",", " ").replace(".", ",")
    return f"{s} {currency}".strip()


def parse_amount(text: str) -> Decimal:
    cleaned = re.sub(r"[^\d,.\-]", "", text or "").replace(",", ".")
    if cleaned.count(".") > 1:  # 1.234.567 -> 1234567
        head, _, tail = cleaned.rpartition(".")
        cleaned = head.replace(".", "") + "." + tail
    try:
        value = Decimal(cleaned).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError) as exc:
        raise ValueError("Не смог распознать сумму. Пример: 125000 или 125 000,50") from exc
    if value <= 0:
        raise ValueError("Сумма должна быть больше нуля.")
    return value


def parse_rate(text: str) -> Decimal:
    raw = parse_amount(text)
    rate = raw / Decimal(100) if raw > 1 else raw
    if not (Decimal(0) < rate <= Decimal(1)):
        raise ValueError("Ставка должна быть в диапазоне 1-100% (например 20 или 0.2).")
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


def kb_back(extra: list[list[InlineKeyboardButton]] | None = None) -> InlineKeyboardMarkup:
    rows = list(extra or [])
    rows.append([InlineKeyboardButton("⬅️ Главное меню", callback_data="nav:main")])
    return InlineKeyboardMarkup(rows)


def kb_artists(artists: list[dict], prefix: str,
               allow_none: bool = False) -> InlineKeyboardMarkup:
    rows: list[list[InlineKeyboardButton]] = []
    for i in range(0, len(artists), 2):
        rows.append(
            [
                InlineKeyboardButton(a["name"][:24], callback_data=f"{prefix}:{a['id']}")
                for a in artists[i : i + 2]
            ]
        )
    if allow_none:
        rows.append([InlineKeyboardButton("— без артиста —", callback_data=f"{prefix}:0")])
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
            InlineKeyboardButton(f"{month:02d}.{year}", callback_data=f"report:run:{year}:{month}")
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
                f"✅ Оплатить #{p['id']} · {p['artist'][:14]}",
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
    rows.append([InlineKeyboardButton("🔄 Пересчитать за месяц", callback_data="pay:recalc")])
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


async def send(update: Update, text: str, keyboard: InlineKeyboardMarkup | None = None) -> None:
    parts = chunks(text)
    chat = update.effective_chat
    for idx, part in enumerate(parts):
        markup = keyboard if idx == len(parts) - 1 else None
        try:
            await chat.send_message(part, reply_markup=markup, parse_mode=ParseMode.HTML)
        except BadRequest:  # некорректный HTML от модели — отправляем как текст
            await chat.send_message(part, reply_markup=markup)


async def show_main_menu(update: Update, prefix: str = "") -> None:
    text = (f"{prefix}\n\n" if prefix else "") + (
        f"<b>💼 Финансы лейбла {Config.LABEL_NAME}</b>\nВыберите действие:"
    )
    await send(update, text, kb_main())


def reset_state(context: ContextTypes.DEFAULT_TYPE) -> None:
    for key in ("state", "draft"):
        context.user_data.pop(key, None)


# -----------------------------------------------------------------------------
# Обработчики: любое сообщение
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
            await asyncio.to_thread(repo.audit, update.effective_user.id, "artist_add", artist)
            reset_state(context)
            await show_main_menu(
                update,
                f"✅ Артист <b>{artist['name']}</b> сохранён. "
                f"Ставка: {Decimal(artist['rate']) * 100:.2f}%",
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
            await send(update, "Дата операции (ДД.ММ.ГГГГ) или «сегодня»:", kb_back())

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
            reset_state(context)
            await show_main_menu(
                update,
                f"✅ {label} {money(draft['amount'], Config.DEFAULT_CURRENCY)} "
                f"от {occurred:%d.%m.%Y} записан.",
            )

        else:
            reset_state(context)
            await show_main_menu(update)

    except ValueError as exc:
        await send(update, f"⚠️ {exc}\nПопробуйте ещё раз:", kb_back())


# -----------------------------------------------------------------------------
# Обработчики: inline-кнопки
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
                f"• <b>{a['name']}</b> — {Decimal(a['rate']) * 100:.2f}%"
                + (f" (@{a['tg_username']})" if a["tg_username"] else "")
                for a in artists
            ]
            await send(update, "🎤 <b>Артисты лейбла</b>\n" + "\n".join(lines), kb_main())
        return

    # ---------- транзакции ----------
    if domain == "tx":
        if action == "new":
            kind = args[0]
            context.user_data["draft"] = {"kind": kind}
            artists = await asyncio.to_thread(repo.list_artists)
            context.user_data["state"] = None
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

            revenue = sum(Decimal(r["revenue"]) for r in breakdown)
            expenses = sum(Decimal(r["expenses"]) for r in breakdown)
            shares = sum(Decimal(r["artist_share"]) for r in breakdown)
            profit = sum(Decimal(r["label_profit"]) for r in breakdown)

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
            "Ожидаемые колонки (в любом порядке и на любом языке): "
            "дата, артист, тип операции, сумма, категория, описание. "
            "Claude сам разберёт структуру и подсветит ошибки.",
            kb_back(),
        )
        return

    # ---------- аномалии ----------
    if domain == "anomaly":
        await update.effective_chat.send_action(ChatAction.TYPING)
        data = await asyncio.to_thread(repo.anomaly_candidates, 6)
        if not any((data["outliers"], data["duplicates"], data["monthly"])):
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
            status = None if args[0] == "all" else args[0]
            payments = await asyncio.to_thread(repo.list_payments, 20, status)
            if not payments:
                await send(update, "Платежей не найдено.", kb_payments([]))
                return
            icons = {"pending": "⏳", "paid": "✅", "canceled": "❌"}
            lines = ["💸 <b>Платежи артистам</b>"]
            for p in payments:
                lines.append(
                    f"{icons.get(p['status'], '•')} #{p['id']} <b>{p['artist']}</b> · "
                    f"{p['period_start']:%m.%Y} · {money(p['amount'], Config.DEFAULT_CURRENCY)} "
                    f"(ставка {Decimal(p['rate']) * 100:.0f}%)"
                )
            total_pending = sum(
                Decimal(p["amount"]) for p in payments if p["status"] == "pending"
            )
            lines.append(f"\nК выплате: <b>{money(total_pending, Config.DEFAULT_CURRENCY)}</b>")
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
# Обработчик Excel-документов
# -----------------------------------------------------------------------------
def read_workbook(blob: bytes, max_rows: int = 500) -> list[list[Any]]:
    wb = load_workbook(BytesIO(blob), data_only=True, read_only=True)
    rows: list[list[Any]] = []
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            if row is None or all(cell in (None, "") for cell in row):
                continue
            rows.append(
                [
                    cell.strftime("%Y-%m-%d") if isinstance(cell, (datetime, date))
                    else (float(cell) if isinstance(cell, Decimal) else cell)
                    for cell in row
                ]
            )
            if len(rows) >= max_rows:
                wb.close()
                return rows
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

    names = sorted({str(r.get("artist")).strip() for r in candidates if r.get("artist")})
    mapping = await asyncio.to_thread(repo.resolve_artist_ids, names)

    prepared: list[dict] = []
    for idx, r in enumerate(candidates):
        try:
            amount = Decimal(str(r["amount"])).quantize(Decimal("0.01"))
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
            errors.append(f"артист «{artist_name}» не найден в базе — операция без привязки")

        prepared.append(
            {
                "artist_id": artist_id,
                "kind": kind,
                "amount": amount,
                "currency": (r.get("currency") or Config.DEFAULT_CURRENCY)[:3].upper(),
                "category": r.get("category") or None,
                "description": r.get("description") or None,
                "occurred_on": occurred,
                "external_key": f"xlsx:{doc.file_unique_id}:{idx}:{kind}:{amount}:{occurred}",
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

    income = sum(r["amount"] for r in prepared if r["kind"] == "income")
    expense = sum(r["amount"] for r in prepared if r["kind"] == "expense")

    lines = [
        f"📥 <b>Импорт «{doc.file_name}» завершён</b>",
        f"Распознано строк: <b>{len(prepared)}</b>, записано новых: <b>{inserted}</b>",
        f"Доходы: {money(income, Config.DEFAULT_CURRENCY)} · "
        f"Расходы: {money(expense, Config.DEFAULT_CURRENCY)}",
    ]
    if parsed.get("summary"):
        lines += ["", f"<i>{parsed['summary']}</i>"]
    if errors:
        lines += ["", "<b>⚠️ Проблемы в данных:</b>"] + [f"• {e}" for e in errors[:15]]
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
# Точка входа
# -----------------------------------------------------------------------------
def main() -> None:
    db = Database(Config.DATABASE_URL, Config.DB_POOL_MIN, Config.DB_POOL_MAX)
    db.init_schema()
    repo = Repo(db)
    claude = ClaudeService(
        Config.ANTHROPIC_API_KEY, Config.CLAUDE_MODEL, Config.CLAUDE_MAX_TOKENS
    )

    app = (
        ApplicationBuilder()
        .token(Config.BOT_TOKEN)
        .rate_limiter(AIORateLimiter())
        .concurrent_updates(True)
        .build()
    )
    app.bot_data["repo"] = repo
    app.bot_data["claude"] = claude

    app.add_handler(CallbackQueryHandler(on_callback))
    app.add_handler(MessageHandler(filters.Document.ALL, on_document))
    # Slash-команд нет: любой текст (включая /start) открывает главное меню
    app.add_handler(MessageHandler(filters.TEXT | filters.COMMAND, on_message))
    app.add_error_handler(on_error)

    log.info("Bot started, model=%s", Config.CLAUDE_MODEL)
    try:
        app.run_polling(allowed_updates=Update.ALL_TYPES, drop_pending_updates=True)
    finally:
        db.close()


if __name__ == "__main__":
    main()
