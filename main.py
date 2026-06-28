import os
import csv
import time
import json
import math
import random
import logging
import argparse
from dataclasses import dataclass
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from typing import Dict, List, Tuple, Any, Optional

import requests
from requests.exceptions import RequestException
from dotenv import load_dotenv

import gspread
from google.oauth2.service_account import Credentials

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ReplyKeyboardMarkup, KeyboardButton
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, CallbackQueryHandler, MessageHandler, ContextTypes, filters

# ------------------- CONFIG -------------------
from pathlib import Path
load_dotenv(dotenv_path=Path(__file__).with_name(".env"))

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
log = logging.getLogger("fleet-reports")

DUBAI_TZ = ZoneInfo("Asia/Almaty")
FLEET_BASE = "https://fleet-api.taxi.yandex.net"

BOT_TOKEN = os.getenv("BOT_TOKEN")
CHAT_ID = os.getenv("CHAT_ID")

YANDEX_CLIENT_ID = os.getenv("YANDEX_CLIENT_ID")
YANDEX_API_KEY = os.getenv("YANDEX_API_KEY")
PARK_ID = os.getenv("PARK_ID")

GOOGLE_SHEET_ID = os.getenv("GOOGLE_SHEET_ID")
GOOGLE_SA_JSON = os.getenv("GOOGLE_SA_JSON")

PARK_COMMISSION_PERCENT = float(os.getenv("PARK_COMMISSION_PERCENT", "5"))
ORDERS_PAGE_LIMIT = 500
MAX_PAGES = 400
MAX_ORDERS_PER_DAY = 8000

missing = [k for k, v in {
    "BOT_TOKEN": BOT_TOKEN,
    "YANDEX_CLIENT_ID": YANDEX_CLIENT_ID,
    "YANDEX_API_KEY": YANDEX_API_KEY,
    "PARK_ID": PARK_ID,
    "GOOGLE_SHEET_ID": GOOGLE_SHEET_ID,
    "GOOGLE_SA_JSON": GOOGLE_SA_JSON,
}.items() if not v]
if missing:
    raise SystemExit(f"Missing .env variables: {', '.join(missing)}")

CHAT_ID_INT: Optional[int] = None
if CHAT_ID and isinstance(CHAT_ID, str) and CHAT_ID.lstrip("-").isdigit():
    CHAT_ID_INT = int(CHAT_ID)

# ------------------- GOOGLE SHEETS -------------------
GS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

def gs_client():
    sa_value = GOOGLE_SA_JSON.strip()
    if sa_value.startswith("{"):
        info = json.loads(sa_value)
        creds = Credentials.from_service_account_info(info, scopes=GS_SCOPES)
    else:
        creds = Credentials.from_service_account_file(sa_value, scopes=GS_SCOPES)
    return gspread.authorize(creds)

def gs_get_or_create_ws(sh, title: str, rows: int = 2000, cols: int = 20):
    try:
        ws = sh.worksheet(title)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=rows, cols=cols)
    return ws

def ws_clear_and_write(ws, values: list[list]):
    ws.clear()
    if values:
        ws.update(values=values, range_name="A1", value_input_option="USER_ENTERED")

# ------------------- YANDEX -------------------
def fleet_headers():
    return {
        "X-Client-ID": YANDEX_CLIENT_ID,
        "X-API-Key": YANDEX_API_KEY,
        "Accept": "application/json",
        "Content-Type": "application/json",
        "Accept-Language": "ru",
    }

def post_with_retry(session: requests.Session, url: str, payload: dict, max_retries: int = 8, timeout: int = 60):
    base_delay = 1.0
    last = None
    for attempt in range(1, max_retries + 1):
        try:
            r = session.post(url, headers=fleet_headers(), json=payload, timeout=timeout)
            last = r
        except RequestException as e:
            if attempt == max_retries:
                raise
            sleep_s = min(base_delay * (2 ** (attempt - 1)), 30.0) + random.uniform(0.2, 0.8)
            log.warning("YANDEX RETRY (network): %s | sleeping %.1fs", e, sleep_s)
            time.sleep(sleep_s)
            continue
        if r.status_code < 400:
            return r
        if 400 <= r.status_code < 500 and r.status_code != 429:
            return r
        retry_after = r.headers.get("Retry-After")
        if retry_after:
            try:
                base_delay = max(base_delay, float(retry_after))
            except ValueError:
                pass
        if attempt < max_retries:
            sleep_s = min(base_delay * (2 ** (attempt - 1)), 30.0) + random.uniform(0.2, 0.9)
            log.warning("YANDEX RETRY: HTTP %s (attempt %s/%s) sleeping %.1fs", r.status_code, attempt, max_retries, sleep_s)
            time.sleep(sleep_s)
            continue
        return r
    return last

def iso_dubai(dt: datetime) -> str:
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=DUBAI_TZ)
    return dt.astimezone(DUBAI_TZ).isoformat()

def dubai_day_range(day_date) -> Tuple[str, str, datetime, datetime]:
    start_dt = datetime(day_date.year, day_date.month, day_date.day, 0, 0, 0, tzinfo=DUBAI_TZ)
    end_dt   = datetime(day_date.year, day_date.month, day_date.day, 23, 59, 59, tzinfo=DUBAI_TZ)
    api_start = start_dt - timedelta(hours=4)
    api_end   = end_dt + timedelta(hours=4)
    return iso_dubai(api_start), iso_dubai(api_end), start_dt, end_dt

def extract_booked_at(order: dict) -> Optional[str]:
    if isinstance(order.get("order"), dict) and order["order"].get("booked_at"):
        return order["order"]["booked_at"]
    if order.get("booked_at"):
        return order["booked_at"]
    return None

def parse_dt(s: str) -> Optional[datetime]:
    if not s:
        return None
    try:
        if s.endswith("Z"):
            s = s.replace("Z", "+00:00")
        return datetime.fromisoformat(s)
    except Exception:
        return None

def normalize_status(o: dict) -> str:
    st = (o.get("status") or "").lower().strip()
    if not st and isinstance(o.get("order"), dict):
        st = str(o["order"].get("status") or "").lower().strip()
    return st

def extract_effective_dt(order: dict) -> Optional[datetime]:
    """Используем ended_at для точного совпадения с Fleet."""
    ended_at = order.get("ended_at")
    if isinstance(order.get("order"), dict):
        ended_at = ended_at or order["order"].get("ended_at")
    if isinstance(ended_at, str) and ended_at:
        dt = parse_dt(ended_at)
        if dt is not None:
            return dt
    booked_at = extract_booked_at(order)
    if booked_at:
        return parse_dt(booked_at)
    return None

def orders_list(session, time_from_iso, time_to_iso, limit, offset=0, cursor=None):
    url = f"{FLEET_BASE}/v1/parks/orders/list"
    limit = min(int(limit), ORDERS_PAGE_LIMIT)
    payload: Dict[str, Any] = {
        "limit": limit,
        "query": {
            "park": {
                "id": str(PARK_ID),
                "order": {
                    "booked_at": {"from": time_from_iso, "to": time_to_iso},
                },
            }
        },
        "fields": {
            "order": ["id", "status", "booked_at", "accepted_at", "started_at", "ended_at"],
            "driver_profile": ["id", "first_name", "last_name", "callsign", "name", "balance"],
            "car": ["callsign", "brand_model", "license"],
            "payment": ["amount", "type"],
            "tip": ["amount"],
            "bonus": ["amount"],
            "promotion": ["amount"],
        },
    }
    if cursor:
        payload["cursor"] = cursor
    else:
        payload["offset"] = int(offset)
    r = post_with_retry(session, url, payload)
    if r.status_code >= 400:
        log.error("YANDEX ORDERS ERROR: %s %s", r.status_code, r.text)
    r.raise_for_status()
    return r.json()

def orders_list_all(session, time_from_iso, time_to_iso, from_dt, to_dt):
    kept = []
    seen_ids = set()
    pages = 0
    cursor = None
    while True:
        pages += 1
        if pages > MAX_PAGES:
            raise RuntimeError(f"Слишком много страниц (> {MAX_PAGES}).")
        data = orders_list(session, time_from_iso, time_to_iso, limit=ORDERS_PAGE_LIMIT,
                           offset=(pages - 1) * ORDERS_PAGE_LIMIT, cursor=cursor)
        items = data.get("orders") or data.get("result", {}).get("orders") or []
        if not isinstance(items, list):
            items = []
        next_cursor = data.get("cursor")
        new_unique = 0
        kept_now = 0
        for o in items:
            oid = o.get("id")
            if oid is None and isinstance(o.get("order"), dict):
                oid = o["order"].get("id")
            if oid is not None and oid in seen_ids:
                continue
            if oid is not None:
                seen_ids.add(oid)
            new_unique += 1
            st = normalize_status(o)
            if st not in ("complete", "completed"):
                continue
            dt = extract_effective_dt(o)
            if dt is None:
                continue
            dt_local = dt.astimezone(DUBAI_TZ)
            if from_dt <= dt_local <= to_dt:
                kept.append(o)
                kept_now += 1
        log.info("Fetched %s orders (page=%s) new_unique=%s kept_in_window=%s total_kept=%s",
                 len(items), pages, new_unique, kept_now, len(kept))
        if len(kept) >= MAX_ORDERS_PER_DAY:
            raise RuntimeError(f"Слишком много заказов за сутки (>= {MAX_ORDERS_PER_DAY}).")
        if len(items) == 0:
            break
        if new_unique == 0:
            log.warning("Pagination seems stuck. Stopping.")
            break
        if not next_cursor:
            break
        cursor = next_cursor
        time.sleep(0.35)
    return kept

def fetch_driver_balances(session: requests.Session) -> Dict[str, float]:
    url = f"{FLEET_BASE}/v1/parks/driver-profiles/list"
    balances: Dict[str, float] = {}
    limit = 200
    offset = 0
    while True:
        payload = {
            "fields": {"account": ["balance"], "driver_profile": ["first_name", "last_name", "middle_name", "id"]},
            "limit": limit, "offset": offset,
            "query": {"park": {"id": str(PARK_ID)}},
        }
        r = post_with_retry(session, url, payload)
        if r is None or r.status_code >= 400:
            break
        data = r.json()
        items = data.get("driver_profiles") or []
        if not items:
            break
        for item in items:
            dp = item.get("driver_profile") or {}
            account = item.get("accounts") or item.get("account") or []
            balance_val = None
            if isinstance(account, list):
                for acc in account:
                    if "balance" in acc:
                        try:
                            balance_val = float(acc["balance"])
                            break
                        except Exception:
                            pass
            elif isinstance(account, dict):
                if "balance" in account:
                    try:
                        balance_val = float(account["balance"])
                    except Exception:
                        pass
            if balance_val is None:
                continue
            first = (dp.get("first_name") or "").strip()
            last = (dp.get("last_name") or "").strip()
            middle = (dp.get("middle_name") or "").strip()
            fio = " ".join([x for x in [last, first, middle] if x]).strip()
            if fio:
                balances[fio] = balance_val
        if len(items) < limit:
            break
        offset += limit
        time.sleep(0.2)
    return balances

def fetch_partner_commission(session: requests.Session, order_ids: List[str], day_from: str = None, day_to: str = None) -> float:
    """Получает реальную комиссию партнёра из Yandex Fleet API."""
    if not order_ids:
        return 0.0
    total = 0.0
    chunk_size = 100
    for i in range(0, len(order_ids), chunk_size):
        chunk = order_ids[i:i + chunk_size]
        payload = {
            "query": {"park": {"id": PARK_ID, "order": {"ids": chunk}}},
            "limit": 500
        }
        try:
            r = post_with_retry(session, "https://fleet-api.taxi.yandex.net/v2/parks/orders/transactions/list", payload)
            if r and r.status_code == 200:
                txs = r.json().get("transactions", [])
                for tx in txs:
                    if tx.get("category_id") not in ("partner_ride_fee", "partner_bonus_fee"):
                        continue
                    if day_from and day_to:
                        event_at = tx.get("event_at", "")
                        if event_at:
                            try:
                                event_dt = parse_dt(event_at)
                                if event_dt:
                                    event_local = event_dt.astimezone(DUBAI_TZ)
                                    from_dt = parse_dt(day_from).astimezone(DUBAI_TZ)
                                    to_dt = parse_dt(day_to).astimezone(DUBAI_TZ)
                                    if not (from_dt <= event_local <= to_dt):
                                        continue
                            except Exception:
                                pass
                    try:
                        total += abs(float(tx.get("amount", 0)))
                    except Exception:
                        pass
            time.sleep(0.2)
        except Exception as e:
            log.warning("fetch_partner_commission error: %s", e)
    return total

# ------------------- TARIFFS -------------------
BULK_TARIFF_IMPORT = [
    ("Алиев Ырысбек Тажимаматович", "Штатный", 5.0, None),
    ("Амангельдиев Бауржан Хасенулы", "Штатный", 5.0, None),
    ("Аскаров Ернар Аскарұлы", "Штатный", 5.0, None),
    ("Асылхан Мұхтархан Маралханұлы", "Штатный", 5.0, None),
    ("Базаров Медель Далелканұлы", "Штатный", 5.0, None),
    ("Баубеков Ермантой Рисдавлат Угли", "Штатный", 5.0, None),
    ("Бауыржанұлы Сержан Serzhan", "Штатный", 5.0, None),
    ("Бахтиярұлы Рустам", "Штатный", 5.0, None),
    ("Бекетаев Жомарт Арипбекович", "Штатный", 5.0, None),
    ("Биназаров Марлен Жанболатович", "Штатный", 5.0, None),
    ("Буглеев Денис Николаевич", "Штатный", 5.0, None),
    ("Әбділдә Арнат Талғатұлы", "Штатный", 5.0, None),
    ("Жұмаханұлы Айдар Айдар", "Штатный", 5.0, None),
    ("Исмаилов Серқожа Азатұлы", "Штатный", 5.0, None),
    ("Исмаилов Оруч Муслимович", "Штатный", 5.0, None),
    ("Казизов Алмаз Даулетбекович", "Штатный", 5.0, None),
    ("Ким Алексей Леонидович", "Штатный", 5.0, None),
    ("Куанишов Сағихан Талғатұлы", "Штатный", 5.0, None),
    ("Құттыбай Арнат Алмасұлы", "Штатный", 5.0, None),
    ("Намазов Елмир Елдарович", "Штатный", 5.0, None),
    ("Николенко Валерий Васильевич", "Штатный", 5.0, None),
    ("Сенцов Михаил Андреевич", "Штатный", 5.0, None),
    ("Табылдиев Ардак Адибаевич", "Штатный", 5.0, None),
    ("Таипов Алишер Ашимжанович", "Штатный", 5.0, None),
    ("Ташпулатов Жасур Садуллаұлы", "Штатный", 5.0, None),
    ("Тургунбаев Ержан Эралиевич", "Штатный", 5.0, None),
    ("Узаков Бахтияр Халмаджанович", "Штатный", 5.0, None),
    ("Шайзинда Нұрдәулет Асанұлы", "Штатный", 5.0, None),
    ("Тютебаев Бауыржан Кайратович", "Штатный2.5", 2.5, None),
    ("Горобец Эльмира Дюсенгалиевна", "Частник", 5.0, None),
    ("Ли Дмитрий Вилларионович", "Частник", 5.0, None),
    ("Аубакиров Ерлан Кажмуратович", "АРА", None, "2026-04-29"),
    ("Байтенов Шералы Султанович", "АРА", None, "2026-05-31"),
    ("Банлоу Муса Закирович", "АРА", None, "2026-04-21"),
    ("Джумангалиев Murat Бакытжанович", "АРА", None, "2026-05-01"),
    ("Дюсембеков Канат Аскарович", "АРА", None, "2026-04-10"),
    ("Ергалиев Нурлан Жолдасович", "АРА", None, "2026-01-06"),
    ("Жорабаев Бейбит Мекембанвич", "АРА", None, "2026-06-03"),
    ("Карабаев Беиик Суярбекович", "АРА", None, "2026-04-27"),
    ("Кожан Магжан Кантореулы", "АРА", None, "2026-04-13"),
    ("Куандыков Ерлан Сапарбекович", "АРА", None, "2026-01-23"),
    ("Кызайбаев Бахытжан Асембаевич", "АРА", None, "2026-03-27"),
    ("Омирсериков Мухит Омирсерикулы", "АРА", None, "2026-04-18"),
    ("Пахриддинұлы Хайриддин", "АРА", None, "2026-04-10"),
    ("Рагчааулы Ерболат", "АРА", None, None),
    ("Саду Аян", "АРА", None, None),
    ("Сламия Әлішер Асанұлы", "АРА", None, "2026-05-10"),
    ("Танаткан Бактияр Маратулы", "АРА", None, "2026-03-05"),
    ("Хамраев Исак", "АРА", None, "2026-04-15"),
    ("Шахметова Айжан Мараткызы", "АРА", None, None),
]

TARIFF_SHEET_NAME = "Тарифы"
TARIFF_HEADER = ["ФИО водителя", "Тариф", "Процент", "Дата начала АРА"]

def load_tariffs() -> Dict[str, dict]:
    import datetime as dt_mod
    result = {}
    try:
        gc = gs_client()
        sh = gc.open_by_key(GOOGLE_SHEET_ID)
        try:
            ws = sh.worksheet(TARIFF_SHEET_NAME)
        except gspread.WorksheetNotFound:
            return result
        rows = ws.get_all_records()
        for r in rows:
            fio = str(r.get("ФИО водителя") or "").strip()
            if not fio:
                continue
            tariff = str(r.get("Тариф") or "").strip()
            percent_raw = r.get("Процент")
            try:
                percent = float(percent_raw) if percent_raw not in (None, "", "авто") else None
            except Exception:
                percent = None
            ara_start_raw = str(r.get("Дата начала АРА") or "").strip()
            ara_start = None
            if ara_start_raw:
                try:
                    ara_start = dt_mod.date.fromisoformat(ara_start_raw)
                except Exception:
                    pass
            result[fio] = {"tariff": tariff, "percent": percent, "ara_start": ara_start}
    except Exception as e:
        log.warning("load_tariffs error: %s", e)
    return result

def get_driver_percent(fio: str, tariffs: Dict[str, dict], for_date=None) -> float:
    import datetime as dt_mod
    info = tariffs.get(fio)
    if not info:
        return PARK_COMMISSION_PERCENT
    tariff = (info.get("tariff") or "").strip().upper()
    if tariff == "АРА":
        ara_start = info.get("ara_start")
        if ara_start:
            check_date = for_date or dt_mod.date.today()
            days_passed = (check_date - ara_start).days
            return 1.0 if days_passed < 14 else 2.0
        return 1.0
    if info.get("percent") is not None:
        return info["percent"]
    return PARK_COMMISSION_PERCENT

def set_tariff(fio: str, tariff: str, percent: Optional[float] = None, ara_start=None) -> None:
    gc = gs_client()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    ws = gs_get_or_create_ws(sh, TARIFF_SHEET_NAME, rows=500, cols=10)
    values = ws.get_all_values()
    has_header = bool(values) and len(values[0]) >= 1 and values[0][0].strip() == TARIFF_HEADER[0]
    if not has_header:
        ws.insert_row(TARIFF_HEADER, index=1, value_input_option="USER_ENTERED")
        values = ws.get_all_values()
    fio_col = 0
    row_idx = None
    for i, row in enumerate(values[1:], start=2):
        if len(row) > fio_col and row[fio_col].strip() == fio.strip():
            row_idx = i
            break
    percent_str = "" if percent is None else str(percent)
    ara_str = "" if not ara_start else str(ara_start)
    new_row = [fio, tariff, percent_str, ara_str]
    if row_idx:
        ws.update(values=[new_row], range_name=f"A{row_idx}:D{row_idx}", value_input_option="USER_ENTERED")
    else:
        ws.append_row(new_row, value_input_option="USER_ENTERED")

def calc_park_income_individual(agg: Dict[str, "DriverAgg"], tariffs: Dict[str, dict], for_date=None) -> float:
    total = 0.0
    for fio, a in agg.items():
        pct = get_driver_percent(fio, tariffs, for_date)
        total += a.base * (pct / 100)
    return total

def get_month_totals_from_sheets(day_date, tariffs: dict = None) -> tuple:
    import datetime as dt_mod
    try:
        gc = gs_client()
        sh = gc.open_by_key(GOOGLE_SHEET_ID)
        month_prefix = day_date.strftime("%Y-%m")
        total_orders = 0
        total_net = 0.0
        total_park = 0.0
        for ws in sh.worksheets():
            title = ws.title.strip()
            if not title.startswith(month_prefix):
                continue
            try:
                sheet_date = dt_mod.date.fromisoformat(title)
                if sheet_date > day_date:
                    continue
            except Exception:
                continue
            values = ws.get_all_values()
            for row in values[1:]:
                if len(row) >= 5:
                    try:
                        fio = str(row[0]).strip()
                        orders_val = int(row[2]) if str(row[2]).strip() else 0
                        net_str = str(row[4]).replace(",", "").replace(" ", "").strip()
                        net_val = float(net_str) if net_str else 0.0
                        total_orders += orders_val
                        total_net += net_val
                        if tariffs and fio:
                            pct = get_driver_percent(fio, tariffs, sheet_date)
                        else:
                            pct = PARK_COMMISSION_PERCENT
                        total_park += net_val * (pct / 100)
                    except Exception:
                        pass
        return total_orders, total_net, total_park
    except Exception as e:
        log.warning("get_month_totals_from_sheets error: %s", e)
        return 0, 0.0, 0.0

# ------------------- REPORT -------------------
def load_driver_directory() -> Dict[str, dict]:
    gc = gs_client()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    try:
        ws = sh.worksheet("Trips Per Driver")
    except Exception:
        return {}
    rows = ws.get_all_records()
    directory = {}
    for r in rows:
        callsign = str(r.get("Позывной") or r.get("callsign") or r.get("Номер") or "").strip()
        fio = str(r.get("ФИО") or r.get("ФИО водителя") or "").strip()
        vu = str(r.get("Номер ВУ") or r.get("ВУ") or "").strip()
        balance = r.get("Баланс")
        if callsign:
            directory[callsign.lower()] = {"fio": fio, "vu": vu, "balance": balance}
    return directory

@dataclass
class DriverAgg:
    fio: str
    callsign: str
    plate: str
    done: int
    net: float       # общая выручка для отображения
    base: float      # база для расчёта % парка
    cash: float
    cashless: float
    line_seconds: float
    balance: float

def _get_payment_obj(o: dict) -> dict:
    if isinstance(o.get("payment"), dict):
        return o["payment"]
    if isinstance(o.get("order"), dict) and isinstance(o["order"].get("payment"), dict):
        return o["order"]["payment"]
    return {}

def safe_amount(o: dict) -> tuple:
    """Возвращает (base_amount, total_amount)."""
    p = _get_payment_obj(o)
    def extract_val(obj):
        if obj is None:
            return 0.0
        if isinstance(obj, dict):
            for k in ("amount", "total", "value", "price"):
                if obj.get(k) is not None:
                    try:
                        return float(obj[k])
                    except Exception:
                        pass
        else:
            try:
                return float(obj)
            except Exception:
                pass
        return 0.0
    payment_amount = extract_val(p.get("amount") or p.get("total") or p.get("value"))
    if payment_amount == 0.0:
        for key in ("cost", "price", "amount"):
            v = (o.get("order") or {}).get(key) or o.get(key)
            payment_amount = extract_val(v)
            if payment_amount > 0:
                break
    tip = extract_val((o.get("tip") or {}).get("amount"))
    bonus = extract_val((o.get("bonus") or {}).get("amount"))
    promo = extract_val((o.get("promotion") or {}).get("amount"))
    return payment_amount, payment_amount + tip + bonus + promo

def pay_type(o: dict) -> str:
    top_level = str(o.get("payment_method") or "").lower().strip()
    if top_level:
        return top_level
    p = _get_payment_obj(o)
    t = p.get("type")
    if isinstance(t, dict):
        t = t.get("id") or t.get("name")
    return str(t).lower().strip() if t else ""

def fio_from_order(o: dict) -> Tuple[str, str]:
    if not isinstance(o, dict):
        return "", ""
    dp = o.get("driver_profile") or {}
    car = o.get("car") or {}
    fio = str(dp.get("name") or "").strip()
    if not fio:
        fio = " ".join([x for x in [dp.get("first_name"), dp.get("last_name")] if x]).strip()
    callsign = str(dp.get("callsign") or car.get("callsign") or "").strip()
    return fio or "", callsign or ""

def plate_from_order(o: dict) -> str:
    car = o.get("car") or {}
    license_obj = car.get("license") or {}
    if isinstance(license_obj, dict):
        return str(license_obj.get("number") or license_obj.get("plate") or "").strip()
    if isinstance(license_obj, str):
        return license_obj.strip()
    return ""

def build_report_rows(day_str, orders, directory, balances=None):
    agg: Dict[str, DriverAgg] = {}
    balances = balances or {}
    for o in orders:
        result = fio_from_order(o)
        fio_api, callsign = result if isinstance(result, tuple) else ("", "")
        plate = plate_from_order(o)
        key_callsign = (callsign or "").lower()
        fio_final = fio_api
        if (not fio_final) and key_callsign in directory:
            fio_final = directory[key_callsign].get("fio", "") or ""
        if not fio_final:
            fio_final = callsign or "неизвестно"
        key = fio_final
        if key not in agg:
            agg[key] = DriverAgg(fio=fio_final, callsign=callsign, plate=plate,
                                  done=0, net=0.0, base=0.0, cash=0.0, cashless=0.0,
                                  line_seconds=0.0, balance=0.0)
        amount_base, amount_total = safe_amount(o)
        driver_balance = None
        dp = o.get("driver_profile") or {}
        bal_val = dp.get("balance")
        if bal_val is not None:
            try:
                driver_balance = float(bal_val)
            except Exception:
                pass
        start_time = o.get("accepted_at") or o.get("started_at")
        end_time = o.get("ended_at")
        if isinstance(o.get("order"), dict):
            start_time = start_time or o["order"].get("accepted_at") or o["order"].get("started_at")
            end_time = end_time or o["order"].get("ended_at")
        duration_sec = 0
        if start_time and end_time:
            dt1 = parse_dt(start_time)
            dt2 = parse_dt(end_time)
            if dt1 and dt2:
                duration_sec = max(0, (dt2 - dt1).total_seconds())
        t = pay_type(o)
        agg[key].done += 1
        agg[key].net += amount_total
        agg[key].base += amount_base
        agg[key].line_seconds += duration_sec
        if driver_balance is not None:
            agg[key].balance = driver_balance
        if "cash" in t or "нал" in t:
            agg[key].cash += amount_total
        else:
            agg[key].cashless += amount_total
    rows_sorted = sorted(agg.values(), key=lambda x: x.net, reverse=True)
    out = []
    for r in rows_sorted:
        hours = round(r.line_seconds / 3600, 2)
        bal = balances.get(r.fio)
        if bal is None:
            bal = r.balance
        if not bal:
            drow = directory.get((r.callsign or "").lower(), {})
            bal = drow.get("balance", "")
        out.append([r.fio, r.plate, r.done, hours, round(r.net, 2), 17000, round(r.net - 17000, 2), bal])
    return out, agg

def write_to_google_sheets(day_str, report_rows):
    gc = gs_client()
    sh = gc.open_by_key(GOOGLE_SHEET_ID)
    ws = gs_get_or_create_ws(sh, day_str, rows=max(2000, len(report_rows) + 10), cols=20)
    header = ["ФИО водителя", "Госномер", "Завершено заказов", "Время на линии",
               "Чистый заработок", "Цель", "Отклонение плана", "Баланс"]
    ws_clear_and_write(ws, [header] + report_rows)

def write_csv(day_str, report_rows):
    filename = f"report_{day_str}.csv"
    with open(filename, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(["ФИО водителя", "Номер ВУ", "Завершено заказов", "Время на линии",
                     "Чистый заработок", "Цель", "Отклонение плана", "Баланс"])
        w.writerows(report_rows)
    return filename

def format_quick_summary(day_str, agg, month_orders=0, month_net=0.0, tariffs=None, month_park_income=0.0, real_park_income=0.0):
    import datetime as dt_mod
    total_orders = sum(a.done for a in agg.values())
    total_net = sum(a.net for a in agg.values())
    if tariffs:
        try:
            report_date = dt_mod.date.fromisoformat(day_str)
        except Exception:
            report_date = dt_mod.date.today()
        park_income_calc = calc_park_income_individual(agg, tariffs, report_date)
    else:
        park_income_calc = total_net * (PARK_COMMISSION_PERCENT / 100)
    park_income = real_park_income if real_park_income > 0 else park_income_calc
    avg_check = round(total_net / total_orders, 2) if total_orders else 0.0
    drivers = list(agg.values())
    top_best = sorted(drivers, key=lambda x: x.net, reverse=True)[:5]
    top_worst = sorted(drivers, key=lambda x: x.net)[:5]
    text = (
        f"📅 <b>Отчёт за {day_str}</b>\n"
        f"✅ Заказов: <b>{total_orders}</b>\n"
        f"💰 Выручка: <b>{total_net:,.0f} ₸</b>\n"
        f"🏦 Доход таксопарка: <b>{park_income:,.0f} ₸</b>\n"
        f"📊 Средний чек: <b>{avg_check:,.0f} ₸</b>\n"
        f"👤 Водителей: <b>{len(agg)}</b>\n"
    )
    groups = {
        "Штатные": {"orders": 0, "net": 0.0, "count": 0},
        "АРА": {"orders": 0, "net": 0.0, "count": 0},
        "Тариф1": {"orders": 0, "net": 0.0, "count": 0},
        "Другие": {"orders": 0, "net": 0.0, "count": 0},
    }
    tariffs_map = tariffs or {}
    for fio, a in agg.items():
        info = tariffs_map.get(fio)
        if info is None:
            group = "Штатные"
        else:
            t = (info.get("tariff") or "").strip().lower()
            if t in ("штатный", "штатный2.5"):
                group = "Штатные"
            elif t == "ара":
                group = "АРА"
            elif t == "тариф1":
                group = "Тариф1"
            else:
                group = "Другие"
        groups[group]["orders"] += a.done
        groups[group]["net"] += a.net
        groups[group]["count"] += 1
    group_emoji = {"Штатные": "👔", "АРА": "🆕", "Тариф1": "2️⃣", "Другие": "📋"}
    text += "\n📂 <b>По тарифам:</b>\n"
    for gname, gdata in groups.items():
        if gdata["count"] == 0:
            continue
        text += f"{group_emoji.get(gname, '•')} {gname} ({gdata['count']}): {gdata['net']:,.0f} ₸\n"
    text += "\n"
    if month_orders > 0:
        month_park = month_park_income if month_park_income > 0 else month_net * (PARK_COMMISSION_PERCENT / 100)
        text += (
            f"\n📆 <b>Итого с начала месяца:</b>\n"
            f"   Поездок: <b>{month_orders:,}</b>\n"
            f"   Выручка: <b>{month_net:,.0f} ₸</b>\n"
            f"   Доход парка: <b>{month_park:,.0f} ₸</b>\n"
        )
    text += "\n🏆 <b>ТОП‑5 водителей</b>\n"
    for d in top_best:
        text += f"• {d.fio} — {d.net:,.0f}\n"
    text += "\n📉 <b>Худшие 5</b>\n"
    for d in top_worst:
        text += f"• {d.fio} — {d.net:,.0f}\n"
    return text

# ------------------- BOT ACTIONS -------------------
CACHE: Dict[str, Dict[str, Any]] = {}

async def _send_report_for_date(day_date, send_to_chat_id, bot):
    day_str = str(day_date)
    time_from, time_to, from_dt, to_dt = dubai_day_range(day_date)
    log.info("Run report for %s Almaty window=%s..%s", day_str, time_from, time_to)
    with requests.Session() as session:
        orders = orders_list_all(session, time_from, time_to, from_dt, to_dt)
        try:
            balances = fetch_driver_balances(session)
        except Exception as e:
            log.warning("fetch_driver_balances failed: %s", e)
            balances = {}
        try:
            order_ids = [o.get("id") for o in orders if o.get("id")]
            real_park_income = fetch_partner_commission(session, order_ids, time_from, time_to)
        except Exception as e:
            log.warning("fetch_partner_commission failed: %s", e)
            real_park_income = 0.0
    try:
        directory = load_driver_directory()
    except Exception as e:
        log.warning("Driver directory unavailable: %s", e)
        directory = {}
    report_rows, agg = build_report_rows(day_str, orders, directory, balances)
    if not orders:
        await bot.send_message(chat_id=send_to_chat_id,
                               text=f"⚠ За выбранный день ({day_str}) не найдено завершённых заказов.")
        return
    no_fio = any((a.fio == a.callsign or a.fio.strip() == "" or a.fio.lower() in ["неизвестно"]) for a in agg.values())
    warn = "\n⚠️ <b>Внимание:</b> Яндекс не отдает ФИО по некоторым водителям." if no_fio else ""
    write_to_google_sheets(day_str, report_rows)
    csv_path = write_csv(day_str, report_rows)
    CACHE[day_str] = {"rows": report_rows, "agg": agg}
    import asyncio
    tariffs = await asyncio.to_thread(load_tariffs)
    month_orders, month_net, month_park_income = await asyncio.to_thread(get_month_totals_from_sheets, day_date, tariffs)
    await bot.send_message(chat_id=send_to_chat_id,
                           text=format_quick_summary(day_str, agg, month_orders, month_net, tariffs, month_park_income, real_park_income) + warn,
                           parse_mode=ParseMode.HTML)
    with open(csv_path, "rb") as f:
        await bot.send_document(chat_id=send_to_chat_id, document=f, caption=f"CSV за {day_str}")

async def run_report_for_date(day_date, send_to_chat_id, context):
    await _send_report_for_date(day_date, send_to_chat_id, context.bot)

async def on_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    q = update.callback_query
    await q.answer()
    chat_id = q.message.chat_id
    now_dubai = datetime.now(DUBAI_TZ).date()
    yesterday = now_dubai - timedelta(days=1)
    try:
        if q.data == "rep:yesterday":
            await q.edit_message_text("⏳ Загружаю вчерашний отчёт...")
            await run_report_for_date(yesterday, chat_id, context)
        elif q.data == "rep:today":
            await q.edit_message_text("⏳ Загружаю сегодняшний отчёт...")
            await run_report_for_date(now_dubai, chat_id, context)
        elif q.data == "rep:month":
            now = datetime.now(DUBAI_TZ)
            buttons = []
            row = []
            for i in range(5, -1, -1):
                month_dt = now.date().replace(day=1)
                for _ in range(i):
                    month_dt = (month_dt - timedelta(days=1)).replace(day=1)
                label = month_dt.strftime("%B %Y")
                cb = f"month:{month_dt.strftime('%Y-%m')}"
                row.append(InlineKeyboardButton(label, callback_data=cb))
                if len(row) == 2:
                    buttons.append(row)
                    row = []
            if row:
                buttons.append(row)
            await q.edit_message_text("Выбери месяц:", reply_markup=InlineKeyboardMarkup(buttons))
        elif q.data.startswith("month:"):
            year_month = q.data.split(":")[1]
            year, month = int(year_month.split("-")[0]), int(year_month.split("-")[1])
            import calendar, asyncio
            month_start = now_dubai.replace(year=year, month=month, day=1)
            last_day = calendar.monthrange(year, month)[1]
            month_end = month_start.replace(day=last_day)
            if month_end > now_dubai:
                month_end = now_dubai
            month_name = month_start.strftime("%B %Y")
            await q.edit_message_text(f"⏳ Считаю отчёт за {month_name}...")
            try:
                directory = load_driver_directory()
            except Exception:
                directory = {}
            tariffs = await asyncio.to_thread(load_tariffs)
            total_orders, total_net, driver_net, park_income, tariff_groups = await fetch_range_data(month_start, month_end, directory, tariffs)
            avg_check = total_net / total_orders if total_orders else 0
            group_emoji = {"Штатные": "👔", "АРА": "🆕", "Тариф1": "2️⃣", "Другие": "📋"}
            tariff_lines = ""
            for gname, gdata in tariff_groups.items():
                if gdata["net"] == 0:
                    continue
                tariff_lines += f"{group_emoji.get(gname, '•')} {gname}: выручка <b>{gdata['net']:,.0f}</b> → доход парка <b>{gdata['park']:,.0f}</b>\n"
            summary = (
                f"📆 <b>Отчёт за {month_name}</b>\n\n"
                f"✅ Заказов: <b>{total_orders}</b>\n"
                f"💰 Выручка: <b>{total_net:,.0f}</b>\n"
                f"🏦 Доход таксопарка: <b>{park_income:,.0f}</b>\n"
                f"📊 Средний чек: <b>{avg_check:,.0f}</b>\n"
                f"👤 Водителей: <b>{len(driver_net)}</b>\n\n"
                f"📂 <b>По тарифам:</b>\n{tariff_lines}"
            )
            await context.bot.send_message(chat_id=chat_id, text=summary, parse_mode=ParseMode.HTML)
            await send_all_drivers(context.bot, chat_id, driver_net, f"Все водители за {month_name}", None)
        elif q.data == "staff:yesterday":
            await q.edit_message_text("⏳ Считаю отчёт по штатным за вчера...")
            await _send_staff_report(yesterday, chat_id, context.bot)
        elif q.data == "staff:today":
            await q.edit_message_text("⏳ Считаю отчёт по штатным за сегодня...")
            await _send_staff_report(now_dubai, chat_id, context.bot)
        elif q.data == "rep:top":
            await q.edit_message_text("⏳ Считаю ТОП водителей...")
            today = str(now_dubai)
            if today not in CACHE:
                await run_report_for_date(now_dubai, chat_id, context)
            agg = CACHE[today]["agg"]
            drivers = list(agg.values())
            top_best = sorted(drivers, key=lambda x: x.net, reverse=True)[:5]
            top_worst = sorted(drivers, key=lambda x: x.net)[:5]
            text = "🏆 <b>ТОП-5 водителей</b>\n\n"
            for d in top_best:
                text += f"• {d.fio} — {d.net:.0f} | {round(d.line_seconds/3600,2)}ч\n"
            text += "\n📉 <b>Худшие 5</b>\n\n"
            for d in top_worst:
                text += f"• {d.fio} — {d.net:.0f} | {round(d.line_seconds/3600,2)}ч\n"
            await context.bot.send_message(chat_id=chat_id, text=text, parse_mode=ParseMode.HTML)
        elif q.data == "rep:compare":
            y = str(yesterday)
            t = str(now_dubai)
            await q.edit_message_text("⏳ Сравниваю...")
            if y not in CACHE:
                await run_report_for_date(yesterday, chat_id, context)
            if t not in CACHE:
                await run_report_for_date(now_dubai, chat_id, context)
            agg_y = CACHE.get(y, {}).get("agg", {})
            agg_t = CACHE.get(t, {}).get("agg", {})
            orders_y = sum(a.done for a in agg_y.values())
            orders_t = sum(a.done for a in agg_t.values())
            net_y = sum(a.net for a in agg_y.values())
            net_t = sum(a.net for a in agg_t.values())
            await context.bot.send_message(chat_id=chat_id,
                text=f"🆚 <b>Сравнение</b>\n\n📅 Вчера ({y}): {orders_y} заказов, {net_y:.2f}\n"
                     f"📅 Сегодня ({t}): {orders_t} заказов, {net_t:.2f}\n\n"
                     f"Δ Заказы: <b>{orders_t - orders_y}</b>\nΔ Сумма: <b>{(net_t - net_y):.2f}</b>",
                parse_mode=ParseMode.HTML)
    except Exception as e:
        log.exception("button error")
        await context.bot.send_message(chat_id=chat_id, text=f"⚠ Ошибка: {e}")

async def job_daily(context: ContextTypes.DEFAULT_TYPE):
    target = CHAT_ID_INT
    if target is None:
        return
    day = datetime.now(DUBAI_TZ).date() - timedelta(days=1)
    try:
        await run_report_for_date(day, target, context)
    except Exception as e:
        await context.bot.send_message(chat_id=target, text=f"⚠ Ошибка авто-отчёта: {e}")

MAIN_KEYBOARD = ReplyKeyboardMarkup(
    keyboard=[
        [KeyboardButton("📊 Отчёт за вчера"), KeyboardButton("📈 Отчёт за сегодня")],
        [KeyboardButton("📅 Отчёт за 10 дней"), KeyboardButton("📆 Отчёт за месяц")],
        [KeyboardButton("🏆 ТОП водителей"), KeyboardButton("🆚 Сравнить вчера и сегодня")],
        [KeyboardButton("👔 Штатные водители")],
    ],
    resize_keyboard=True,
)

async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Выбери действие:", reply_markup=MAIN_KEYBOARD)

async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = (
        "📋 <b>Список команд</b>\n\n"
        "/start — Главное меню\n"
        "/settariff ФИО Процент — установить тариф\n"
        "/setara ФИО [дата] — тариф АРА\n"
        "/tariffs — список тарифов\n"
        "/importtariffs — массовый импорт тарифов\n"
        "/staff — штатные водители\n"
        "/help — это сообщение"
    )
    await update.message.reply_text(text, parse_mode=ParseMode.HTML, reply_markup=MAIN_KEYBOARD)

async def cmd_settariff(update: Update, context: ContextTypes.DEFAULT_TYPE):
    import asyncio
    args = context.args
    if len(args) < 2:
        await update.message.reply_text("Использование: <code>/settariff ФИО Процент</code>", parse_mode=ParseMode.HTML)
        return
    last_arg = args[-1].lower()
    fio = " ".join(args[:-1])
    if last_arg in ("штатный", "штат", "5"):
        tariff_name, percent = "Штатный", 5.0
    elif last_arg in ("штатный2.5", "2.5"):
        tariff_name, percent = "Штатный2.5", 2.5
    elif last_arg in ("тариф1", "тариф 1", "2"):
        tariff_name, percent = "Тариф1", 2.0
    else:
        try:
            percent = float(last_arg.replace(",", "."))
            tariff_name = "Кастом"
        except ValueError:
            await update.message.reply_text("⚠ Не понял процент.")
            return
    try:
        await asyncio.to_thread(set_tariff, fio, tariff_name, percent, None)
        await update.message.reply_text(f"✅ Тариф для <b>{fio}</b>: <b>{tariff_name} ({percent}%)</b>", parse_mode=ParseMode.HTML)
    except Exception as e:
        await update.message.reply_text(f"⚠ Ошибка: {e}")

async def cmd_setara(update: Update, context: ContextTypes.DEFAULT_TYPE):
    import asyncio, datetime as dt_mod
    args = context.args
    if len(args) < 1:
        await update.message.reply_text("Использование: <code>/setara ФИО [YYYY-MM-DD]</code>", parse_mode=ParseMode.HTML)
        return
    last_arg = args[-1]
    try:
        ara_date = dt_mod.date.fromisoformat(last_arg)
        fio = " ".join(args[:-1])
    except ValueError:
        fio = " ".join(args)
        ara_date = datetime.now(DUBAI_TZ).date()
    try:
        await asyncio.to_thread(set_tariff, fio, "АРА", None, ara_date)
        await update.message.reply_text(f"✅ АРА для <b>{fio}</b> с {ara_date}. Первые 14 дней — 1%, потом 2%.", parse_mode=ParseMode.HTML)
    except Exception as e:
        await update.message.reply_text(f"⚠ Ошибка: {e}")

async def cmd_importtariffs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    import asyncio, datetime as dt_mod
    await update.message.reply_text(f"⏳ Импортирую {len(BULK_TARIFF_IMPORT)} тарифов...")
    ok_count = 0
    fail_count = 0
    for fio, tariff, percent, date_str in BULK_TARIFF_IMPORT:
        try:
            ara_date = dt_mod.date.fromisoformat(date_str) if date_str else None
            await asyncio.to_thread(set_tariff, fio, tariff, percent, ara_date)
            ok_count += 1
        except Exception as e:
            log.warning("importtariffs failed for %s: %s", fio, e)
            fail_count += 1
        await asyncio.sleep(3)
    await update.message.reply_text(f"✅ Импорт завершён.\nУспешно: {ok_count}\nОшибок: {fail_count}", reply_markup=MAIN_KEYBOARD)

async def cmd_tariffs(update: Update, context: ContextTypes.DEFAULT_TYPE):
    import asyncio, datetime as dt_mod
    try:
        tariffs = await asyncio.to_thread(load_tariffs)
    except Exception as e:
        await update.message.reply_text(f"⚠ Ошибка: {e}")
        return
    if not tariffs:
        await update.message.reply_text("Список тарифов пуст.")
        return
    today = datetime.now(DUBAI_TZ).date()
    lines = ["📋 <b>Тарифы водителей</b>\n"]
    for fio, info in tariffs.items():
        tariff = info.get("tariff") or "—"
        if tariff.upper() == "АРА":
            ara_start = info.get("ara_start")
            if ara_start:
                pct = get_driver_percent(fio, tariffs, today)
                days = (today - ara_start).days
                lines.append(f"• {fio} — АРА (с {ara_start}, день {days+1}) → <b>{pct}%</b>")
            else:
                lines.append(f"• {fio} — АРА → <b>1%</b>")
        else:
            pct = info.get("percent")
            pct_str = f"{pct}%" if pct is not None else f"{PARK_COMMISSION_PERCENT}% (дефолт)"
            lines.append(f"• {fio} — {tariff} → <b>{pct_str}</b>")
    await update.message.reply_text("\n".join(lines), parse_mode=ParseMode.HTML)

async def _send_staff_report(day_date, chat_id, bot):
    import asyncio
    day_str = str(day_date)
    await bot.send_message(chat_id=chat_id, text=f"⏳ Считаю отчёт по штатным за {day_str}...")
    time_from, time_to, from_dt, to_dt = dubai_day_range(day_date)
    try:
        with requests.Session() as session:
            orders = await asyncio.to_thread(orders_list_all, session, time_from, time_to, from_dt, to_dt)
            balances = await asyncio.to_thread(fetch_driver_balances, session)
    except Exception as e:
        await bot.send_message(chat_id=chat_id, text=f"⚠ Ошибка: {e}")
        return
    try:
        directory = await asyncio.to_thread(load_driver_directory)
    except Exception:
        directory = {}
    _, agg = build_report_rows(day_str, orders, directory, balances)
    tariffs = await asyncio.to_thread(load_tariffs)
    staff_agg = {}
    for fio, a in agg.items():
        info = tariffs.get(fio)
        is_staff = (info is None) or ((info.get("tariff") or "").strip().lower() in ("штатный", "штатный2.5"))
        if is_staff:
            staff_agg[fio] = a
    if not staff_agg:
        await bot.send_message(chat_id=chat_id, text="Штатных водителей не найдено.", reply_markup=MAIN_KEYBOARD)
        return
    total_orders = sum(a.done for a in staff_agg.values())
    total_net = sum(a.net for a in staff_agg.values())
    park_income = calc_park_income_individual(staff_agg, tariffs, day_date)
    avg_check = round(total_net / total_orders, 2) if total_orders else 0.0
    text = (
        f"👔 <b>Штатные водители за {day_str}</b>\n\n"
        f"✅ Заказов: <b>{total_orders}</b>\n"
        f"💰 Выручка: <b>{total_net:,.0f} ₸</b>\n"
        f"🏦 Доход таксопарка: <b>{park_income:,.0f} ₸</b>\n"
        f"📊 Средний чек: <b>{avg_check:,.0f} ₸</b>\n"
        f"👤 Водителей: <b>{len(staff_agg)}</b>"
    )
    await bot.send_message(chat_id=chat_id, text=text, parse_mode=ParseMode.HTML)
    await send_staff_drivers_with_balance(bot, chat_id, staff_agg, balances, f"Штатные водители за {day_str}", MAIN_KEYBOARD)
    try:
        all_staff_fios = [fio for fio, info in tariffs.items()
                          if (info.get("tariff") or "").strip().lower() in ("штатный", "штатный2.5")]
        not_worked = [fio for fio in all_staff_fios if fio not in staff_agg]
        if not_worked:
            lines = "\n".join(f"• {fio}" for fio in sorted(not_worked))
            await bot.send_message(chat_id=chat_id,
                text=f"😴 <b>Не работали {day_str} ({len(not_worked)} чел.):</b>\n\n{lines}",
                parse_mode=ParseMode.HTML, reply_markup=MAIN_KEYBOARD)
    except Exception as e:
        log.exception("not_worked error: %s", e)

async def cmd_staff(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    args = context.args
    period = (args[0].lower() if args else "вчера")
    now_dubai = datetime.now(DUBAI_TZ).date()
    day_date = now_dubai if "сегодня" in period else now_dubai - timedelta(days=1)
    await _send_staff_report(day_date, chat_id, context.bot)

async def send_all_drivers(bot, chat_id, driver_net, title, keyboard):
    sorted_drivers = sorted(driver_net.items(), key=lambda x: x[1], reverse=True)
    chunk_size = 20
    for i in range(0, len(sorted_drivers), chunk_size):
        chunk = sorted_drivers[i:i+chunk_size]
        num_start = i + 1
        lines = "\n".join(f"{num_start + j}. {fio} — {int(net):,}".replace(",", " ") for j, (fio, net) in enumerate(chunk))
        part_num = i // chunk_size + 1
        total_parts = (len(sorted_drivers) + chunk_size - 1) // chunk_size
        header = f"👥 <b>{title}</b> (часть {part_num}/{total_parts})\n\n" if total_parts > 1 else f"👥 <b>{title}</b>\n\n"
        await bot.send_message(chat_id=chat_id, text=header + lines, parse_mode="HTML",
                               reply_markup=keyboard if i + chunk_size >= len(sorted_drivers) else None)

async def send_staff_drivers_with_balance(bot, chat_id, staff_agg, balances, title, keyboard):
    sorted_drivers = sorted(staff_agg.items(), key=lambda x: x[1].net, reverse=True)
    chunk_size = 20
    for i in range(0, len(sorted_drivers), chunk_size):
        chunk = sorted_drivers[i:i+chunk_size]
        num_start = i + 1
        lines_list = []
        for j, (fio, a) in enumerate(chunk):
            bal = balances.get(fio) or a.balance
            net_str = f"{int(a.net):,}".replace(",", " ")
            bal_str = f"{bal:,.0f}".replace(",", " ") if bal else "—"
            lines_list.append(f"{num_start + j}. {fio} — {net_str} ₸ | Баланс: {bal_str} ₸")
        lines = "\n".join(lines_list)
        part_num = i // chunk_size + 1
        total_parts = (len(sorted_drivers) + chunk_size - 1) // chunk_size
        header = f"👥 <b>{title}</b> (часть {part_num}/{total_parts})\n\n" if total_parts > 1 else f"👥 <b>{title}</b>\n\n"
        await bot.send_message(chat_id=chat_id, text=header + lines, parse_mode="HTML",
                               reply_markup=keyboard if i + chunk_size >= len(sorted_drivers) else None)

async def fetch_day_data(d, directory, tariffs=None):
    import asyncio
    def _sync():
        with requests.Session() as session:
            time_from, time_to, from_dt, to_dt = dubai_day_range(d)
            orders = orders_list_all(session, time_from, time_to, from_dt, to_dt)
            rows, agg = build_report_rows(str(d), orders, directory)
            return agg, d
    return await asyncio.to_thread(_sync)

async def fetch_range_data(start_date, end_date, directory, tariffs=None):
    import asyncio
    days = []
    d = start_date
    while d <= end_date:
        days.append(d)
        d += timedelta(days=1)
    results = await asyncio.gather(*[fetch_day_data(d, directory) for d in days], return_exceptions=True)
    total_orders = 0
    total_net = 0.0
    total_park = 0.0
    driver_net = {}
    tariff_groups = {
        "Штатные": {"net": 0.0, "park": 0.0},
        "АРА": {"net": 0.0, "park": 0.0},
        "Тариф1": {"net": 0.0, "park": 0.0},
        "Другие": {"net": 0.0, "park": 0.0},
    }
    tariffs_map = tariffs or {}
    for result in results:
        if isinstance(result, Exception):
            log.warning("Day fetch error: %s", result)
            continue
        agg, day_date = result
        for a in agg.values():
            total_orders += a.done
            total_net += a.net
            driver_net[a.fio] = driver_net.get(a.fio, 0.0) + a.net
            pct = get_driver_percent(a.fio, tariffs_map, day_date) if tariffs_map else PARK_COMMISSION_PERCENT
            park_val = a.base * (pct / 100)
            total_park += park_val
            info = tariffs_map.get(a.fio)
            if info is None:
                group = "Штатные"
            else:
                t = (info.get("tariff") or "").strip().lower()
                if t in ("штатный", "штатный2.5"):
                    group = "Штатные"
                elif t == "ара":
                    group = "АРА"
                elif t == "тариф1":
                    group = "Тариф1"
                else:
                    group = "Другие"
            tariff_groups[group]["net"] += a.net
            tariff_groups[group]["park"] += park_val
    return total_orders, total_net, driver_net, total_park, tariff_groups

async def on_text_button(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = update.message.text.strip()
    chat_id = update.effective_chat.id
    bot = context.bot
    now_dubai = datetime.now(DUBAI_TZ).date()
    yesterday = now_dubai - timedelta(days=1)
    try:
        if "вчера" in text.lower() and "сравн" not in text.lower():
            await update.message.reply_text("⏳ Загружаю вчерашний отчёт...")
            await _send_report_for_date(yesterday, chat_id, bot)
        elif "сегодня" in text.lower() and "сравн" not in text.lower():
            await update.message.reply_text("⏳ Загружаю сегодняшний отчёт...")
            await _send_report_for_date(now_dubai, chat_id, bot)
        elif "10 дней" in text.lower():
            import asyncio
            await update.message.reply_text("⏳ Считаю отчёт за 10 дней...")
            start_date = now_dubai - timedelta(days=9)
            end_date = now_dubai
            try:
                directory = load_driver_directory()
            except Exception:
                directory = {}
            tariffs_10 = await asyncio.to_thread(load_tariffs)
            total_orders, total_net, driver_net, park_income, tariff_groups = await fetch_range_data(start_date, end_date, directory, tariffs_10)
            avg_check = total_net / total_orders if total_orders else 0
            date_range = f"{start_date} – {end_date}"
            group_emoji = {"Штатные": "👔", "АРА": "🆕", "Тариф1": "2️⃣", "Другие": "📋"}
            tariff_lines = ""
            for gname, gdata in tariff_groups.items():
                if gdata["net"] == 0:
                    continue
                tariff_lines += f"{group_emoji.get(gname, '•')} {gname}: выручка <b>{gdata['net']:,.0f}</b> → доход парка <b>{gdata['park']:,.0f}</b>\n"
            summary = (
                f"📅 <b>Отчёт за 10 дней ({date_range})</b>\n\n"
                f"✅ Заказов: <b>{total_orders}</b>\n"
                f"💰 Выручка: <b>{total_net:,.0f}</b>\n"
                f"🏦 Доход таксопарка: <b>{park_income:,.0f}</b>\n"
                f"📊 Средний чек: <b>{avg_check:,.0f}</b>\n"
                f"👤 Водителей: <b>{len(driver_net)}</b>\n\n"
                f"📂 <b>По тарифам:</b>\n{tariff_lines}"
            )
            await update.message.reply_text(summary, parse_mode=ParseMode.HTML)
            await send_all_drivers(bot, chat_id, driver_net, "Все водители за 10 дней", MAIN_KEYBOARD)
        elif "месяц" in text.lower():
            buttons = []
            row = []
            for i in range(5, -1, -1):
                month_dt = now_dubai.replace(day=1)
                for _ in range(i):
                    month_dt = (month_dt - timedelta(days=1)).replace(day=1)
                label = month_dt.strftime("%B %Y")
                cb = f"month:{month_dt.strftime('%Y-%m')}"
                row.append(InlineKeyboardButton(label, callback_data=cb))
                if len(row) == 2:
                    buttons.append(row)
                    row = []
            if row:
                buttons.append(row)
            await update.message.reply_text("Выбери месяц:", reply_markup=InlineKeyboardMarkup(buttons))
        elif "штатные" in text.lower():
            kb = InlineKeyboardMarkup([[
                InlineKeyboardButton("📊 За вчера", callback_data="staff:yesterday"),
                InlineKeyboardButton("📈 За сегодня", callback_data="staff:today"),
            ]])
            await update.message.reply_text("Выбери период:", reply_markup=kb)
        elif "топ" in text.lower():
            await update.message.reply_text("⏳ Считаю ТОП водителей...")
            today = str(now_dubai)
            if today not in CACHE:
                await _send_report_for_date(now_dubai, chat_id, bot)
            agg = CACHE.get(today, {}).get("agg", {})
            if not agg:
                await update.message.reply_text("Нет данных за сегодня.", reply_markup=MAIN_KEYBOARD)
                return
            drivers = list(agg.values())
            top_best = sorted(drivers, key=lambda x: x.net, reverse=True)[:5]
            top_worst = sorted(drivers, key=lambda x: x.net)[:5]
            text_out = "🏆 <b>ТОП-5 водителей</b>\n\n"
            for d in top_best:
                text_out += f"• {d.fio} — {d.net:.0f} | {round(d.line_seconds/3600,2)}ч\n"
            text_out += "\n📉 <b>Худшие 5</b>\n\n"
            for d in top_worst:
                text_out += f"• {d.fio} — {d.net:.0f} | {round(d.line_seconds/3600,2)}ч\n"
            await update.message.reply_text(text_out, parse_mode=ParseMode.HTML, reply_markup=MAIN_KEYBOARD)
        elif "сравн" in text.lower():
            await update.message.reply_text("⏳ Сравниваю...")
            y = str(yesterday)
            t = str(now_dubai)
            if y not in CACHE:
                await _send_report_for_date(yesterday, chat_id, bot)
            if t not in CACHE:
                await _send_report_for_date(now_dubai, chat_id, bot)
            agg_y = CACHE.get(y, {}).get("agg", {})
            agg_t = CACHE.get(t, {}).get("agg", {})
            orders_y = sum(a.done for a in agg_y.values())
            orders_t = sum(a.done for a in agg_t.values())
            net_y = sum(a.net for a in agg_y.values())
            net_t = sum(a.net for a in agg_t.values())
            await update.message.reply_text(
                f"🆚 <b>Сравнение</b>\n\n📅 Вчера ({y}): {orders_y} заказов, {net_y:.2f}\n"
                f"📅 Сегодня ({t}): {orders_t} заказов, {net_t:.2f}\n\n"
                f"Δ Заказы: <b>{orders_t - orders_y}</b>\nΔ Сумма: <b>{(net_t - net_y):.2f}</b>",
                parse_mode=ParseMode.HTML, reply_markup=MAIN_KEYBOARD)
        else:
            await update.message.reply_text("Выбери действие:", reply_markup=MAIN_KEYBOARD)
    except Exception as e:
        log.exception("text button error")
        await update.message.reply_text(f"⚠ Ошибка: {e}", reply_markup=MAIN_KEYBOARD)

async def setup_bot_commands(app: Application):
    from telegram import BotCommand
    commands = [
        BotCommand("start", "Главное меню"),
        BotCommand("help", "Список команд"),
        BotCommand("settariff", "Установить тариф"),
        BotCommand("setara", "Тариф АРА"),
        BotCommand("tariffs", "Список тарифов"),
        BotCommand("importtariffs", "Массовый импорт тарифов"),
        BotCommand("staff", "Штатные водители"),
    ]
    await app.bot.set_my_commands(commands)

def build_app() -> Application:
    app = Application.builder().token(BOT_TOKEN).post_init(setup_bot_commands).build()
    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("help", cmd_help))
    app.add_handler(CommandHandler("settariff", cmd_settariff))
    app.add_handler(CommandHandler("setara", cmd_setara))
    app.add_handler(CommandHandler("tariffs", cmd_tariffs))
    app.add_handler(CommandHandler("staff", cmd_staff))
    app.add_handler(CommandHandler("importtariffs", cmd_importtariffs))
    app.add_handler(CallbackQueryHandler(on_button))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, on_text_button))
    from datetime import timezone, time as dtime
    report_time_utc = dtime(hour=3, minute=0, tzinfo=timezone.utc)
    app.job_queue.run_daily(job_daily, time=report_time_utc, days=(0,1,2,3,4,5,6), name="daily_report")
    return app

if __name__ == "__main__":
    app = build_app()
    log.info("Bot started. Авто-отчёт в 08:00 Almaty.")
    app.run_polling(close_loop=False)