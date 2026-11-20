#!/usr/bin/env python3
"""
Garmin Connect -> SQLite: ежедневная выгрузка данных здоровья.

Запуск:
  python garmin_sync.py --login              первый вход (email, пароль, код MFA), сохраняет токены
  python garmin_sync.py --start 2024-01-01   загрузка истории с указанной даты
  python garmin_sync.py                      обычный ежедневный запуск (докачивает пропущенные дни)
  python garmin_sync.py --summary-only       только пересобрать сводную таблицу и CSV

Данные:
  data/garmin.db           SQLite: raw_daily (сырой JSON по дням и метрикам) + daily_summary
  data/daily_summary.csv   сводка по дням для Excel
  data/sync.log            лог

Google Диск через API (без программ на ПК):
  python garmin_sync.py --drive-login        один раз; нужен credentials.json (см. README)
  Дальше при каждом запуске в папку GarminSync на Диске загружаются garmin.db, CSV
  и обновляется Google-таблица «Garmin — сводка по дням».

Google Диск через программу Google Drive для компьютеров (если установлена):
  После каждого запуска снимок garmin.db и daily_summary.csv копируется в папку
  "Мой диск\\GarminSync" (Google Drive для компьютеров ищется автоматически).
  Другая папка: --export-dir "D:\\путь" или переменная GARMIN_EXPORT_DIR; отключить: --export-dir none
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import getpass
import json
import logging
import os
import re
import sqlite3
import sys
import time
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DB_PATH = DATA_DIR / "garmin.db"
CSV_PATH = DATA_DIR / "daily_summary.csv"
LOG_PATH = DATA_DIR / "sync.log"
TOKEN_DIR = Path(os.environ.get("GARMIN_TOKEN_DIR", Path.home() / ".garminconnect"))

REQUEST_DELAY = 1.0      # пауза между запросами, сек (не стоит уменьшать)
REFETCH_DAYS = 3         # последние N дней всегда перезапрашиваются (данные дозаливаются после синхронизации)
CSV_SEP = ";"            # для Excel с эстонской/русской локалью
CSV_DECIMAL = ","

# Метрика -> как её получить за один день. Отсутствующие в библиотеке методы просто пропускаются.
DAILY_METRICS = {
    "stats":              lambda c, d: c.get_stats(d),
    "sleep":              lambda c, d: c.get_sleep_data(d),
    "spo2":               lambda c, d: c.get_spo2_data(d),
    "hrv":                lambda c, d: c.get_hrv_data(d),
    "respiration":        lambda c, d: c.get_respiration_data(d),
    "stress":             lambda c, d: c.get_stress_data(d),
    "heart_rate":         lambda c, d: c.get_heart_rates(d),
    "body_battery":       lambda c, d: c.get_body_battery(d, d),
    "training_readiness": lambda c, d: c.get_training_readiness(d),
    "training_status":    lambda c, d: c.get_training_status(d),
    "max_metrics":        lambda c, d: c.get_max_metrics(d),
    "activities":         lambda c, d: c.get_activities_by_date(d, d),
    "race_predictions":   lambda c, d: c.get_race_predictions(d, d, "daily"),
}

# Данные «на сейчас», а не за конкретный день: забираются один раз за запуск.
# Порог лактата дополнительно пишется в raw_daily за сегодняшнюю дату — так копится история.
SNAPSHOT_METRICS = {
    "personal_information":      lambda c: c.get_personal_information(),
    "personal_records":          lambda c: c.get_personal_record(),
    "personal_record_types":     lambda c: c.get_personal_record_types(),
    "lactate_threshold":         lambda c: c.get_lactate_threshold(latest=True),
    "lactate_threshold_history": lambda c: c.get_lactate_threshold(
        latest=False, start_date=(dt.date.today() - dt.timedelta(days=365)).isoformat(),
        end_date=dt.date.today().isoformat()),
}

# Пульсовые зоны для разбивки пробежек (как на дашборде): лёгкая — до EASY, пороговая — от THRESHOLD.
EASY_HR_MAX = int(os.environ.get("GARMIN_EASY_HR_MAX", 152))
THRESHOLD_HR_MIN = int(os.environ.get("GARMIN_THRESHOLD_HR_MIN", 170))
SERIES_MAX_PER_RUN = 400   # сколько пробежек догружать за запуск (первый раз — вся история понемногу)

# Ошибки, при которых нет смысла продолжать (имена классов garminconnect)
FATAL_ERRORS = ("Authentication", "TooManyRequests")

log = logging.getLogger("garmin_sync")


# ---------------------------------------------------------------- setup

def setup_logging() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    fmt = logging.Formatter("%(asctime)s %(levelname)s %(message)s")
    fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
    fh.setFormatter(fmt)
    log.addHandler(fh)
    if sys.stderr is not None:  # под pythonw.exe консоли нет
        sh = logging.StreamHandler()
        sh.setFormatter(fmt)
        log.addHandler(sh)
    log.setLevel(logging.INFO)


def open_db() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
    con.execute("""
        CREATE TABLE IF NOT EXISTS raw_daily (
            day        TEXT NOT NULL,
            metric     TEXT NOT NULL,
            payload    TEXT,
            fetched_at TEXT NOT NULL,
            PRIMARY KEY (day, metric)
        )""")
    con.execute("""
        CREATE TABLE IF NOT EXISTS snapshots (
            key        TEXT PRIMARY KEY,
            payload    TEXT,
            fetched_at TEXT NOT NULL
        )""")
    con.execute("""
        CREATE TABLE IF NOT EXISTS activity_series (
            activity_id INTEGER PRIMARY KEY,
            payload     TEXT,
            fetched_at  TEXT NOT NULL
        )""")
    return con


def make_client(interactive: bool):
    try:
        from garminconnect import Garmin
    except ImportError:
        log.error("Не установлена библиотека: pip install --upgrade garminconnect curl_cffi ua-generator")
        sys.exit(1)

    TOKEN_DIR.mkdir(parents=True, exist_ok=True)
    if interactive:
        email = input("Email Garmin Connect: ").strip()
        password = getpass.getpass("Пароль: ")
        client = Garmin(email, password, prompt_mfa=lambda: input("Код MFA: ").strip())
    else:
        client = Garmin()
    client.login(str(TOKEN_DIR))
    return client


# ---------------------------------------------------------------- fetch

def is_fatal(exc: Exception) -> bool:
    return any(name in type(exc).__name__ for name in FATAL_ERRORS)


def daterange(start: dt.date, end: dt.date):
    d = start
    while d <= end:
        yield d
        d += dt.timedelta(days=1)


def default_start(con: sqlite3.Connection, today: dt.date) -> dt.date:
    """С последнего загруженного дня (минус запас), чтобы закрыть пропуски, если ПК был выключен."""
    row = con.execute("SELECT MAX(day) FROM raw_daily").fetchone()
    recent = today - dt.timedelta(days=REFETCH_DAYS - 1)
    if row and row[0]:
        last = dt.date.fromisoformat(row[0]) - dt.timedelta(days=REFETCH_DAYS - 1)
        return min(last, recent)
    return recent


def sync(client, con: sqlite3.Connection, start: dt.date, end: dt.date, force: bool) -> None:
    today = dt.date.today()
    recent_from = today - dt.timedelta(days=REFETCH_DAYS - 1)
    existing = {(d, m) for d, m in con.execute("SELECT day, metric FROM raw_daily")}
    unavailable: set[str] = set()
    fetched = skipped = failed = 0

    for day in daterange(start, end):
        ds = day.isoformat()
        for metric, fn in DAILY_METRICS.items():
            if metric in unavailable:
                continue
            if not force and day < recent_from and (ds, metric) in existing:
                skipped += 1
                continue
            try:
                data = fn(client, ds)
            except AttributeError:
                log.warning("Метод для '%s' отсутствует в этой версии garminconnect — пропускаю", metric)
                unavailable.add(metric)
                continue
            except Exception as exc:  # noqa: BLE001
                if is_fatal(exc):
                    log.error("Остановка: %s: %s", type(exc).__name__, exc)
                    con.commit()
                    raise
                log.warning("%s %s: %s: %s", ds, metric, type(exc).__name__, exc)
                failed += 1
                time.sleep(REQUEST_DELAY)
                continue

            con.execute(
                "INSERT OR REPLACE INTO raw_daily(day, metric, payload, fetched_at) VALUES (?,?,?,?)",
                (ds, metric, json.dumps(data, ensure_ascii=False),
                 dt.datetime.now().isoformat(timespec="seconds")),
            )
            fetched += 1
            time.sleep(REQUEST_DELAY)
        con.commit()
        log.info("%s готово", ds)

    log.info("Итого: загружено %d, пропущено (уже есть) %d, ошибок %d", fetched, skipped, failed)


def sync_snapshots(client, con: sqlite3.Connection) -> None:
    """Рекорды и порог лактата — один запрос на метрику за запуск."""
    now = dt.datetime.now().isoformat(timespec="seconds")
    for key, fn in SNAPSHOT_METRICS.items():
        try:
            data = fn(client)
        except AttributeError:
            log.warning("Метод для '%s' отсутствует в этой версии garminconnect — пропускаю", key)
            continue
        except Exception as exc:  # noqa: BLE001
            if is_fatal(exc):
                log.error("Остановка: %s: %s", type(exc).__name__, exc)
                raise
            log.warning("%s: %s: %s", key, type(exc).__name__, exc)
            time.sleep(REQUEST_DELAY)
            continue
        payload = json.dumps(data, ensure_ascii=False)
        con.execute("INSERT OR REPLACE INTO snapshots(key, payload, fetched_at) VALUES (?,?,?)", (key, payload, now))
        if key == "lactate_threshold":
            con.execute("INSERT OR REPLACE INTO raw_daily(day, metric, payload, fetched_at) VALUES (?,?,?,?)",
                        (dt.date.today().isoformat(), key, payload, now))
        time.sleep(REQUEST_DELAY)
    con.commit()
    log.info("Профиль, рекорды и порог лактата обновлены")


def _series_from_details(details) -> dict | None:
    """Из get_activity_details оставляем только время (с), пульс и скорость (м/с) — компактно."""
    if not isinstance(details, dict):
        return None
    idx = {d.get("key"): d.get("metricsIndex") for d in details.get("metricDescriptors") or [] if isinstance(d, dict)}
    i_hr, i_v = idx.get("directHeartRate"), idx.get("directSpeed")
    i_t, scale = idx.get("sumDuration"), 1.0
    if i_t is None and idx.get("directTimestamp") is not None:
        i_t, scale = idx["directTimestamp"], 0.001
    if i_hr is None or i_t is None:
        return None
    t, hr, v = [], [], []
    for row in details.get("activityDetailMetrics") or []:
        m = row.get("metrics") if isinstance(row, dict) else None
        if not isinstance(m, list):
            continue
        tt = m[i_t] if i_t < len(m) else None
        hh = m[i_hr] if i_hr < len(m) else None
        vv = m[i_v] if i_v is not None and i_v < len(m) else None
        if tt is None:
            continue
        t.append(round(tt * scale, 1))
        hr.append(int(round(hh)) if isinstance(hh, (int, float)) else None)
        v.append(round(vv, 2) if isinstance(vv, (int, float)) else None)
    if len(t) < 10:
        return None
    t0 = t[0]
    return {"t": [round(x - t0, 1) for x in t], "hr": hr, "v": v}


def _is_run(typ: str) -> bool:
    return any(k in (typ or "") for k in ("running", "treadmill", "trail", "track"))


def sync_activity_series(client, con: sqlite3.Connection) -> None:
    """Пульс и скорость по ходу пробежки — для минут в зонах и дрейфа пульса. Каждая пробежка загружается один раз."""
    have = {r[0] for r in con.execute("SELECT activity_id FROM activity_series")}
    todo = [a for a in build_activities(con) if _is_run(a["type"]) and a["activity_id"] and a["activity_id"] not in have]
    todo.sort(key=lambda a: a["day"], reverse=True)   # сначала свежие
    done = 0
    for a in todo[:SERIES_MAX_PER_RUN]:
        aid = a["activity_id"]
        try:
            details = client.get_activity_details(aid, maxchart=2000, maxpoly=0)
        except AttributeError:
            log.warning("get_activity_details нет в этой версии garminconnect — пропускаю пульс по ходу пробежек")
            return
        except Exception as exc:  # noqa: BLE001
            if is_fatal(exc):
                log.error("Остановка: %s: %s", type(exc).__name__, exc)
                con.commit()
                raise
            log.warning("Пробежка %s (%s): %s: %s", aid, a["day"], type(exc).__name__, exc)
            time.sleep(REQUEST_DELAY)
            continue
        series = _series_from_details(details)
        con.execute("INSERT OR REPLACE INTO activity_series(activity_id, payload, fetched_at) VALUES (?,?,?)",
                    (aid, json.dumps(series) if series else None, dt.datetime.now().isoformat(timespec="seconds")))
        done += 1
        if done % 20 == 0:
            con.commit()
            log.info("Пульс по ходу пробежек: %d из %d", done, len(todo))
        time.sleep(REQUEST_DELAY)
    con.commit()
    if todo:
        log.info("Пульс по ходу пробежек: загружено %d, осталось %d", done, max(0, len(todo) - done))


def hr_zone_minutes(series: dict) -> tuple:
    """Минуты в зонах: лёгкая (≤ EASY_HR_MAX), серая, пороговая (≥ THRESHOLD_HR_MIN). Паузы > 10 с не считаются."""
    t, hr = series.get("t") or [], series.get("hr") or []
    z = [0.0, 0.0, 0.0]
    for i in range(len(t) - 1):
        h = hr[i]
        d = t[i + 1] - t[i]
        if h is None or d <= 0 or d > 10:
            continue
        z[0 if h <= EASY_HR_MAX else 2 if h >= THRESHOLD_HR_MIN else 1] += d
    if sum(z) < 60:
        return None, None, None
    return tuple(round(x / 60, 1) for x in z)


def hr_drift(series: dict):
    """Дрейф пульса (Pa:HR): насколько упала эффективность (скорость/пульс) во второй половине. Без первых 10 мин разминки.
    Считается для пробежек от 40 минут движения; < 5% — хорошая аэробная база."""
    t, hr, v = series.get("t") or [], series.get("hr") or [], series.get("v") or []
    pts = [(t[i], hr[i], v[i]) for i in range(len(t))
           if hr[i] and v[i] is not None and v[i] > 1.0 and t[i] >= 600]
    if len(pts) < 20 or pts[-1][0] - pts[0][0] < 30 * 60:
        return None
    mid = (pts[0][0] + pts[-1][0]) / 2
    def ef(part):
        return (sum(p[2] for p in part) / len(part)) / (sum(p[1] for p in part) / len(part))
    a, b = [p for p in pts if p[0] < mid], [p for p in pts if p[0] >= mid]
    if len(a) < 10 or len(b) < 10:
        return None
    return round((ef(a) / ef(b) - 1) * 100, 1)


# ---------------------------------------------------------------- summary

def dig(obj, path: str):
    """Достаёт значение по пути 'a.b.0.c'; None, если чего-то нет."""
    for key in path.split("."):
        if obj is None:
            return None
        if isinstance(obj, list):
            try:
                obj = obj[int(key)]
            except (ValueError, IndexError):
                return None
        elif isinstance(obj, dict):
            obj = obj.get(key)
        else:
            return None
    return obj


def first(raw: dict, *candidates: str):
    """Первое непустое значение из списка 'метрика:путь'."""
    for cand in candidates:
        metric, path = cand.split(":", 1)
        val = dig(raw.get(metric), path)
        if val not in (None, ""):
            return val
    return None


def hours(sec):
    return round(sec / 3600, 2) if isinstance(sec, (int, float)) else None


def activities_stats(raw: dict):
    acts = raw.get("activities")
    if not isinstance(acts, list):
        return None, None, None
    dist = sum((a.get("distance") or 0) for a in acts if isinstance(a, dict))
    dur = sum((a.get("duration") or 0) for a in acts if isinstance(a, dict))
    return len(acts), round(dist / 1000, 2), round(dur / 60, 1)


def find_key(obj, key: str):
    """Первое непустое значение ключа на любой глубине. Garmin кладёт статус тренированности
    в словари с id устройства в качестве ключа, поэтому точный путь заранее неизвестен."""
    if isinstance(obj, dict):
        if obj.get(key) not in (None, ""):
            return obj[key]
        for v in obj.values():
            r = find_key(v, key)
            if r is not None:
                return r
    elif isinstance(obj, list):
        for v in obj:
            r = find_key(v, key)
            if r is not None:
                return r
    return None


def status_code(phrase):
    """PRODUCTIVE_6 -> PRODUCTIVE"""
    if not isinstance(phrase, str):
        return None
    return re.sub(r"_\d+$", "", phrase)


def num1(v):
    return round(v, 1) if isinstance(v, (int, float)) else None


def race_seconds(raw: dict, key: str):
    """Прогноз Garmin в секундах: ответ бывает списком дней или одним объектом."""
    val = first(raw, f"race_predictions:0.{key}", f"race_predictions:{key}")
    return int(round(val)) if isinstance(val, (int, float)) and val > 0 else None


SUMMARY_COLUMNS = [
    "day",
    "spo2_avg", "spo2_min", "spo2_sleep_avg",
    "sleep_h", "deep_h", "light_h", "rem_h", "awake_h", "sleep_score",
    "hrv_night", "hrv_week", "hrv_status",
    "rhr", "resp_sleep", "stress_avg",
    "bb_max", "bb_min",
    "steps", "kcal", "intensity_moderate_min", "intensity_vigorous_min",
    "readiness", "vo2max",
    "activities", "activities_km", "activities_min",
    "race_5k_s", "race_10k_s", "race_hm_s", "race_marathon_s",
    "floors_up", "floors_down",
    "training_status", "acute_load", "chronic_load", "acwr", "acwr_status",
    "load_aerobic_low", "load_aerobic_high", "load_anaerobic", "load_balance",
]


def summarize_day(day: str, raw: dict) -> dict:
    n_act, km, mins = activities_stats(raw)
    ts = raw.get("training_status")
    return {
        "day": day,
        "spo2_avg": first(raw, "spo2:averageSpO2", "sleep:dailySleepDTO.averageSpO2Value"),
        "spo2_min": first(raw, "spo2:lowestSpO2", "sleep:dailySleepDTO.lowestSpO2Value"),
        "spo2_sleep_avg": first(raw, "spo2:avgSleepSpO2"),
        "sleep_h": hours(first(raw, "sleep:dailySleepDTO.sleepTimeSeconds")),
        "deep_h": hours(first(raw, "sleep:dailySleepDTO.deepSleepSeconds")),
        "light_h": hours(first(raw, "sleep:dailySleepDTO.lightSleepSeconds")),
        "rem_h": hours(first(raw, "sleep:dailySleepDTO.remSleepSeconds")),
        "awake_h": hours(first(raw, "sleep:dailySleepDTO.awakeSleepSeconds")),
        "sleep_score": first(raw, "sleep:dailySleepDTO.sleepScores.overall.value"),
        "hrv_night": first(raw, "hrv:hrvSummary.lastNightAvg"),
        "hrv_week": first(raw, "hrv:hrvSummary.weeklyAvg"),
        "hrv_status": first(raw, "hrv:hrvSummary.status"),
        "rhr": first(raw, "stats:restingHeartRate", "heart_rate:restingHeartRate"),
        "resp_sleep": first(raw, "respiration:avgSleepRespirationValue",
                            "sleep:dailySleepDTO.averageRespirationValue"),
        "stress_avg": first(raw, "stats:averageStressLevel"),
        "bb_max": first(raw, "stats:bodyBatteryHighestValue"),
        "bb_min": first(raw, "stats:bodyBatteryLowestValue"),
        "steps": first(raw, "stats:totalSteps"),
        "kcal": first(raw, "stats:totalKilocalories"),
        "intensity_moderate_min": first(raw, "stats:moderateIntensityMinutes"),
        "intensity_vigorous_min": first(raw, "stats:vigorousIntensityMinutes"),
        "readiness": first(raw, "training_readiness:0.score", "training_readiness:score"),
        "vo2max": first(raw, "max_metrics:0.generic.vo2MaxPreciseValue",
                        "max_metrics:0.generic.vo2MaxValue",
                        "max_metrics:generic.vo2MaxPreciseValue"),
        "activities": n_act,
        "activities_km": km,
        "activities_min": mins,
        # Прогноз Garmin (Race Predictor), секунды
        "race_5k_s": race_seconds(raw, "time5K"),
        "race_10k_s": race_seconds(raw, "time10K"),
        "race_hm_s": race_seconds(raw, "timeHalfMarathon"),
        "race_marathon_s": race_seconds(raw, "timeMarathon"),
        # Этажи (барометр часов)
        "floors_up": num1(first(raw, "stats:floorsAscended")),
        "floors_down": num1(first(raw, "stats:floorsDescended")),
        # Статус тренированности и нагрузка
        "training_status": status_code(find_key(ts, "trainingStatusFeedbackPhrase")),
        "acute_load": num1(find_key(ts, "dailyTrainingLoadAcute")),
        "chronic_load": num1(find_key(ts, "dailyTrainingLoadChronic")),
        "acwr": num1(find_key(ts, "dailyAcuteChronicWorkloadRatio")),
        "acwr_status": find_key(ts, "acwrStatus"),
        "load_aerobic_low": num1(find_key(ts, "monthlyLoadAerobicLow")),
        "load_aerobic_high": num1(find_key(ts, "monthlyLoadAerobicHigh")),
        "load_anaerobic": num1(find_key(ts, "monthlyLoadAnaerobic")),
        "load_balance": status_code(find_key(ts, "trainingBalanceFeedbackPhrase")),
    }


def build_summary(con: sqlite3.Connection) -> list[dict]:
    by_day: dict[str, dict] = {}
    for day, metric, payload in con.execute("SELECT day, metric, payload FROM raw_daily ORDER BY day"):
        try:
            by_day.setdefault(day, {})[metric] = json.loads(payload) if payload else None
        except json.JSONDecodeError:
            continue

    rows = [summarize_day(day, raw) for day, raw in sorted(by_day.items())]

    con.execute("DROP TABLE IF EXISTS daily_summary")
    cols_sql = ", ".join(f"{c} {'TEXT PRIMARY KEY' if c == 'day' else ''}" for c in SUMMARY_COLUMNS)
    con.execute(f"CREATE TABLE daily_summary ({cols_sql})")
    con.executemany(
        f"INSERT INTO daily_summary VALUES ({','.join('?' * len(SUMMARY_COLUMNS))})",
        [[r[c] for c in SUMMARY_COLUMNS] for r in rows],
    )
    con.commit()

    def fmt(v):
        if isinstance(v, float):
            return str(v).replace(".", CSV_DECIMAL)
        return "" if v is None else v

    with open(CSV_PATH, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=CSV_SEP)
        w.writerow(SUMMARY_COLUMNS)
        for r in rows:
            w.writerow([fmt(r[c]) for c in SUMMARY_COLUMNS])
    return rows


# ---------------------------------------------------------------- extra tables

def _hms(sec) -> str | None:
    if not isinstance(sec, (int, float)) or sec <= 0:
        return None
    s = int(round(sec))
    h, m, x = s // 3600, s % 3600 // 60, s % 60
    return f"{h}:{m:02d}:{x:02d}" if h else f"{m}:{x:02d}"


ACTIVITY_COLUMNS = [
    "activity_id", "day", "start_local", "type", "name",
    "distance_km", "duration_min", "moving_min", "pace_s_km", "pace",
    "avg_hr", "max_hr", "avg_cadence", "avg_power", "elev_gain_m",
    "aerobic_te", "anaerobic_te", "te_label", "training_load", "vo2max", "kcal",
    "hr_easy_min", "hr_grey_min", "hr_hard_min", "hr_drift_pct",
]


def build_activities(con: sqlite3.Connection) -> list[dict]:
    """Одна строка на тренировку — из уже загруженных списков тренировок по дням."""
    rows, seen = [], set()
    series = {}
    try:
        series = {aid: p for aid, p in con.execute("SELECT activity_id, payload FROM activity_series") if p}
    except sqlite3.OperationalError:
        pass
    for day, payload in con.execute("SELECT day, payload FROM raw_daily WHERE metric='activities' ORDER BY day"):
        try:
            acts = json.loads(payload) if payload else []
        except json.JSONDecodeError:
            continue
        for a in acts if isinstance(acts, list) else []:
            if not isinstance(a, dict):
                continue
            aid = a.get("activityId")
            if aid in seen:
                continue
            seen.add(aid)
            typ = dig(a, "activityType.typeKey") or ""
            dist = a.get("distance") or 0
            speed = a.get("averageSpeed")
            running = any(k in typ for k in ("running", "treadmill", "trail", "track"))
            pace = 1000 / speed if running and isinstance(speed, (int, float)) and speed > 0 else None
            rows.append({
                "activity_id": aid,
                "day": (a.get("startTimeLocal") or day)[:10],
                "start_local": a.get("startTimeLocal"),
                "type": typ,
                "name": a.get("activityName"),
                "distance_km": round(dist / 1000, 2) if dist else None,
                "duration_min": num1((a.get("duration") or 0) / 60) or None,
                "moving_min": num1(a["movingDuration"] / 60) if a.get("movingDuration") else None,
                "pace_s_km": int(round(pace)) if pace else None,
                "pace": _hms(pace),
                "avg_hr": a.get("averageHR"),
                "max_hr": a.get("maxHR"),
                "avg_cadence": num1(a.get("averageRunningCadenceInStepsPerMinute")),
                "avg_power": num1(a.get("avgPower")),
                "elev_gain_m": num1(a.get("elevationGain")),
                "aerobic_te": num1(a.get("aerobicTrainingEffect")),
                "anaerobic_te": num1(a.get("anaerobicTrainingEffect")),
                "te_label": a.get("trainingEffectLabel"),
                "training_load": num1(a.get("activityTrainingLoad")),
                "vo2max": a.get("vO2MaxValue"),
                "kcal": num1(a.get("calories")),
            })
            zones, drift = (None, None, None), None
            if aid in series:
                try:
                    s = json.loads(series[aid])
                    zones, drift = hr_zone_minutes(s), hr_drift(s)
                except (json.JSONDecodeError, TypeError, IndexError):
                    pass
            rows[-1].update({"hr_easy_min": zones[0], "hr_grey_min": zones[1], "hr_hard_min": zones[2], "hr_drift_pct": drift})
    return rows


RECORD_COLUMNS = ["type_id", "record", "value", "display", "day", "activity_id", "activity_name"]
RECORD_NAMES = {1: "1 км", 2: "1 миля", 3: "5 км", 4: "10 км", 5: "Полумарафон", 6: "Марафон",
                7: "Самая длинная пробежка", 8: "Самая длинная велопоездка", 9: "Максимальный набор высоты",
                10: "Велосипед 20 км", 11: "Велосипед 40 км",
                12: "Больше всего шагов за день", 13: "Больше всего шагов за неделю",
                14: "Больше всего шагов за месяц", 15: "Самая длинная серия дней с целью по шагам",
                16: "Текущая серия дней с целью по шагам"}
RECORD_TIME_IDS = {1, 2, 3, 4, 5, 6, 10, 11}      # значение — секунды
RECORD_DISTANCE_IDS = {7, 8}                       # значение — метры


def _snapshot(con: sqlite3.Connection, key: str):
    row = con.execute("SELECT payload FROM snapshots WHERE key=?", (key,)).fetchone()
    try:
        return json.loads(row[0]) if row and row[0] else None
    except json.JSONDecodeError:
        return None


def build_records(con: sqlite3.Connection) -> list[dict]:
    prs = _snapshot(con, "personal_records")
    if isinstance(prs, dict):
        prs = prs.get("personalRecords") or prs.get("records") or []
    if not isinstance(prs, list):
        return []
    types = _snapshot(con, "personal_record_types") or []
    type_keys = {t.get("id"): t.get("key") for t in types if isinstance(t, dict)} if isinstance(types, list) else {}
    rows = []
    for p in prs:
        if not isinstance(p, dict) or p.get("value") is None:
            continue
        tid = p.get("typeId")
        val = p["value"]
        if tid in RECORD_TIME_IDS:
            display = _hms(val)
        elif tid in RECORD_DISTANCE_IDS:
            display = f"{val / 1000:.2f} км".replace(".", ",")
        elif tid == 9:
            display = f"{round(val)} м"
        elif tid in (15, 16):
            display = f"{int(val)} дн."
        else:
            display = str(int(round(val))) if isinstance(val, (int, float)) else str(val)
        day = (p.get("prStartTimeLocalFormatted") or p.get("actStartDateTimeInGMTFormatted")
               or p.get("prStartTimeGmtFormatted") or "")[:10] or None
        rows.append({
            "type_id": tid,
            "record": RECORD_NAMES.get(tid) or type_keys.get(tid) or f"Рекорд {tid}",
            "value": val, "display": display, "day": day,
            "activity_id": p.get("activityId") or None, "activity_name": p.get("activityName"),
        })
    rows.sort(key=lambda r: (r["type_id"] is None, r["type_id"] or 0))
    return rows


LT_COLUMNS = ["day", "lt_hr", "lt_speed_ms", "lt_pace_s_km", "lt_pace", "lt_power_w"]


def _lt_speed(v):
    """Garmin отдаёт скорость ПАНО в десятых долях: 0,36 означает 3,6 м/с."""
    if not isinstance(v, (int, float)) or v <= 0:
        return None
    return round(v * 10 if v < 1.5 else v, 3)


def build_lactate(con: sqlite3.Connection) -> list[dict]:
    by_day: dict[str, dict] = {}

    def put(day, field, value):
        if day and value is not None:
            by_day.setdefault(str(day)[:10], {})[field] = value

    # История за год одним запросом: списки точек {дата, значение}
    hist = _snapshot(con, "lactate_threshold_history") or {}
    for field, src in (("lt_speed_ms", "speed"), ("lt_hr", "heart_rate"), ("lt_power_w", "power")):
        items = hist.get(src) if isinstance(hist, dict) else None
        for it in items if isinstance(items, list) else []:
            if not isinstance(it, dict):
                continue
            day = it.get("calendarDate") or it.get("from") or it.get("date")
            val = it.get("value")
            if val is None:
                val = it.get("functionalThresholdPower") if field == "lt_power_w" else None
            put(day, field, _lt_speed(val) if field == "lt_speed_ms" else val)

    # Ежедневные снимки «последнего» значения (копятся с каждым запуском)
    for _, payload in con.execute("SELECT day, payload FROM raw_daily WHERE metric='lactate_threshold'"):
        try:
            lt = json.loads(payload) if payload else {}
        except json.JSONDecodeError:
            continue
        sh = lt.get("speed_and_heart_rate") or {}
        day = sh.get("calendarDate")
        put(day, "lt_hr", sh.get("heartRate"))
        put(day, "lt_speed_ms", _lt_speed(sh.get("speed")))
        pw = lt.get("power") or {}
        put(pw.get("calendarDate") or day, "lt_power_w", pw.get("functionalThresholdPower"))

    rows = []
    for day in sorted(by_day):
        r = {c: None for c in LT_COLUMNS}
        r.update(by_day[day])
        r["day"] = day
        sp = r.get("lt_speed_ms")
        if sp:
            r["lt_pace_s_km"] = int(round(1000 / sp))
            r["lt_pace"] = _hms(1000 / sp)
        rows.append(r)
    return rows


SLEEP_STAGE_COLUMNS = ["day", "seq", "start_local", "minutes", "stage"]
SLEEP_STAGE_NIGHTS = 120          # сколько последних ночей выгружать на лист «Фазы сна»
# Garmin: activityLevel в sleepLevels — 0 глубокий, 1 лёгкий, 2 REM, 3 бодрствование
SLEEP_STAGE_NAMES = {0: "deep", 1: "light", 2: "rem", 3: "awake"}


def _parse_ts(v):
    """Время из Garmin: миллисекунды эпохи или строка ISO 'YYYY-MM-DDTHH:MM:SS.0'."""
    if isinstance(v, (int, float)):
        return dt.datetime.fromtimestamp(v / 1000, tz=dt.timezone.utc).replace(tzinfo=None)
    if isinstance(v, str):
        try:
            return dt.datetime.fromisoformat(v.replace("Z", "").split(".")[0])
        except ValueError:
            return None
    return None


def build_sleep_stages(con: sqlite3.Connection) -> list[dict]:
    """Последовательность фаз сна за ночь: по порядку, с длительностью каждого отрезка."""
    rows_db = con.execute(
        "SELECT day, payload FROM raw_daily WHERE metric='sleep' ORDER BY day DESC LIMIT ?",
        (SLEEP_STAGE_NIGHTS,)).fetchall()
    out = []
    for day, payload in reversed(rows_db):
        try:
            sleep = json.loads(payload) if payload else None
        except json.JSONDecodeError:
            continue
        levels = find_key(sleep, "sleepLevels")
        if not isinstance(levels, list) or not levels:
            continue
        # сдвиг местного времени относительно GMT — из начала сна
        dto = (sleep or {}).get("dailySleepDTO") or {}
        g, l = dto.get("sleepStartTimestampGMT"), dto.get("sleepStartTimestampLocal")
        offset = dt.timedelta(milliseconds=(l - g)) if isinstance(g, (int, float)) and isinstance(l, (int, float)) else dt.timedelta(0)
        segs = []
        for it in levels:
            if not isinstance(it, dict):
                continue
            s, e = _parse_ts(it.get("startGMT")), _parse_ts(it.get("endGMT"))
            lvl = it.get("activityLevel")
            if s is None or e is None or lvl is None or e <= s:
                continue
            segs.append((s, e, SLEEP_STAGE_NAMES.get(int(round(lvl)), str(lvl))))
        segs.sort()
        # склеиваем соседние отрезки одной фазы
        merged = []
        for s, e, st in segs:
            if merged and merged[-1][2] == st and merged[-1][1] >= s:
                merged[-1] = (merged[-1][0], max(merged[-1][1], e), st)
            else:
                merged.append((s, e, st))
        for i, (s, e, st) in enumerate(merged):
            out.append({"day": day, "seq": i + 1, "start_local": (s + offset).strftime("%H:%M"),
                        "minutes": round((e - s).total_seconds() / 60), "stage": st})
    return out


PROFILE_COLUMNS = ["age", "sex", "height_cm", "weight_kg", "updated"]


def build_profile(con: sqlite3.Connection) -> list[dict]:
    """Возраст и пол для сравнения с возрастными нормами. Дату рождения не сохраняем — только возраст."""
    info = _snapshot(con, "personal_information")
    if not isinstance(info, dict):
        return []
    birth = find_key(info, "birthDate")
    age = None
    if isinstance(birth, str) and len(birth) >= 10:
        try:
            b = dt.date.fromisoformat(birth[:10])
            t = dt.date.today()
            age = t.year - b.year - ((t.month, t.day) < (b.month, b.day))
        except ValueError:
            pass
    sex = find_key(info, "genderType") or find_key(info, "gender")
    sex = {"MALE": "male", "FEMALE": "female"}.get(str(sex).upper(), sex) if sex else None
    height = find_key(info, "height")
    weight = find_key(info, "weight")
    if isinstance(weight, (int, float)) and weight > 1000:  # Garmin хранит вес в граммах
        weight = weight / 1000
    row = con.execute("SELECT fetched_at FROM snapshots WHERE key='personal_information'").fetchone()
    return [{"age": age, "sex": sex, "height_cm": num1(height), "weight_kg": num1(weight),
             "updated": row[0] if row else None}]


# Дополнительные таблицы: имя файла/таблицы SQLite, лист в Google-таблице, столбцы, построитель
EXTRA_TABLES = [
    ("activities", "Тренировки", ACTIVITY_COLUMNS, build_activities),
    ("records", "Рекорды", RECORD_COLUMNS, build_records),
    ("lactate_threshold", "Порог лактата", LT_COLUMNS, build_lactate),
    ("profile", "Профиль", PROFILE_COLUMNS, build_profile),
    ("sleep_stages", "Фазы сна", SLEEP_STAGE_COLUMNS, build_sleep_stages),
]


def write_csv(path: Path, columns: list[str], rows: list[dict]) -> None:
    def fmt(v):
        if isinstance(v, float):
            return str(v).replace(".", CSV_DECIMAL)
        return "" if v is None else v
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=CSV_SEP)
        w.writerow(columns)
        for r in rows:
            w.writerow([fmt(r.get(c)) for c in columns])


def build_extras(con: sqlite3.Connection) -> list[dict]:
    """Строит дополнительные таблицы, пишет их в SQLite и CSV. Ошибка в одной не мешает другим."""
    extras = []
    for name, tab, columns, builder in EXTRA_TABLES:
        try:
            rows = builder(con)
        except Exception as exc:  # noqa: BLE001
            log.warning("Таблица %s не собрана: %s: %s", name, type(exc).__name__, exc)
            continue
        con.execute(f"DROP TABLE IF EXISTS {name}")
        con.execute(f"CREATE TABLE {name} ({', '.join(columns)})")
        con.executemany(f"INSERT INTO {name} VALUES ({','.join('?' * len(columns))})",
                        [[r.get(c) for c in columns] for r in rows])
        path = DATA_DIR / f"{name}.csv"
        write_csv(path, columns, rows)
        extras.append({"name": name, "tab": tab, "columns": columns, "rows": rows, "csv": path})
        log.info("Таблица «%s»: %d строк", tab, len(rows))
    con.commit()
    return extras


# ---------------------------------------------------------------- export

DRIVE_ROOT_NAMES = ("My Drive", "Мой диск", "Minu ketas", "Google Drive")
EXPORT_SUBFOLDER = "GarminSync"


def find_google_drive() -> Path | None:
    """Ищет папку Google Drive для компьютеров: G:\\My Drive (или локализованное имя) либо %USERPROFILE%\\Google Drive."""
    candidates = []
    if os.name == "nt":
        import string
        for letter in string.ascii_uppercase:
            for name in DRIVE_ROOT_NAMES:
                candidates.append(Path(f"{letter}:\\{name}"))
    for name in DRIVE_ROOT_NAMES:
        candidates.append(Path.home() / name)
    for c in candidates:
        try:
            if c.is_dir():
                return c
        except OSError:
            continue
    return None


def resolve_export_dir(setting: str) -> Path | None:
    if not setting or setting.lower() == "none":
        return None
    if setting.lower() == "auto":
        root = find_google_drive()
        if root is None:
            log.debug("Google Drive для компьютеров не найден — копирование на Диск пропущено "
                     "(укажите папку через --export-dir)")
            return None
        return root / EXPORT_SUBFOLDER
    return Path(setting).expanduser()


def export_to_drive(con: sqlite3.Connection, setting: str, extras: list[dict] | None = None) -> None:
    """Кладёт в папку Диска целостный снимок базы и CSV.
    Пишем во временный файл и подменяем — так Диск никогда не заливает недописанный файл."""
    dst = resolve_export_dir(setting)
    if dst is None:
        return
    try:
        dst.mkdir(parents=True, exist_ok=True)

        tmp_db = dst / "garmin.db.tmp"
        tmp_db.unlink(missing_ok=True)
        snap = sqlite3.connect(tmp_db)
        with snap:
            con.backup(snap)
        snap.close()
        os.replace(tmp_db, dst / "garmin.db")

        for src in [CSV_PATH] + [e["csv"] for e in extras or []]:
            tmp_csv = dst / (src.name + ".tmp")
            tmp_csv.write_bytes(src.read_bytes())
            os.replace(tmp_csv, dst / src.name)

        log.info("Скопировано на Google Диск: %s", dst)
    except OSError as exc:
        # Диск может быть не смонтирован (Drive не запущен) — локальные данные при этом в порядке
        log.warning("Не удалось скопировать в %s: %s", dst, exc)


# ---------------------------------------------------------------- Google Drive API

GDRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.file"]  # доступ только к файлам, созданным этим скриптом
GDRIVE_CLIENT_SECRET = BASE_DIR / "credentials.json"
GDRIVE_TOKEN = Path.home() / ".garmin-sync" / "gdrive_token.json"
GDRIVE_STATE = DATA_DIR / "gdrive_state.json"   # id папки и файлов на Диске
GDRIVE_FOLDER = "GarminSync"
GDRIVE_SHEET_TITLE = "Garmin — сводка по дням"


def gdrive_credentials(interactive: bool):
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials

    creds = None
    if GDRIVE_TOKEN.exists():
        creds = Credentials.from_authorized_user_file(str(GDRIVE_TOKEN), GDRIVE_SCOPES)
    if creds and not creds.valid and creds.refresh_token:
        try:
            creds.refresh(Request())
        except Exception as exc:  # noqa: BLE001
            log.warning("Токен Google устарел (%s)", exc)
            creds = None
    if not creds or not creds.valid:
        if not interactive:
            raise RuntimeError("нет действующего токена Google — запустите: garmin_sync.py --drive-login")
        if not GDRIVE_CLIENT_SECRET.exists():
            raise FileNotFoundError(f"не найден {GDRIVE_CLIENT_SECRET} (см. README, раздел про Google Drive API)")
        from google_auth_oauthlib.flow import InstalledAppFlow
        flow = InstalledAppFlow.from_client_secrets_file(str(GDRIVE_CLIENT_SECRET), GDRIVE_SCOPES)
        creds = flow.run_local_server(port=0)
    GDRIVE_TOKEN.parent.mkdir(parents=True, exist_ok=True)
    GDRIVE_TOKEN.write_text(creds.to_json(), encoding="utf-8")
    return creds


def _alive(drive, file_id: str | None) -> bool:
    if not file_id:
        return False
    from googleapiclient.errors import HttpError
    try:
        f = drive.files().get(fileId=file_id, fields="id,trashed").execute()
        return not f.get("trashed")
    except HttpError as exc:
        if exc.resp.status == 404:
            return False
        raise


def _put_file(drive, state: dict, key: str, name: str, path: Path, mime: str, folder_id: str) -> None:
    from googleapiclient.http import MediaFileUpload
    media = MediaFileUpload(str(path), mimetype=mime, resumable=True)
    if _alive(drive, state.get(key)):
        drive.files().update(fileId=state[key], media_body=media).execute()
    else:
        f = drive.files().create(body={"name": name, "parents": [folder_id]},
                                 media_body=media, fields="id").execute()
        state[key] = f["id"]


def _sheet_values(columns: list[str], rows: list[dict]) -> list[list]:
    return [columns] + [["" if r.get(c) is None else r.get(c) for c in columns] for r in rows]


def _update_tabs(sheets, sid: str, extras: list[dict]) -> None:
    """Каждая дополнительная таблица — отдельный лист; недостающие листы создаются."""
    if not extras:
        return
    meta = sheets.spreadsheets().get(spreadsheetId=sid, fields="sheets.properties.title").execute()
    titles = {s["properties"]["title"] for s in meta.get("sheets", [])}
    missing = [e["tab"] for e in extras if e["tab"] not in titles]
    if missing:
        sheets.spreadsheets().batchUpdate(spreadsheetId=sid, body={
            "requests": [{"addSheet": {"properties": {"title": t}}} for t in missing]}).execute()
    for e in extras:
        rng = f"'{e['tab']}'"
        sheets.spreadsheets().values().clear(spreadsheetId=sid, range=f"{rng}!A:ZZ", body={}).execute()
        sheets.spreadsheets().values().update(
            spreadsheetId=sid, range=f"{rng}!A1", valueInputOption="RAW",
            body={"values": _sheet_values(e["columns"], e["rows"])}).execute()


def _update_sheet(drive, sheets, state: dict, folder_id: str, rows: list[dict]) -> None:
    if not _alive(drive, state.get("sheet_id")):
        sp = sheets.spreadsheets().create(
            body={"properties": {"title": GDRIVE_SHEET_TITLE}}, fields="spreadsheetId").execute()
        sid = sp["spreadsheetId"]
        parents = drive.files().get(fileId=sid, fields="parents").execute().get("parents", [])
        drive.files().update(fileId=sid, addParents=folder_id,
                             removeParents=",".join(parents), fields="id").execute()
        state["sheet_id"] = sid
    sid = state["sheet_id"]
    values = [SUMMARY_COLUMNS] + [["" if r[c] is None else r[c] for c in SUMMARY_COLUMNS] for r in rows]
    sheets.spreadsheets().values().clear(spreadsheetId=sid, range="A:ZZ", body={}).execute()
    sheets.spreadsheets().values().update(
        spreadsheetId=sid, range="A1", valueInputOption="RAW", body={"values": values}).execute()


def upload_to_gdrive(con: sqlite3.Connection, rows: list[dict], extras: list[dict] | None = None) -> None:
    """Загрузка на Google Диск через API. Включается после однократного --drive-login."""
    if not GDRIVE_TOKEN.exists():
        return
    try:
        from googleapiclient.discovery import build
    except ImportError:
        log.warning("Для загрузки на Диск: pip install google-api-python-client google-auth-oauthlib")
        return
    try:
        creds = gdrive_credentials(interactive=False)
        drive = build("drive", "v3", credentials=creds, cache_discovery=False)
        sheets = build("sheets", "v4", credentials=creds, cache_discovery=False)

        state = json.loads(GDRIVE_STATE.read_text(encoding="utf-8")) if GDRIVE_STATE.exists() else {}
        if not _alive(drive, state.get("folder_id")):
            f = drive.files().create(body={"name": GDRIVE_FOLDER,
                                           "mimeType": "application/vnd.google-apps.folder"},
                                     fields="id").execute()
            state = {"folder_id": f["id"]}  # папка новая — старые id файлов не нужны
        folder_id = state["folder_id"]

        snapshot = DATA_DIR / "garmin_snapshot.db"
        snapshot.unlink(missing_ok=True)
        snap = sqlite3.connect(snapshot)
        with snap:
            con.backup(snap)
        snap.close()

        _put_file(drive, state, "db_id", "garmin.db", snapshot, "application/x-sqlite3", folder_id)
        _put_file(drive, state, "csv_id", "daily_summary.csv", CSV_PATH, "text/csv", folder_id)
        for e in extras or []:
            _put_file(drive, state, f"csv_{e['name']}_id", e["csv"].name, e["csv"], "text/csv", folder_id)
        _update_sheet(drive, sheets, state, folder_id, rows)
        _update_tabs(sheets, state["sheet_id"], extras or [])

        GDRIVE_STATE.write_text(json.dumps(state, indent=2), encoding="utf-8")
        log.info("Загружено на Google Диск (папка %s, таблица «%s»)", GDRIVE_FOLDER, GDRIVE_SHEET_TITLE)
    except Exception as exc:  # noqa: BLE001  — локальные данные в порядке, в следующий раз попробуем снова
        log.warning("Загрузка на Google Диск не удалась: %s: %s", type(exc).__name__, exc)


# ---------------------------------------------------------------- main

def main() -> int:
    ap = argparse.ArgumentParser(description="Выгрузка данных Garmin Connect в SQLite")
    ap.add_argument("--login", action="store_true", help="интерактивный вход и сохранение токенов")
    ap.add_argument("--start", type=dt.date.fromisoformat, help="начальная дата YYYY-MM-DD")
    ap.add_argument("--end", type=dt.date.fromisoformat, help="конечная дата YYYY-MM-DD (по умолчанию сегодня)")
    ap.add_argument("--force", action="store_true", help="перезапросить даже уже загруженные дни")
    ap.add_argument("--summary-only", action="store_true", help="только пересобрать сводку и CSV")
    ap.add_argument("--export-dir", default=os.environ.get("GARMIN_EXPORT_DIR", "auto"),
                    help="куда копировать результат (папка Google Диска); "
                         "'auto' — найти Google Drive для компьютеров, 'none' — не копировать")
    ap.add_argument("--drive-login", action="store_true",
                    help="разрешить загрузку на Google Диск через API (открывает браузер, один раз)")
    args = ap.parse_args()

    setup_logging()
    con = open_db()

    if args.drive_login:
        try:
            gdrive_credentials(interactive=True)
        except Exception as exc:  # noqa: BLE001
            log.error("Не удалось подключить Google Диск: %s: %s", type(exc).__name__, exc)
            return 4
        log.info("Google Диск подключён, токен сохранён в %s", GDRIVE_TOKEN)
        upload_to_gdrive(con, build_summary(con), build_extras(con))
        return 0

    if not args.summary_only:
        try:
            client = make_client(interactive=args.login)
        except Exception as exc:  # noqa: BLE001
            log.error("Не удалось войти (%s: %s). Запустите: python garmin_sync.py --login",
                      type(exc).__name__, exc)
            return 2
        if args.login:
            log.info("Вход выполнен, токены сохранены в %s", TOKEN_DIR)

        today = dt.date.today()
        end = args.end or today
        start = args.start or default_start(con, today)
        log.info("Синхронизация %s … %s", start, end)
        try:
            sync(client, con, start, end, args.force)
            sync_snapshots(client, con)
            sync_activity_series(client, con)
        except Exception:  # noqa: BLE001  (фатальная ошибка уже залогирована)
            rows, extras = build_summary(con), build_extras(con)
            export_to_drive(con, args.export_dir, extras)
            upload_to_gdrive(con, rows, extras)
            return 3

    rows = build_summary(con)
    extras = build_extras(con)
    log.info("Сводка: %d дней -> %s", len(rows), CSV_PATH)
    export_to_drive(con, args.export_dir, extras)
    upload_to_gdrive(con, rows, extras)
    return 0


if __name__ == "__main__":
    sys.exit(main())
