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
}

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
]


def summarize_day(day: str, raw: dict) -> dict:
    n_act, km, mins = activities_stats(raw)
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


def export_to_drive(con: sqlite3.Connection, setting: str) -> None:
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

        tmp_csv = dst / "daily_summary.csv.tmp"
        tmp_csv.write_bytes(CSV_PATH.read_bytes())
        os.replace(tmp_csv, dst / "daily_summary.csv")

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


def upload_to_gdrive(con: sqlite3.Connection, rows: list[dict]) -> None:
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
        _update_sheet(drive, sheets, state, folder_id, rows)

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
        upload_to_gdrive(con, build_summary(con))
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
        except Exception:  # noqa: BLE001  (фатальная ошибка уже залогирована)
            rows = build_summary(con)
            export_to_drive(con, args.export_dir)
            upload_to_gdrive(con, rows)
            return 3

    rows = build_summary(con)
    log.info("Сводка: %d дней -> %s", len(rows), CSV_PATH)
    export_to_drive(con, args.export_dir)
    upload_to_gdrive(con, rows)
    return 0


if __name__ == "__main__":
    sys.exit(main())
