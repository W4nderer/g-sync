# G-Sync — ежедневная выгрузка данных Garmin Connect (Windows)

> `index.html` и `privacy.html` — домашняя страница и политика конфиденциальности для экрана согласия Google OAuth (GitHub Pages: https://w4nderer.github.io/g-sync/).
>
> **Никогда не коммитьте** `credentials.json`, токены и папку `data\` — они уже в `.gitignore`.

Каждый день забирает из Garmin Connect все основные данные здоровья и тренировок и складывает их локально:

- `data\garmin.db` — SQLite
  - `raw_daily` — сырой JSON по каждому дню и метрике (stats, sleep, spo2, hrv, respiration, stress, heart_rate, body_battery, training_readiness, training_status, max_metrics, activities). Это «первоисточник»: из него всегда можно пересчитать что угодно.
  - `daily_summary` — сводка: одна строка на день (SpO2 средний/минимум, сон по фазам, sleep score, HRV, пульс покоя, дыхание, стресс, Body Battery, шаги, readiness, VO2max, тренировки).
- `data\daily_summary.csv` — та же сводка для Excel (разделитель `;`, десятичная запятая).
- `data\sync.log` — лог.

## Google Диск — вариант А: через API, без программ (рекомендуется)

Скрипт сам загружает данные в папку **GarminSync** на вашем Google Диске:
- `garmin.db` — снимок базы;
- `daily_summary.csv`;
- Google-таблица **«Garmin — сводка по дням»**. Она перезаписывается при каждом запуске, числа в ней — настоящие числа, так что графики и формулы можно строить прямо в Google Таблицах.

Доступ ограничен областью `drive.file`: скрипт видит только файлы, которые создал сам, а остальное содержимое Диска ему недоступно.

**Однократная настройка (~10 минут):**
1. Откройте https://console.cloud.google.com/ → создайте проект (например, `garmin-sync`).
2. **APIs & Services → Library**: включите **Google Drive API** и **Google Sheets API**.
3. **APIs & Services → OAuth consent screen** (или **Google Auth Platform**): тип **External**, укажите название и свой email. В **Audience / Test users** добавьте свой Gmail.
4. Там же нажмите **Publish app → In production**. Это важно: в режиме «Testing» Google отзывает токен каждые 7 дней, и загрузка перестанет работать. Проверка (verification) для области `drive.file` не нужна.
5. **Credentials → Create credentials → OAuth client ID → Desktop app**. Скачайте JSON, переименуйте в `credentials.json` и положите рядом с `garmin_sync.py`.
6. Подключите Диск (откроется браузер, войдите и разрешите доступ):
   ```powershell
   .\.venv\Scripts\python.exe garmin_sync.py --drive-login
   ```
   Если появится «Google не проверил это приложение», нажмите **Дополнительно → Перейти** (это ваше собственное приложение).

Дальше загрузка идёт автоматически при каждом запуске. Токен хранится в `%USERPROFILE%\.garmin-sync\gdrive_token.json`. Чтобы отключить загрузку, удалите этот файл.

## Google Диск — вариант Б: через Google Drive для компьютеров

Если программа Google Drive для компьютеров уже установлена, ничего настраивать не надо. После каждого запуска свежие копии `garmin.db` и `daily_summary.csv` кладутся в **Мой диск → GarminSync**, и Google Drive для компьютеров сам заливает их в облако.

- Нужно установить Google Drive для компьютеров (https://www.google.com/drive/download/) и войти в аккаунт. Скрипт сам находит диск `G:\My Drive` (или «Мой диск» / «Minu ketas» — в зависимости от языка), а также папку `%USERPROFILE%\Google Drive`.
- Рабочая база остаётся локально в `data\`, а на Диск уходит целостный снимок. Так файл не повредится, если Drive начнёт заливку посреди записи.
- Другая папка: `garmin_sync.py --export-dir "G:\My Drive\Health\Garmin"` или переменная окружения `GARMIN_EXPORT_DIR`. Отключить копирование: `--export-dir none`.
- Если Drive не запущен, локальные данные всё равно обновятся, а в логе появится предупреждение. Следующий запуск скопирует всё заново.

## Установка

Нужен Python 3.10+ (https://www.python.org/downloads/, при установке отметьте «Add python.exe to PATH»).

1. Скачайте репозиторий: `git clone https://github.com/W4nderer/g-sync C:\GarminSync` (или **Code → Download ZIP** на GitHub и распакуйте в `C:\GarminSync`) и откройте эту папку в PowerShell. Обновить потом: `git pull`.
2. Окружение и зависимости:
   ```powershell
   powershell -ExecutionPolicy Bypass -File setup.ps1
   ```
3. Первый вход (email, пароль, код 2FA, если включён). Пароль нигде не сохраняется — только токены в `%USERPROFILE%\.garminconnect`:
   ```powershell
   .\.venv\Scripts\python.exe garmin_sync.py --login
   ```
4. Загрузка истории (долго: ~12 запросов на день с паузой 1 с, год ≈ 1,5 часа; можно прервать и продолжить — уже загруженное пропускается):
   ```powershell
   .\.venv\Scripts\python.exe garmin_sync.py --start 2024-01-01
   ```
5. Ежедневная задача (в 10:00 и 21:00; если компьютер был выключен — запустится при включении и догрузит пропущенные дни):
   ```powershell
   powershell -ExecutionPolicy Bypass -File install_task.ps1
   ```

## Что попадает в Google-таблицу

- **Лист 1 — сводка по дням**: сон, SpO2, HRV, пульс покоя, стресс, Body Battery, шаги, VO2max, тренировки за день, прогноз Garmin на забег, этажи (`floors_up`, `floors_down`), статус тренированности (`training_status`), острая и хроническая нагрузка (`acute_load`, `chronic_load`), их соотношение (`acwr`, `acwr_status`) и баланс нагрузки за 4 недели (`load_aerobic_low`, `load_aerobic_high`, `load_anaerobic`, `load_balance`).
- **«Тренировки»**: одна строка на тренировку — дистанция, время, темп, пульс, каденс, мощность, набор высоты, Training Effect, нагрузка.
- **«Рекорды»**: личные рекорды Garmin (1 км, 5 км, 10 км, полумарафон, самая длинная пробежка и т. п.).
- **«Профиль»**: возраст, пол, рост и вес из профиля Garmin — для сравнения с возрастными нормами. Дата рождения не сохраняется, только возраст.
- **«Порог лактата»**: пульс, темп и мощность на пороге по датам. Garmin определяет порог только на пробежках с нагрудным пульсометром, поэтому без него лист может быть пустым.

Те же таблицы лежат в `data\garmin.db` и CSV-файлах в папке `data\`.

## Полезное

- Прогноз Garmin на забег (Race Predictor) хранится в столбцах `race_5k_s`, `race_10k_s`, `race_hm_s` и `race_marathon_s` (в секундах). Новые дни он собирает сам. Чтобы догрузить прогнозы за прошлые дни, один раз запустите загрузку истории: уже загруженные метрики будут пропущены, запросы пойдут только за прогнозами.
  ```powershell
  .\.venv\Scripts\python.exe garmin_sync.py --start 2025-10-21
  ```

- Последние 3 дня перезапрашиваются при каждом запуске — Garmin дозаливает данные после синхронизации часов.
- `--force` — перезапросить уже загруженные дни; `--summary-only` — только пересобрать сводку/CSV.
- Если в логе «Не удалось войти» — Garmin сбросил токены: повторите шаг 3.
- Если сломалось после обновления у Garmin: `.\.venv\Scripts\python.exe -m pip install --upgrade garminconnect` — библиотека неофициальная и периодически чинится под изменения.
- Удалить задачу: `Unregister-ScheduledTask -TaskName "Garmin Sync" -Confirm:$false`.
