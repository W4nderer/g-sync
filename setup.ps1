# Создаёт виртуальное окружение и ставит зависимости.
# Запуск: powershell -ExecutionPolicy Bypass -File setup.ps1
$ErrorActionPreference = "Stop"
$Dir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Dir

py -3 -m venv .venv
& "$Dir\.venv\Scripts\python.exe" -m pip install --upgrade pip
& "$Dir\.venv\Scripts\python.exe" -m pip install --upgrade -r requirements.txt

Write-Host ""
Write-Host "Готово. Теперь первый вход:" -ForegroundColor Green
Write-Host "  .\.venv\Scripts\python.exe garmin_sync.py --login"
