# Регистрирует ежедневную задачу в Планировщике заданий Windows.
# Запуск: powershell -ExecutionPolicy Bypass -File install_task.ps1
# Удалить задачу: Unregister-ScheduledTask -TaskName "Garmin Sync" -Confirm:$false
$ErrorActionPreference = "Stop"
$Dir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$Py     = Join-Path $Dir ".venv\Scripts\pythonw.exe"   # pythonw — без окна консоли
$Script = Join-Path $Dir "garmin_sync.py"
$Time   = "10:00"                                      # к этому времени ночные данные обычно уже синхронизированы

if (-not (Test-Path $Py)) { throw "Не найден $Py — сначала запустите setup.ps1" }

$Action   = New-ScheduledTaskAction -Execute $Py -Argument "`"$Script`"" -WorkingDirectory $Dir
$Triggers = @(
    (New-ScheduledTaskTrigger -Daily -At $Time),
    (New-ScheduledTaskTrigger -Daily -At "21:00")      # второй прогон: подхватит дневные тренировки
)
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
    -RunOnlyIfNetworkAvailable `
    -ExecutionTimeLimit (New-TimeSpan -Hours 3)

Register-ScheduledTask -TaskName "Garmin Sync" -Action $Action -Trigger $Triggers `
    -Settings $Settings -Description "Ежедневная выгрузка данных Garmin Connect в SQLite" -Force | Out-Null

Write-Host "Задача 'Garmin Sync' создана: каждый день в $Time и 21:00 (если ПК был выключен — при следующем включении)." -ForegroundColor Green
Write-Host "Проверить прямо сейчас: Start-ScheduledTask -TaskName 'Garmin Sync'; лог — data\sync.log"
