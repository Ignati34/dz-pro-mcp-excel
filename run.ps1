$ErrorActionPreference = 'Stop'
$project = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $project
$python = Join-Path $project '.venv\Scripts\python.exe'
if (!(Test-Path $python)) {
    throw 'Не найдено .venv. Сначала выполните команды установки из README.md.'
}
Write-Host 'ДЗ Pro запустится на http://127.0.0.1:8001'
& $python -m uvicorn app:app --host 127.0.0.1 --port 8001
