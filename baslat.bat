@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
  echo Ilk kurulum yapiliyor...
  py -3 -m venv .venv || exit /b 1
  .venv\Scripts\python.exe -m pip install -q -r requirements.txt || exit /b 1
)
.venv\Scripts\python.exe app.py
