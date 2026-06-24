@echo off
chcp 65001 >nul
cd /d "%~dp0"
start "" http://127.0.0.1:5000
py web_app.py
if errorlevel 1 (
  echo.
  echo Не удалось запустить веб-приложение.
  echo Установите зависимости через install_dependencies.bat
  pause
)
