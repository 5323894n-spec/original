@echo off
chcp 65001 >nul
cd /d "%~dp0"
py app.py
if errorlevel 1 (
  echo.
  echo Не удалось запустить программу.
  echo Установите зависимости через install_dependencies.bat
  pause
)
