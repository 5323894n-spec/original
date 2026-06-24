@echo off
chcp 65001 >nul
cd /d "%~dp0"
py -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Ошибка установки зависимостей.
  pause
  exit /b 1
)
echo.
echo Зависимости успешно установлены.
pause
