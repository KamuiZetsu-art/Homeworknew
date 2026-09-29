@echo off
chcp 65001 >nul
echo ====================================
echo 📚 ЗАПУСК СИСТЕМЫ РАСПИСАНИЯ
echo ====================================
echo.

cd /d "%~dp0"

echo Проверка зависимостей...
python -c "import flask" 2>nul
if errorlevel 1 (
    echo Установка зависимостей...
    pip install -r requirements.txt
)

echo.
echo Запуск приложения...
python app.py
pause