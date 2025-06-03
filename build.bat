@echo off
chcp 65001 > nul
echo Встановлення залежностей...

REM Перевіряємо наявність pip
pip --version > nul 2>&1
if errorlevel 1 (
    echo pip не знайдено. Переконайтеся, що Python встановлено та додано до PATH.
    pause
    exit /b
)

REM Перевіряємо файл requirements.txt і встановлюємо залежності, якщо потрібно
if exist requirements.txt (
    echo Перевірка та встановлення залежностей з requirements.txt...
    pip install --requirement requirements.txt --exists-action i
) else (
    echo Файл requirements.txt не знайдено. Пропускаємо установку залежностей.
)

REM Перевірити чи встановлено pyinstaller
pyinstaller --version > nul 2>&1
if errorlevel 1 (
    echo Встановлення pyinstaller...
    pip install pyinstaller
) else (
    echo PyInstaller вже встановлений.
)

echo Компіляція програми...
pyinstaller --onefile --windowed --name "MeterGenerator" --icon=icon.ico meter_generator.py

echo Готово! EXE файл знаходиться в папці dist/
pause
