@echo off
echo Встановлення залежностей...
pip install -r requirements.txt
pip install pyinstaller

echo Компіляція програми...
pyinstaller --onefile --windowed --name "MeterGenerator" --icon=icon.ico meter_generator.py

echo Готово! EXE файл знаходиться в папці dist/
pause