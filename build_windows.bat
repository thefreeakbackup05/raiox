@echo off
setlocal
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
pyinstaller --clean --noconfirm --onefile --windowed --name GERADOR_RAIO_X app.py
echo.
echo Build concluido. O executavel esta em dist\GERADOR_RAIO_X.exe
pause
