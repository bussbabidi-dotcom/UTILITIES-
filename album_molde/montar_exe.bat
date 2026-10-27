@echo off
cd /d "%~dp0"
py -3 -m pip install pyinstaller
py -3 -m PyInstaller --noconfirm --clean --onedir --windowed --name Album --add-data "modelos;modelos" --add-data "molde;molde" album_ui.pyw
echo.
echo Pasta pronta para copiar: dist\Album
pause
