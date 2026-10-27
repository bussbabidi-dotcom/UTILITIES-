@echo off
cd /d "%~dp0"
where pyw >nul 2>&1 && start "" pyw -3 "%~dp0album_ui.pyw" && exit /b 0
start "" pythonw "%~dp0album_ui.pyw"
