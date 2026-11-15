@echo off
setlocal
cd /d "%~dp0"
title Build Automatizador Canjes
py -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto error
py -m PyInstaller --noconfirm --clean --onefile --windowed --name Automatizador_Canjes --collect-all certifi --collect-all babel --hidden-import keyring.backends.Windows main.py
if errorlevel 1 goto error
echo.
echo Listo: dist\Automatizador_Canjes.exe
echo Copia correos.xlsx junto al EXE (no dentro del repositorio).
pause
exit /b 0
:error
echo ERROR: fallo la instalacion o compilacion.
pause
exit /b 1
