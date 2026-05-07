@echo off
setlocal
cd /d "%~dp0"

echo ===========================================
echo COMPILACION DEL .EXE (PyInstaller)
echo ===========================================

if not exist ".venv\Scripts\python.exe" (
    echo No se encontro .venv. Ejecuta primero INSTALAR_PROYECTO.bat
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"

python -m pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

if exist "build" rmdir /S /Q "build"
if exist "dist" rmdir /S /Q "dist"

pyinstaller --noconfirm --clean RenombrePanacea.spec
if %errorlevel% neq 0 (
    echo Error al compilar el .exe
    pause
    exit /b 1
)

echo.
echo .EXE generado en: dist\RenombrePanacea.exe
pause
exit /b 0
