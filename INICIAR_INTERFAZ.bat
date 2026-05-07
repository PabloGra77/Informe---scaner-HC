@echo off
setlocal
cd /d "%~dp0"

echo ===========================================
echo INICIO DEL PROCESADOR DE PDF (INTERFAZ)
echo ===========================================

if exist ".venv\Scripts\python.exe" (
    set "PYTHON_CMD=.venv\Scripts\python.exe"
) else (
    where py >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=py"
    ) else (
        where python >nul 2>nul
        if %errorlevel%==0 (
            set "PYTHON_CMD=python"
        ) else (
            echo No se encontro Python.
            echo Ejecuta primero INSTALAR_PROYECTO.bat
            pause
            exit /b 1
        )
    )
)

"%PYTHON_CMD%" "app.py"
if %errorlevel% neq 0 (
    echo.
    echo El programa termino con error.
    pause
    exit /b 1
)

exit /b 0
