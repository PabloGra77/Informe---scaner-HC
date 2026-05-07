@echo off
setlocal
cd /d "%~dp0"

echo ===========================================
echo INSTALACION DEL PROYECTO
echo ===========================================

set "PYTHON_CMD="
where py >nul 2>nul
if %errorlevel%==0 (
    set "PYTHON_CMD=py"
) else (
    where python >nul 2>nul
    if %errorlevel%==0 (
        set "PYTHON_CMD=python"
    )
)

if "%PYTHON_CMD%"=="" (
    echo No se encontro Python en el equipo.
    echo Instala Python 3.10 o superior y vuelve a intentar.
    pause
    exit /b 1
)

echo Creando entorno virtual .venv ...
%PYTHON_CMD% -m venv .venv
if %errorlevel% neq 0 (
    echo Error al crear el entorno virtual.
    pause
    exit /b 1
)

call ".venv\Scripts\activate.bat"
if %errorlevel% neq 0 (
    echo Error al activar el entorno virtual.
    pause
    exit /b 1
)

echo Actualizando pip ...
python -m pip install --upgrade pip
if %errorlevel% neq 0 (
    echo Error al actualizar pip.
    pause
    exit /b 1
)

echo Instalando dependencias ...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Error al instalar dependencias.
    pause
    exit /b 1
)

echo.
echo Instalacion completada correctamente.
echo Usa INICIAR_INTERFAZ.bat para ejecutar el programa.
pause
exit /b 0
