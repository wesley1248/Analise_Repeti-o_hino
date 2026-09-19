@echo off
title Compilando Analisador para Executavel (.exe)
cls

cd /d "%~dp0"

echo ======================================================================
echo    COMPILADOR DO ANALISADOR DE OCORRENCIAS (PyInstaller)
echo ======================================================================
echo.

:: Verificar se PyInstaller esta instalado
python -m PyInstaller --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [1/3] Instalando PyInstaller...
    python -m pip install pyinstaller
) else (
    echo [1/3] PyInstaller ja instalado.
)

echo.
echo [2/3] Compilando executavel unico (sem console)...
echo Aguarde alguns instantes...
echo.

python -m PyInstaller --noconsole --onefile --name "Analisador_Ocorrencias" --clean analisador.py

if %errorlevel% neq 0 (
    echo.
    echo [ERRO] Falha ao compilar o executavel.
    pause
    exit /b %errorlevel%
)

echo.
echo ======================================================================
echo [3/3] SUCESSO! Executavel gerado na pasta:
echo       %~dp0dist\Analisador_Ocorrencias.exe
echo ======================================================================
echo.

:: Abrir a pasta dist automaticamente
explorer.exe "%~dp0dist"

pause
