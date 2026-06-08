@echo off
setlocal enabledelayedexpansion

REM Script para instalar dependências automaticamente
REM Executa antes de rodar a aplicação

echo.
echo ====================================================
echo   Instalando Dependencias do NovaBusca...
echo ====================================================
echo.

REM Verifica se o Python está acessivel
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRO] Python nao encontrado no PATH
    echo Certifique-se que Python 3.10+ esta instalado
    pause
    exit /b 1
)

echo [INFO] Python encontrado. Instalando pacotes...
echo.

REM Instala os pacotes do requirements.txt
python -m pip install --upgrade pip >nul 2>&1
python -m pip install -r requirements.txt

if %errorlevel% equ 0 (
    echo.
    echo [OK] Todas as dependencias foram instaladas com sucesso!
    echo.
    echo Agora voce pode executar: Abrir_NovaBusca_UI.bat
) else (
    echo.
    echo [ERRO] Falha ao instalar algumas dependencias
    echo Tente executar novamente ou instale manualmente:
    echo   pip install -r requirements.txt
)

echo.
pause
