@echo off
setlocal enabledelayedexpansion

echo ===== Organizador de Planilhas BrasPub - Compilacao =====
echo.

REM Verificar se Python está instalado
echo Verificando instalacao do Python...
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo Python nao encontrado. Por favor, instale o Python 3.7 ou superior.
    echo Voce pode baixa-lo em: https://www.python.org/downloads/
    goto :error
)

REM Verificar se Node.js está instalado
echo Verificando instalacao do Node.js...
node --version > nul 2>&1
if %errorlevel% neq 0 (
    echo Node.js nao encontrado. Por favor, instale o Node.js 14 ou superior.
    echo Voce pode baixa-lo em: https://nodejs.org/
    goto :error
)

REM Verificar se o ambiente virtual existe, se não, criá-lo
if not exist ".venv" (
    echo Criando ambiente virtual Python...
    python -m venv .venv
    if %errorlevel% neq 0 goto :error
)

REM Ativar o ambiente virtual
call .venv\Scripts\activate

REM Criar pasta dist se não existir
if not exist "dist" mkdir dist

REM Compilar o backend
echo.
echo ===== Compilando o backend (Python) =====
cd src\backend
echo Instalando dependencias do Python no ambiente virtual...
pip install -r requirements.txt
if %errorlevel% neq 0 goto :error

echo Instalando PyInstaller no ambiente virtual...
pip install pyinstaller
if %errorlevel% neq 0 goto :error

echo Compilando o backend para um executavel...
python compilar_backend.py
if %errorlevel% neq 0 goto :error

cd ..\..
echo Backend compilado com sucesso!

REM Compilar o frontend
echo.
echo ===== Compilando o frontend (React + Electron) =====
cd src\frontend

echo Instalando dependencias do Node.js...
call npm install
if %errorlevel% neq 0 goto :error

echo Compilando o aplicativo...
call npm run electron:build
if %errorlevel% neq 0 goto :error

echo Frontend compilado com sucesso!
cd ..\..

REM Copiar o instalador para a pasta raiz
echo.
echo ===== Finalizando a compilacao =====
if exist "src\frontend\dist\*.exe" (
    copy "src\frontend\dist\*.exe" "dist\"
    echo Instalador copiado para a pasta dist\
) else (
    echo AVISO: Instalador nao encontrado.
)

REM Desativar o ambiente virtual
call venv\Scripts\deactivate

echo.
echo ===== Compilacao concluida com sucesso! =====
echo Os arquivos compilados estao na pasta dist\
echo.
goto :eof

:error
echo.
echo ===== ERRO: A compilacao falhou =====
exit /b 1 