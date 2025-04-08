@echo off
setlocal enabledelayedexpansion

echo ===== Organizador de Planilhas BrasPub - Modo Desenvolvimento =====
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

REM Instalar dependências do backend
echo.
echo ===== Instalando dependencias do backend =====
cd src\backend
echo Instalando pacotes Python no ambiente virtual...
call ..\..\.venv\Scripts\activate
pip install -r requirements.txt
if %errorlevel% neq 0 goto :error
cd ..\..

REM Instalar dependências do frontend
echo.
echo ===== Instalando dependencias do frontend =====
cd src\frontend
echo Instalando pacotes Node.js...
call npm install
if %errorlevel% neq 0 goto :error

REM Iniciar o backend em segundo plano
echo.
echo ===== Iniciando o backend (Python Flask) =====
start cmd /k "cd ..\..\src\backend && ..\..\.venv\Scripts\activate && python api.py"

REM Aguardar o backend iniciar
echo Aguardando o backend iniciar...
timeout /t 5 /nobreak > nul

REM Iniciar o frontend
echo.
echo ===== Iniciando o frontend (React + Electron) =====
echo Para parar a aplicacao, pressione Ctrl+C nesta janela
call npm run electron-dev
if %errorlevel% neq 0 goto :error

cd ..\..
goto :eof

:error
echo.
echo ===== ERRO: A execucao falhou =====
exit /b 1 