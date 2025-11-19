@echo off
setlocal

REM ============================================================
REM Build script for the GridLamEdit Windows executable.
REM 1. (Opcional) Ative o ambiente virtual:
REM        call .venv\Scripts\activate
REM 2. Instale as dependencias:
REM        python -m pip install --upgrade pip
REM        pip install -r requirements.txt
REM 3. Execute este script sempre que precisar gerar um novo build.
REM ============================================================

echo Limpando e gerando executavel GridLamEdit...

pyinstaller ^
  --noconfirm ^
  --clean ^
  GridLamEdit.spec

REM Opcional: inclua --icon caminho\para\icone.ico no comando acima quando possuir um arquivo .ico.

if errorlevel 1 (
  echo.
  echo Falha ao gerar o executavel. Veja as mensagens acima.
  exit /b 1
)

echo.
echo Build concluido. O executavel unico esta em dist\GridLamEdit.exe
exit /b 0
