@echo off
:: O comando abaixo força o prompt a abrir na exata pasta onde este arquivo .bat está
cd /d "%~dp0"

echo =======================================================
echo Iniciando o processamento do SGF...
echo =======================================================

:: Verifica se a pasta do ambiente virtual existe
if not exist ".venv\Scripts\activate.bat" (
    echo [ERRO] O ambiente virtual '.venv' nao foi encontrado nesta pasta!
    echo Onde o terminal tentou procurar: %cd%
    pause
    exit /b
)

:: Ativa o ambiente virtual
call ".venv\Scripts\activate.bat"

:: Executa o script Python (ajustado para o nome exato do seu arquivo)
python extract.py

echo.
echo =======================================================
echo Processamento concluido! O arquivo consolidado foi gerado.
echo =======================================================
pause