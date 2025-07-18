@echo off
echo Ativando ambiente virtual...
call .venv\Scripts\activate.bat

REM Inicia a contagem do tempo de execução
set "start_time=%time%"

echo Atualizando planilha...
python atualizar_planilha.py

echo Executando previsao de atrasos...
python irf.py

REM Calcula o tempo de execução em minutos e segundos
set "end_time=%time%"

REM Extrai horas, minutos, segundos e centésimos do início e fim
for /f "tokens=1-4 delims=:.," %%a in ("%start_time%") do (
    set /a "start_total=(%%a*3600)+(%%b*60)+%%c"
    set /a "start_centesimos=%%d"
)
for /f "tokens=1-4 delims=:.," %%a in ("%end_time%") do (
    set /a "end_total=(%%a*3600)+(%%b*60)+%%c"
    set /a "end_centesimos=%%d"
)

REM Calcula a diferença total em segundos e centésimos
set /a "elapsed_total=end_total-start_total"
set /a "elapsed_centesimos=end_centesimos-start_centesimos"

REM Ajusta se os centésimos ficarem negativos
if %elapsed_centesimos% lss 0 (
    set /a "elapsed_centesimos+=100"
    set /a "elapsed_total-=1"
)

REM Converte para minutos e segundos
set /a "elapsed_min=elapsed_total/60"
set /a "elapsed_sec=elapsed_total%%60"

REM Exibe o tempo de execução em minutos e segundos
echo Tempo de execucao: %elapsed_min% minuto(s) e %elapsed_sec% segundo(s).


echo.
echo Pressione qualquer tecla para sair...
pause >nul 