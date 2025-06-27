@echo off
echo Ativando ambiente virtual...
call .venv\Scripts\activate.bat

echo Executando previsao de atrasos...
python irf.py

echo.
echo Pressione qualquer tecla para sair...
pause >nul 