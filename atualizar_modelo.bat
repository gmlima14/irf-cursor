@echo off
echo Ativando ambiente virtual...
call .venv\Scripts\activate.bat

echo Executando treinamento do modelo...
python modelo_irf.py

echo.
echo Treinamento finalizado. Pressione qualquer tecla para sair...
pause >nul
