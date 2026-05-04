@echo off
cd /d "%~dp0"
echo Iniciando Relatorios Tecnicos Agres...
echo.
echo App local: http://localhost:8501
echo Coleta offline: http://localhost:8501/app/static/offline/index.html
echo.
echo Mantenha esta janela aberta enquanto estiver usando o app local.
echo Para encerrar, feche esta janela.
echo.
"C:\Users\henrique.chaves\AppData\Local\Programs\Python\Python312\python.exe" -m streamlit run app.py --server.port 8501 --server.address 0.0.0.0 --server.headless true
pause
