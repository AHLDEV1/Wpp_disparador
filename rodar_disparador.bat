@echo off
cd /d "C:\whatsapp_disparador"

echo.
echo ==============================================
echo Carregando Aguarde!!
echo ==============================================
echo.

call venv\Scripts\activate.bat

py whatsapp_disparador.py

pause
