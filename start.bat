@echo off
cd /d "%~dp0"
echo.
echo   Reporte de seguridad  ·  fecha y turno
echo   http://127.0.0.1:8765
echo.
py -3 -m pip install python-pptx Pillow --quiet
py -3 server.py
if errorlevel 1 (
  echo.
  echo  No pude iniciar con py. Probando python...
  python -m pip install python-pptx Pillow --quiet
  python server.py
)
pause
