@echo off
call activate .\env
start http://127.0.0.1:7860/
python .\app.py
pause
