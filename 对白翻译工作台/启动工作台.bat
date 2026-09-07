@echo off
chcp 65001 >nul
cd /d E:\desk\对白翻译工作台
start "" http://127.0.0.1:8100
python server.py
pause
