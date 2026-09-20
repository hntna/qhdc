@echo off
title Khoi dong Web Masterlist QHDC 2027-2028
cd /d "%~dp0"
start "" python server.py
timeout /t 1 >nul
start http://localhost:8080/index.html
