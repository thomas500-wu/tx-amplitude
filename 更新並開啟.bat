@echo off
chcp 65001 >nul
cd /d "%~dp0"
git pull --quiet || "C:\Users\thoma\AppData\Local\Programs\Python\Python312\python.exe" update.py
start "" "%~dp0index.html"
