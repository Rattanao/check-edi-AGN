@echo off
chcp 65001 >nul
title Check EDI AGN
cd /d "%~dp0"
rem ปิดเว็บตัวเก่าที่ยังค้างอยู่ก่อน กันเปิดซ้อนหลายตัวบนพอร์ตเดียวกัน
powershell -NoProfile -Command "Get-CimInstance Win32_Process -Filter \"Name='python.exe'\" | Where-Object { $_.CommandLine -match 'webapp[\\/]app\.py' } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }" >nul 2>&1
python webapp\app.py
pause
