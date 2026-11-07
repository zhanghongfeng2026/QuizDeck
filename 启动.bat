@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动题练...
echo 启动后请用浏览器打开 http://127.0.0.1:5000
python app.py
pause
