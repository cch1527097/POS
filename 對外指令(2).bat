@echo off
powershell -NoProfile -Command npm install -g localtunnel
powershell -NoProfile -Command lt --port 3000
pause