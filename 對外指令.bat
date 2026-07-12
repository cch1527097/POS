@echo off
powershell -NoProfile -Command "ngrok config add-authtoken 3GIMSvKqSQRM6aagQIvMCSk656H_4Uc8tPYDwMc3n8BGLQ9cB"
powershell -NoProfile -Command "ngrok http 3000"
pause