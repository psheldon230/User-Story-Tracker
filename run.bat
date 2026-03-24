@echo off
echo Pulling latest updates...
git pull

echo Starting app...
start http://localhost:5000
python -m flask run
pause
