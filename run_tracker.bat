@echo off
cd "C:\Users\dadod\OneDrive\Desktop\clash-leaderboard"
python fetch_wars.py

REM Push updates to GitHub
git add leaderboard_*.json
git commit -m "Update leaderboard data - %date% %time%"
git push