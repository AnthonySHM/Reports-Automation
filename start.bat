@echo off
set SHM_GOOGLE_CREDENTIALS=%~dp0Credentials\focal-set-486609-s9-f052f8c1a756.json
echo Starting SHM Report Generator...
echo Drive credentials: %SHM_GOOGLE_CREDENTIALS%
python run.py serve
