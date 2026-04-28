@echo off
cd /d "D:\Проект офис БООТ"

if not exist "backup" mkdir "backup"

set dt=%date:~-4,4%-%date:~-7,2%-%date:~-10,2%_%time:~0,2%-%time:~3,2%

copy "bot.py" "backup\bot_%dt%.py" >nul

echo Backup saved
echo Starting bot...

python bot.py

pause