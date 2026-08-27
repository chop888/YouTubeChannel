@echo off
REM Entry point for Windows Task Scheduler.
REM Output is logged to research\generate-themes.log by the script itself.
REM python.exe is called via full path since PATH may not be set under Task Scheduler.
"C:\Users\mikic\AppData\Local\Programs\Python\Python313\python.exe" "%~dp0generate-themes.py"
