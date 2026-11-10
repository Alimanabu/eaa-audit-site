@echo off
chcp 65001 >nul
echo Pered zapuskom skopiruj parol yashchika hello@eaa-audit.net v bufer obmena.
"%LOCALAPPDATA%\Programs\Python\Python314\python.exe" "%~dp0mail_watch.py" save
