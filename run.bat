@echo off

set PROGRAM_PATH=%~dp0Program\
set PYTHON_EXE=%PROGRAM_PATH%Python\python.exe
set MY_SCRIPT_PATH=%PROGRAM_PATH%MyScript\

REM run script
%PYTHON_EXE% %MY_SCRIPT_PATH%run.py

pause