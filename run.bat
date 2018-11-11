@echo off

set PROGRAM_PATH=%~dp0program\
set PYTHON_EXE=%PROGRAM_PATH%python\python.exe
set MY_SCRIPT_PATH=%PROGRAM_PATH%my_script\

REM run script
%PYTHON_EXE% %MY_SCRIPT_PATH%run.py

pause