@echo off

REM set path variable
set PROGRAM_PATH=%~dp0program\
set PYTHON_ROOT=%PROGRAM_PATH%python\

REM remove origin python path
echo 기존 파이썬 경로 삭제를 시작합니다. (현재 사용자의 환경변수)
setx PYTHONPATH ""
REG delete HKCU\Environment /F /V PYTHONPATH
echo 기존 파이썬 경로 삭제를 완료했습니다.

REM set new python path
echo 파이썬 경로 추가를 시작합니다. (현재 사용자의 환경변수)
setx PYTHONPATH "%PYTHON_ROOT%;%PYTHON_ROOT%Lib;%PYTHON_ROOT%Lib\site-packages"
echo 파이썬 경로 추가를 완료했습니다.

pause