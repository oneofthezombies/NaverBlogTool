@echo off

REM set path variable
set PROGRAM_PATH=%~dp0program\
set PYTHON_ROOT=%PROGRAM_PATH%python\

REM remove origin python path
echo ���� ���̽� ��� ������ �����մϴ�. (���� ������� ȯ�溯��)
setx PYTHONPATH ""
REG delete HKCU\Environment /F /V PYTHONPATH
echo ���� ���̽� ��� ������ �Ϸ��߽��ϴ�.

REM set new python path
echo ���̽� ��� �߰��� �����մϴ�. (���� ������� ȯ�溯��)
setx PYTHONPATH "%PYTHON_ROOT%;%PYTHON_ROOT%Lib;%PYTHON_ROOT%Lib\site-packages"
echo ���̽� ��� �߰��� �Ϸ��߽��ϴ�.

pause