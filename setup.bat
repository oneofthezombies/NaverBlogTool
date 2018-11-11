@echo off

REM get os architecture
reg Query "HKLM\Hardware\Description\System\CentralProcessor\0" | find /i "x86" > NUL && set OS=32BIT || set OS=64BIT

REM print os architecture
if %OS%==32BIT echo 32bit �ü���� ������Դϴ�.
if %OS%==64BIT echo 64bit �ü���� ������Դϴ�.

REM set PYTHON_ZIP_FILE name
if %OS%==32BIT (
  set PYTHON_ZIP_FILE=python-3.6.7-embed-win32.zip 
  echo 32bit ���̽��� �����߽��ϴ�.
)
if %OS%==64BIT (
  set PYTHON_ZIP_FILE=python-3.6.7-embed-amd64.zip 
  echo 64bit ���̽��� �����߽��ϴ�.
)

REM set path variable
set PROGRAM_PATH=%~dp0program\
set PYTHON_ROOT=%PROGRAM_PATH%python\

REM unzip python 
echo ���̽� ���������� �����մϴ�.
powershell.exe -nologo -noprofile -command "& { Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('%PROGRAM_PATH%%PYTHON_ZIP_FILE%', '%PYTHON_ROOT%'); }"
echo ���̽� ���������� �Ϸ��߽��ϴ�.

REM get pip
echo pip �ٿ�ε带 �����մϴ�.
%PYTHON_ROOT%python.exe %PROGRAM_PATH%get-pip.py
echo pip �ٿ�ε带 �Ϸ��߽��ϴ�.

REM save old python path
echo ���̽� ������� ����� �����մϴ�.
ren %PYTHON_ROOT%python36._pth python36._pth.save
echo ���̽� ������� ����� �Ϸ��߽��ϴ�.

REM set new python path
echo ���̽� ��� �߰��� �����մϴ�. (���� ������� ȯ�溯��)
setx PYTHONPATH "%PYTHON_ROOT%;%PYTHON_ROOT%Lib;%PYTHON_ROOT%Lib\site-packages"
echo ���̽� ��� �߰��� �Ϸ��߽��ϴ�.

REM get selenium
echo selenium ��ġ�� �����մϴ�.
%PYTHON_ROOT%Scripts\pip3.exe install selenium
echo selenium ��ġ�� �Ϸ��߽��ϴ�.

REM get python3_anticaptcha
echo python3_anticaptcha ��ġ�� �����մϴ�.
%PYTHON_ROOT%Scripts\pip3.exe install python3-anticaptcha
echo python3_anticaptcha ��ġ�� �Ϸ��߽��ϴ�.

REM get openpyxl
echo openpyxl ��ġ�� �����մϴ�.
%PYTHON_ROOT%Scripts\pip3.exe install openpyxl
echo openpyxl ��ġ�� �Ϸ��߽��ϴ�.

REM get PyQt5
echo PyQt5 ��ġ�� �����մϴ�.
%PYTHON_ROOT%Scripts\pip3.exe install PyQt5
echo PyQt5 ��ġ�� �Ϸ��߽��ϴ�.

pause