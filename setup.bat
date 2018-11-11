@echo off

REM get os architecture
reg Query "HKLM\Hardware\Description\System\CentralProcessor\0" | find /i "x86" > NUL && set OS=32BIT || set OS=64BIT

REM print os architecture
if %OS%==32BIT echo 32bit 운영체제를 사용중입니다.
if %OS%==64BIT echo 64bit 운영체제를 사용중입니다.

REM set PYTHON_ZIP_FILE name
if %OS%==32BIT (
  set PYTHON_ZIP_FILE=python-3.6.7-embed-win32.zip 
  echo 32bit 파이썬을 선택했습니다.
)
if %OS%==64BIT (
  set PYTHON_ZIP_FILE=python-3.6.7-embed-amd64.zip 
  echo 64bit 파이썬을 선택했습니다.
)

REM set path variable
set PROGRAM_PATH=%~dp0program\
set PYTHON_ROOT=%PROGRAM_PATH%python\

REM unzip python 
echo 파이썬 압축해제를 시작합니다.
powershell.exe -nologo -noprofile -command "& { Add-Type -A 'System.IO.Compression.FileSystem'; [IO.Compression.ZipFile]::ExtractToDirectory('%PROGRAM_PATH%%PYTHON_ZIP_FILE%', '%PYTHON_ROOT%'); }"
echo 파이썬 압축해제를 완료했습니다.

REM get pip
echo pip 다운로드를 시작합니다.
%PYTHON_ROOT%python.exe %PROGRAM_PATH%get-pip.py
echo pip 다운로드를 완료했습니다.

REM save old python path
echo 파이썬 경로파일 백업을 시작합니다.
ren %PYTHON_ROOT%python36._pth python36._pth.save
echo 파이썬 경로파일 백업을 완료했습니다.

REM set new python path
echo 파이썬 경로 추가를 시작합니다. (현재 사용자의 환경변수)
setx PYTHONPATH "%PYTHON_ROOT%;%PYTHON_ROOT%Lib;%PYTHON_ROOT%Lib\site-packages"
echo 파이썬 경로 추가를 완료했습니다.

REM get selenium
echo selenium 설치를 시작합니다.
%PYTHON_ROOT%Scripts\pip3.exe install selenium
echo selenium 설치를 완료했습니다.

REM get python3_anticaptcha
echo python3_anticaptcha 설치를 시작합니다.
%PYTHON_ROOT%Scripts\pip3.exe install python3-anticaptcha
echo python3_anticaptcha 설치를 완료했습니다.

REM get openpyxl
echo openpyxl 설치를 시작합니다.
%PYTHON_ROOT%Scripts\pip3.exe install openpyxl
echo openpyxl 설치를 완료했습니다.

REM get PyQt5
echo PyQt5 설치를 시작합니다.
%PYTHON_ROOT%Scripts\pip3.exe install PyQt5
echo PyQt5 설치를 완료했습니다.

pause