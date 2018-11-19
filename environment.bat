@echo off


echo 현재 cmd 창에서'만' 사용할 임시 환경변수를 추가합니다.
echo Add temporarily environment variables to use in the current cmd window.


REM __MY_PYTHON_PATH__를 python.exe 파일의 경로로 수정해주세요.
REM Edit __MY_PYTHON_PATH__ to the path to the python.exe file.
REM e.g.) set __MY_PYTHON_PATH__=C:\my\python\directory


set __MY_PYTHON_PATH__=C:\Users\hunho\source\repos\oneofthezombies\NaverBlogTool\program\python





REM 수정 금지
REM No modification


set PATH=%PATH%;%__MY_PYTHON_PATH__%;%__MY_PYTHON_PATH__%\Scripts;
set PYTHONPATH=%PYTHONPATH%;%__MY_PYTHON_PATH__%;%__MY_PYTHON_PATH__%\Lib;%__MY_PYTHON_PATH__%\Lib\site-packages;


echo 임시 환경변수 추가를 완료했습니다.
echo You have successfully added a temporary environment variable.


REM 만약 다른 batch 스크립트에서 호출한다면 그 스크립트 내에 call /path/to/environemnt.bat 를 실행해주세요.
REM If called by another batch script, run call /path/to/environemnt.bat" in that script.