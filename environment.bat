@echo off


echo ���� cmd â����'��' ����� �ӽ� ȯ�溯���� �߰��մϴ�.
echo Add temporarily environment variables to use in the current cmd window.


REM __MY_PYTHON_PATH__�� python.exe ������ ��η� �������ּ���.
REM Edit __MY_PYTHON_PATH__ to the path to the python.exe file.
REM e.g.) set __MY_PYTHON_PATH__=C:\my\python\directory


set __MY_PYTHON_PATH__=C:\Users\hunho\source\repos\oneofthezombies\NaverBlogTool\program\python





REM ���� ����
REM No modification


set PATH=%PATH%;%__MY_PYTHON_PATH__%;%__MY_PYTHON_PATH__%\Scripts;
set PYTHONPATH=%PYTHONPATH%;%__MY_PYTHON_PATH__%;%__MY_PYTHON_PATH__%\Lib;%__MY_PYTHON_PATH__%\Lib\site-packages;


echo �ӽ� ȯ�溯�� �߰��� �Ϸ��߽��ϴ�.
echo You have successfully added a temporary environment variable.


REM ���� �ٸ� batch ��ũ��Ʈ���� ȣ���Ѵٸ� �� ��ũ��Ʈ ���� call /path/to/environemnt.bat �� �������ּ���.
REM If called by another batch script, run call /path/to/environemnt.bat" in that script.