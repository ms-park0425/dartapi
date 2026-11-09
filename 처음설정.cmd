@echo off
chcp 65001 >nul
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" goto dependencies
where py >nul 2>nul
if errorlevel 1 goto python_command
py -3 -m venv .venv
if errorlevel 1 goto failed
goto dependencies
:python_command
python -m venv .venv
if errorlevel 1 goto failed
:dependencies
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
if not exist ".env" copy /y ".env.example" ".env" >nul
echo.
echo Python 패키지 준비가 끝났습니다.
echo README.md의 처음 설정 안내에 따라 .env에 DART 키를 저장하세요.
echo AI는 본인 계정으로 로그인하고 이 폴더를 로컬 프로젝트로 등록하세요.
echo 분기 선택 화면을 이용하려면 Codex CLI 설치와 로그인도 필요합니다.
pause
exit /b 0
:failed
echo.
echo 설정에 실패했습니다. Python 3.10 이상 설치와 인터넷 연결을 확인하세요.
echo 자세한 오류는 이 창의 위쪽에 있습니다. IT 담당자에게 이 내용을 전달하세요.
pause
exit /b 1
