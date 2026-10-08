@echo off
chcp 65001 >nul
REM emedi 수집기 - Windows 앱 만들기 (이 파일을 더블클릭)
cd /d "%~dp0"
python -m venv .venv || goto :err
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt pyinstaller || goto :err
pyinstaller --noconfirm --clean --windowed --name "emedi_crawler" ^
  --collect-data selenium ^
  --exclude-module tkinter --exclude-module matplotlib ^
  app.py || goto :err
echo.
echo 완료: dist\emedi_crawler\emedi_crawler.exe
echo dist\emedi_crawler 폴더 전체를 원하는 곳에 두고, exe의 바로가기를 바탕화면에 만드세요.
pause
exit /b 0
:err
echo 빌드 실패. 위 오류 메시지를 확인하세요.
pause
exit /b 1
