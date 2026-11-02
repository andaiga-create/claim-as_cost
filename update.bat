@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo [1/3] 데이터 변환 (data 폴더 xls -^> docs\data.json)...
python -W ignore scripts\build_data.py
if errorlevel 1 (echo 변환 실패. 중단합니다. & pause & exit /b 1)
echo [2/3] git 커밋...
git add .gitignore docs scripts update.bat
git commit -m "데이터 갱신 %date%"
echo [3/3] GitHub 업로드...
git push -u origin main
echo 완료. 1~2분 뒤 사이트에 반영됩니다.
pause
