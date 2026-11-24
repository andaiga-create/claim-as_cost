@echo off
setlocal
cd /d "%~dp0"
echo ==========================================
echo  CLAIM/AS dashboard data update
echo ==========================================

rem --- 1. find a python that has pandas + xlrd (several pythons may be installed) ---
set "PY="
for /f "delims=" %%P in ('where python 2^>nul') do (
  if not defined PY ("%%P" -c "import pandas, xlrd, xlsxwriter" >nul 2>&1 && set "PY=%%P")
)
if not defined PY (
  py -3 -c "import pandas, xlrd, xlsxwriter" >nul 2>&1 && set "PY=py -3"
)
if not defined PY (
  echo [!] pandas/xlrd not found. Installing...
  for /f "delims=" %%P in ('where python 2^>nul') do (
    if not defined PY ("%%P" -m pip install pandas xlrd xlsxwriter && "%%P" -c "import pandas, xlrd, xlsxwriter" >nul 2>&1 && set "PY=%%P")
  )
)
if not defined PY (
  echo [ERROR] No usable Python. Install Python 3 then run: pip install pandas xlrd xlsxwriter
  goto :fail
)
echo [1/3] Converting xls to docs\data.json + excel report  (python: %PY%)
%PY% -W ignore scripts\build_data.py
if errorlevel 1 (
  echo [ERROR] Conversion failed. Check the xls files in the data folder.
  goto :fail
)
if /i "%~1"=="build" (
  echo Build only. Skipped git.
  goto :end
)

rem --- 2. commit ---
echo [2/3] git commit
git add .gitignore .gitattributes README.md docs scripts update.bat
git commit -m "data update %date%"
if errorlevel 1 echo (nothing new to commit - continuing)

rem --- 3. push ---
echo [3/3] git push
git push -u origin main
if errorlevel 1 (
  echo [ERROR] Push failed. Check internet / GitHub login.
  goto :fail
)
echo.
echo Done. The site updates in 1-2 minutes (Ctrl+F5 to refresh).
goto :end

:fail
echo.
echo *** FAILED - see messages above ***
:end
if /i not "%~1"=="build" pause
endlocal
