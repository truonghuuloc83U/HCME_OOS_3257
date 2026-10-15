@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title UPDATE OOS DASHBOARD + GITHUB PAGES

echo =====================================================
echo    UPDATE OOS DASHBOARD - NPP HAI THIEN LONG 3257
 echo =====================================================
echo.

echo [1/5] Dang cap nhat Dashboard tu DATA_INPUT.xlsx...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0UPDATE_OOS.ps1"
if errorlevel 1 (
  echo.
  echo CAP NHAT THAT BAI. Hay dong file Excel va thu lai.
  pause
  exit /b 1
)

echo [2/5] Dang tao index.html cho GitHub Pages...
copy /Y "%~dp0OOS_Dashboard.html" "%~dp0index.html" >nul
if errorlevel 1 (
  echo ERROR: Khong tao duoc index.html.
  pause
  exit /b 1
)

where git >nul 2>&1
if errorlevel 1 (
  echo.
  echo DASHBOARD LOCAL DA CAP NHAT, NHUNG GIT CHUA DUOC CAI.
  start "" "%~dp0OOS_Dashboard.html"
  pause
  exit /b 1
)

if not exist "%~dp0.git" (
  echo.
  echo DASHBOARD LOCAL DA CAP NHAT, NHUNG CHUA SETUP GITHUB.
  echo Hay chay SETUP_GITHUB_ONCE.bat mot lan.
  start "" "%~dp0OOS_Dashboard.html"
  pause
  exit /b 1
)

echo [3/5] Dang ghi nhan phien ban moi...
git add index.html .nojekyll
git diff --cached --quiet
if not errorlevel 1 (
  echo Khong co thay doi moi de publish.
  goto OPEN_SITE
)

for /f "tokens=1-4 delims=/-. " %%a in ('date /t') do set TODAY=%%a-%%b-%%c-%%d
for /f "tokens=1-2 delims=: " %%a in ('time /t') do set NOW=%%a-%%b
git commit -m "Update OOS Dashboard %TODAY% %NOW%"
if errorlevel 1 (
  echo ERROR: Git commit that bai.
  pause
  exit /b 1
)

echo [4/5] Dang push len GitHub...
git push origin main
if errorlevel 1 (
  echo.
  echo DASHBOARD LOCAL DA CAP NHAT NHUNG GITHUB PUSH THAT BAI.
  echo Kiem tra Internet, dang nhap GitHub va quyen repository.
  start "" "%~dp0OOS_Dashboard.html"
  pause
  exit /b 1
)

:OPEN_SITE
echo [5/5] HOAN TAT.
echo =====================================================
echo DASHBOARD DA CAP NHAT VA PUBLISH LEN GITHUB
 echo =====================================================
start "" "%~dp0OOS_Dashboard.html"
pause
