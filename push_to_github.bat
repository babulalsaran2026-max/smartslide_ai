@echo off
cd /d "%~dp0"
echo ========================================================
echo Pushing SmartSlide AI to GitHub...
echo If Git Credential Manager opens, please sign in.
echo ========================================================
"C:\Program Files\Git\cmd\git.exe" push -u origin main
echo.
pause
