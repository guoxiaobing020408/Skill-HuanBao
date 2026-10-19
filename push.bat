@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1
cd /d "%~dp0"

set "SRC=%USERPROFILE%\.workbuddy\skills\guo6-param-sheet-check"
set "DST=%~dp0guo6-param-sheet-check"
set "REPO=guoxiaobing020408/Skill-HuanBao"
set "PY=C:\Users\Administrator\.workbuddy\binaries\python\versions\3.13.12\python.exe"
if not exist "%PY%" set "PY=python"

echo.
echo ==========================================================
echo   guo6-param-sheet-check  --^>  GitHub (%REPO%)
echo ==========================================================
echo.

echo [1/3] Sync skill files from WorkBuddy ...
if exist "%SRC%\SKILL.md" (
    robocopy "%SRC%" "%DST%" /E /NFL /NDL /NJH /NJS /NP /XF *.pyc >nul
    if errorlevel 8 (
        echo   ERROR: robocopy failed. Abort.
        goto :end
    )
    echo   OK: %SRC%
) else (
    echo   SKIP: %SRC% not found, upload local folder as-is.
)
echo.

for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value 2^>nul ^| find "="') do set "DT=%%I"
if "%DT%"=="" set "DT=update"
set "MSG=update skill %DT:~0,8%-%DT:~8,4%"

echo [2/3] Check token ...
if not exist "%USERPROFILE%\.workbuddy\github_token.txt" (
    if "%GITHUB_TOKEN%"=="" (
        echo.
        echo   No token found. Put your GitHub Personal Access Token in:
        echo     %USERPROFILE%\.workbuddy\github_token.txt
        echo   Create one at: https://github.com/settings/tokens
        echo   Scope needed: repo  ^(classic token^)
        echo.
        goto :end
    )
)
echo   OK.
echo.

echo [3/3] Upload via API ...
"%PY%" "%~dp0push_api.py" --repo %REPO% --dir "%~dp0." --branch main --message "%MSG%"
if errorlevel 1 (
    echo.
    echo   UPLOAD FAILED. See message above.
) else (
    echo.
    echo   DONE: https://github.com/%REPO%
)

:end
echo.
pause
