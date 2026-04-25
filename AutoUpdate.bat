@echo off
setlocal
cd /d "%USERPROFILE%\OneDrive\Documents\Code\Python tkinter\To-Do Games List"

set "VERSION=1.2.0"
set "DIST_DIR=%CD%\dist"
set "ZIP_DIR=%CD%\zip"
set "STAGE_DIR=%CD%\_zip_stage"

echo === Building executable with PyInstaller ===
pyinstaller ^
  --noconsole ^
  --onefile ^
  --icon=icon.ico ^
  --hidden-import=help ^
  --hidden-import=settings ^
  --hidden-import="To-Do List" ^
  --hidden-import="To-Do List Modal" ^
  --hidden-import=tkcalendar ^
  "To-Do Games List.py"

if errorlevel 1 (
  echo Build failed.
  exit /b 1
)

echo === Preparing staging folder ===
if exist "%STAGE_DIR%" rmdir /s /q "%STAGE_DIR%"
mkdir "%STAGE_DIR%"

echo === Copying dist files except To-Do Games List.json ===
robocopy "%DIST_DIR%" "%STAGE_DIR%" /E /XF "To-Do Games List.json" >nul
if %ERRORLEVEL% GEQ 8 (
  echo Copy failed.
  exit /b 1
)

echo === Creating ZIP folder ===
if not exist "%ZIP_DIR%" mkdir "%ZIP_DIR%"

echo === Creating ZIP archive ===
powershell -NoProfile -Command "Compress-Archive -Path '%STAGE_DIR%\*' -DestinationPath '%ZIP_DIR%\To-Do_Games_List_%VERSION%.zip' -Force"

if errorlevel 1 (
  echo ZIP creation failed.
  exit /b 1
)

echo === Cleaning up ===
rmdir /s /q "%STAGE_DIR%"

echo === Done! ===
endlocal