@echo off
cd "%USERPROFILE%\OneDrive\Documents\Code\Python tkinter\To-Do Games List"

set "VERSION=1.1.0"

echo === Building executable with PyInstaller ===
pyinstaller ^
  --noconsole ^
  --onefile ^
  --icon=icon.ico ^
  --hidden-import=help ^
  --hidden-import=settings ^
  "To-Do Games List.py"

echo === Creating ZIP folder ===
if not exist "zip" mkdir "zip"

echo === Creating ZIP archive ===
powershell -NoProfile -Command "Compress-Archive -Path '.\dist\*' -DestinationPath '.\zip\To-Do_Games_List_%VERSION%.zip' -Force"

echo === Done! ===