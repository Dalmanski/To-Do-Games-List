@echo off
cd "%USERPROFILE%\OneDrive\Documents\Code\Python tkinter\To-Do Games List"

echo === Building executable with PyInstaller ===
pyinstaller ^
  --noconsole ^
  --onefile ^
  --icon=icon.ico ^
  --add-data "settings.json;." ^
  --hidden-import=help ^
  --hidden-import=settings ^
  "To-Do Games List.py"

echo === Creating ZIP archive ===
powershell -Command "Compress-Archive -Path 'dist\*' -DestinationPath 'To-Do_Games_List.zip' -Force"

echo === Done! ===