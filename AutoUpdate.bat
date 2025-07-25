@echo off
cd "%USERPROFILE%\OneDrive\Documents\Code\Python tkinter\To-Do Games List"
pyinstaller ^
  --noconsole ^
  --onefile ^
  --icon=icon.ico ^
  --add-data "settings.json;." ^
  --hidden-import=help ^
  --hidden-import=settings ^
  "To-Do Games List.py"
