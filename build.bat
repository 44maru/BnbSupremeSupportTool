@echo off

call venv\Scripts\activate
rem pyInstaller --onefile --noconsole jsonMaker.py
pyInstaller jsonmaker.spec
move dist\jsonMaker.exe "dist\��a�^BNB Supreme�⏕�c�[��.exe"

pause