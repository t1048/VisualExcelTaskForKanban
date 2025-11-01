@echo off
REM Miniconda

SETLOCAL

CALL "%USERPROFILE%\miniconda3\Scripts\activate.bat" py310

python backend\backend.py --excel .\data\task.xlsx --sheet ^É^ÉXÉN
popd

ENDLOCAL
pause
