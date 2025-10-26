@echo off
REM Miniconda 実行バッチ

SETLOCAL

REM Miniconda 環境
CALL "%USERPROFILE%\miniconda3\Scripts\activate.bat" py310

python backend\backend.py --excel .\data\task.xlsx --sheet ^タスク
popd

ENDLOCAL
pause
