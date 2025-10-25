@echo off
REM Fletアプリ起動用バッチファイル（Miniconda版）

SETLOCAL

REM Minicondaの flet 環境をアクティベート
CALL "%USERPROFILE%\miniconda3\Scripts\activate.bat" py310

REM アプリの起動（main.pyを実行）
python backend.py --excel ./task.xlsx --sheet タスク

ENDLOCAL
pause
