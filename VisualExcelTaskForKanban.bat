@echo off
REM Miniconda 環境を有効化

SETLOCAL

REM Miniconda の仮想環境をアクティブにする
CALL "%USERPROFILE%\miniconda3\Scripts\activate.bat" py310

REM リポジトリルートに移動してバックエンドを起動
pushd "%~dp0.."
python backend\backend.py --excel .\data\task.xlsx --sheet ^タスク
popd

ENDLOCAL
pause
