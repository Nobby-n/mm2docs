@echo off
setlocal
pushd "%~dp0"
set PYPATH="C:\Python312\python.exe"
set PYSCRIPT="USDM2docs.py"

REM ドラッグ＆ドロップされたファイルをforでまわす
for %%A in (%*) do (
    REM 拡張子が .xlsx だったら処理続行
    if /i "%%~xA"==".xlsx" (
        REM Markdown に変換
        echo "%%A" convert to "%%~dpnA_SRS.md"
        %PYPATH% %PYSCRIPT% %%A -o "%%~dpnA_SRS.md"
    )
)
popd
pause
