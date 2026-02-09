@echo off
setlocal
pushd "%~dp0"
set PYPATH="C:\Python312\python.exe"
set PYSCRIPT="USDM2mm.py"

REM ドラッグ＆ドロップされたファイルをforでまわす
for %%A in (%*) do (
    REM 拡張子が .xlsx だったら処理続行
    if /i "%%~xA"==".xlsx" (
        REM Freeplane に変換
        echo "%%A" convert to "%%~dpnA_reverse.mm"
        %PYPATH% %PYSCRIPT% %%A -o "%%~dpnA_reverse.mm"
    )
)
popd
pause
