@echo off
setlocal
pushd "%~dp0"
set PYPATH="C:\Python312\python.exe"
set PYSCRIPT="mm2docs.py"

REM ドラッグ＆ドロップされたファイルをforでまわす
for %%A in (%*) do (
    REM 拡張子が .xmind または .mm だったら処理続行
    for %%E in (.xmind .mm) do if /i "%%~xA"=="%%E" (
        REM 直接 Markdown に変換
        echo "%%A" convert to "%%~dA%%~pA%%~nA_SRS.md"
        %PYPATH% %PYSCRIPT% %%A -o "%%~dA%%~pA%%~nA_SRS.md"
    )
)
popd
pause
