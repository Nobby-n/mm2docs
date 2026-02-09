@echo off
setlocal
pushd "%~dp0"
set PYPATH="C:\Python312\python.exe"
set PYSCRIPT="USDM2docs.py"
set TEMPLATE=".\templates\SpecTemplate.docx"

REM ドラッグ＆ドロップされたファイルをforでまわす
for %%A in (%*) do (
    REM 拡張子が .xlsx だったら処理続行
    if /i "%%~xA"==".xlsx" (
        REM Word に変換
        echo "%%A" convert to "%%~dpnA_要求仕様書_r0.docx"
        %PYPATH% %PYSCRIPT% %%A -t %TEMPLATE% -o "%%~dpnA_要求仕様書_r0.docx"
    )
)
popd
pause
