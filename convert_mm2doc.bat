@echo off
setlocal
pushd "%~dp0"
set PYPATH="C:\Python312\python.exe"
set PYSCRIPT="mm2docs.py"
set TEMPLATE=".\templates\SpecTemplate.docx"

REM ドラッグ＆ドロップされたファイルをforでまわす
for %%A in (%*) do (
    REM 拡張子が .xmind または .mm だったら処理続行
    for %%E in (.xmind .mm) do if /i "%%~xA"=="%%E" (
        REM 直接 Word に変換
        echo "%%A" convert to "%%~dpnA_機能仕様書_r0.docx"
        %PYPATH% %PYSCRIPT% %%A -t %TEMPLATE% -o "%%~dpnA_機能仕様書_r0.docx"
    )
)
popd
pause

