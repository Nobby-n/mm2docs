@echo off
setlocal
pushd "%~dp0"
set PYPATH="C:\Python312\python.exe"
set PYSCRIPT="mm2docs.py"
set TEMPLATE=".\templates\USDM_Template.xlsx"

REM ドラッグ＆ドロップされたファイルをforでまわす
for %%A in (%*) do (
    REM 拡張子が .xmind または .mm だったら処理続行
    for %%E in (.xmind .mm) do if /i "%%~xA"=="%%E" (
        REM アウトプットしたいフルパスをセットする
        REM %%~dA = ドライブ名, %%~pA = パス, %%~nA = ファイル名
        echo "%%A" convert to "%%~dA%%~pA%%~nA_USDM.xlsx"
        %PYPATH% %PYSCRIPT% %%A -t %TEMPLATE% -o "%%~dA%%~pA%%~nA_USDM.xlsx"
    )
)
popd
pause
