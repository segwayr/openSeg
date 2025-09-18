@echo off
setlocal enabledelayedexpansion

:: 引数が無ければ終了
if "%~1"=="" (
    echo 使用方法: %~nx0 対象フォルダ
    exit /b 1
)

:: 探索対象フォルダ
set TARGET=%~1

:: 最初に対象フォルダ自体
echo [%TARGET%]
for %%F in ("%TARGET%\*") do (
    if exist "%%F" echo %%~nxF
)
echo.

:: サブフォルダを順番に処理
for /R "%TARGET%" /D %%D in (*) do (
    echo [%%~fD]
    for %%F in ("%%D\*") do (
        if exist "%%F" echo %%~nxF
    )
    echo.
)

endlocal

pause
