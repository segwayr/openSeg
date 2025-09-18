@echo off
set "BASE=%~1"


REM 最後に \ があると処理しにくいので削除
if "%BASE:~-1%"=="\" set "BASE=%BASE:~0,-1%"

for /R "%BASE%" %%F in (*) do (
    REM %%F から BASE を取り除いて相対パスにする
    set "REL=%%F"
    setlocal enabledelayedexpansion
    echo !REL:%BASE%\=!
    endlocal
)

pause
