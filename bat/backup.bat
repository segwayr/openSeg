chcp 65001
set dater=%date:~0,4%%date:~5,2%%date:~8,2%
if not exist %~dp1@bk mkdir %~dp1@bk

rem 大量の引数をシフトしながらグルグル回す処理とか
:LOOP
if "%~n1dummy" neq "dummy" (
    echo %date:~0,4%年%date:~5,2%月%date:~8,2%日%time:~0,2%時%time:~3,2%分%time:~6,2%秒に%~n1をバックアップしました>>%~dp1@bk\log.txt
)
set cnt=1
if %1=="" goto eof
if not exist "%~dp1@bk\%~n1_%dater%%~x1" (
    copy %1 "%~dp1@bk\%~n1_%dater%%~x1"
) else (
    :DoWhile
    if not exist "%~dp1@bk\%~n1_%dater%_%cnt%%~x1" (
        copy %1 "%~dp1@bk\%~n1_%dater%_%cnt%%~x1"
    ) else (
        set /a cnt=%cnt%+1
        goto DoWhile
    )
)
shift
goto LOOP
