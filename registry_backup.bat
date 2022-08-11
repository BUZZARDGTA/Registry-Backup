@echo off

setlocal DisableDelayedExpansion
set @STRIP_WHITE_SPACES=^
(for %%a in (4096 2048 1024 512 256 128 64 32 16 8 4 2 1) do (^
    (set "_s=!?:~0,%%a!"^&^
    if "!_s: =!"==""^
        set^^ ?=!?:~%%a!)^&^
    (set "_s=!?:~-%%a!"^&^
    if "!_s: =!"==""^
        set^^ ?=!?:~0,-%%a!)^
))^&^
set _s=
setlocal EnableDelayedExpansion

>nul 2>&1 powershell /?
if !errorlevel!==0 (
    set powershell=1
) else (
    if defined powershell (
        set powershell=
    )
)
for /f %%A in ('2^>nul set registry_backup[') do (
    set %%A=
)
set registry_backup[#]=0
set "reg_file=%~1"
if defined reg_file (
    set "reg_file=!reg_file:"=!"
)
if not exist "!reg_file!" (
    echo ERROR: Reg file "!reg_file!" doesn't exist.
    echo        You have to pass or drag and drop an existing ".reg" file over this script.
    echo Exiting.
    goto :END
)
for %%A in ("!reg_file!") do (
    if /i not "%%~xA"==".reg" (
        echo ERROR: Passed file "!reg_file!" doesn't have the ".reg" file extension.
        echo        You have to pass or drag and drop an existing ".reg" file over this script.
        echo Exiting as a protection against it.
        goto :END
    )
    set "reg_pass=%%~A"
    set "reg_file=%%~nxA"
)
echo Started backing up "!reg_file!" file.
for /f "usebackqdelims=" %%A in ("!reg_pass!") do (
    if defined first_line (
        if /i not "!first_line!"=="Windows Registry Editor Version 5.00" (
            echo ERROR: First line from "!reg_file!" file doesn't match "Windows Registry Editor Version 5.00".
            echo Aborting script execution for safety reasons.
            goto :END
        )
        set "registry_path=%%A"
        %@STRIP_WHITE_SPACES:?=registry_path%
        if defined registry_path (
            if "!registry_path:~0,1!"=="[" (
                if "!registry_path:~-1!"=="]" (
                    set "registry_path=!registry_path:~1!"
                    set "registry_path=!registry_path:~0,-1!"
                    echo Processing backup of "[!registry_path!]" ...
                    if defined registry_path (
                        call :GET_DATE_TIME || (
                            echo ERROR fatal: Can't generate "date_time" variable.
                            echo Aborting script execution.
                            goto :END
                        )
                        if exist "registry_backup_!date_time!.reg" (
                            echo ERROR fatal: File "registry_backup_!date_time!.reg" already existing.
                            echo Aborting script execution for safety reasons.
                            goto :END
                        )
                        set /a registry_backup[#]+=1
                        set "registry_backup[!registry_backup[#]!]=registry_backup_!date_time!.reg"
                        >nul reg export "!registry_path!" "registry_backup_!date_time!.reg" /y || (
                            echo ERROR fatal: Reg export for "[!registry_path!]" to "registry_backup_!date_time!.reg" failed.
                            echo Aborting script execution for safety reasons.
                            goto :END
                        )
                    )
                )
            )
        )
    ) else (
        set "first_line=%%A"
        %@STRIP_WHITE_SPACES:?=first_line%
    )
)
if not defined registry_backup[1] (
    echo ERROR: Noting found to backup.
    echo Aborting script execution.
    goto :END
)
call :GET_DATE_TIME || (
    echo ERROR fatal: Can't generate "date_time" variable.
    echo Aborting script execution.
    goto :END
)
if exist "[registry_backup]_!reg_file!_!date_time!.reg" (
    echo ERROR fatal: File "[registry_backup]_!reg_file!_!date_time!.reg" already exist.
    echo Aborting script execution for safety reasons.
    goto :END
)
echo Compressing processed backup(s) in a single file ...
>"[registry_backup]_!reg_file!_!date_time!.reg" (
    echo Windows Registry Editor Version 5.00
)
for /l %%A in (1,1,!registry_backup[#]!) do (
    >>"[registry_backup]_!reg_file!_!date_time!.reg" (
        type !registry_backup[%%A]! | find /v /i "Windows Registry Editor Version 5.00"
        del !registry_backup[%%A]!
    )
)
echo:
echo Finished backing up all reg values from "!reg_file!" ...
>nul timeout /t 2 /nobreak
echo Starting generated backup file "[registry_backup]_!reg_file!_!date_time!.reg".
start notepad "[registry_backup]_!reg_file!_!date_time!.reg"

:END
echo:
echo Press {ANY KEY} to exit ...
>nul pause
endlocal
exit /b 0

:GET_DATE_TIME
if defined date_time (
    set date_time=
)
for /f "tokens=2delims==" %%A in ('2^>nul wmic os get Localdatetime /value') do (
    set "date_time=%%A"
    set "date_time=!date_time:~-26,4!-!date_time:~-22,2!-!date_time:~-20,2!_!date_time:~-18,2!-!date_time:~-16,2!-!date_time:~-14,2!.!date_time:~-11,3!"
)
call :CHECK_DATE_TIME date_time && (
    exit /b 0
)
if defined powershell (
    if defined date_time (
        set date_time=
    )
    for /f "delims=" %%A in ('2^>nul powershell get-date -format "'yyyy-MM-dd_HH-mm-ss.fff'"') do (
        set "date_time=%%A"
    )
    call :CHECK_DATE_TIME date_time && (
        exit /b 0
    )
)
exit /b 1

:CHECK_DATE_TIME
if not defined %1 (
    exit /b 1
)
if not "!%1:~4,1!"=="-" (
    exit /b 1
)
if not "!%1:~7,1!"=="-" (
    exit /b 1
)
if not "!%1:~10,1!"=="_" (
    exit /b 1
)
if not "!%1:~13,1!"=="-" (
    exit /b 1
)
if not "!%1:~16,1!"=="-" (
    exit /b 1
)
if not "!%1:~19,1!"=="." (
    exit /b 1
)
for /f "delims=0123456789-_." %%A in ("!%1!") do (
    exit /b 1
)
for /f "tokens=1-7delims=-_." %%A in ("!%1!") do (
    call :CHECK_NUMBER "%%A" && (
        if "%%B"=="01" (
            set y1=31
        ) else if "%%B"=="02" (
            set "years=%%A"
            call :IS_LEAP_YEAR_OR_NOT
            if !leap!==1 (
                set y1=29
            ) else (
                set y1=28
            )
        ) else if "%%B"=="03" (
            set y1=31
        ) else if "%%B"=="04" (
            set y1=30
        ) else if "%%B"=="05" (
            set y1=31
        ) else if "%%B"=="06" (
            set y1=30
        ) else if "%%B"=="07" (
            set y1=31
        ) else if "%%B"=="08" (
            set y1=31
        ) else if "%%B"=="09" (
            set y1=30
        ) else if "%%B"=="10" (
            set y1=31
        ) else if "%%B"=="11" (
            set y1=30
        ) else if "%%B"=="12" (
            set y1=31
        )
        if defined y1 (
            call :CHECK_NUMBER_BETWEEN_CUSTOM "%%C" 01-!y1! && (
                call :CHECK_NUMBER_BETWEEN_CUSTOM "%%D" 00-23 && (
                    call :CHECK_NUMBER_BETWEEN_CUSTOM "%%E" 00-59 && (
                        call :CHECK_NUMBER_BETWEEN_CUSTOM "%%F" 00-59 && (
                            call :CHECK_NUMBER_BETWEEN_CUSTOM "%%G" 000-999 && (
                                exit /b 0
                            )
                        )
                    )
                )
            )
        )
    )
)
exit /b 1

:CHECK_NUMBER
set data=%1
if "!data:~0,1!!data:~-1!"=="""" (
set "data=%~1"
) else set "data=!%1!"
if not defined data exit /b 1
for /f "delims=0123456789" %%A in ("!data!") do exit /b 1
exit /b 0

:CHECK_NUMBER_BETWEEN_CUSTOM
if "%~1"=="" exit /b 1
for /f "delims=0123456789" %%A in ("%~1") do exit /b 1
for /f "tokens=1,2delims=-" %%A in ("%~2") do (
    if %~1 lss %%A exit /b 1
    if %~1 gtr %%B exit /b 1
)
exit /b 0

:IS_LEAP_YEAR_OR_NOT
::https://stackoverflow.com/questions/35157817/batch-file-leap-year
set /a "leap=^!(years%%4) + (^!^!(years%%100)-^!^!(years%%400))"
exit /b

