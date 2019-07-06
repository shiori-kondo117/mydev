@echo off
:proc-Begin
	if "%1" == "" (
		echo Usage:%~nx0 [src-file]
		pause
		exit /b 1
	)

	if not "%~nx1" == "VR_MOVIE.VRO" (
		call :proc-RunLog "E" "illegal file!!  [file=%~nx1]"
		pause
		exit /b 2
	)
	set vLogDir=%userprofile%\documents\logs\
	set vLogFile=%vLogDir%%~n0.log
:proc-Main
	set vCmdHome=%~dp0
	set vAppHome=C:\Program Files (x86)\CPRMDecrypter\
	set vApp=c2dec.exe
	set vAppOpt=-R
	set /p vCprmKey=type-in CPRM-KEY:
	set vFile=%~nx1
	set vSrcDir=%~dp1
	set vSrcFile=%~dpnx1
	set vDestFile=%vSrcDir%..\%vFile%

	call :proc-RunLog "I" "command starting... [cmd=%vApp% src-dir=%vSrcDir%]"

	"%vAppHome%%vApp%" %vAppOpt% %vCprmKey% %vSrcFile% %vDestFile%

	if %errorlevel% neq 0 (
		call :proc-RunLog "E" "app execute failer!! [cmd=%vApp% rc=%errorlevel%]"
		exit /b 2
	)

	ren %vSrcFile% _%vFile%

	if %errorlevel% neq 0 (
		call :proc-RunLog "E" "rename command failer!! [cmd=ren rc=%errorlevel%]"
		exit /b 3
	)

	move %vDestFile% %vSrcDir%

	if %errorlevel% neq 0 (
		call :proc-RunLog "E" "move command failer!! [cmd=move rc=%errorlevel%]"
		exit /b 4
	)

	call :proc-RunLog "I" "command sucessfully [cmd=%vApp% src-dir=%vSrcDir%]"
:proc-End
	powershell -command cat %vLogFile% -tail 5
	pause
	goto :eof

:proc-RunLog
	set vDateTime=%date: =0% %time: =0%
	if not exist %vLogFile% (
		echo [%~1],%vDateTime%,%~2 > %vLogFile% 2>&1
	) else (
		echo [%~1],%vDateTime%,%~2 >> %vLogFile% 2>&1
	)
	
	goto :eof
