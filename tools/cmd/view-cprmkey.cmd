@echo off
@rem //****************************************************
@rem // @procedure proc-Begin
@rem //****************************************************
:proc-Begin
	@rem -- set configuration
	set vLogDir=%userprofile%\documents\logs\
	set vLogFile=%vLogDir%%~n0.log
	
	if "%~1" == "" (
		echo Usage:%~nx0 [src-drive]
		set vRc=1
		goto :proc-End
	)

	if not exist %~d1% (
		call :proc_RunLog "E","not exist src-drive [src-drv=%~d1%]"
		set vRc=2
		goto :proc-End
	)

	set vCmdHome=%~dp0
	set vDriveName=%~d1
	
@rem //****************************************************
@rem // @procedure proc-Main
@rem //****************************************************
:proc-Main
	call :proc-RunLog "I" "cmd=%~nx0 starting..."
	@rem -- get cprm key
	set vCmd=get-cprmkey.cmd
	call %vCmdHome%%vCmd% %vDriveName% vCprmKey_
	
	set vRc=%errorlevel%
	if %vRc% neq 0 (
		echo "command execute failer!! [cmd=%vCmd% rc=%vRc%]"
		call :proc-RunLog "E" "command execute failer!! [cmd=%vCmd% rc=%vRc%]"
		goto :proc-End
	)
	
	@rem -- get date time
	set vCmd=get-date-time.cmd
	call %vCmdHome%%vCmd% vDate_ vTime_
	
	set vRc=%errorlevel%
	if %vRc% neq 0 (
		call :proc-RunLog "E" "command execute failer!! [cmd=%vCmd% rc=%vRc%]"
		goto :proc-End
	)
	
	@rem -- get drive name
	set vCmd=get-vol.cmd
	call %vCmdHome%%vCmd% %vDriveName% vDriveName_ vVolName_ vVolSer_
	
	set vRc=%errorlevel%
	if %vRc% neq 0 (
		call :proc-RunLog "E" "command execute failer!! [cmd=%vCmd% rc=%vRc%]"
		goto :proc-End
	)
	
	set vViewData=[VOLSER:%vVolSer_%] [CPRM-KEY:%vCprmKey_%] [DRIVE-NAME:%vDriveName%]
	
	echo %vViewData%

	call :proc-RunLog "I" "cmd=%~nx0 ended rc=%vRc%"

@rem //****************************************************
@rem // @procedure proc-End
@rem //****************************************************
:proc-End
	if exist %vLogFile% (
@rem --		powershell -command cat %vLogFile% -tail 5
	)

	pause

	goto :eof
	
@rem //****************************************************
@rem // @procedure proc-RunLog
@rem //****************************************************
:proc-RunLog
	set vDate=%date: =0%
	set vTime=%time: =0%
	set vLogMsg=[%~1],%vDate% %vTime%,%~2
	if not exist %vLogFile% (
		echo %vLogMsg% > %vLogFile% 2>&1
	) else (
		echo %vLogMsg% >> %vLogFile% 2>&1
	)
@rem --	echo %vLogMsg%
	goto :eof
