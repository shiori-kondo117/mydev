@echo off
@rem //****************************************************
@rem // @procedure proc-Begin
@rem //****************************************************
:proc-Begin
	@rem -- set configuration
	call %~dp0profile.cmd
	set vDestDir=%vCsvHome%
	set vDestFile=cprm_keys.csv
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
	
	@rem -- get id data last
	set vCmd=get-id-data-last.cmd
	call %vCmdHome%%vCmd% %vDestDir%%vDestFile% vId_
	
	set vRc=%errorlevel%
	if %vRc% neq 0 (
		call :proc-RunLog "E" "command execute failer!! [cmd=%vCmd% rc=%vRc%]"
		goto :proc-End
	)
	
	set /a vId_+= 1
	
	set vCsvData=%vId_%,"%vVolSer_%","%vCprmKey_%","%vDriveName%","%vDate_% %vTime_%","",""
	
	echo %vCsvData%>> %vDestDir%%vDestFile%

	set vRc=%errorlevel%
	if %vRc% neq 0 (
		call :proc-RunLog "E" "csv file output failer!! [dest=%vDestDir%%vDestFile% rc=%vRc%]"
		goto :proc-End
	)
	
	call :proc-RunLog "I" "cmd=%~nx0 ended rc=%vRc%"

@rem //****************************************************
@rem // @procedure proc-End
@rem //****************************************************
:proc-End
	powershell -command cat %vLogFile% -tail 5

	pause

	goto :eof
	
@rem //****************************************************
@rem // @procedure proc-RunLog
@rem //****************************************************
:proc-RunLog
	set vDate=%date: =0%
	set vTime=%time: =0%
	if not exist %vLogFile% (
		echo [%~1],%vDate% %vTime%,%~2 > %vLogFile% 2>&1
	) else (
		echo [%~1],%vDate% %vTime%,%~2 >> %vLogFile% 2>&1
	)
	goto :eof
